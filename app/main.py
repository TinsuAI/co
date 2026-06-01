from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import os
import re
from contextlib import asynccontextmanager
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from urllib.parse import quote

import httpx
from fastapi import File, Form, HTTPException, Request, UploadFile
from fastapi import FastAPI
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse, RedirectResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app import co_auth
from app.bom_store import attach_case_bom_snapshot
from app.bom_service import bom_service
from app.co_case_store import (
    MAX_SUPPORTING_FILE_BYTES,
    CaseClosedError,
    CaseHasActiveClaimsError,
    acquire_origin_calculation_lock,
    active_origin_calculation_lock,
    build_case_criteria_rows,
    case_from_record,
    co_case_delete_block_reason,
    co_case_is_completed,
    create_case_record,
    create_case_workbook,
    declaration_refs,
    delete_case_record,
    get_case_record,
    get_case_workspace,
    get_supporting_file,
    invoice_keys,
    json_safe,
    match_case_bcct_exports,
    safe_filename,
    save_supporting_file,
    release_origin_calculation_lock,
    update_case_record,
)
from app.co_forms import (
    COMMON_MARKET_PRESETS,
    common_market_guidance,
    criteria_preview_for_hs,
    form_candidates_for_market,
    prioritized_form_lanes,
    recommended_form_lane,
)
from app.co_form_config_store import (
    co_form_config_path,
    load_co_form_config,
    reset_co_form_config,
    sanitize_co_form_config,
    save_co_form_config,
    unique_text_list,
)
from app import co_stock_adjustments_store, co_stock_eligibility, co_stock_events_store, co_stock_ledger, co_stock_materializer
from app.co_stock_template import CoStockTemplateError, read_standard_co_stock, write_standard_co_stock
from app.co_market_hints import infer_market_from_invoice_matches
from app.client_registry import get_client as registry_get_client
from app.client_registry import get_client_case
from app.customs_fx_store import CUSTOMS_FX_CLIENT_ID, get_customs_fx_store, refresh_customs_exchange_rates
from app.database import apply_migrations, database_url
from app.data_hub_client import (
    DataHubClient,
    bom_product_code_from_material_identity,
    reset_current_data_hub_token,
    set_current_data_hub_token,
)
from app.data_hub_settings import (
    DATA_HUB_LINK_ENV_KEYS,
    DataHubLinkSettings,
    data_hub_config_path,
    data_hub_link_settings,
    load_data_hub_overrides,
    save_data_hub_overrides,
)
from app.demo_data import (
    DEMO_CASE,
    SOURCE_NOTES,
    attach_results,
    clone_case,
    update_products_from_form,
)
from app.origin import evaluate_tariff_shift
from app import material_search
from app.app_state_store import get_app_state_store
from app.portfolio import SourceBackendUnavailable, portfolio_app, portfolio_service
from app.source_store import (
    attach_case_source_snapshot,
    co_stock_rows_from_bcct,
    enrich_client_with_source_workspace,
)
from app.table_view import build_table_view
from app.workbook_io import (
    WorkbookParseError,
    create_dossier_zip,
    create_evidence_workbook,
    create_hq_bang_ke_workbook,
    create_input_workbook,
    parse_input_workbook,
)


PSR_STATUS_LABELS = {
    "pending_trong_tin_confirmation": "Chờ Trọng Tín xác nhận",
    "extracted_from_local_corpus_pending_trong_tin_confirmation": "Extract từ corpus, chờ xác nhận",
    "needs_2026_evfta_refresh": "Cần đối chiếu EVFTA 2026",
    "requires_manual_lookup": "Cần tra thủ công",
    "confirmed_by_trong_tin": "Đã xác nhận bởi Trọng Tín",
}

ROOT = Path(__file__).resolve().parent

THEME_COOKIE = "co_theme"
SUPPORTED_THEMES = {"light", "dark"}

ORIGIN_SHEET_STATUS_LABELS = {
    "draft": "Chưa tính",
    "calculating": "Đang tính",
    "calculated": "Đã tính",
    "locked": "Chốt",
    "stale": "Cần tính lại",
}


def normalize_theme(value: str | None) -> str:
    return value if value in SUPPORTED_THEMES else "light"


def theme_context(request: Request) -> dict[str, str]:
    theme = normalize_theme(request.cookies.get(THEME_COOKIE))
    user = co_auth.current_user(request)
    return {
        "theme": theme,
        "next_theme": "light" if theme == "dark" else "dark",
        "co_user": user,
        "auth_required": co_auth.auth_required(),
        "show_login": (co_auth.auth_required() or co_auth.data_hub_source_mode_enabled()) and request.url.path != "/auth/logout",
        "can_view_technical_settings": co_auth.can_view_technical_settings(user),
        "can_delete_co_cases": co_auth.can_delete_co_cases(user),
    }


def origin_lock_actor(request: Request) -> dict[str, str]:
    user = co_auth.current_user(request)
    if not user:
        return {"id": "local", "label": "Local user"}
    return {
        "id": user.user_id,
        "label": user.name or user.email or user.user_id,
        "name": user.name,
        "email": user.email,
    }


@asynccontextmanager
async def lifespan(_app: FastAPI):
    if database_url():
        apply_migrations()
    yield


app = FastAPI(title="Barry CO Demo", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")
app.mount("/portfolio", portfolio_app, name="portfolio")


@app.exception_handler(CaseClosedError)
async def _case_closed_handler(request: Request, exc: CaseClosedError):
    """Mutating route hit a closed case → 409 with the friendly Vietnamese message."""
    return JSONResponse(status_code=409, content={"detail": str(exc)})


@app.exception_handler(SourceBackendUnavailable)
async def _source_backend_unavailable_handler(request: Request, exc: SourceBackendUnavailable):
    """Data Hub is the source of truth but unavailable, and local fallback is not
    allowed → 503 with a clear message instead of a silently-empty page."""
    return PlainTextResponse(str(exc), status_code=503)


@app.exception_handler(httpx.TransportError)
async def _data_hub_unreachable_handler(request: Request, exc: httpx.TransportError):
    """Data Hub is enabled but the API is unreachable (connection refused /
    timeout). Surface a clear 503 instead of a generic 500 so the operator knows
    it's a Data Hub outage, not a CO bug. Only fires for transport errors that
    propagate unhandled — local try/except (e.g. 404 fallbacks) still wins."""
    return PlainTextResponse(
        "Data Hub không phản hồi (kết nối thất bại/timeout). CO không dùng dữ liệu "
        "local backup; kiểm tra Data Hub rồi thử lại.",
        status_code=503,
    )


@app.exception_handler(httpx.HTTPStatusError)
async def _data_hub_error_status_handler(request: Request, exc: httpx.HTTPStatusError):
    """Data Hub returned an error status that no route handled → 502 (bad
    gateway): the upstream source failed, not CO. Routes that intentionally
    handle Data Hub statuses (e.g. 404 → fallback) catch the error themselves
    and never reach this handler."""
    upstream = exc.response.status_code if exc.response is not None else "?"
    return PlainTextResponse(
        f"Data Hub trả lỗi ({upstream}). CO không dùng dữ liệu local backup; "
        "kiểm tra Data Hub rồi thử lại.",
        status_code=502,
    )

templates = Jinja2Templates(directory=ROOT / "templates", context_processors=[theme_context])


async def large_request_form(request: Request):
    try:
        return await request.form(max_fields=100000, max_files=2000)
    except TypeError:
        return await request.form()


def origin_case_revision(case: dict) -> str:
    payload = json_safe(
        {
            "source_snapshot": case.get("source_snapshot", {}),
            "bom_snapshot": case.get("bom_snapshot", {}),
            "origin_snapshot": case.get("origin_snapshot", {}),
            "origin_product_order": origin_product_order(case),
            "origin_sheet_states": case.get("origin_sheet_states", {}),
            "bom_product_artifact_overrides": case.get("bom_product_artifact_overrides", {}),
        }
    )
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:16]


def persisted_origin_case(client: dict, case_id: str) -> dict:
    record = get_case_record(client, case_id)
    case = case_from_record(default_client_case(client), client, record)
    case["persisted_case_id"] = case.get("persisted_case_id") or case_id
    return case


def merge_origin_action_payload(case: dict, payload: dict) -> dict:
    if not isinstance(payload, dict):
        return case
    prepared = dict(case)
    order = payload.get("origin_product_order")
    if isinstance(order, str):
        prepared["origin_product_order"] = origin_product_order({"origin_product_order": order})
    elif isinstance(order, list):
        prepared["origin_product_order"] = [str(code).strip() for code in order if str(code).strip()]
    bom_artifact_id = str(payload.get("bom_artifact_id") or payload.get("bom_version_id") or "").strip()
    if bom_artifact_id:
        prepared["bom_artifact_id"] = bom_artifact_id
        prepared["bom_version_id"] = bom_artifact_id
    overrides = dict(prepared.get("bom_product_artifact_overrides") or {})
    legacy_overrides = dict(prepared.get("bom_product_version_overrides") or {})
    incoming_overrides = payload.get("bom_product_artifact_overrides")
    if isinstance(incoming_overrides, dict):
        for code, artifact_id in incoming_overrides.items():
            code = str(code or "").strip()
            artifact_id = str(artifact_id or "").strip()
            if code and artifact_id:
                overrides[code] = artifact_id
                legacy_overrides[code] = artifact_id
    products_by_code = {
        str(product.get("code") or "").strip(): dict(product)
        for product in prepared.get("products", [])
        if str(product.get("code") or "").strip()
    }
    product_sheet_states: dict[str, dict] = {}
    for incoming in payload.get("products") or []:
        if not isinstance(incoming, dict):
            continue
        code = str(incoming.get("code") or incoming.get("product_code") or "").strip()
        if not code:
            continue
        product = products_by_code.get(code, {"code": code})
        for key in [
            "name",
            "finished_hs",
            "quantity",
            "unit",
            "currency",
            "source_declaration_no",
            "source_line_no",
            "invoice_ref",
            "fob",
            "non_origin_value",
            "rvc_threshold",
            "lvc_threshold",
            "bom_product_code",
            "bom_product_artifact_id",
            "bom_product_artifact_no",
            "bom_product_version_id",
            "bom_product_version_no",
            "origin_sheet_status",
            "origin_sheet_status_label",
        ]:
            if key in incoming:
                product[key] = incoming.get(key)
        if isinstance(incoming.get("cost_buildup"), dict):
            existing_cb = product.get("cost_buildup") if isinstance(product.get("cost_buildup"), dict) else {}
            cb_whitelist = {
                "wages", "welfare", "rent", "depreciation", "other_mfg", "transport_storage",
                "profit",
                "labor", "overhead", "other",  # legacy 4-key shape, still accepted
            }
            product["cost_buildup"] = {**existing_cb, **{k: str(v or "") for k, v in incoming["cost_buildup"].items() if k in cb_whitelist}}
        if isinstance(incoming.get("materials"), list):
            product["materials"] = incoming["materials"]
        if incoming.get("origin_sheet_status") or incoming.get("origin_sheet_status_label"):
            product_sheet_states[code] = {
                "status": str(incoming.get("origin_sheet_status") or "").strip(),
                "status_label": str(incoming.get("origin_sheet_status_label") or "").strip(),
            }
        artifact_id = str(
            product.get("bom_product_artifact_id") or product.get("bom_product_version_id") or ""
        ).strip()
        bom_product_code = str(product.get("bom_product_code") or code).strip()
        if artifact_id:
            overrides[code] = artifact_id
            legacy_overrides[code] = artifact_id
            if bom_product_code:
                overrides[bom_product_code] = artifact_id
                legacy_overrides[bom_product_code] = artifact_id
        products_by_code[code] = product
    if products_by_code:
        ordered = origin_product_order(prepared)
        remainder = [code for code in products_by_code if code not in ordered]
        prepared["products"] = [products_by_code[code] for code in ordered + remainder if code in products_by_code]
    sheet_states = payload.get("origin_sheet_states")
    if not isinstance(sheet_states, dict) and product_sheet_states:
        sheet_states = product_sheet_states
    if isinstance(sheet_states, dict):
        existing_states = prepared.get("origin_sheet_states") if isinstance(prepared.get("origin_sheet_states"), dict) else {}
        merged_states: dict[str, dict] = {
            str(code): dict(state)
            for code, state in existing_states.items()
            if isinstance(state, dict)
        }
        for code, state in sheet_states.items():
            if not isinstance(state, dict):
                continue
            previous = existing_states.get(str(code)) if isinstance(existing_states.get(str(code)), dict) else {}
            merged_states[str(code)] = {**previous, **state}
        prepared["origin_sheet_states"] = merged_states
    prepared["bom_product_artifact_overrides"] = overrides
    prepared["bom_product_version_overrides"] = legacy_overrides
    return prepared


async def origin_case_from_request(request: Request, client: dict, case_id: str) -> tuple[dict, dict]:
    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        try:
            payload = await request.json()
        except json.JSONDecodeError:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
        case = persisted_origin_case(client, case_id)
        expected_revision = str(payload.get("expected_revision") or "").strip()
        if expected_revision and expected_revision != origin_case_revision(case):
            raise HTTPException(status_code=409, detail="Origin case state changed; reload before saving.")
        return merge_origin_action_payload(case, payload), payload
    form = await large_request_form(request)
    case = update_products_from_form({key: str(value) for key, value in form.items()})
    case["persisted_case_id"] = case.get("persisted_case_id") or case_id
    return case, {key: str(value) for key, value in form.items()}


def format_number_display(value, max_decimals: int = 2) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    try:
        decimal = Decimal(text.replace(",", ""))
    except (InvalidOperation, ValueError):
        return text
    if decimal == decimal.to_integral():
        return f"{int(decimal):,}"
    max_decimals = max(0, min(int(max_decimals), 8))
    quant = Decimal("1").scaleb(-max_decimals)
    rounded = decimal.quantize(quant, rounding=ROUND_HALF_UP)
    return f"{rounded:,.{max_decimals}f}".rstrip("0").rstrip(".")


templates.env.filters["number"] = format_number_display


@app.get("/healthz")
async def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.middleware("http")
async def require_data_hub_auth(request: Request, call_next):
    redirect = co_auth.guard_response(request)
    if redirect:
        return redirect
    co_auth.load_optional_user(request)
    token_context = None
    user = co_auth.current_user(request)
    if user and user.access_token:
        token_context = set_current_data_hub_token(user.access_token)
    try:
        return await call_next(request)
    finally:
        if token_context:
            reset_current_data_hub_token(token_context)

CATALOG_VIEWS = {
    "materials": {
        "module": "material_catalog",
        "title": "DS NVL DK HQ",
        "subtitle": "Nguyên vật liệu theo danh mục đăng ký hải quan.",
        "template_url": "material-template.xlsx",
        "template_label": "Tải template DS NVL",
        "columns": [
            {"key": "customs_code", "label": "Mã HQ", "class": "mono"},
            {"key": "name", "label": "Tên"},
            {"key": "unit", "label": "ĐVT", "class": "mono"},
            {"key": "hs_code", "label": "HS", "class": "mono"},
            {"key": "purpose", "label": "Mục đích"},
            {"key": "origin_default", "label": "Xuất xứ mặc định"},
            {"key": "status", "label": "Trạng thái"},
        ],
        "filters": [
            {"name": "status", "field": "status", "label": "Trạng thái"},
            {"name": "unit", "field": "unit", "label": "ĐVT"},
            {"name": "hs", "field": "hs_code", "label": "HS"},
            {"name": "purpose", "field": "purpose", "label": "Mục đích"},
        ],
        "summary_fields": [
            {"field": "status", "label": "Trạng thái"},
            {"field": "unit", "label": "ĐVT"},
        ],
        "default_sort": "customs_code",
    },
    "products": {
        "module": "product_catalog",
        "title": "DS SP DK HQ",
        "subtitle": "Thành phẩm theo danh mục đăng ký hải quan.",
        "template_url": "product-template.xlsx",
        "template_label": "Tải template DS SP",
        "columns": [
            {"key": "product_code", "label": "Mã SP", "class": "mono"},
            {"key": "name", "label": "Tên"},
            {"key": "unit", "label": "ĐVT", "class": "mono"},
            {"key": "hs_code", "label": "HS", "class": "mono"},
            {"key": "purpose", "label": "Mục đích"},
            {"key": "status", "label": "Trạng thái"},
        ],
        "filters": [
            {"name": "status", "field": "status", "label": "Trạng thái"},
            {"name": "unit", "field": "unit", "label": "ĐVT"},
            {"name": "hs", "field": "hs_code", "label": "HS"},
            {"name": "purpose", "field": "purpose", "label": "Mục đích"},
        ],
        "summary_fields": [
            {"field": "status", "label": "Trạng thái"},
            {"field": "unit", "label": "ĐVT"},
        ],
        "default_sort": "product_code",
    },
}

BCCT_COLUMNS = [
    {"key": "coverage_period", "label": "Kỳ"},
    {"key": "declaration_no", "label": "Tờ khai", "class": "mono"},
    {"key": "line_no", "label": "STT", "class": "mono"},
    {"key": "declaration_type", "label": "LH", "class": "mono"},
    {"key": "direction_label", "label": "Luồng"},
    {"key": "item_code", "label": "Mã hàng", "class": "mono"},
    {"key": "hs_code", "label": "HS", "class": "mono"},
    {"key": "quantity", "label": "Số lượng", "class": "num"},
    {"key": "unit", "label": "ĐVT", "class": "mono"},
    {"key": "customs_value", "label": "Trị giá", "class": "num"},
    {"key": "invoice_ref", "label": "Hóa đơn", "class": "mono"},
]

CO_CASE_WORKFLOW_STEPS = [
    {
        "key": "shipment",
        "label": "Lô hàng",
        "short_label": "1",
        "description": "Thông tin shipment, invoice, B/L và thị trường.",
    },
    {
        "key": "documents",
        "label": "Chứng từ",
        "short_label": "2",
        "description": "Upload BL, Invoice/Packing và các chứng từ bổ sung. TKX query sau khi chốt bảng kê.",
    },
    {
        "key": "origin",
        "label": "Bảng kê C/O",
        "short_label": "3",
        "description": "Tính tuần tự từng sheet, override tiêu chí/ngưỡng, thay NVL, chốt và sinh BOM artifact mới.",
    },
    {
        "key": "exports",
        "label": "TKX / TKN",
        "short_label": "4",
        "description": "Sau khi chốt bảng kê: query TKX/TKN từ Data Hub, bổ sung phần thiếu.",
    },
    {
        "key": "review",
        "label": "Review & Xuất",
        "short_label": "5",
        "description": "Kiểm tra dossier và xuất .zip tổng hợp (chứng từ + TKX/TKN + bảng kê HQ).",
    },
]
CO_CASE_WORKFLOW_STEP_KEYS = {step["key"] for step in CO_CASE_WORKFLOW_STEPS}
CO_CASE_STEP_STATUS_LABELS = {
    "ready": "Đủ",
    "todo": "Thiếu",
    "review": "Cần soát",
    "preview": "Preview",
}

CO_STOCK_COLUMNS = [
    {"key": "source_row", "label": "Dòng nguồn", "class": "mono"},
    {"key": "import_declaration_no", "label": "Tờ khai nhập", "class": "mono"},
    {"key": "line_no", "label": "STT", "class": "mono"},
    {"key": "declaration_type", "label": "LH", "class": "mono"},
    {"key": "customs_item_code", "label": "Mã HQ", "class": "mono"},
    {"key": "allocation_code", "label": "Mã phân bổ", "class": "mono"},
    {"key": "available_qty", "label": "Tồn CO", "class": "num"},
    {"key": "used_qty", "label": "Đã dùng", "class": "num"},
    {"key": "remaining_qty", "label": "Còn lại", "class": "num"},
    {"key": "status_label", "label": "Trạng thái"},
    {"key": "stock_reason_label", "label": "Lý do"},
    {"key": "history_action", "label": "Lịch sử", "kind": "history", "sortable": False},
]

CUSTOMS_FX_COLUMNS = [
    {"key": "currency_code", "label": "Nguyên tệ", "class": "mono"},
    {"key": "currency_name", "label": "Tên ngoại tệ"},
    {"key": "effective_date", "label": "Ngày hiệu lực", "class": "mono"},
    {"key": "rate_display", "label": "Tỷ giá", "class": "num", "sortable": False},
    {"key": "source_endpoint", "label": "Nguồn API", "class": "mono"},
    {"key": "fetched_at", "label": "Lần lấy", "class": "mono"},
]

BOM_PRODUCT_COLUMNS = [
    {"key": "product_code", "label": "Mã TP", "class": "mono", "link_key": "view_href"},
    {"key": "product_artifact_no", "label": "TP artifact", "class": "mono"},
    {"key": "row_count", "label": "Dòng BOM", "class": "num"},
    {"key": "status", "label": "Trạng thái"},
    {"key": "version_hash_short", "label": "Hash", "class": "mono"},
]

BOM_LINE_COLUMNS = [
    {"key": "material_code", "label": "Mã NVL", "class": "mono"},
    {"key": "material_name", "label": "Tên NVL"},
    {"key": "qty_per", "label": "Định mức", "class": "num"},
    {"key": "uom", "label": "ĐVT", "class": "mono"},
    {"key": "scrap_rate", "label": "Hao hụt", "class": "num"},
    {"key": "source", "label": "Nguồn"},
    {"key": "row_class", "label": "Trạng thái"},
]


@app.post("/settings/theme")
async def set_theme(theme: str = Form("light"), next_url: str = Form("/clients")):
    target = next_url if next_url.startswith("/") and not next_url.startswith("//") else "/clients"
    response = RedirectResponse(target, status_code=303)
    response.set_cookie(
        THEME_COOKIE,
        normalize_theme(theme),
        max_age=60 * 60 * 24 * 365,
        httponly=True,
        samesite="lax",
    )
    return response


def require_technical_settings_dev(request: Request) -> None:
    if not co_auth.can_view_technical_settings(co_auth.current_user(request)):
        raise HTTPException(status_code=403, detail="Technical settings require a dev Data Hub session.")


def mask_secret(value: str) -> str:
    if not value:
        return "Chưa cấu hình"
    return f"Đã cấu hình ({len(value)} ký tự)"


def data_hub_settings_context(request: Request, *, saved: bool = False, error: str = "", test_result: dict | None = None) -> dict:
    settings = data_hub_link_settings()
    overrides = load_data_hub_overrides()
    env_values = os.environ
    token_source = (
        "environment"
        if env_values.get("DATA_HUB_SERVICE_TOKEN")
        else "local override"
        if overrides.get("DATA_HUB_SERVICE_TOKEN")
        else "missing"
    )
    rows = [
        {"key": "DATA_HUB_ENABLED", "label": "Dùng Data Hub cho source/master data", "value": "1" if settings.source_enabled else "0", "type": "checkbox"},
        {"key": "CO_AUTH_REQUIRED", "label": "Bắt buộc Data Hub SSO cho CO", "value": "1" if settings.auth_required else "0", "type": "checkbox"},
        {"key": "DATA_HUB_BASE_URL", "label": "Data Hub browser URL", "value": settings.data_hub_base_url, "type": "url"},
        {"key": "DATA_HUB_API_BASE_URL", "label": "Data Hub API URL", "value": settings.data_hub_api_base_url, "type": "url"},
        {"key": "DATA_HUB_ISSUER_URL", "label": "JWT issuer", "value": settings.issuer_url, "type": "url"},
        {"key": "DATA_HUB_JWKS_URL", "label": "JWKS URL", "value": settings.jwks_url, "type": "url"},
        {"key": "CO_PUBLIC_BASE_URL", "label": "CO public URL", "value": settings.co_public_base_url, "type": "url"},
        {"key": "CO_FORCE_HTTPS_COOKIE", "label": "Secure cookie HTTPS", "value": "1" if settings.force_https_cookie else "0", "type": "checkbox"},
        {"key": "DATA_HUB_REQUEST_TIMEOUT_SECONDS", "label": "Request timeout seconds", "value": str(settings.request_timeout_seconds), "type": "number"},
        {"key": "DATA_HUB_CLIENT_CLAIM_KEYS", "label": "Client claim keys", "value": ",".join(settings.client_claim_keys), "type": "text"},
        {"key": "DATA_HUB_ADMIN_ROLES", "label": "Admin roles", "value": ",".join(sorted(settings.admin_roles)), "type": "text"},
        {"key": "CO_CASE_DELETE_ROLES", "label": "Roles được xoá hồ sơ C/O", "value": ",".join(sorted(settings.co_case_delete_roles)), "type": "text"},
    ]
    for row in rows:
        row["source"] = "environment" if row["key"] in env_values else "local override" if row["key"] in overrides else "default"
    return {
        "saved": saved,
        "error": error,
        "settings": settings,
        "rows": rows,
        "token": {
            "configured": bool(settings.api_token),
            "masked": mask_secret(settings.api_token),
            "source": token_source,
            "has_local_override": "DATA_HUB_SERVICE_TOKEN" in overrides,
        },
        "config_path": str(data_hub_config_path()),
        "test_result": test_result,
    }


def data_hub_override_payload(form, current_overrides: dict[str, str]) -> dict[str, str]:
    payload: dict[str, str] = {
        "DATA_HUB_ENABLED": "1" if form.get("DATA_HUB_ENABLED") == "1" else "0",
        "CO_AUTH_REQUIRED": "1" if form.get("CO_AUTH_REQUIRED") == "1" else "0",
        "CO_FORCE_HTTPS_COOKIE": "1" if form.get("CO_FORCE_HTTPS_COOKIE") == "1" else "0",
    }
    for key in DATA_HUB_LINK_ENV_KEYS:
        if key in payload or key in {"DATA_HUB_SERVICE_TOKEN", "CO_FORCE_HTTPS_COOKIE"}:
            continue
        payload[key] = str(form.get(key, "")).strip()
    token = str(form.get("DATA_HUB_SERVICE_TOKEN", "")).strip()
    if token:
        payload["DATA_HUB_SERVICE_TOKEN"] = token
    elif form.get("CLEAR_DATA_HUB_SERVICE_TOKEN") != "1":
        if current_overrides.get("DATA_HUB_SERVICE_TOKEN"):
            payload["DATA_HUB_SERVICE_TOKEN"] = current_overrides["DATA_HUB_SERVICE_TOKEN"]
    return payload


def data_hub_link_check() -> dict:
    settings = data_hub_link_settings()
    checks: list[dict] = []
    try:
        jwks = co_auth.fetch_data_hub_jwks(settings.jwks_url)
        key_count = len(jwks.get("keys", [])) if isinstance(jwks, dict) else 0
        checks.append({"name": "JWKS", "status": "success", "detail": f"{key_count} signing keys"})
    except Exception as exc:
        checks.append({"name": "JWKS", "status": "error", "detail": str(exc)})

    if not settings.source_enabled:
        checks.append({"name": "Source API", "status": "warning", "detail": "DATA_HUB_ENABLED đang tắt"})
    else:
        client = DataHubClient(
            base_url=settings.data_hub_api_base_url,
            token=settings.api_token,
            timeout=settings.request_timeout_seconds,
        )
        try:
            clients = client.list_clients()
            checks.append({"name": "Source API", "status": "success", "detail": f"{len(clients)} clients"})
        except Exception as exc:
            checks.append({"name": "Source API", "status": "error", "detail": str(exc)})
        finally:
            client.close()
    return {"ok": all(check["status"] != "error" for check in checks), "checks": checks}


def co_form_settings_context(request: Request, *, saved: bool = False, error: str = "") -> dict:
    config = load_co_form_config()
    display_config = {
        **config,
        "forms": [
            {
                **form,
                "verification_status_label": co_form_status_label(form.get("verification_status", "")),
            }
            for form in config.get("forms", [])
        ],
    }
    form_codes = [row["form_code"] for row in config["forms"] if row.get("enabled")]
    psr_rule_counts = {
        form_code: len([rule for rule in config.get("psr_rules", []) if rule.get("form_code") == form_code])
        for form_code in form_codes
    }
    active_tab = request.query_params.get("tab") or "overview"
    if active_tab not in {"overview", "forms", "markets", "psr"}:
        active_tab = "overview"
    psr_selected_form = request.query_params.get("psr_form") or (form_codes[0] if form_codes else "")
    psr_query = str(request.query_params.get("psr_query") or "").strip()
    psr_status = str(request.query_params.get("psr_status") or "").strip()
    psr_filtered_rules = filter_psr_rules(config.get("psr_rules", []), psr_selected_form, psr_query, psr_status)
    psr_status_values = [
        str(rule.get("status") or "")
        for rule in config.get("psr_rules", [])
        if rule.get("status")
    ]
    form_status_values = [
        str(form.get("verification_status") or "")
        for form in config.get("forms", [])
        if form.get("verification_status")
    ]
    return {
        "saved": saved,
        "error": error,
        "config": display_config,
        "config_path": str(co_form_config_path()),
        "form_codes": form_codes,
        "psr_rule_counts": psr_rule_counts,
        "active_tab": active_tab,
        "settings_tabs": [
            {"id": "overview", "label": "Overview"},
            {"id": "forms", "label": "Forms"},
            {"id": "markets", "label": "Markets"},
            {"id": "psr", "label": "HS Criteria"},
        ],
        "enabled_form_count": len([row for row in config["forms"] if row.get("enabled")]),
        "enabled_market_count": len([row for row in config["market_presets"] if row.get("enabled")]),
        "picker_market_count": len([row for row in config["market_presets"] if row.get("enabled") and row.get("show_in_picker")]),
        "psr_selected_form": psr_selected_form,
        "psr_query": psr_query,
        "psr_status": psr_status,
        "psr_filtered_count": len(psr_filtered_rules),
        "psr_visible_rules": [with_co_form_status_label(rule) for rule in psr_filtered_rules[:120]],
        "form_status_options": co_form_status_options(form_status_values),
        "psr_status_options": co_form_status_options(psr_status_values),
    }


def co_form_config_from_form(form) -> dict:
    current = load_co_form_config()
    source_note = current.get("source_note", "")
    if "source_note" in form:
        source_note = form.get("source_note", source_note)
    form_priority = current.get("form_priority", [])
    if "form_priority" in form:
        form_priority = unique_text_list(form.get("form_priority", ""))

    forms = current.get("forms", [])
    if "form_count" in form:
        form_count = int(str(form.get("form_count") or "0") or "0")
        forms = []
        for index in range(form_count):
            form_code = str(form.get(f"form_{index}_form_code") or "").strip()
            if not form_code:
                continue
            forms.append({
                "form_code": form_code,
                "display_name": form.get(f"form_{index}_display_name", ""),
                "agreement": form.get(f"form_{index}_agreement", ""),
                "instrument": form.get(f"form_{index}_instrument", ""),
                "instrument_note": form.get(f"form_{index}_instrument_note", ""),
                "source_label": form.get(f"form_{index}_source_label", ""),
                "source_url": form.get(f"form_{index}_source_url", ""),
                "verification_status": form.get(f"form_{index}_verification_status", ""),
                "enabled": form.get(f"form_{index}_enabled") == "1",
            })

    market_presets = current.get("market_presets", [])
    if "market_count" in form:
        market_count = int(str(form.get("market_count") or "0") or "0")
        market_presets = []
        for index in range(market_count + 1):
            market = str(form.get(f"market_{index}_market") or "").strip()
            form_code = str(form.get(f"market_{index}_form_code") or "").strip()
            if not market or not form_code:
                continue
            market_presets.append({
                "market": market,
                "label": form.get(f"market_{index}_label", ""),
                "form_code": form_code,
                "aliases": unique_text_list(form.get(f"market_{index}_aliases", "")),
                "selection_reason": form.get(f"market_{index}_selection_reason", ""),
                "source_label": form.get(f"market_{index}_source_label", ""),
                "enabled": form.get(f"market_{index}_enabled") == "1",
                "show_in_picker": form.get(f"market_{index}_show_in_picker") == "1",
            })

    psr_rules = current.get("psr_rules", [])
    if "psr_count" in form:
        psr_count = int(str(form.get("psr_count") or "0") or "0")
        psr_rules = []
        for index in range(psr_count + 1):
            form_code = str(form.get(f"psr_{index}_form_code") or "").strip()
            hs_scope = str(form.get(f"psr_{index}_hs_scope") or "").strip()
            criteria = str(form.get(f"psr_{index}_criteria") or "").strip()
            if not form_code or not hs_scope or not criteria:
                continue
            psr_rules.append({
                "form_code": form_code,
                "hs_scope": hs_scope,
                "criteria": criteria,
                "source_reference": form.get(f"psr_{index}_source_reference", ""),
                "note": form.get(f"psr_{index}_note", ""),
                "status": form.get(f"psr_{index}_status", ""),
                "enabled": form.get(f"psr_{index}_enabled") == "1",
            })
    elif "psr_visible_count" in form:
        psr_rules = list(psr_rules)
        psr_visible_count = int(str(form.get("psr_visible_count") or "0") or "0")
        for index in range(psr_visible_count + 1):
            form_code = str(form.get(f"psr_{index}_form_code") or "").strip()
            hs_scope = str(form.get(f"psr_{index}_hs_scope") or "").strip()
            criteria = str(form.get(f"psr_{index}_criteria") or "").strip()
            if not form_code or not hs_scope or not criteria:
                continue
            rule = {
                "form_code": form_code,
                "hs_scope": hs_scope,
                "criteria": criteria,
                "source_reference": form.get(f"psr_{index}_source_reference", ""),
                "note": form.get(f"psr_{index}_note", ""),
                "status": form.get(f"psr_{index}_status", ""),
                "enabled": form.get(f"psr_{index}_enabled") == "1",
            }
            original_index = str(form.get(f"psr_{index}_original_index") or "").strip()
            if original_index.isdigit() and int(original_index) < len(psr_rules):
                psr_rules[int(original_index)] = rule
            else:
                psr_rules.append(rule)

    return sanitize_co_form_config({
        **current,
        "source_note": source_note,
        "form_priority": form_priority,
        "forms": forms,
        "market_presets": market_presets,
        "psr_rules": psr_rules,
    })


def filter_psr_rules(rules: list[dict], form_code: str, query: str, status: str) -> list[dict]:
    query_key = co_form_filter_key(query)
    output = []
    for index, rule in enumerate(rules):
        if form_code and rule.get("form_code") != form_code:
            continue
        if status and rule.get("status") != status:
            continue
        if query_key:
            haystack = co_form_filter_key(" ".join([
                str(rule.get("hs_scope") or ""),
                str(rule.get("criteria") or ""),
                str(rule.get("source_reference") or ""),
                str(rule.get("note") or ""),
                str(rule.get("status") or ""),
                co_form_status_label(str(rule.get("status") or "")),
            ]))
            if query_key not in haystack:
                continue
        output.append({**rule, "original_index": index})
    return output


def with_co_form_status_label(row: dict) -> dict:
    status = str(row.get("status") or "")
    return {**row, "status_label": co_form_status_label(status)}


def co_form_status_options(statuses: list[str]) -> list[dict]:
    output = []
    seen = set()
    for status in list(PSR_STATUS_LABELS) + sorted(set(statuses)):
        if not status or status in seen:
            continue
        output.append({"value": status, "label": co_form_status_label(status)})
        seen.add(status)
    return output


def co_form_status_label(status: str) -> str:
    if not status:
        return ""
    return PSR_STATUS_LABELS.get(status, status.replace("_", " ").strip().capitalize())


def co_form_filter_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


@app.get("/user", response_class=HTMLResponse)
async def user_page(request: Request, logged_out: str = ""):
    user = co_auth.current_user(request)
    return templates.TemplateResponse(
        request=request,
        name="user.html",
        context={
            "user": user,
            "logged_out": logged_out == "1",
            "visible_clients": sorted(co_auth.visible_client_ids(user) or []) if user and co_auth.visible_client_ids(user) is not None else [],
            "all_clients": bool(user and co_auth.visible_client_ids(user) is None),
        },
    )


@app.get("/settings", response_class=HTMLResponse)
async def settings_page(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="settings.html",
        context={
            "can_view_technical_settings": co_auth.can_view_technical_settings(co_auth.current_user(request)),
        },
    )


@app.get("/settings/co-forms", response_class=HTMLResponse)
async def co_form_settings_page(request: Request, saved: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="co_form_settings.html",
        context=co_form_settings_context(request, saved=saved == "1"),
    )


@app.post("/settings/co-forms", response_class=HTMLResponse)
async def save_co_form_settings(request: Request):
    form = await request.form()
    active_tab = str(form.get("active_tab") or "").strip()
    try:
        save_co_form_config(co_form_config_from_form(form))
    except Exception as exc:
        return templates.TemplateResponse(
            request=request,
            name="co_form_settings.html",
            context=co_form_settings_context(request, error=str(exc)),
            status_code=400,
        )
    suffix = f"&tab={quote(active_tab)}" if active_tab else ""
    return RedirectResponse(f"/settings/co-forms?saved=1{suffix}", status_code=303)


@app.post("/settings/co-forms/reset")
async def reset_co_form_settings(request: Request):
    reset_co_form_config()
    return RedirectResponse("/settings/co-forms?saved=1", status_code=303)


@app.get("/settings/technical", response_class=HTMLResponse)
@app.get("/settings/data-hub", response_class=HTMLResponse)
async def data_hub_settings_page(request: Request, saved: str = ""):
    require_technical_settings_dev(request)
    return templates.TemplateResponse(
        request=request,
        name="data_hub_settings.html",
        context=data_hub_settings_context(request, saved=saved == "1"),
    )


@app.post("/settings/technical", response_class=HTMLResponse)
@app.post("/settings/data-hub", response_class=HTMLResponse)
async def save_data_hub_settings(request: Request):
    require_technical_settings_dev(request)
    form = await request.form()
    payload = data_hub_override_payload(form, load_data_hub_overrides())
    try:
        DataHubLinkSettings.from_env(payload)
        DataHubLinkSettings.from_env({**payload, **os.environ})
    except RuntimeError as exc:
        return templates.TemplateResponse(
            request=request,
            name="data_hub_settings.html",
            context=data_hub_settings_context(request, error=str(exc)),
            status_code=400,
        )
    save_data_hub_overrides(payload)
    return RedirectResponse("/settings/technical?saved=1", status_code=303)


@app.post("/settings/technical/test", response_class=HTMLResponse)
@app.post("/settings/data-hub/test", response_class=HTMLResponse)
async def test_data_hub_settings(request: Request):
    require_technical_settings_dev(request)
    return templates.TemplateResponse(
        request=request,
        name="data_hub_settings.html",
        context=data_hub_settings_context(request, test_result=data_hub_link_check()),
    )


@app.get("/auth/login")
async def auth_login(request: Request, next: str = "/clients"):
    redirect_uri = f"{co_auth.co_public_base_url(request)}/auth/callback"
    return RedirectResponse(
        co_auth.data_hub_authorize_url(redirect_uri=redirect_uri, state=next),
        status_code=303,
    )


@app.post("/auth/logout")
async def auth_logout(request: Request, next_url: str = Form("/clients")):
    next_path = co_auth.safe_next_path(next_url)
    if co_auth.auth_required():
        request.state.co_user = None
        response = templates.TemplateResponse(
            request=request,
            name="sso_logout.html",
            context={
                "data_hub_logout_url": co_auth.data_hub_logout_url(),
                "next_path": next_path,
            },
        )
        co_auth.clear_session_cookie(response)
        return response
    target = f"{next_path}?logged_out=1" if next_path == "/user" else next_path
    response = RedirectResponse(target, status_code=303)
    co_auth.clear_session_cookie(response)
    return response


@app.get("/auth/callback", name="auth_callback")
async def auth_callback(request: Request, code: str = "", state: str = "/clients"):
    next_url = co_auth.safe_next_path(state)
    if not code:
        return RedirectResponse(f"/auth/login?next={quote(next_url, safe='/')}", status_code=303)
    redirect_uri = f"{co_auth.co_public_base_url(request)}/auth/callback"
    try:
        payload = co_auth.exchange_data_hub_sso_code(code, redirect_uri=redirect_uri)
        token = str(payload["access_token"])
        verifier = co_auth.DataHubTokenVerifier(
            issuer=co_auth.data_hub_issuer_urls(),
            jwks_provider=lambda: co_auth.fetch_data_hub_jwks(co_auth.data_hub_jwks_url()),
        )
        verifier.verify(token)
    except Exception:
        response = PlainTextResponse(
            "Data Hub login failed. Check Technical Settings for issuer/JWKS and Data Hub SSO config.",
            status_code=401,
        )
        co_auth.clear_session_cookie(response)
        return response
    response = RedirectResponse(next_url, status_code=303)
    co_auth.set_session_cookie(response, token, int(payload.get("expires_in") or 600))
    return response


def require_local_source_writes() -> None:
    if co_auth.data_hub_source_mode_enabled():
        raise HTTPException(
            status_code=409,
            detail="Shared source data is read-only in CO when DATA_HUB_ENABLED is active. Use Data Hub for source changes.",
        )


# Local CO state uses short client IDs (e.g. "johnson") while Data Hub stores
# them suffixed with a country code (e.g. "johnson-vn"). When the literal
# Data Hub lookup 404s, retry with these suffix variants before falling back
# to the local registry. Edit when new tenants join.
_CLIENT_ID_FALLBACK_SUFFIXES: tuple[str, ...] = ("-vn",)


def resolve_client(client_id: str) -> dict:
    """Resolve a CO client by short ID, mapping to Data Hub's suffix variant.

    Data Hub stores tenants with a country suffix (e.g. `growatt-vn`) while
    CO local state and URLs use the short form (`growatt`). The Data Hub
    client wrapper itself retries 404s on the suffix variant, so the lookup
    succeeds — but the returned record carries the Data Hub ID. Force the
    short ID back onto the result so downstream lookups (`get_case_record`,
    Postgres queries) stay consistent with CO state.

    Identity fields the agency edits in CO (legal_name + tax_code, used on
    the bảng kê HQ render) are stored locally and overlay whatever Data Hub
    returns — Data Hub does not yet expose `legal_name`, and tax_code is
    often "Chưa nhập" upstream.
    """
    service_client = getattr(portfolio_service, "client", None)
    if callable(service_client):
        try:
            resolved = service_client(client_id)
        except httpx.HTTPStatusError as exc:
            if exc.response.status_code != 404:
                raise
            return _apply_co_identity_overlay(client_id, registry_get_client(client_id))
        if isinstance(resolved, dict):
            resolved = dict(resolved)
            resolved["id"] = client_id
            resolved["client_id"] = client_id
        return _apply_co_identity_overlay(client_id, resolved)
    return _apply_co_identity_overlay(client_id, registry_get_client(client_id))


def effective_min_gap_days(client: dict | None, client_config: dict | None = None) -> int:
    """Resolve the final 2-day rule threshold for a client.

    1. CO-local `co_stock_overrides.min_days_before_export` wins — it's
       the value the operator set via the CO client-config form.
    2. Falls back to `client_config["co_stock"]["min_days_before_export"]`
       (which today lives in Data Hub when DH source mode is enabled).
    3. Falls back to `DEFAULT_MIN_GAP_DAYS` (2).

    Centralised so the calculate / substitute paths all see the same
    answer without each one re-implementing the lookup.
    """
    if isinstance(client, dict):
        overrides = client.get("co_stock_overrides")
        if isinstance(overrides, dict) and "min_days_before_export" in overrides:
            try:
                value = int(overrides["min_days_before_export"])
                if value >= 0:
                    return value
            except (TypeError, ValueError):
                pass
    return co_stock_eligibility.min_gap_days_from_config(client_config)


def _apply_co_identity_overlay(client_id: str, client: dict) -> dict:
    if not isinstance(client, dict):
        return client
    store = get_app_state_store()
    if not store:
        return client
    try:
        local = store.client(client_id)
    except KeyError:
        return client
    legal_name = str(local.get("legal_name") or "").strip()
    tax_code = str(local.get("tax_code") or "").strip()
    if legal_name:
        client["legal_name"] = legal_name
    if tax_code and tax_code != "Chưa nhập":
        client["tax_code"] = tax_code
    overrides = local.get("co_stock_overrides")
    if isinstance(overrides, dict) and overrides:
        client["co_stock_overrides"] = dict(overrides)
    return client


def default_client_case(client: dict) -> dict:
    case = clone_case(DEMO_CASE)
    case.update(
        {
            "id": f"{client['id']}-empty-co-case",
            "customer": client.get("legal_name") or client["name"],
            "customer_legal_name": client.get("legal_name", ""),
            "customer_tax_code": client.get("tax_code", ""),
            "case_code": "Chưa tạo",
            "title": f"Hồ sơ C/O {client['name']}",
            "destination_market": "Chưa nhập",
            "agreement": "Chưa nhập",
            "co_form_type": "Chưa nhập",
            "source_label": "Chưa có dữ liệu C/O",
            "products": [],
        }
    )
    return attach_results(case)


def client_case(client: dict) -> dict:
    try:
        return get_client_case(client["id"])
    except KeyError:
        return default_client_case(client)


def client_context(client_id: str, active: str, **extra):
    client = resolve_client(client_id)
    case = extra.pop("case", client_case(client))
    source_workspace, source_backend = source_workspace_for_client(client)
    client = enrich_client_with_source_workspace(client, source_workspace)
    bom_workspace = bom_service.workspace(client)
    case = attach_case_bom_snapshot(case, bom_workspace)
    case = attach_case_source_snapshot(case, source_workspace)
    # Deep-link to the Data Hub page for tabs that have a canonical DH surface.
    # None for everything else (incl. /config) so the template guard hides the
    # "Mở trên Data Hub" button instead of rendering an empty href.
    data_hub_target_url = None
    if source_backend == "data-hub" and active in {"catalog", "bom", "bcct"}:
        dh_base = data_hub_link_settings().data_hub_base_url
        data_hub_target_url = f"{dh_base.rstrip('/')}/clients/{client_id}/{active}"
    return {
        "client": client,
        "case": case,
        "active": active,
        "data_hub_target_url": data_hub_target_url,
        "bom_workspace": bom_workspace,
        "source_workspace": source_workspace,
        "client_config": source_workspace["client_config"],
        "case_workspace": extra.pop("case_workspace", get_case_workspace(client)),
        "form_candidates": extra.pop("form_candidates", form_candidates_for_market(case.get("destination_market", ""))),
        "form_lanes": prioritized_form_lanes(case.get("destination_market", ""), case_finished_hs_codes(case)),
        "recommended_form_lane": recommended_form_lane(
            prioritized_form_lanes(case.get("destination_market", ""), case_finished_hs_codes(case))
        ),
        "common_market_presets": COMMON_MARKET_PRESETS,
        "common_market_guidance": common_market_guidance(),
        "co_form_options": [
            {"form_code": row["form_code"], "display_name": row.get("display_name") or row["form_code"]}
            for row in load_co_form_config().get("forms", [])
            if row.get("enabled")
        ],
        "invoice_matches": extra.pop("invoice_matches", []),
        "invoice_criteria_rows": extra.pop("invoice_criteria_rows", []),
        "criteria_rows": extra.pop("criteria_rows", []),
        "source_notes": SOURCE_NOTES,
        "source_backend": source_backend,
        **extra,
    }


def source_workspace_for_client(client: dict) -> tuple[dict, str]:
    return portfolio_service.source_workspace(client)


def co_case_light_context(client_id: str, case: dict, current_step: str, **extra) -> dict:
    client = resolve_client(client_id)
    fast_origin_context = bool(extra.pop("fast_origin_context", False))
    cached_case_context = fast_origin_context or bool(extra.pop("cached_case_context", False))
    force_source_refresh = bool(extra.pop("force_source_refresh", False))
    use_cached_context = cached_case_context and not force_source_refresh and bool(case.get("source_snapshot"))
    if use_cached_context:
        source_context = cached_origin_source_context(client, case)
    elif current_step == "origin":
        # Origin tab-load: stock from the materialized CO-stock snapshot +
        # narrow export invoice_matches, never the ~40s full BCCT pull.
        source_context = origin_source_context(client, case)
    else:
        # Non-origin steps only need source_summary + invoice_matches; skip the
        # heavy materials + BCCT pagination so the shipment tab (the default
        # landing tab) renders fast instead of ~21s for big clients.
        source_context = co_case_source_context(client, case, skip_heavy_context=True)
    if not use_cached_context:
        # Warm the TTL cache so the next substitute-modal open in this session
        # reuses the same Data Hub fetch (avoids 30s re-pagination for Johnson).
        # Only warm when a real materials pull happened: the converged origin
        # path returns material_rows=[] and must not overwrite the cache the
        # substitute-modal HS heuristic relies on.
        if source_context.get("material_rows"):
            import time
            persisted = case.get("persisted_case_id") or case.get("id") or ""
            shipment = case.get("shipment") or {}
            fingerprint = (
                client.get("id", ""),
                str(persisted),
                str(shipment.get("invoice_no") or ""),
                ",".join(sorted(shipment.get("export_declaration_nos") or [])),
                str(len(case.get("products") or [])),
            )
            _CO_CASE_SOURCE_CACHE[fingerprint] = (time.time(), source_context)
    source_summary = source_context["source_summary"]
    invoice_matches = source_context["invoice_matches"]
    reference_warnings = shipment_reference_warnings(case.get("shipment", {}), invoice_matches)
    case_workspace = extra.pop("case_workspace")
    form_candidates = extra.pop("form_candidates")
    criteria_rows = extra.pop("criteria_rows")
    if current_step == "origin":
        bom_product_codes = co_case_bom_product_codes(case, invoice_matches)
        picker_case_id = str(case.get("persisted_case_id") or case.get("id") or "")
        bom_workspace = (
            bom_service.workspace(client, product_codes=bom_product_codes, case_id=picker_case_id)
            if bom_product_codes
            else minimal_bom_workspace()
        )
        if use_cached_context and case.get("products") and not bom_workspace.get("product_versions"):
            bom_workspace = bom_workspace_from_case_snapshot(case)
    else:
        bom_workspace = minimal_bom_workspace()
    origin_demo_allowed = extra.pop("origin_demo_allowed", True)
    preserve_origin_products = extra.pop("preserve_origin_products", False)
    origin_calculation_blocked = bool(extra.get("origin_calculation_blocked", False))
    client = enrich_client_with_source_summary(client, source_summary)
    case = attach_case_source_summary_snapshot(case, source_summary)
    if not use_cached_context:
        case["source_invoice_matches"] = json_safe(invoice_matches)
    if current_step == "origin" and not origin_calculation_blocked and not use_cached_context:
        selected_lane = recommended_form_lane(
            prioritized_form_lanes(case.get("destination_market", ""), co_case_hs_codes(case, invoice_matches))
        )
        case = prepare_case_origin_product_shells(
            case,
            invoice_matches,
            bom_workspace,
            selected_lane,
            preserve_existing=preserve_origin_products,
        )
        case = attach_case_bom_snapshot(case, bom_workspace)
        case = attach_origin_bom_product_codes(case, bom_workspace)
        if case.get("products"):
            case = attach_origin_readiness(case)
            case = attach_results(case)
            case = attach_origin_sheet_states(case)
            criteria_rows = build_case_criteria_rows(case, form_candidates)
    elif current_step == "origin" and case.get("products"):
        if not use_cached_context:
            case = attach_case_bom_snapshot(case, bom_workspace)
            case = attach_origin_bom_product_codes(case, bom_workspace)
        case = attach_origin_readiness(case)
        case = attach_results(case)
        case = attach_origin_sheet_states(case)
        criteria_rows = build_case_criteria_rows(case, form_candidates)
    elif current_step == "origin":
        case = attach_case_bom_snapshot(case, bom_workspace)
    origin_demo_active = origin_demo_allowed and should_show_origin_demo(current_step, case, invoice_matches)
    if origin_demo_active:
        case = attach_origin_demo(case)
        case = attach_origin_readiness(case)
        case = attach_origin_sheet_states(case)
        criteria_rows = build_case_criteria_rows(case, form_candidates)
    form_lanes = prioritized_form_lanes(case.get("destination_market", ""), co_case_hs_codes(case, invoice_matches))
    selected_form_lane = recommended_form_lane(form_lanes)
    invoice_criteria_rows = invoice_match_criteria_rows(invoice_matches, selected_form_lane)
    invoice_lookup_preview = invoice_preview_from_matches(
        case.get("shipment", {}).get("invoice_no", ""),
        invoice_matches,
    )
    if not criteria_rows and invoice_criteria_rows:
        criteria_rows = invoice_criteria_rows
    context = {
        "client": client,
        "case": case,
        "active": "co-case",
        "bom_workspace": bom_workspace,
        "source_workspace": {},
        "client_config": source_summary["client_config"],
        "case_workspace": case_workspace,
        "form_candidates": form_candidates,
        "form_lanes": form_lanes,
        "recommended_form_lane": selected_form_lane,
        "invoice_lookup_preview": invoice_lookup_preview,
        "common_market_presets": COMMON_MARKET_PRESETS,
        "common_market_guidance": common_market_guidance(),
        "co_form_options": [
            {"form_code": row["form_code"], "display_name": row.get("display_name") or row["form_code"]}
            for row in load_co_form_config().get("forms", [])
            if row.get("enabled")
        ],
        "tkx_tkn_summary": case_tkx_tkn_summary(
            case,
            invoice_matches,
            source_context.get("stock_rows") or [],
            source_context.get("declaration_file_counts") or {},
        ),
        "data_hub_base_url": data_hub_link_settings().data_hub_base_url,
        "invoice_matches": invoice_matches,
        "origin_source_context": source_context,
        "shipment_reference_warnings": reference_warnings,
        "invoice_criteria_rows": invoice_criteria_rows,
        "criteria_rows": criteria_rows,
        "origin_demo_active": origin_demo_active,
        "origin_demo_material_count": origin_material_count(case) if origin_demo_active else 0,
        "origin_case_revision": origin_case_revision,
        "source_notes": SOURCE_NOTES,
        "source_backend": source_context["source_backend"],
        **extra,
    }
    context["co_case_active_step"] = current_step
    context["co_case_steps"] = co_case_workflow_steps(
        client_id,
        context["case"],
        current_step,
        invoice_matches=invoice_matches,
        criteria_rows=criteria_rows,
        origin_demo_active=origin_demo_active,
        tkx_tkn_summary=context.get("tkx_tkn_summary"),
    )
    return context


_CO_CASE_SOURCE_CACHE: dict[tuple[str, str], tuple[float, dict]] = {}
_CO_CASE_SOURCE_CACHE_TTL_SECONDS = 90.0


def co_case_source_context(client: dict, case: dict, *, skip_heavy_context: bool = False) -> dict:
    return portfolio_service.co_case_source_context(
        client, case, skip_heavy_context=skip_heavy_context
    )


def co_case_source_context_cached(client: dict, case: dict) -> dict:
    """TTL-cached source_context for high-frequency endpoints (substitute modal,
    typeahead) where the underlying Data Hub catalog rarely changes within a
    session. Cache key includes case_id + shipment fingerprint so different
    cases / mutated shipments don't collide.

    For Johnson the underlying call paginates 11k materials + 65k BCCT rows
    (multi-second). Without this cache, every modal open re-paginated.

    Stock rows returned by this function ALWAYS reflect the current ledger
    state (used_qty / remaining_qty net of locked claims across all cases).
    """
    import time

    persisted = case.get("persisted_case_id") or case.get("id") or ""
    shipment = case.get("shipment") or {}
    fingerprint = (
        client.get("id", ""),
        str(persisted),
        str(shipment.get("invoice_no") or ""),
        ",".join(sorted(shipment.get("export_declaration_nos") or [])),
        str(len(case.get("products") or [])),
    )
    now = time.time()
    cached = _CO_CASE_SOURCE_CACHE.get(fingerprint)
    if cached and now - cached[0] < _CO_CASE_SOURCE_CACHE_TTL_SECONDS:
        context = cached[1]
    else:
        context = co_case_source_context(client, case)
        _CO_CASE_SOURCE_CACHE[fingerprint] = (now, context)
        if len(_CO_CASE_SOURCE_CACHE) > 32:
            oldest = sorted(_CO_CASE_SOURCE_CACHE.items(), key=lambda kv: kv[1][0])[0][0]
            _CO_CASE_SOURCE_CACHE.pop(oldest, None)
    # Always re-apply the ledger + manual adjustments — claim state changes
    # outside the cache window (sheet lock/unlock invalidates entry, adjustments
    # can be imported any time, but be defensive in case caller bypasses).
    client_id = client.get("id", "")
    used_by_lot = co_stock_ledger.used_qty_by_lot(client_id)
    adjustments = co_stock_adjustments_store.aggregate_by_lookup_key(client_id)
    if (used_by_lot or adjustments) and context.get("stock_rows"):
        # Don't mutate the cached list in place if someone else holds a reference;
        # snapshot a new list with applied ledger + adjustment values.
        rows = [dict(row) for row in context["stock_rows"]]
        if used_by_lot:
            rows = co_stock_ledger.apply_used_qty(rows, used_by_lot)
        if adjustments:
            rows = co_stock_adjustments_store.apply_adjustments(rows, adjustments)
        context = {**context, "stock_rows": rows}
    return context


def origin_source_context(client: dict, case: dict) -> dict:
    """Converged source context for the origin tab-load (cold path).

    Replaces the heavy co_case_source_context — whose dominant cost is the full
    list_bcct pull (~40s for Johnson) — with:
      - stock_rows from the materialized CO-stock snapshot (the same source the
        /calculate path uses, net of the ledger), so the preview matches the
        computed result;
      - invoice_matches from a narrow Data Hub fetch (per-declaration / by-codes
        export), persisted on the case downstream for the warm path to reuse;
      - material_rows = [] (unused at tab render; the substitute modal self-fetches).

    This is the COLD path: the warm reuse of case["source_invoice_matches"]
    lives upstream in cached_origin_source_context (gated on use_cached_context).
    We always re-fetch here so a force_source_refresh actually refreshes — never
    serve possibly-stale cached matches when the caller asked to bypass the cache.

    Falls back to the legacy full pull when the snapshot is unusable (no DB /
    empty for this client) so an operator never sees an empty stock preview.
    """
    # The converged snapshot path is a Data-Hub-mode optimization: stock comes
    # from the materialized co_stock snapshot built off Data Hub BCCT. In
    # file-store mode (tests / offline dev) there is no such snapshot, so use
    # the cheap in-memory heavy path directly and never touch the materializer.
    if getattr(portfolio_service, "data_hub", None) is None:
        return co_case_source_context(client, case)
    snapshot_stock = _calculate_stock_rows_from_snapshot(client)
    if snapshot_stock is None:
        return co_case_source_context(client, case)

    source_summary, source_backend = portfolio_service.source_summary(client)
    client_config = source_summary.get("client_config", {}) if isinstance(source_summary, dict) else {}
    invoice_matches = portfolio_service.origin_invoice_matches(client, case, client_config)

    declaration_file_counts = {"export": {}, "import": {}}
    if hasattr(portfolio_service, "declaration_file_counts"):
        declaration_file_counts = portfolio_service.declaration_file_counts(
            client.get("id", ""), case, invoice_matches
        )
    return {
        "source_backend": source_backend,
        "source_summary": source_summary,
        "invoice_matches": invoice_matches,
        "material_rows": [],
        "stock_rows": snapshot_stock,
        "declaration_file_counts": declaration_file_counts,
    }


def record_sheet_lock_claims(client_id: str, case_id: str, product_code: str, case: dict) -> int:
    """Persist allocation lines from a locked sheet to the cross-case stock ledger.

    Reads the sheet's products[*].materials[*].allocation_lines and writes one
    claim per (source_row, material_code) so other cases see remaining_qty drop.
    Idempotent: re-locking the same sheet replaces prior claims for it.

    `case` may be the form-rebuilt case (which strips allocation_lines), so
    fall back to the persisted case from disk to get the canonical allocations.
    """
    target = _sheet_with_allocations(client_id, case_id, product_code, case)
    if not target:
        return 0
    allocations: list[dict] = []
    for material_index, material in enumerate(target.get("materials", []) or []):
        material_code = str(material.get("material_code") or material.get("internal_material_code") or "").strip()
        for line in material.get("allocation_lines", []) or []:
            allocations.append({
                "source_row": line.get("source_row", ""),
                "material_code": material_code,
                "material_index": material_index,
                "claimed_qty": line.get("allocated_qty", "0"),
                "declaration_no": line.get("import_declaration_no", ""),
                "line_no": line.get("import_line_no", ""),
                "customs_code": line.get("customs_material_code", "") or line.get("customs_code", ""),
            })
    return co_stock_ledger.record_sheet_lock(client_id, case_id, product_code, allocations)


def _sheet_with_allocations(client_id: str, case_id: str, product_code: str, case: dict) -> dict | None:
    """Pick the product entry, preferring the in-memory case but falling back to
    the persisted record on disk if its materials lack allocation_lines."""

    def _find(case_obj: dict | None) -> dict | None:
        if not case_obj:
            return None
        return next(
            (p for p in case_obj.get("products", []) if str(p.get("code") or "").strip() == product_code),
            None,
        )

    target = _find(case)
    if target and any(material.get("allocation_lines") for material in target.get("materials", []) or []):
        return target
    try:
        client = resolve_client(client_id)
        persisted = persisted_origin_case(client, case_id)
    except Exception:  # noqa: BLE001
        return target
    persisted_target = _find(persisted)
    return persisted_target or target


def invalidate_co_case_source_cache(client_id: str = "", case_id: str = "") -> None:
    """Clear cache entries — call when case mutates (lock, override, etc.)."""
    if not client_id and not case_id:
        _CO_CASE_SOURCE_CACHE.clear()
        return
    keys_to_drop = [
        key for key in _CO_CASE_SOURCE_CACHE
        if (not client_id or key[0] == client_id) and (not case_id or key[1] == case_id)
    ]
    for key in keys_to_drop:
        _CO_CASE_SOURCE_CACHE.pop(key, None)


def invoice_matches_only(client: dict, shipment: dict) -> list[dict]:
    """Lightweight invoice-match lookup for invoice-preview / search dropdown.

    File-store mode: falls through to match_case_bcct_exports (in-memory,
    surfaces invoice/declaration mismatch warnings).
    Data Hub mode: calls invoice_matches adapter directly (~300-500ms),
    avoiding the full materials + BCCT pagination that co_case_source_context
    would otherwise trigger on every keystroke for big clients.
    """
    invoice_no = str(shipment.get("invoice_no") or "").strip()
    declaration_nos = declaration_refs(shipment.get("export_declaration_nos"))
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is None or not hasattr(data_hub, "invoice_matches"):
        # File-store / test mode: defer to the existing co_case_source_context
        # (in-memory or fake), which surfaces market_hint + reference_warning
        # via the canonical match path.
        try:
            source_context = co_case_source_context(client, {"shipment": shipment})
        except Exception:  # noqa: BLE001
            source_context = {}
        return source_context.get("invoice_matches") or []
    # Data Hub mode: exact-invoice lookup is indexed.
    try:
        client_config = portfolio_service.get_client_config(client) if hasattr(portfolio_service, "get_client_config") else {}
    except Exception:  # noqa: BLE001
        client_config = {}
    relevant_types = list(client_config.get("bcct", {}).get("relevant_export_declaration_types", []))
    matches: list[dict] = []
    if invoice_no:
        try:
            matches = data_hub.invoice_matches(client["id"], invoice_no, relevant_types)
        except Exception:  # noqa: BLE001
            matches = []
    if declaration_nos and not matches:
        for declaration in declaration_nos:
            matches.extend(declaration_invoice_matches(client, declaration, exact=True, include_invoice=False))
    elif declaration_nos and matches:
        wanted = {re.sub(r"[^A-Z0-9]", "", str(decl).upper()) for decl in declaration_nos}
        matches = [
            row for row in matches
            if re.sub(r"[^A-Z0-9]", "", str(row.get("declaration_no") or "").upper()) in wanted
        ] or matches
    return matches


def declaration_file_count(
    declaration_file_counts: dict,
    direction: str,
    declaration_no: str,
) -> int:
    key = str(declaration_no or "").strip()
    if not key:
        return 0
    if (direction, key) in declaration_file_counts:
        return int(declaration_file_counts.get((direction, key)) or 0)
    direction_counts = declaration_file_counts.get(direction) if isinstance(declaration_file_counts, dict) else {}
    if isinstance(direction_counts, dict):
        return int(direction_counts.get(key) or 0)
    return 0


def case_tkx_tkn_summary(
    case: dict,
    invoice_matches: list[dict],
    stock_rows: list[dict],
    declaration_file_counts: dict | None = None,
) -> dict:
    """Aggregate TKX (export) and TKN (import) declarations referenced by the case.

    TKX/TKN presence means the declaration file exists, not merely that a BCCT
    row exists. BCCT rows only identify which declarations the case references.
    """
    invoice_matches = invoice_matches or []
    stock_rows = stock_rows or []
    declaration_file_counts = declaration_file_counts or {}
    tkx: dict[str, dict] = {}
    for row in invoice_matches:
        key = str(row.get("declaration_no") or "").strip()
        if not key:
            continue
        file_count = declaration_file_count(declaration_file_counts, "export", key)
        entry = tkx.setdefault(key, {
            "declaration_no": key,
            "in_data_hub": file_count > 0,
            "file_count": file_count,
            "lines": [],
            "declaration_type": str(row.get("declaration_type") or ""),
        })
        entry["lines"].append({
            "line_no": row.get("line_no", ""),
            "item_code": row.get("item_code", ""),
            "hs_code": row.get("hs_code", ""),
            "quantity": row.get("quantity", ""),
            "invoice_ref": row.get("invoice_ref", ""),
        })
    for declared in case.get("shipment", {}).get("export_declaration_nos", []) or []:
        declared_key = str(declared or "").strip()
        if not declared_key or declared_key in tkx:
            continue
        file_count = declaration_file_count(declaration_file_counts, "export", declared_key)
        tkx[declared_key] = {
            "declaration_no": declared_key,
            "in_data_hub": file_count > 0,
            "file_count": file_count,
            "lines": [],
            "declaration_type": "",
        }

    tkn: dict[str, dict] = {}
    for product in case.get("products", []) or []:
        if str(product.get("origin_sheet_status") or "").strip() != "locked":
            continue
        product_code = str(product.get("code") or "").strip()
        for material in product.get("materials", []) or []:
            for line in material.get("allocation_lines", []) or []:
                key = str(line.get("import_declaration_no") or "").strip()
                if not key:
                    continue
                file_count = declaration_file_count(declaration_file_counts, "import", key)
                entry = tkn.setdefault(key, {
                    "declaration_no": key,
                    "in_data_hub": file_count > 0,
                    "file_count": file_count,
                    "lines": [],
                    "products": set(),
                })
                entry["products"].add(product_code)
                entry["lines"].append({
                    "product_code": product_code,
                    "material_code": material.get("material_code", ""),
                    "line_no": line.get("import_line_no", ""),
                    "allocated_qty": line.get("allocated_qty", ""),
                })
    for entry in tkn.values():
        entry["products"] = sorted(entry["products"])
    return {
        "tkx": sorted(tkx.values(), key=lambda e: e["declaration_no"]),
        "tkn": sorted(tkn.values(), key=lambda e: e["declaration_no"]),
        "missing_tkx": [e for e in tkx.values() if not e["in_data_hub"]],
        "missing_tkn": [e for e in tkn.values() if not e["in_data_hub"]],
    }


def preload_co_case_origin_context(client_id: str, case_id: str) -> None:
    client = resolve_client(client_id)
    try:
        record = get_case_record(client, case_id)
    except KeyError:
        return
    if record.get("source_snapshot") and record.get("products"):
        return
    context = co_case_context(
        client_id,
        case_id,
        current_step="origin",
        origin_demo_allowed=False,
        force_source_refresh=True,
    )
    if context.get("origin_demo_active"):
        return
    try:
        update_case_record(client, context["case"])
    except KeyError:
        return


def cached_origin_source_context(client: dict, case: dict) -> dict:
    cached_matches = case.get("source_invoice_matches") if isinstance(case.get("source_invoice_matches"), list) else []
    return {
        "source_backend": "case-snapshot",
        "source_summary": source_summary_from_case_snapshot(client, case),
        "invoice_matches": cached_matches or [origin_match_from_existing_product(product) for product in case.get("products", [])],
        "material_rows": [],
        "stock_rows": [],
    }


def source_summary_from_case_snapshot(client: dict, case: dict) -> dict:
    snapshot = case.get("source_snapshot") if isinstance(case.get("source_snapshot"), dict) else {}
    counts = client.get("counts") if isinstance(client.get("counts"), dict) else {}
    return {
        "client_config": {
            "config_version": snapshot.get("client_config_version", ""),
            "config_hash": snapshot.get("client_config_hash", ""),
        },
        "material_catalog": {
            "published_row_count": int(counts.get("materials") or 0),
            "latest_version": {
                "version_id": snapshot.get("material_catalog_version_id", ""),
                "version_no": snapshot.get("material_catalog_version_no", ""),
            },
        },
        "product_catalog": {
            "published_row_count": int(counts.get("products") or 0),
            "latest_version": {
                "version_id": snapshot.get("product_catalog_version_id", ""),
                "version_no": snapshot.get("product_catalog_version_no", ""),
            },
        },
        "bcct": {
            "published_row_count": int(counts.get("bcct") or snapshot.get("bcct_reviewed_row_count") or 0),
            "reviewed_row_count": int(snapshot.get("bcct_reviewed_row_count") or 0),
            "correction_candidate_count": int(snapshot.get("correction_candidate_count") or 0),
            "latest_version": {
                "version_id": snapshot.get("bcct_version_id", ""),
                "version_no": snapshot.get("bcct_version_no", ""),
            },
        },
        "co_stock_row_count": int(counts.get("co_stock") or 0),
    }


def bom_workspace_from_case_snapshot(case: dict) -> dict:
    snapshot = case.get("bom_snapshot") if isinstance(case.get("bom_snapshot"), dict) else {}
    composition = [dict(row) for row in snapshot.get("composition", []) if isinstance(row, dict)]
    product_versions = []
    seen = set()
    for row in composition:
        product_code = str(row.get("product_code") or "").strip()
        version_id = str(row.get("product_artifact_id") or row.get("product_version_id") or row.get("artifact_id") or row.get("version_id") or "").strip()
        if not product_code or not version_id or version_id in seen:
            continue
        seen.add(version_id)
        product_versions.append({
            "product_code": product_code,
            "product_artifact_id": version_id,
            "product_version_id": version_id,
            "version_id": version_id,
            "product_artifact_no": row.get("product_artifact_no") or row.get("product_version_no") or row.get("artifact_no") or row.get("version_no") or "",
            "product_version_no": row.get("product_version_no") or row.get("version_no") or "",
            "version_no": row.get("product_version_no") or row.get("version_no") or "",
            "row_count": row.get("row_count", 0),
            "status": row.get("status") or "snapshot",
            "rows": None,
        })
    for product in case.get("products", []):
        product_code = str(product.get("bom_product_code") or product.get("code") or "").strip()
        version_id = str(product.get("bom_product_artifact_id") or product.get("bom_product_version_id") or "").strip()
        if not product_code or not version_id or version_id in seen:
            continue
        seen.add(version_id)
        product_versions.append({
            "product_code": product_code,
            "product_artifact_id": version_id,
            "product_version_id": version_id,
            "version_id": version_id,
            "product_artifact_no": product.get("bom_product_artifact_no") or product.get("bom_product_version_no", ""),
            "product_version_no": product.get("bom_product_version_no", ""),
            "version_no": product.get("bom_product_version_no", ""),
            "row_count": len(product.get("materials", []) or []),
            "status": "snapshot",
            "rows": None,
        })
    options_by_code: dict[str, list[dict]] = {}
    for version in product_versions:
        options_by_code.setdefault(version["product_code"], []).append(version)
    aggregate = {
        "artifact_id": snapshot.get("aggregate_artifact_id") or snapshot.get("aggregate_version_id", ""),
        "artifact_no": snapshot.get("aggregate_artifact_no") or snapshot.get("aggregate_version_no", ""),
        "version_id": snapshot.get("aggregate_artifact_id") or snapshot.get("aggregate_version_id", ""),
        "version_no": snapshot.get("aggregate_artifact_no") or snapshot.get("aggregate_version_no", ""),
        "product_versions": composition,
        "rows": [],
    }
    return {
        "versions": [aggregate] if aggregate["version_id"] else [],
        "product_versions": product_versions,
        "product_version_options_by_code": options_by_code,
        "latest_version": aggregate,
        "latest_rows": [],
    }


def co_case_bom_product_codes(case: dict, invoice_matches: list[dict]) -> list[str]:
    codes = []
    for row in invoice_matches:
        bom_product_code = bom_product_code_from_material_identity(row) or str(row.get("bom_product_code") or "").strip()
        add_bom_code(
            codes,
            bom_product_code or str(row.get("item_code") or row.get("product_code") or row.get("customs_code") or "").strip(),
        )
    for product in case.get("products", []):
        add_bom_code(
            codes,
            str(product.get("bom_product_code") or product.get("code") or product.get("product_code") or "").strip(),
        )
    override_codes = {
        **(case.get("bom_product_version_overrides", {}) or {}),
        **(case.get("bom_product_artifact_overrides", {}) or {}),
    }
    for code in override_codes:
        add_bom_code(codes, str(code or "").strip())
    return codes


def add_bom_code(codes: list[str], code: str) -> None:
    code = str(code or "").strip()
    if code and code not in codes:
        codes.append(code)

def bom_code_candidates(code: str) -> list[str]:
    code = str(code or "").strip()
    if not code:
        return []
    return [code]


def invoice_lookup_payload(client: dict, invoice_no: str, query: str = "", export_declaration_nos: str | list[str] = "") -> dict:
    invoice_no = str(invoice_no or "").strip()
    declaration_nos = declaration_refs(export_declaration_nos)
    query = str(query or invoice_no or (declaration_nos[0] if declaration_nos else "")).strip()
    options = invoice_search_options(client, query)
    if not invoice_no and not declaration_nos:
        return {
            "status": "empty",
            "invoice_no": "",
            "reference_warnings": [],
            "options": options,
            "match_count": 0,
            "matches": [],
            "summary": {},
            "market_inference": market_inference_view({"status": "missing", "destination_market": "", "hints": []}),
            "suggested_forms": [],
        }
    resolved = resolve_shipment_reference(client, invoice_no, declaration_nos)
    lookup_invoice_no = resolved["invoice_no"]
    try:
        # Lightweight match path — invoice-preview only needs invoice_matches.
        # Calling co_case_source_context here would full-paginate materials + BCCT
        # from Data Hub on every keystroke (seconds per request for large clients).
        invoice_matches = invoice_matches_only(client, resolved["shipment"])
    except Exception as exc:
        return {
            "status": "error",
            "invoice_no": lookup_invoice_no,
            "reference_warnings": [],
            "options": options,
            "match_count": 0,
            "matches": [],
            "summary": {},
            "market_inference": market_inference_view({"status": "missing", "destination_market": "", "hints": []}),
            "suggested_forms": [],
            "message": f"Không tra được invoice: {exc}",
        }
    payload = invoice_preview_from_matches(lookup_invoice_no, invoice_matches)
    payload["reference_warnings"] = shipment_reference_warnings(
        resolved["shipment"],
        invoice_matches,
    )
    payload["options"] = options
    if resolved.get("source_reference"):
        payload["source_reference"] = resolved["source_reference"]
        payload["source_reference_type"] = resolved["source_reference_type"]
        payload["reference_label"] = primary_shipment_reference(resolved["shipment"])
        if not payload.get("invoice_no"):
            payload["invoice_no"] = resolved["source_reference"]
    return payload


def shipment_reference_warnings(shipment: dict, invoice_matches: list[dict]) -> list[str]:
    warnings = [
        str(row.get("reference_warning") or "").strip()
        for row in invoice_matches
        if str(row.get("reference_warning") or "").strip()
    ]
    declarations = declaration_refs(shipment.get("export_declaration_nos"))
    if declarations and not invoice_matches:
        warnings.append(
            f"Chưa thấy dòng BCCT xuất khẩu đã duyệt cho tờ khai {', '.join(declarations)}."
        )
    return unique_texts(warnings)


def resolve_shipment_reference(client: dict, reference: str, export_declaration_nos: str | list[str] = "") -> dict:
    reference = str(reference or "").strip()
    explicit_declarations = declaration_refs(export_declaration_nos)
    if explicit_declarations:
        matches = []
        for declaration in explicit_declarations:
            matches.extend(declaration_invoice_matches(client, declaration, exact=True, include_invoice=False))
        invoice_refs = sorted({row.get("invoice_ref", "") for row in matches if row.get("invoice_ref")})
        invoice_no = reference if reference and not declaration_invoice_matches(client, reference, exact=True, include_invoice=False) else ""
        if len(invoice_refs) == 1:
            invoice_no = invoice_no or invoice_refs[0]
        return {
            "invoice_no": invoice_no,
            "export_declaration_nos": explicit_declarations,
            "source_reference": ", ".join(explicit_declarations),
            "source_reference_type": "declaration",
            "shipment": {"invoice_no": invoice_no, "export_declaration_nos": explicit_declarations},
        }
    if not reference:
        return {
            "invoice_no": "",
            "export_declaration_nos": [],
            "source_reference": "",
            "source_reference_type": "",
            "shipment": {"invoice_no": "", "export_declaration_nos": []},
        }
    matches = declaration_invoice_matches(client, reference, exact=True, include_invoice=False)
    if matches:
        invoice_refs = sorted({row["invoice_ref"] for row in matches if row.get("invoice_ref")})
        invoice_no = invoice_refs[0] if len(invoice_refs) == 1 else ""
        return {
            "invoice_no": invoice_no,
            "export_declaration_nos": [reference],
            "source_reference": reference,
            "source_reference_type": "declaration",
            "shipment": {"invoice_no": invoice_no, "export_declaration_nos": [reference]},
        }
    return {
        "invoice_no": reference,
        "export_declaration_nos": [],
        "source_reference": "",
        "source_reference_type": "",
        "shipment": {"invoice_no": reference, "export_declaration_nos": []},
    }


def resolve_invoice_reference(client: dict, reference: str) -> dict:
    resolved = resolve_shipment_reference(client, reference)
    return {
        "invoice_no": resolved["invoice_no"],
        "source_reference": resolved["source_reference"],
        "source_reference_type": resolved["source_reference_type"],
    }


def primary_shipment_reference(shipment: dict) -> str:
    declarations = declaration_refs(shipment.get("export_declaration_nos"))
    if declarations:
        return "Tờ khai " + ", ".join(declarations)
    invoice_no = str(shipment.get("invoice_no") or "").strip()
    return f"Invoice {invoice_no}" if invoice_no else "Chưa nhập"


def has_shipment_reference(shipment: dict) -> bool:
    return bool(str(shipment.get("invoice_no") or "").strip() or declaration_refs(shipment.get("export_declaration_nos")))


def invoice_preview_from_matches(invoice_no: str, invoice_matches: list[dict]) -> dict:
    invoice_no = str(invoice_no or "").strip()
    inference = infer_market_from_invoice_matches(invoice_matches)
    hs_codes = co_case_hs_codes({"shipment": {"invoice_no": invoice_no}}, invoice_matches)
    suggested_forms = []
    if inference.get("status") == "ready":
        suggested_forms = [
            invoice_form_lane_view(row)
            for row in prioritized_form_lanes(inference["destination_market"], hs_codes)
        ]
    summary = invoice_match_summary(invoice_matches)
    return {
        "status": "found" if invoice_matches else "not_found" if invoice_no else "empty",
        "invoice_no": invoice_no,
        "match_count": len(invoice_matches),
        "matches": [invoice_match_preview_row(row) for row in invoice_matches[:12]],
        "summary": summary,
        "market_inference": market_inference_view(inference),
        "suggested_forms": suggested_forms,
        "options": [],
    }


def invoice_search_options(client: dict, query: str, limit: int = 10) -> list[dict]:
    query = str(query or "").strip()
    if len(query) < 2:
        return []
    # Use Data Hub invoice-matches adapter when available — it's an indexed
    # query (~300-500ms typical) instead of pulling the full BCCT pagination
    # for every keystroke. Falls back to local indexed lookup only when the
    # Data Hub adapter is not present (file-store mode).
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is not None and hasattr(data_hub, "invoice_matches"):
        try:
            rows = data_hub.invoice_matches(client["id"], query, [])
        except Exception:  # noqa: BLE001
            rows = []
        if not rows and hasattr(data_hub, "list_bcct"):
            # Try declaration-number lookup (declaration_no exact prefix).
            compact = re.sub(r"[^A-Z0-9]", "", query.upper())
            try:
                bcct_rows = data_hub.list_bcct(client["id"], declaration_no=compact, direction="export")
            except TypeError:
                bcct_rows = []
            except Exception:  # noqa: BLE001
                bcct_rows = []
            rows = [row for row in bcct_rows if row.get("review_status") in ("", "reviewed")]
    else:
        rows = declaration_invoice_matches(client, query, exact=False)
    groups: dict[str, dict] = {}
    for row in rows:
        invoice_ref = row.get("invoice_ref", "")
        option_value = invoice_ref or row.get("declaration_no", "")
        group = groups.setdefault(
            option_value,
            {
                "invoice_no": option_value,
                "row_count": 0,
                "declarations": set(),
                "hs_codes": set(),
                "item_codes": set(),
            },
        )
        group["row_count"] += 1
        if row.get("declaration_no"):
            group["declarations"].add(str(row.get("declaration_no")))
        if row.get("hs_code"):
            group["hs_codes"].add(str(row.get("hs_code")))
        if row.get("item_code"):
            group["item_codes"].add(str(row.get("item_code")))
    options = []
    for group in groups.values():
        options.append({
            "invoice_no": group["invoice_no"],
            "row_count": group["row_count"],
            "declaration_count": len(group["declarations"]),
            "declarations": sorted(group["declarations"])[:4],
            "hs_codes": sorted(group["hs_codes"])[:6],
            "item_codes": sorted(group["item_codes"])[:4],
        })
    return sorted(options, key=lambda row: (-int(row["row_count"]), row["invoice_no"]))[:limit]


def declaration_invoice_matches(client: dict, query: str, exact: bool, include_invoice: bool = True) -> list[dict]:
    query = str(query or "").strip()
    if len(query) < 2:
        return []
    # Fast path for Data Hub mode: indexed invoice/declaration lookup, no
    # full-catalog pagination. Falls back to in-memory workspace for file-store
    # / test mode.
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is not None and hasattr(data_hub, "invoice_matches"):
        try:
            client_config = portfolio_service.get_client_config(client) if hasattr(portfolio_service, "get_client_config") else {}
        except Exception:  # noqa: BLE001
            client_config = {}
        relevant_types_list = list(client_config.get("bcct", {}).get("relevant_export_declaration_types", []))
        rows: list[dict] = []
        if include_invoice:
            try:
                rows = list(data_hub.invoice_matches(client["id"], query, relevant_types_list))
            except Exception:  # noqa: BLE001
                rows = []
        # Declaration lookup — list_bcct accepts declaration_no kwarg on Data Hub.
        if hasattr(data_hub, "list_bcct"):
            compact = re.sub(r"[^A-Z0-9]", "", query.upper())
            try:
                decl_rows = list(data_hub.list_bcct(client["id"], declaration_no=compact, direction="export"))
            except Exception:  # noqa: BLE001
                decl_rows = []
            seen = {(row.get("declaration_no"), row.get("line_no"), row.get("transaction_key")) for row in rows}
            for row in decl_rows:
                key = (row.get("declaration_no"), row.get("line_no"), row.get("transaction_key"))
                if key not in seen:
                    rows.append(row)
                    seen.add(key)
        return rows
    try:
        source_workspace, _source_backend = source_workspace_for_client(client)
    except Exception:
        return []
    client_config = source_workspace.get("client_config", {})
    relevant_types = set(client_config.get("bcct", {}).get("relevant_export_declaration_types", []))
    query_keys = invoice_keys(query)
    query_compact = next(iter(query_keys), re.sub(r"[^A-Z0-9]", "", query.upper()))
    matches = []
    for row in source_workspace.get("bcct", {}).get("published_rows", []):
        if row.get("direction") != "export":
            continue
        if row.get("review_status") not in ("", "reviewed"):
            continue
        if relevant_types and row.get("declaration_type") not in relevant_types:
            continue
        invoice_ref = str(row.get("invoice_ref") or "").strip()
        row_keys = invoice_keys(invoice_ref)
        row_compact = re.sub(r"[^A-Z0-9]", "", invoice_ref.upper())
        declaration_compact = re.sub(r"[^A-Z0-9]", "", str(row.get("declaration_no") or "").upper())
        invoice_matches = include_invoice and query_compact and (query_compact in row_compact or bool(query_keys.intersection(row_keys)))
        declaration_matches = (
            query_compact == declaration_compact
            if exact
            else query_compact and query_compact in declaration_compact
        )
        if not invoice_matches and not declaration_matches:
            continue
        matches.append({**row, "invoice_ref": invoice_ref})
    return matches


def invoice_match_summary(invoice_matches: list[dict]) -> dict:
    declarations = sorted({
        str(row.get("declaration_no") or "")
        for row in invoice_matches
        if row.get("declaration_no")
    })
    hs_codes = sorted({
        str(row.get("hs_code") or "")
        for row in invoice_matches
        if row.get("hs_code")
    })
    item_codes = sorted({
        str(row.get("item_code") or "")
        for row in invoice_matches
        if row.get("item_code")
    })
    invoice_refs = sorted({
        str(row.get("invoice_ref") or "")
        for row in invoice_matches
        if row.get("invoice_ref")
    })
    return {
        "declaration_count": len(declarations),
        "declarations": declarations[:8],
        "hs_codes": hs_codes[:12],
        "item_codes": item_codes[:8],
        "invoice_refs": invoice_refs[:4],
    }


def invoice_match_preview_row(row: dict) -> dict:
    return {
        "declaration_no": row.get("declaration_no", ""),
        "line_no": row.get("line_no", ""),
        "declaration_type": row.get("declaration_type", ""),
        "item_code": row.get("item_code", ""),
        "description": row.get("description", ""),
        "hs_code": row.get("hs_code", ""),
        "quantity": row.get("quantity", ""),
        "unit": row.get("unit", ""),
        "customs_value": row.get("customs_value") or row.get("total_value", ""),
        "value_currency": row.get("value_currency") or row.get("currency", ""),
        "invoice_ref": row.get("invoice_ref", ""),
        "unloading_location": row.get("unloading_location") or row.get("destination_location_name", ""),
        "consignee_name": row.get("consignee_name", ""),
    }


def market_inference_view(inference: dict) -> dict:
    status = inference.get("status", "missing")
    hints = inference.get("hints", [])
    if status == "ready" and hints:
        hint = hints[0]
        source_field = str(hint.get("source_field") or "market_hint")
        source_value = str(hint.get("source_value") or hint.get("country_name") or hint.get("country_code") or "")
        explanation = (
            f"Gợi ý từ {source_field} = {source_value}. "
            "Các dòng invoice chỉ có một quốc gia đích đủ độ tin cậy cao."
        )
        action_label = f"Dùng thị trường {inference.get('destination_market', '')}"
    elif status == "conflict":
        markets = ", ".join(
            str(hint.get("country_name") or hint.get("country_code") or "")
            for hint in hints
            if hint.get("country_name") or hint.get("country_code")
        )
        explanation = f"Không tự chọn vì invoice có nhiều gợi ý thị trường: {markets}."
        action_label = ""
    else:
        explanation = "Chưa có market hint đủ tin cậy từ dữ liệu invoice; cần chọn thị trường thủ công."
        action_label = ""
    return {
        **inference,
        "explanation": explanation,
        "action_label": action_label,
    }


def invoice_form_lane_view(row: dict) -> dict:
    return {
        "form_code": row.get("form_code", ""),
        "display_name": row.get("display_name", ""),
        "agreement": row.get("agreement", ""),
        "instrument": row.get("instrument", ""),
        "reason": row.get("reason", ""),
        "recommended": bool(row.get("recommended")),
        "criteria_preview": row.get("criteria_preview", [])[:4],
    }


def should_show_origin_demo(current_step: str, case: dict, invoice_matches: list[dict]) -> bool:
    return current_step == "origin" and not case.get("products") and not invoice_matches


def attach_origin_demo(case: dict) -> dict:
    demo = attach_results(clone_case(DEMO_CASE))
    case = dict(case)
    case["products"] = demo["products"]
    if not case.get("documents"):
        case["documents"] = demo["documents"]
    case["summary"] = demo["summary"]
    case["mode"] = "Demo tự nạp trong tab Xuất xứ"
    case["mode_note"] = "Dùng khi hồ sơ chưa có đủ invoice/BCCT/BOM để tính thật; không ghi vào hồ sơ lưu."
    return case


def origin_material_count(case: dict) -> int:
    return sum(len(product.get("materials", [])) for product in case.get("products", []))


def case_finished_hs_codes(case: dict) -> list[str]:
    return [
        str(product.get("finished_hs", ""))
        for product in case.get("products", [])
        if str(product.get("finished_hs", "")).strip()
    ]


def co_case_hs_codes(case: dict, invoice_matches: list[dict]) -> list[str]:
    product_hs = case_finished_hs_codes(case)
    if product_hs:
        return product_hs
    return [
        str(row.get("hs_code", ""))
        for row in invoice_matches
        if str(row.get("hs_code", "")).strip()
    ]


def prepare_case_origin_products(
    case: dict,
    invoice_matches: list[dict],
    bom_workspace: dict,
    form_lane: dict,
    material_rows: list[dict],
    stock_rows: list[dict],
    *,
    preserve_existing: bool = False,
) -> dict:
    source_matches = (
        invoice_matches
        if invoice_matches
        else [origin_match_from_existing_product(product) for product in case.get("products", [])]
    )
    if not source_matches:
        return case

    ordered_invoice_matches = order_invoice_matches_for_origin(case, source_matches) if invoice_matches else source_matches
    bom_rows_by_product = selected_bom_rows_by_product(case, bom_workspace)
    build_signature = origin_build_signature(ordered_invoice_matches, bom_rows_by_product, material_rows, stock_rows, form_lane)
    if (
        case.get("products")
        and (
            preserve_existing
            or case.get("origin_snapshot", {}).get("build_signature") == build_signature
        )
    ):
        return case

    material_index = material_catalog_index(material_rows)
    stock_pool = case_allocation_pool(case, ordered_invoice_matches, stock_rows)
    products = []
    for product_sequence, match in enumerate(ordered_invoice_matches, start=1):
        product_code = str(match.get("item_code", "")).strip()
        if not product_code:
            continue
        bom_product_code = bom_product_code_from_material_identity(match) or resolve_bom_product_code(
            product_code,
            bom_workspace,
        )
        product_rows = bom_rows_by_product.get(product_code) or bom_rows_by_product.get(bom_product_code, [])
        products.append(origin_product_from_invoice_match(
            match,
            product_rows,
            form_lane,
            material_index,
            stock_pool,
            product_sequence=product_sequence,
            bom_product_code=bom_product_code,
        ))
    if not products:
        return case

    prepared = dict(case)
    prepared["products"] = products
    prepared["mode"] = "Invoice + BCCT + BOM snapshot"
    prepared["mode_note"] = "Sản phẩm lấy từ BCCT xuất khẩu khớp invoice; NVL lấy từ BOM snapshot hiện hành. Đơn giá NVL ưu tiên từ tồn CO/BCCT nhập, nếu thiếu mới fallback danh mục NVL."
    prepared["origin_snapshot"] = {
        "source": "invoice_bcct_bom",
        "build_signature": build_signature,
        "invoice_no": prepared.get("shipment", {}).get("invoice_no", ""),
        "invoice_match_count": len(ordered_invoice_matches),
        "product_order": [product.get("code", "") for product in products],
        "product_count": len(products),
        "material_count": sum(len(product.get("materials", [])) for product in products),
        "stock_row_count": len(stock_rows),
    }
    return prepared


def prepare_case_origin_product_shells(
    case: dict,
    invoice_matches: list[dict],
    bom_workspace: dict,
    form_lane: dict,
    *,
    preserve_existing: bool = True,
) -> dict:
    """Create per-product origin sheets without calculating BOM/material rows.

    The origin page should show the workbook and let staff explicitly load each
    sheet. Allocation and VNM calculation only happen in prepare_case_origin_sheet.
    """
    source_matches = (
        invoice_matches
        if invoice_matches
        else [origin_match_from_existing_product(product) for product in case.get("products", [])]
    )
    if not source_matches:
        return case

    ordered_invoice_matches = order_invoice_matches_for_origin(case, source_matches) if invoice_matches else source_matches
    bom_rows_by_product = selected_bom_rows_by_product(case, bom_workspace)
    existing_by_code = {
        str(product.get("code") or product.get("product_code") or "").strip(): dict(product)
        for product in case.get("products", [])
        if str(product.get("code") or product.get("product_code") or "").strip()
    }
    products = []
    for product_sequence, match in enumerate(ordered_invoice_matches, start=1):
        product_code = str(match.get("item_code") or match.get("product_code") or "").strip()
        if not product_code:
            continue
        existing = existing_by_code.get(product_code)
        if preserve_existing and existing and existing.get("materials"):
            product = dict(existing)
            product["allocation_sequence"] = str(product_sequence)
            products.append(product)
            continue
        bom_product_code = bom_product_code_from_material_identity(match) or resolve_bom_product_code(
            product_code,
            bom_workspace,
        )
        product_rows = bom_rows_by_product.get(product_code) or bom_rows_by_product.get(bom_product_code, [])
        shell = origin_product_shell_from_invoice_match(
            match,
            product_rows,
            form_lane,
            product_sequence=product_sequence,
            bom_product_code=bom_product_code,
        )
        if preserve_existing and existing:
            shell = {
                **shell,
                "materials": existing.get("materials", []),
                "origin_sheet_material_overrides": existing.get("origin_sheet_material_overrides", {}),
            }
        products.append(shell)
    if not products:
        return case

    prepared = dict(case)
    prepared["products"] = products
    prepared["mode"] = "Invoice + BCCT + BOM snapshot"
    prepared["mode_note"] = "Sản phẩm lấy từ BCCT xuất khẩu khớp invoice; bấm Load BOM trên từng sheet để tính NVL và tồn CO."
    prepared["origin_snapshot"] = {
        **dict(prepared.get("origin_snapshot") or {}),
        "source": "invoice_bcct_bom",
        "invoice_no": prepared.get("shipment", {}).get("invoice_no", ""),
        "invoice_match_count": len(ordered_invoice_matches),
        "product_order": [product.get("code", "") for product in products],
        "product_count": len(products),
        "material_count": sum(len(product.get("materials", [])) for product in products),
    }
    return prepared


def prepare_case_origin_sheet(
    case: dict,
    product_code: str,
    invoice_matches: list[dict],
    bom_workspace: dict,
    form_lane: dict,
    material_rows: list[dict],
    stock_rows: list[dict],
    *,
    min_gap_days: int | None = None,
) -> dict:
    target_code = str(product_code or "").strip()
    if not target_code:
        return case
    source_matches = (
        invoice_matches
        if invoice_matches
        else [origin_match_from_existing_product(product) for product in case.get("products", [])]
    )
    ordered_invoice_matches = order_invoice_matches_for_origin(case, source_matches) if invoice_matches else source_matches
    target_match = None
    target_sequence = 0
    products_by_code = {str(product.get("code") or product.get("product_code") or "").strip(): product for product in case.get("products", [])}
    stock_pool = case_allocation_pool(case, ordered_invoice_matches, stock_rows, min_gap_days=min_gap_days)
    for sequence, match in enumerate(ordered_invoice_matches, start=1):
        match_code = str(match.get("item_code") or match.get("product_code") or "").strip()
        if match_code == target_code:
            target_match = match
            target_sequence = sequence
            break
        existing_product = products_by_code.get(match_code)
        if existing_product:
            apply_existing_origin_product_consumption(existing_product, stock_pool)
    if not target_match:
        return case

    bom_rows_by_product = selected_bom_rows_by_product(case, bom_workspace)
    material_index = material_catalog_index(material_rows)
    bom_product_code = bom_product_code_from_material_identity(target_match) or resolve_bom_product_code(
        target_code,
        bom_workspace,
    )
    product_rows = bom_rows_by_product.get(target_code) or bom_rows_by_product.get(bom_product_code, [])
    recalculated_product = origin_product_from_invoice_match(
        target_match,
        product_rows,
        form_lane,
        material_index,
        stock_pool,
        product_sequence=target_sequence,
        bom_product_code=bom_product_code,
    )
    products = []
    changed = False
    for product in case.get("products", []):
        code = str(product.get("code") or product.get("product_code") or "").strip()
        if code == target_code and not changed:
            products.append(recalculated_product)
            changed = True
        else:
            products.append(product)
    if not changed:
        products.append(recalculated_product)
    prepared = dict(case)
    prepared["products"] = products
    return prepared


def origin_product_shell_from_invoice_match(
    match: dict,
    bom_rows: list[dict],
    form_lane: dict,
    *,
    product_sequence: int | None = None,
    bom_product_code: str = "",
) -> dict:
    product_code = str(match.get("item_code") or match.get("product_code") or "").strip()
    bom_product_code = str(bom_product_code or product_code).strip()
    finished_hs = str(match.get("hs_code", "")).strip()
    preview = criteria_preview_for_hs(form_lane.get("form_code", ""), finished_hs) if form_lane else {}
    criterion = preview.get("criteria") or "Cần tra cứu PSR theo HS"
    threshold = lvc_threshold_from_criterion(criterion)
    quantity = decimal_value(match.get("quantity", "0"))
    product_value = origin_product_value(match)
    fob = product_value["value"]
    first_row = bom_rows[0] if bom_rows else {}
    product = {
        "code": product_code,
        "bom_product_code": bom_product_code,
        "allocation_sequence": str(product_sequence or ""),
        "name": match.get("description") or product_code,
        "finished_hs": finished_hs,
        "quantity": decimal_text(quantity),
        "unit": match.get("unit", ""),
        "currency": product_value["currency"],
        "fob_currency": product_value["currency"] or match.get("currency", ""),
        "declared_currency": match.get("currency", ""),
        "value_source": product_value["source"],
        "source_declaration_no": match.get("declaration_no", ""),
        "source_declaration_date": match.get("declaration_date") or match.get("registration_date", ""),
        "source_line_no": match.get("line_no", ""),
        "invoice_ref": match.get("invoice_ref", ""),
        "fob": decimal_text(fob) if fob is not None else "",
        "non_origin_value": "",
        "rvc_threshold": decimal_text(threshold) if threshold is not None else "",
        "documented_result": criterion,
        "lvc_percentage": "",
        "lvc_status": "review",
        "lvc_status_label": "Chưa tính",
        "lvc_threshold": decimal_text(threshold) if threshold is not None else "",
        "vnm_value": "",
        "bom_product_artifact_id": first_row.get("product_artifact_id") or first_row.get("product_version_id", ""),
        "bom_product_artifact_no": first_row.get("product_artifact_no") or first_row.get("product_version_no", ""),
        "bom_product_version_id": first_row.get("product_artifact_id") or first_row.get("product_version_id", ""),
        "bom_product_version_no": first_row.get("product_artifact_no") or first_row.get("product_version_no", ""),
        "materials": [],
    }
    return enrich_origin_product(product)


def recalculate_origin_sheet_edits(client: dict, case: dict, product_code: str, *, min_gap_days: int | None = None) -> dict:
    """Recompute one sheet from its saved sheet edits, without changing BOM selection."""
    prepared = attach_origin_sheet_states(case)
    products = prepared.get("products", [])
    target_index = next(
        (index for index, product in enumerate(products) if str(product.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        return prepared
    target = products[target_index]
    overrides = target.get("origin_sheet_material_overrides") or {}
    if not overrides:
        return prepared

    sheet_rows = sheet_edit_bom_rows(target, overrides)
    material_codes = sorted({
        str(row.get("material_code") or "").strip()
        for row in sheet_rows
        if str(row.get("material_code") or "").strip()
    })
    source_context = {"material_rows": [], "stock_rows": []}
    stock_rows: list[dict] = []
    try:
        narrow_rows = portfolio_service.list_bcct_by_codes(client.get("id", ""), material_codes, direction="import")
    except Exception:  # noqa: BLE001
        narrow_rows = []
    if narrow_rows:
        try:
            client_config = portfolio_service.get_client_config(client) if hasattr(portfolio_service, "get_client_config") else {}
            stock_rows = co_stock_rows_from_bcct(narrow_rows, client_config)
        except Exception:  # noqa: BLE001
            stock_rows = []
    if not stock_rows and not co_auth.data_hub_source_mode_enabled():
        try:
            source_context = co_case_source_context_cached(client, prepared)
            stock_rows = source_context.get("stock_rows") or []
        except Exception:  # noqa: BLE001
            source_context = {"material_rows": [], "stock_rows": []}
            stock_rows = []
    material_index = material_catalog_index(source_context.get("material_rows") or [])
    cached_matches = case.get("source_invoice_matches") if isinstance(case.get("source_invoice_matches"), list) else []
    stock_pool = case_allocation_pool(prepared, cached_matches, stock_rows, min_gap_days=min_gap_days)
    for previous in products[:target_index]:
        apply_existing_origin_product_consumption(previous, stock_pool)

    form_lane = recommended_form_lane(
        prioritized_form_lanes(prepared.get("destination_market", ""), [str(target.get("finished_hs") or "")])
    )
    recalculated = origin_product_from_invoice_match(
        origin_match_from_existing_product(target),
        sheet_rows,
        form_lane,
        material_index,
        stock_pool,
        product_sequence=target_index + 1,
        bom_product_code=str(target.get("bom_product_code") or target.get("code") or ""),
    )
    for key in [
        "bom_product_artifact_id",
        "bom_product_artifact_no",
        "bom_product_version_id",
        "bom_product_version_no",
    ]:
        if target.get(key) and not recalculated.get(key):
            recalculated[key] = target.get(key)

    updated_products = [dict(product) for product in products]
    updated_products[target_index] = recalculated
    prepared["products"] = updated_products
    return attach_origin_sheet_states(prepared)


def sheet_edit_bom_rows(product: dict, overrides: dict) -> list[dict]:
    rows: list[dict] = []
    materials = product.get("materials") or []
    for index, material in enumerate(materials):
        override = overrides.get(str(index)) if isinstance(overrides.get(str(index)), dict) else {}
        if override.get("deleted"):
            continue
        replacement_code = str(override.get("material_code") or "").strip()
        original_code = str(material.get("material_code") or material.get("internal_material_code") or "").strip()
        material_code = replacement_code or original_code
        if not material_code:
            continue
        row = {
            "product_code": product.get("bom_product_code") or product.get("code") or "",
            "material_code": material_code,
            "qty_per": str(override.get("norm_per_unit") or material.get("bom_qty_per") or "0"),
            "uom": str(override.get("uom") or material.get("uom") or ""),
            "material_name": str(override.get("name") or ("" if replacement_code else material.get("material_description")) or ""),
            "hs_code": str(override.get("hs_code") or ("" if replacement_code else material.get("hs_code")) or ""),
            "source": material.get("bom_source") or material.get("source_document_ref") or "sheet_edit",
            "row_class": material.get("bom_row_class") or "",
        }
        if not replacement_code:
            row["unit_value"] = material.get("unit_value", "")
        rows.append(row)
    added_items = [
        (key, value)
        for key, value in overrides.items()
        if str(key).startswith("added_") and isinstance(value, dict) and value.get("material_code")
    ]
    added_items.sort(key=lambda item: numeric_sort_text(str(item[0]).split("_", 1)[1] if "_" in str(item[0]) else "0"))
    for _key, value in added_items:
        rows.append({
            "product_code": product.get("bom_product_code") or product.get("code") or "",
            "material_code": str(value.get("material_code") or "").strip(),
            "qty_per": str(value.get("norm_per_unit") or "0"),
            "uom": str(value.get("uom") or ""),
            "material_name": str(value.get("name") or ""),
            "hs_code": str(value.get("hs_code") or ""),
            "source": "sheet_edit_added",
            "row_class": "added",
        })
    return rows


def apply_existing_origin_product_consumption(product: dict, stock_pool: dict[str, list[dict]]) -> None:
    for material in product.get("materials", []) or []:
        material_code = str(material.get("material_code") or material.get("internal_material_code") or "").strip()
        if not material_code:
            continue
        for line in material.get("allocation_lines", []) or []:
            allocated_qty = decimal_value(line.get("allocated_qty"))
            if allocated_qty <= 0:
                continue
            stock = stock_for_existing_allocation_line(stock_pool, material_code, line)
            if not stock:
                continue
            remaining_qty = stock_allocation_remaining_qty(stock)
            stock["_allocation_remaining_qty"] = remaining_qty - allocated_qty
            allocation_context = {
                "product_sequence": line.get("product_sequence") or product.get("allocation_sequence", ""),
                "product_code": line.get("product_code") or product.get("code", ""),
                "product_name": product.get("name", ""),
                "material_sequence": line.get("material_sequence") or material.get("material_sequence", ""),
                "material_code": material_code,
                "material_uom": material.get("uom", ""),
            }
            stock.setdefault("_allocation_consumptions", []).append(stock_allocation_consumption(line, allocation_context))


def stock_for_existing_allocation_line(stock_pool: dict[str, list[dict]], material_code: str, line: dict) -> dict:
    candidates = stock_candidates_for_material(stock_pool, material_code)
    for stock in candidates:
        if allocation_line_matches_stock(line, stock):
            return stock
    return {}


def allocation_line_matches_stock(line: dict, stock: dict) -> bool:
    checks = [
        ("source_row", "source_row"),
        ("import_declaration_no", "import_declaration_no"),
        ("import_line_no", "line_no"),
        ("allocation_code", "allocation_code"),
    ]
    matched = False
    for line_key, stock_key in checks:
        line_value = str(line.get(line_key) or "").strip()
        stock_value = str(stock.get(stock_key) or "").strip()
        if line_value and stock_value:
            if line_value != stock_value:
                return False
            matched = True
    return matched


def order_invoice_matches_for_origin(case: dict, invoice_matches: list[dict]) -> list[dict]:
    order = origin_product_order(case)
    if not order:
        return list(invoice_matches)
    rank = {code: index for index, code in enumerate(order)}

    def sort_key(item: tuple[int, dict]) -> tuple[int, int]:
        index, row = item
        code = str(row.get("item_code") or row.get("product_code") or row.get("customs_code") or "").strip()
        return rank.get(code, len(rank) + index), index

    return [row for _, row in sorted(enumerate(invoice_matches), key=sort_key)]


def origin_product_order(case: dict) -> list[str]:
    raw_order = case.get("origin_product_order") or case.get("origin_snapshot", {}).get("product_order", [])
    if isinstance(raw_order, str):
        candidates = re.split(r"[|,\n]", raw_order)
    elif isinstance(raw_order, (list, tuple)):
        candidates = raw_order
    else:
        candidates = []
    output = []
    for candidate in candidates:
        code = str(candidate or "").strip()
        if code and code not in output:
            output.append(code)
    return output


def _hydrate_product_export_declaration_dates(case: dict, client: dict | None = None) -> None:
    """Backfill missing `product.source_declaration_date` from cached matches,
    falling back to Data Hub `list_declarations` for cases saved before
    `enrich_invoice_matches_with_bcct` started forwarding `declaration_date`.

    Read-only patch — operator overrides on `case["products"]` are untouched.
    Cached `case["source_invoice_matches"]` is updated in place so the next
    `update_case_record` call (typically right after export prep) persists the
    backfill, making future renders free.
    """
    products = [p for p in (case.get("products") or []) if not p.get("source_declaration_date")]
    if not products:
        return

    matches = case.get("source_invoice_matches") if isinstance(case.get("source_invoice_matches"), list) else []
    by_decl: dict[str, str] = {}
    for row in matches:
        decl = str(row.get("declaration_no") or "").strip()
        if not decl or decl in by_decl:
            continue
        value = str(row.get("declaration_date") or row.get("registration_date") or "").strip()
        if value:
            by_decl[decl] = value

    missing: list[str] = []
    for product in products:
        decl = str(product.get("source_declaration_no") or "").strip()
        if decl and decl not in by_decl:
            missing.append(decl)

    if missing and client and client.get("id"):
        dates = _fetch_export_declaration_dates(client["id"], sorted(set(missing)))
        by_decl.update({k: v for k, v in dates.items() if v})
        if dates and matches:
            for row in matches:
                decl = str(row.get("declaration_no") or "").strip()
                if decl and not row.get("declaration_date") and dates.get(decl):
                    row["declaration_date"] = dates[decl]

    if not by_decl:
        return
    for product in products:
        decl = str(product.get("source_declaration_no") or "").strip()
        if decl and decl in by_decl:
            product["source_declaration_date"] = by_decl[decl]


def _fetch_export_declaration_dates(client_id: str, declaration_nos: list[str]) -> dict[str, str]:
    """Resolve `earliest_bcct_date` per export declaration_no from Data Hub.

    Returns mapping {declaration_no: "YYYY-MM-DD"}. Empty dict on any error or
    when the active portfolio service doesn't wrap a Data Hub client.
    """
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is None or not declaration_nos:
        return {}
    try:
        rows = data_hub.list_declarations(
            client_id,
            direction="export",
            declaration_nos=declaration_nos,
        )
    except Exception:  # noqa: BLE001 — best-effort backfill; export must not block
        return {}
    out: dict[str, str] = {}
    for row in rows or []:
        decl = str(row.get("declaration_no") or "").strip()
        date = str(row.get("earliest_bcct_date") or "").strip()
        if decl and date:
            out[decl] = _to_vietnamese_date(date)
    return out


def _to_vietnamese_date(value: str) -> str:
    """Convert "YYYY-MM-DD" (Data Hub ISO) to "DD/MM/YYYY" (bảng kê format)."""
    text = value.strip()
    if not text:
        return ""
    try:
        from datetime import date
        d = date.fromisoformat(text[:10])
        return d.strftime("%d/%m/%Y")
    except ValueError:
        return text


def _hydrate_material_dates_from_stock(case: dict, client: dict) -> None:
    """Backfill missing `import_declaration_date` on materials + allocation lines.

    Materials/allocations saved before `co_stock_rows_from_bcct` started
    copying `registration_date` out of the BCCT payload have empty date
    fields, which leaves col "Ngày" blank on the exported xlsx. Re-running
    Calculate would refresh the snapshot but also wipes operator overrides
    (delete/substitute/norm/added rows), so we look up dates from the
    materialized stock table here instead — read-only, no override loss.
    """
    client_id = str(client.get("id") or "").strip()
    if not client_id:
        return
    rows_to_lookup: set[str] = set()
    for product in case.get("products") or []:
        for material in product.get("materials") or []:
            if not (material.get("import_declaration_date")
                    or material.get("declaration_date")
                    or material.get("registration_date")):
                source_row = str(material.get("source_row") or "").strip()
                if source_row:
                    rows_to_lookup.add(source_row)
            for allocation in material.get("allocation_lines") or []:
                if not (allocation.get("import_declaration_date")
                        or allocation.get("declaration_date")
                        or allocation.get("registration_date")):
                    source_row = str(allocation.get("source_row") or "").strip()
                    if source_row:
                        rows_to_lookup.add(source_row)
    if not rows_to_lookup:
        return
    dates = co_stock_materializer.registration_dates_for_source_rows(client_id, list(rows_to_lookup))
    if not dates:
        return
    for product in case.get("products") or []:
        for material in product.get("materials") or []:
            if not material.get("import_declaration_date"):
                joined = []
                source_row = str(material.get("source_row") or "").strip()
                for piece in [p.strip() for p in source_row.split(",") if p.strip()]:
                    value = dates.get(piece)
                    if value and value not in joined:
                        joined.append(value)
                if joined:
                    material["import_declaration_date"] = ", ".join(joined)
            for allocation in material.get("allocation_lines") or []:
                if not allocation.get("import_declaration_date"):
                    value = dates.get(str(allocation.get("source_row") or "").strip())
                    if value:
                        allocation["import_declaration_date"] = value


def _attach_fob_vnd(product: dict) -> None:
    """Compute product.fob_vnd from fob × FX rate at the product's declaration date.

    fob_currency falls back to product.currency (set when the product was built
    from a BCCT match). For VND-native cases, fob_vnd = fob. For non-VND with
    no FX hit, fob_vnd stays empty + fob_fx_source = "missing" so the renderer
    can show a warning instead of a wrong number.
    """
    fob_raw = str(product.get("fob") or "").strip()
    if not fob_raw:
        product["fob_vnd"] = ""
        product["fob_fx_source"] = "missing"
        return
    fob_currency = str(product.get("fob_currency") or product.get("currency") or "").strip().upper()
    try:
        fob_dec = Decimal(fob_raw)
    except (InvalidOperation, ValueError):
        product["fob_vnd"] = ""
        product["fob_fx_source"] = "missing"
        return
    if not fob_currency or fob_currency == "VND":
        product["fob_vnd"] = decimal_text(fob_dec)
        product["fob_fx_source"] = "vnd_native"
        return
    target_date = str(
        product.get("source_declaration_date")
        or product.get("export_declaration_date")
        or product.get("invoice_date")
        or ""
    ).strip()
    try:
        from app.customs_fx_store import CUSTOMS_FX_CLIENT_ID, get_customs_fx_store, lookup_exchange_rate
        rows = get_customs_fx_store().rows(CUSTOMS_FX_CLIENT_ID)
        hit = lookup_exchange_rate(rows, fob_currency, target_date) if target_date else None
    except Exception:  # noqa: BLE001 — fob_vnd is optional; never block render
        hit = None
    if hit and hit.get("rate_vnd_per_unit"):
        try:
            rate = Decimal(str(hit["rate_vnd_per_unit"]))
        except (InvalidOperation, ValueError):
            rate = None
        if rate and rate > 0:
            product["fob_vnd"] = decimal_text(fob_dec * rate)
            product["fob_fx_rate"] = decimal_text(rate)
            product["fob_fx_source"] = "customs_lookup"
            return
    product["fob_vnd"] = ""
    product["fob_fx_source"] = "missing"


def attach_origin_sheet_states(case: dict) -> dict:
    products = case.get("products", [])
    existing = case.get("origin_sheet_states") if isinstance(case.get("origin_sheet_states"), dict) else {}
    normalized = {}
    market = case.get("destination_market", "")
    for product in products:
        code = str(product.get("code") or "").strip()
        if not code:
            continue
        raw_state = existing.get(code) if isinstance(existing.get(code), dict) else {}
        # Default = "draft" (Chưa tính). A sheet only becomes "calculated"
        # after staff explicitly clicks "Tính bảng kê" (which sets it via
        # set_origin_sheet_status). Never auto-mark calculated even when the
        # underlying snapshot has data — staff has to confirm intent.
        default_status = "draft"
        status = str(raw_state.get("status") or default_status).strip()
        if status not in ORIGIN_SHEET_STATUS_LABELS:
            status = default_status
        recommendation = sheet_form_recommendation(market, product.get("finished_hs", ""))
        form_override = str(raw_state.get("form_override") or "").strip()
        criteria_override = str(raw_state.get("criteria_override") or "").strip()
        lvc_threshold_override = normalize_threshold(raw_state.get("lvc_threshold_override"))
        rvc_threshold_override = normalize_threshold(raw_state.get("rvc_threshold_override"))
        currency_mode = str(raw_state.get("currency_mode") or "").strip().lower()
        if currency_mode not in SHEET_CURRENCY_MODES:
            currency_mode = "native"
        optimization_mode = str(raw_state.get("optimization_mode") or "").strip().lower()
        if optimization_mode not in SHEET_OPTIMIZATION_MODES:
            optimization_mode = "max_lvc"
        effective_form = form_override or recommendation.get("form_code", "")
        effective_criteria = criteria_override or recommendation.get("criteria_text", "")
        effective_lvc_threshold = lvc_threshold_override or str(product.get("lvc_threshold") or "").strip()
        effective_rvc_threshold = rvc_threshold_override or str(product.get("rvc_threshold") or "").strip()
        state = {
            "status": status,
            "status_label": ORIGIN_SHEET_STATUS_LABELS[status],
            "form_override": form_override,
            "criteria_override": criteria_override,
            "lvc_threshold_override": lvc_threshold_override,
            "rvc_threshold_override": rvc_threshold_override,
            "currency_mode": currency_mode,
            "optimization_mode": optimization_mode,
            "recommended_form_code": recommendation.get("form_code", ""),
            "recommended_form_label": recommendation.get("form_label", ""),
            "recommended_criteria_text": recommendation.get("criteria_text", ""),
            "recommendation_source": recommendation.get("source", ""),
            "effective_form_code": effective_form,
            "effective_criteria_text": effective_criteria,
            "effective_lvc_threshold": effective_lvc_threshold,
            "effective_rvc_threshold": effective_rvc_threshold,
        }
        normalized[code] = state
        product["origin_sheet_state"] = state
        product["origin_sheet_status"] = state["status"]
        product["origin_sheet_status_label"] = state["status_label"]
        product["origin_sheet_form_override"] = form_override
        product["origin_sheet_criteria_override"] = criteria_override
        product["origin_sheet_lvc_threshold_override"] = lvc_threshold_override
        product["origin_sheet_rvc_threshold_override"] = rvc_threshold_override
        product["origin_sheet_currency_mode"] = currency_mode
        product["origin_sheet_optimization_mode"] = optimization_mode
        _attach_fob_vnd(product)
        product["origin_sheet_recommended_form_code"] = state["recommended_form_code"]
        product["origin_sheet_recommended_form_label"] = state["recommended_form_label"]
        product["origin_sheet_recommended_criteria_text"] = state["recommended_criteria_text"]
        product["origin_sheet_effective_form_code"] = effective_form
        product["origin_sheet_effective_criteria_text"] = effective_criteria
        product["origin_sheet_effective_lvc_threshold"] = effective_lvc_threshold
        product["origin_sheet_effective_rvc_threshold"] = effective_rvc_threshold
        material_overrides = raw_state.get("material_overrides") if isinstance(raw_state.get("material_overrides"), dict) else {}
        # Carry overrides on the sheet state so they round-trip through save/calculate.
        state["material_overrides"] = {str(k): dict(v) for k, v in material_overrides.items() if isinstance(v, dict)}
        diff_added = sum(1 for v in state["material_overrides"].values() if v.get("added"))
        diff_removed = sum(1 for v in state["material_overrides"].values() if v.get("deleted"))
        diff_replaced = sum(
            1 for v in state["material_overrides"].values()
            if not v.get("added") and not v.get("deleted") and v.get("material_code") and not v.get("norm_edit_only")
        )
        diff_norm_only = sum(
            1 for v in state["material_overrides"].values()
            if v.get("norm_edit_only") and not v.get("added") and not v.get("deleted")
        )
        state["material_diff_added"] = diff_added
        state["material_diff_removed"] = diff_removed
        state["material_diff_replaced"] = diff_replaced
        state["material_diff_norm_only"] = diff_norm_only
        state["material_diff_total"] = diff_added + diff_removed + diff_replaced + diff_norm_only
        product["origin_sheet_material_overrides"] = state["material_overrides"]
        product["origin_sheet_has_material_overrides"] = state["material_diff_total"] > 0
        product["origin_sheet_material_diff_added"] = diff_added
        product["origin_sheet_material_diff_removed"] = diff_removed
        product["origin_sheet_material_diff_replaced"] = diff_replaced
        product["origin_sheet_material_diff_norm_only"] = diff_norm_only
        product["origin_sheet_material_diff_total"] = state["material_diff_total"]
        proposed_artifact_id = str(raw_state.get("proposed_artifact_id") or "").strip()
        proposed_proposal_id = str(raw_state.get("proposed_proposal_id") or "").strip()
        proposed_status = str(raw_state.get("proposed_status") or "").strip()
        state["proposed_artifact_id"] = proposed_artifact_id
        state["proposed_proposal_id"] = proposed_proposal_id
        state["proposed_status"] = proposed_status
        product["origin_sheet_proposed_artifact_id"] = proposed_artifact_id
        product["origin_sheet_proposed_proposal_id"] = proposed_proposal_id
        product["origin_sheet_proposed_status"] = proposed_status
    for index, product in enumerate(products):
        code = str(product.get("code") or "").strip()
        status = product.get("origin_sheet_status")
        previous_unlocked = [
            str(previous.get("code") or "")
            for previous in products[:index]
            if previous.get("origin_sheet_status") != "locked"
        ]
        later_locked = [
            str(later.get("code") or "")
            for later in products[index + 1:]
            if later.get("origin_sheet_status") == "locked"
        ]
        sequence_reason = ""
        if previous_unlocked:
            sequence_reason = f"Cần chốt các bước trước: {', '.join(previous_unlocked[:5])}."
        elif later_locked:
            sequence_reason = f"Cần mở chốt các bước sau trước: {', '.join(later_locked[:5])}."
        product["origin_can_calculate"] = bool(code and status != "locked" and not sequence_reason)
        product["origin_calculate_block_reason"] = (
            f"Bảng kê {code} đã chốt; cần mở chốt trước khi tính lại."
            if status == "locked"
            else sequence_reason
        )
        product["origin_can_lock"] = bool(code and status == "calculated" and not sequence_reason)
        product["origin_lock_block_reason"] = (
            "" if product["origin_can_lock"] else sequence_reason or f"Chỉ chốt được bảng kê {code} sau khi đã tính."
        )
        product["origin_can_reopen"] = bool(code and status == "locked" and not later_locked)
        product["origin_reopen_block_reason"] = (
            "" if product["origin_can_reopen"] else f"Chỉ được mở chốt từ bước cuối cùng; cần mở chốt {', '.join(later_locked[:5])} trước." if later_locked else ""
        )
    prepared = dict(case)
    prepared["origin_sheet_states"] = normalized
    return prepared


def set_origin_sheet_status(case: dict, product_code: str, status: str) -> dict:
    if status not in ORIGIN_SHEET_STATUS_LABELS:
        status = "draft"
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    states[product_code] = {
        **previous,
        "status": status,
        "status_label": ORIGIN_SHEET_STATUS_LABELS[status],
    }
    prepared = dict(case)
    prepared["origin_sheet_states"] = states
    return attach_origin_sheet_states(prepared)


def reject_if_sheet_locked(case: dict, product_code: str) -> None:
    """Refuse material/norm mutations on a sheet whose status is `locked`.

    The UI hides the edit buttons when locked (`co_case.html` + JS gate from
    commit `0ca012a`), but those guards can be bypassed by direct POST. Without
    this server check, mutating a locked sheet would leave the ledger holding
    `co_stock_claims` for the old materials while the persisted sheet now lists
    the new ones — a quiet Tồn CO leak. Operator must Mở chốt the sheet first.
    """
    state = (case.get("origin_sheet_states") or {}).get(product_code, {}) or {}
    if state.get("status") == "locked":
        raise HTTPException(
            status_code=409,
            detail=f"Sheet {product_code} đã chốt; mở chốt trước khi sửa NVL.",
        )


def mark_origin_sheets_stale(case: dict, from_index: int) -> dict:
    prepared = attach_origin_sheet_states(case)
    states = dict(prepared.get("origin_sheet_states") or {})
    for index, product in enumerate(prepared.get("products", [])):
        code = str(product.get("code") or "").strip()
        current_status = str(product.get("origin_sheet_status") or "").strip()
        if code and index >= max(from_index, 0) and current_status != "draft":
            previous = states.get(code) if isinstance(states.get(code), dict) else {}
            states[code] = {
                **previous,
                "status": "stale",
                "status_label": ORIGIN_SHEET_STATUS_LABELS["stale"],
            }
    prepared["origin_sheet_states"] = states
    return attach_origin_sheet_states(prepared)


SHEET_CURRENCY_MODES = {"native", "vnd"}
SHEET_OPTIMIZATION_MODES = {"max_lvc", "min_lvc"}


def set_origin_sheet_config_override(
    case: dict, product_code: str, overrides: dict
) -> dict:
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    sanitized = {**previous}
    if "form_override" in overrides:
        sanitized["form_override"] = str(overrides.get("form_override") or "").strip()
    if "criteria_override" in overrides:
        sanitized["criteria_override"] = str(overrides.get("criteria_override") or "").strip()
    if "lvc_threshold_override" in overrides:
        sanitized["lvc_threshold_override"] = normalize_threshold(overrides.get("lvc_threshold_override"))
    if "rvc_threshold_override" in overrides:
        sanitized["rvc_threshold_override"] = normalize_threshold(overrides.get("rvc_threshold_override"))
    if "currency_mode" in overrides:
        mode = str(overrides.get("currency_mode") or "").strip().lower()
        sanitized["currency_mode"] = mode if mode in SHEET_CURRENCY_MODES else "native"
    if "optimization_mode" in overrides:
        mode = str(overrides.get("optimization_mode") or "").strip().lower()
        sanitized["optimization_mode"] = mode if mode in SHEET_OPTIMIZATION_MODES else "max_lvc"
    states[product_code] = sanitized
    prepared = dict(case)
    prepared["origin_sheet_states"] = states
    return attach_origin_sheet_states(prepared)


def normalize_threshold(value) -> str:
    text = str(value or "").strip().rstrip("%").strip()
    if not text:
        return ""
    try:
        decimal_value = Decimal(text)
    except (InvalidOperation, ValueError):
        return ""
    if decimal_value < 0 or decimal_value > 100:
        return ""
    return str(decimal_value.quantize(Decimal("0.01")).normalize())


def sheet_form_recommendation(market: str, finished_hs: str) -> dict:
    market = str(market or "").strip()
    finished_hs = str(finished_hs or "").strip()
    if not market or market.lower() == "chưa nhập":
        return {"form_code": "", "form_label": "", "criteria_text": "", "source": "missing_market"}
    lanes = prioritized_form_lanes(market, [finished_hs] if finished_hs else [])
    selected = recommended_form_lane(lanes)
    if not selected:
        return {"form_code": "", "form_label": "", "criteria_text": "", "source": "no_lane"}
    criteria_rows = selected.get("criteria_preview") or []
    criteria_text = ""
    if criteria_rows:
        first = criteria_rows[0]
        criteria_text = str(first.get("criteria") or "").strip()
    return {
        "form_code": str(selected.get("form_code") or "").strip(),
        "form_label": str(selected.get("display_name") or "").strip(),
        "criteria_text": criteria_text,
        "source": "engine",
    }


def origin_sheet_export_blockers(case: dict) -> list[str]:
    blockers = []
    for product in attach_origin_sheet_states(case).get("products", []):
        status = product.get("origin_sheet_status")
        if status in {"draft", "stale", "calculating"}:
            blockers.append(str(product.get("code") or "sheet"))
    return blockers


def origin_sheet_action_error(case: dict, product_code: str, action: str) -> str:
    prepared = attach_origin_sheet_states(case)
    products = prepared.get("products", [])
    target_index = next(
        (index for index, product in enumerate(products) if str(product.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        return f"Không tìm thấy bảng kê {product_code}."
    target = products[target_index]
    status = target.get("origin_sheet_status")
    previous_unlocked = [
        str(product.get("code") or "")
        for product in products[:target_index]
        if product.get("origin_sheet_status") != "locked"
    ]
    later_locked = [
        str(product.get("code") or "")
        for product in products[target_index + 1:]
        if product.get("origin_sheet_status") == "locked"
    ]
    if action in {"calculate", "lock"} and previous_unlocked:
        return f"Cần chốt các bước trước trước khi xử lý {product_code}: {', '.join(previous_unlocked[:5])}."
    if action in {"calculate", "lock"} and later_locked:
        return f"Cần mở chốt các bước sau trước khi xử lý lại {product_code}: {', '.join(later_locked[:5])}."
    if action == "calculate" and status == "locked":
        return f"Bảng kê {product_code} đã chốt; cần mở chốt trước khi tính lại."
    if action == "lock" and status != "calculated":
        return f"Chỉ chốt được bảng kê {product_code} sau khi đã tính."
    if action == "reopen":
        if status != "locked":
            return f"Bảng kê {product_code} chưa chốt."
        if later_locked:
            return f"Chỉ được mở chốt từ bước cuối cùng; cần mở chốt {', '.join(later_locked[:5])} trước."
    return ""


def origin_build_signature(
    invoice_matches: list[dict],
    bom_rows_by_product: dict[str, list[dict]],
    material_rows: list[dict],
    stock_rows: list[dict],
    form_lane: dict,
) -> str:
    payload = {
        "form": {
            "form_code": form_lane.get("form_code", ""),
            "display_name": form_lane.get("display_name", ""),
        },
        "invoice_matches": [
            compact_origin_signature_row(
                row,
                [
                    "transaction_key",
                    "declaration_no",
                    "line_no",
                    "item_code",
                    "hs_code",
                    "quantity",
                    "unit",
                    "customs_value",
                    "foreign_currency_value",
                    "total_value",
                    "currency",
                    "value_currency",
                    "invoice_ref",
                ],
            )
            for row in invoice_matches
        ],
        "bom_rows": [
            compact_origin_signature_row(
                row,
                [
                    "product_code",
                    "product_version_id",
                    "product_version_no",
                    "material_code",
                    "qty_per",
                    "uom",
                    "hs_code",
                    "unit_value",
                    "unit_price",
                ],
            )
            for product_code in sorted(bom_rows_by_product)
            for row in bom_rows_by_product[product_code]
        ],
        "materials": [
            compact_origin_signature_row(
                row,
                ["material_code", "customs_code", "internal_code", "origin_default", "origin_status", "unit_price", "taxable_unit_price"],
            )
            for row in material_rows
        ],
        "stock_rows": [
            compact_origin_signature_row(
                row,
                [
                    "material_code",
                    "allocation_code",
                    "customs_item_code",
                    "source_row",
                    "import_declaration_no",
                    "line_no",
                    "remaining_qty",
                    "available_qty",
                    "customs_value",
                    "currency",
                    "value_currency",
                    "unit_value",
                    "unit_price",
                    "taxable_unit_price",
                    "eligibility_status",
                ],
            )
            for row in stock_rows
        ],
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(encoded.encode()).hexdigest()


def compact_origin_signature_row(row: dict, fields: list[str]) -> dict:
    return {field: str(row.get(field, "")) for field in fields if row.get(field, "") not in (None, "")}


def selected_bom_rows_by_product(
    case: dict,
    bom_workspace: dict,
) -> dict[str, list[dict]]:
    selected_version_id = (
        case.get("bom_artifact_id")
        or case.get("bom_version_id")
        or bom_workspace.get("latest_version", {}).get("artifact_id")
        or bom_workspace.get("latest_version", {}).get("version_id", "")
    )
    aggregate = next(
        (
            version
            for version in bom_workspace.get("versions", [])
            if (version.get("artifact_id") or version.get("version_id")) == selected_version_id
        ),
        bom_workspace.get("latest_version", {}),
    )
    rows = aggregate.get("rows")
    if rows is None:
        rows = bom_workspace.get("latest_rows", [])
    output: dict[str, list[dict]] = {}
    for row in rows or []:
        product_code = str(row.get("product_code", "")).strip()
        if product_code:
            output.setdefault(product_code, []).append(dict(row))

    version_index = {}
    for version in bom_workspace.get("product_versions", []):
        for artifact_key in (version.get("product_artifact_id"), version.get("product_version_id"), version.get("artifact_id"), version.get("version_id")):
            if artifact_key:
                version_index[str(artifact_key)] = version
    composition_by_product = {
        row.get("product_code", ""): row.get("product_artifact_id") or row.get("product_version_id", "")
        for row in aggregate.get("product_versions", [])
    }
    overrides = {
        **dict(case.get("bom_product_version_overrides", {})),
        **dict(case.get("bom_product_artifact_overrides", {})),
    }
    for product in case.get("products", []):
        product_code = str(product.get("code", "")).strip()
        bom_product_code = resolve_bom_product_code(
            str(product.get("bom_product_code") or product_code),
            bom_workspace,
        )
        selected_product_version_id = (
            product.get("bom_product_artifact_id")
            or product.get("bom_product_version_id")
            or overrides.get(product_code)
            or overrides.get(bom_product_code)
            or composition_by_product.get(bom_product_code, "")
            or composition_by_product.get(product_code, "")
        )
        selected_product_version = version_index.get(selected_product_version_id)
        if product_code and not usable_bom_product_version(selected_product_version):
            fallback_version_id = composition_by_product.get(bom_product_code, "") or composition_by_product.get(product_code, "")
            selected_product_version = version_index.get(fallback_version_id) or latest_usable_product_version(
                bom_workspace,
                bom_product_code or product_code,
            )
        if product_code and selected_product_version and selected_product_version.get("rows") is not None:
            output[product_code] = [dict(row) for row in selected_product_version.get("rows", [])]
    return output


def attach_origin_bom_product_codes(
    case: dict,
    bom_workspace: dict,
) -> dict:
    products = []
    changed = False
    for product in case.get("products", []):
        display_code = str(product.get("code") or product.get("product_code") or "").strip()
        bom_product_code = resolve_bom_product_code(
            str(product.get("bom_product_code") or display_code),
            bom_workspace,
        )
        if bom_product_code and bom_product_code != product.get("bom_product_code"):
            updated = dict(product)
            updated["bom_product_code"] = bom_product_code
            products.append(updated)
            changed = True
        else:
            products.append(product)
    if not changed:
        return case
    prepared = dict(case)
    prepared["products"] = products
    return prepared


def resolve_bom_product_code(
    code: str,
    bom_workspace: dict,
) -> str:
    options_by_code = bom_workspace.get("product_version_options_by_code", {})
    latest_product_codes = {
        str(row.get("product_code") or "").strip()
        for row in bom_workspace.get("latest_rows", [])
        if str(row.get("product_code") or "").strip()
    }
    for candidate in bom_code_candidates(code):
        if candidate in options_by_code or candidate in latest_product_codes:
            return candidate
    candidates = bom_code_candidates(code)
    return candidates[0] if candidates else ""


def usable_bom_product_version(version: dict | None) -> bool:
    if not version:
        return False
    if version.get("flatten_status") == "non_flattened":
        return False
    return bool(version.get("rows"))


def latest_usable_product_version(bom_workspace: dict, product_code: str) -> dict:
    versions = [
        version
        for version in bom_workspace.get("product_versions", [])
        if version.get("product_code") == product_code and usable_bom_product_version(version)
    ]
    return max(versions, key=lambda version: int(version.get("product_version_no") or 0), default={})


def material_catalog_index(material_rows: list[dict]) -> dict[str, dict]:
    output = {}
    for row in material_rows:
        for key in [row.get("material_code", ""), row.get("customs_code", ""), row.get("internal_code", "")]:
            if str(key).strip():
                output[str(key).strip()] = row
    return output


def co_stock_index(stock_rows: list[dict]) -> dict[str, dict]:
    output = {}
    for row in stock_rows:
        for key in co_stock_key_candidates(row):
            existing = output.get(key)
            if existing is None or co_stock_rank(row) > co_stock_rank(existing):
                output[key] = row
    return output


def co_stock_key_candidates(row: dict) -> list[str]:
    keys = []
    for value in [row.get("material_code"), row.get("allocation_code"), row.get("customs_item_code")]:
        key = str(value or "").strip()
        if key and key not in keys:
            keys.append(key)
    return keys


def co_stock_rank(row: dict) -> tuple[bool, bool, bool]:
    return (
        co_stock_is_usable(row),
        decimal_value(row.get("remaining_qty") or row.get("available_qty") or "0") > 0,
        co_stock_has_value(row),
    )


def case_allocation_pool(
    case: dict,
    invoice_matches: list[dict],
    stock_rows: list[dict],
    *,
    min_gap_days: int | None = None,
) -> dict[str, list[dict]]:
    """Convenience: resolve the export anchor date for this case and build
    an allocation pool that applies the 2-day rule. Callers that already
    know the threshold can pass `min_gap_days`; otherwise the default
    (`DEFAULT_MIN_GAP_DAYS`) is used."""
    export_date = case_export_anchor_date(case, invoice_matches)
    gap = co_stock_eligibility.DEFAULT_MIN_GAP_DAYS if min_gap_days is None else min_gap_days
    return co_stock_allocation_pool(stock_rows, export_date=export_date, min_gap_days=gap)


def case_export_anchor_date(case: dict, invoice_matches: list[dict]) -> date | None:
    """Earliest BCCT registration_date across the case's matched export
    declarations — the anchor for the 2-day gap rule.

    Returns None when the case has no export anchor (no
    `shipment.export_declaration_nos` set OR no matching BCCT row
    found). The rule then becomes a no-op for this case — we don't
    fabricate a date from invoice_date or today() because either
    could overreject lots.
    """
    if not isinstance(case, dict) or not invoice_matches:
        return None
    shipment_nos = {
        str(value or "").strip()
        for value in (case.get("shipment") or {}).get("export_declaration_nos") or []
        if str(value or "").strip()
    }
    relevant = [
        match for match in invoice_matches
        if isinstance(match, dict)
        and (not shipment_nos or str(match.get("declaration_no") or "").strip() in shipment_nos)
    ]
    return co_stock_eligibility.earliest_export_date(relevant)


def co_stock_allocation_pool(
    stock_rows: list[dict],
    *,
    export_date: date | None = None,
    min_gap_days: int = co_stock_eligibility.DEFAULT_MIN_GAP_DAYS,
) -> dict[str, list[dict]]:
    """Group + sort candidate stock rows by allocation key.

    `export_date` + `min_gap_days` apply the regulatory 2-day rule
    (see `co_stock_eligibility.is_stock_lot_eligible`). Rejected rows
    are NOT dropped — they stay in the pool with an
    `_eligibility_reason` annotation so the substitute modal can
    surface why a candidate was filtered out. Sorting pushes
    rejected rows to the bottom.
    """
    output: dict[str, list[dict]] = {}
    for index, row in enumerate(stock_rows):
        stock = dict(row)
        stock["_allocation_sequence"] = index
        stock["_allocation_remaining_qty"] = stock_available_qty(stock)
        verdict = co_stock_eligibility.is_stock_lot_eligible(
            stock, export_date=export_date, min_gap_days=min_gap_days,
        )
        stock["_eligibility_ok"] = verdict.ok
        stock["_eligibility_reason"] = verdict.reason
        for key in co_stock_key_candidates(stock):
            output.setdefault(key, []).append(stock)
    for rows in output.values():
        rows.sort(key=co_stock_allocation_sort_key)
    return output


def co_stock_allocation_sort_key(row: dict) -> tuple:
    return (
        not co_stock_is_usable(row),
        stock_allocation_remaining_qty(row) <= 0,
        not co_stock_has_value(row),
        str(row.get("declaration_date") or row.get("import_declaration_date") or ""),
        str(row.get("import_declaration_no", "")),
        numeric_sort_text(row.get("line_no", "")),
        int(row.get("_allocation_sequence", 0)),
        str(row.get("source_row", "")),
    )


def numeric_sort_text(value) -> tuple[int, str]:
    text = str(value or "").strip()
    try:
        return int(Decimal(text)), text
    except (InvalidOperation, ValueError):
        return 0, text


def co_stock_is_usable(
    row: dict,
    *,
    export_date: date | None = None,
    min_gap_days: int | None = None,
) -> bool:
    """Boolean wrapper around `co_stock_eligibility.is_stock_lot_eligible`.

    When the row carries the `_eligibility_ok` annotation written by
    `co_stock_allocation_pool`, trust it — the pool has already done
    the work with the right `export_date` / `min_gap_days` context.
    Otherwise compute fresh with the (optional) caller-supplied params.
    """
    if "_eligibility_ok" in row:
        return bool(row["_eligibility_ok"])
    gap = co_stock_eligibility.DEFAULT_MIN_GAP_DAYS if min_gap_days is None else min_gap_days
    verdict = co_stock_eligibility.is_stock_lot_eligible(
        row, export_date=export_date, min_gap_days=gap,
    )
    return verdict.ok


def co_stock_has_value(row: dict) -> bool:
    return bool(first_non_empty([
        row.get("unit_value", ""),
        row.get("unit_price", ""),
        row.get("taxable_unit_price", ""),
        row.get("customs_value", ""),
    ]))


def stock_available_qty(row: dict) -> Decimal:
    return decimal_value(row.get("remaining_qty") or row.get("available_qty") or "0")


def stock_allocation_remaining_qty(row: dict) -> Decimal:
    if "_allocation_remaining_qty" in row:
        value = row.get("_allocation_remaining_qty")
        return value if isinstance(value, Decimal) else decimal_value(value)
    return stock_available_qty(row)


def stock_candidates_for_material(stock_pool: dict, material_code: str) -> list[dict]:
    candidates = stock_pool.get(material_code, [])
    if isinstance(candidates, dict):
        candidates = [candidates]
    output = []
    seen = set()
    for row in candidates or []:
        marker = id(row)
        if marker in seen:
            continue
        output.append(row)
        seen.add(marker)
    return sorted(output, key=co_stock_allocation_sort_key)


def origin_product_from_invoice_match(
    match: dict,
    bom_rows: list[dict],
    form_lane: dict,
    material_index: dict[str, dict],
    stock_pool: dict[str, list[dict]],
    *,
    product_sequence: int | None = None,
    bom_product_code: str = "",
) -> dict:
    product_code = str(match.get("item_code", "")).strip()
    bom_product_code = str(bom_product_code or product_code).strip()
    finished_hs = str(match.get("hs_code", "")).strip()
    preview = criteria_preview_for_hs(form_lane.get("form_code", ""), finished_hs) if form_lane else {}
    criterion = preview.get("criteria") or "Cần tra cứu PSR theo HS"
    threshold = lvc_threshold_from_criterion(criterion)
    quantity = decimal_value(match.get("quantity", "0"))
    product_value = origin_product_value(match)
    fob = product_value["value"]
    materials = [
        origin_material_from_bom_row(
            row,
            quantity,
            material_index,
            stock_pool,
            product_sequence=product_sequence,
            product_code=product_code,
            product_name=match.get("description") or product_code,
            material_sequence=material_sequence,
        )
        for material_sequence, row in enumerate(bom_rows, start=1)
    ]
    vnm = sum(
        decimal_value(material.get("non_origin_cif_value"))
        for material in materials
    )
    missing_material_values = any(
        material.get("unit_value_missing") or material.get("allocation_status") == "shortage"
        for material in materials
    )
    lvc = calculate_lvc_result(fob, vnm, threshold, missing_material_values, missing_bom_materials=not materials)
    product = {
        "code": product_code,
        "bom_product_code": bom_product_code,
        "allocation_sequence": str(product_sequence or ""),
        "name": match.get("description") or product_code,
        "finished_hs": finished_hs,
        "quantity": decimal_text(quantity),
        "unit": match.get("unit", ""),
        "currency": product_value["currency"],
        "fob_currency": product_value["currency"] or match.get("currency", ""),
        "declared_currency": match.get("currency", ""),
        "value_source": product_value["source"],
        "source_declaration_no": match.get("declaration_no", ""),
        "source_declaration_date": match.get("declaration_date") or match.get("registration_date", ""),
        "source_line_no": match.get("line_no", ""),
        "invoice_ref": match.get("invoice_ref", ""),
        "fob": decimal_text(fob) if fob is not None else "",
        "non_origin_value": decimal_text(vnm) if materials else "",
        "rvc_threshold": decimal_text(threshold) if threshold is not None else "",
        "documented_result": criterion,
        "lvc_percentage": lvc["percentage"],
        "lvc_status": lvc["status"],
        "lvc_status_label": lvc["status_label"],
        "lvc_threshold": decimal_text(threshold) if threshold is not None else "",
        "vnm_value": decimal_text(vnm) if materials else "",
        "bom_product_artifact_id": first_non_empty(row.get("product_artifact_id") or row.get("product_version_id", "") for row in bom_rows),
        "bom_product_artifact_no": first_non_empty(row.get("product_artifact_no") or row.get("product_version_no", "") for row in bom_rows),
        "bom_product_version_id": first_non_empty(row.get("product_artifact_id") or row.get("product_version_id", "") for row in bom_rows),
        "bom_product_version_no": first_non_empty(row.get("product_artifact_no") or row.get("product_version_no", "") for row in bom_rows),
        "materials": materials,
    }
    return enrich_origin_product(product)


def origin_match_from_existing_product(product: dict) -> dict:
    return {
        "item_code": product.get("code", ""),
        "description": product.get("name", ""),
        "hs_code": product.get("finished_hs", ""),
        "quantity": product.get("quantity", ""),
        "unit": product.get("unit") or product.get("export_unit", ""),
        "fob_value": product.get("fob", ""),
        "fob_currency": product.get("currency", ""),
        "currency": product.get("currency", ""),
        "declaration_no": product.get("source_declaration_no", ""),
        "line_no": product.get("source_line_no", ""),
        "invoice_ref": product.get("invoice_ref", ""),
    }


def origin_product_value(match: dict) -> dict:
    value_sources = [
        ("fob_value", match.get("fob_value"), match.get("fob_currency") or match.get("value_currency") or match.get("currency", "")),
        ("customs_value", match.get("customs_value"), match.get("value_currency") or "VND"),
        ("total_value", match.get("total_value"), match.get("value_currency") or "VND"),
        ("foreign_currency_value", match.get("foreign_currency_value"), match.get("currency", "")),
        ("invoice_value", match.get("invoice_value"), match.get("currency", "")),
    ]
    for source, value, currency in value_sources:
        if value not in (None, ""):
            return {"value": decimal_value(value), "currency": currency, "source": source}
    return {"value": None, "currency": "", "source": ""}


def origin_material_from_bom_row(
    row: dict,
    export_quantity: Decimal,
    material_index: dict[str, dict],
    stock_pool: dict,
    *,
    product_sequence: int | None = None,
    product_code: str = "",
    product_name: str = "",
    material_sequence: int | None = None,
) -> dict:
    material_code = str(row.get("material_code", "")).strip()
    material = material_index.get(material_code, {})
    qty_per = decimal_value(row.get("qty_per", "0"))
    consumed_qty = export_quantity * qty_per
    stock_candidates = stock_candidates_for_material(stock_pool, material_code)
    stock = stock_candidates[0] if stock_candidates else {}
    allocation_context = {
        "product_sequence": str(product_sequence or ""),
        "product_code": product_code,
        "product_name": product_name,
        "material_sequence": str(material_sequence or ""),
        "material_code": material_code,
        "material_uom": row.get("uom", ""),
    }
    allocation_lines, shortage_qty, shortage_trace = allocate_material_stock(
        material_code,
        consumed_qty,
        stock_candidates,
        row,
        material,
        allocation_context,
    )
    origin_details = origin_status_details_from_material(material)
    origin_status = origin_details["status"]
    fallback_unit_value, fallback_unit_value_source = first_decimal_source(
        ("bom", row.get("unit_value")),
        ("bom", row.get("unit_price")),
        ("material_catalog", material.get("unit_price")),
        ("material_catalog", material.get("taxable_unit_price")),
    )

    allocated_values = [
        decimal_value(line.get("material_value"))
        for line in allocation_lines
        if line.get("material_value") not in (None, "")
    ]
    mixed_allocation_currency = len(unique_texts(line.get("currency", "") for line in allocation_lines)) > 1
    if allocated_values and not mixed_allocation_currency:
        material_value = sum(allocated_values, Decimal("0"))
    elif not stock_candidates and fallback_unit_value is not None:
        material_value = consumed_qty * fallback_unit_value
    else:
        material_value = None
    # VND-base aggregates — even when mixed currencies make material_value
    # ambiguous, the per-line *_vnd values still sum cleanly because every
    # line was already converted to VND via its own exchange_rate_to_vnd.
    allocated_values_vnd = [
        decimal_value(line.get("material_value_vnd"))
        for line in allocation_lines
        if line.get("material_value_vnd") not in (None, "")
    ]
    if allocated_values_vnd and len(allocated_values_vnd) == len(allocation_lines):
        material_value_vnd = sum(allocated_values_vnd, Decimal("0"))
    else:
        material_value_vnd = None
    vnm_value = material_value if origin_status == "non_origin" and material_value is not None else None
    vnm_value_vnd = material_value_vnd if origin_status == "non_origin" and material_value_vnd is not None else None
    line_fx_sources = unique_texts(line.get("exchange_rate_source", "") for line in allocation_lines)
    aggregated_fx_source = line_fx_sources[0] if len(line_fx_sources) == 1 else ("mixed" if line_fx_sources else "")
    line_unit_missing = any(not line.get("unit_value") for line in allocation_lines)
    unit_value_missing = material_value is None or line_unit_missing
    allocation_status = "covered" if shortage_qty <= 0 else "shortage"
    if mixed_allocation_currency:
        valuation_status = "partial_valuation"
        unit_value_missing = True
    elif unit_value_missing:
        valuation_status = "missing_unit_value"
    elif allocation_status == "shortage":
        valuation_status = "partial_allocation"
    else:
        valuation_status = "ready"
    material_warnings = []
    material_description = row.get("material_name") or material.get("name", "") or stock.get("material_description", "")
    hs_code = row.get("hs_code") or material.get("hs_code", "") or stock.get("hs_code", "")
    if valuation_status == "missing_unit_value":
        material_warnings.append(f"{material_code}: thiếu đơn giá để tính trị giá NVL/VNM.")
    if mixed_allocation_currency:
        material_warnings.append(f"{material_code}: nhiều tiền tệ trong các dòng tồn, chưa cộng VNM tự động.")
    if allocation_status == "shortage" and consumed_qty > 0 and allocation_lines:
        material_warnings.append(
            f"{material_code}: thiếu tồn CO {decimal_text(shortage_qty)} {row.get('uom', '')} để phủ lượng dùng."
        )
    if allocation_status == "shortage" and shortage_trace:
        material_warnings.append(f"{material_code}: tồn CO đã dùng ở bước trước: {shortage_trace}.")
    if origin_details["source"] == "default_conservative":
        material_warnings.append(f"{material_code}: chưa có phân loại xuất xứ, đang tính bảo thủ là không xuất xứ.")
    if not material_description:
        material_warnings.append(f"{material_code}: thiếu tên NVL từ BOM, danh mục NVL và BCCT nhập.")
    unit_value_text = allocation_unit_value_summary(allocation_lines)
    if not unit_value_text and fallback_unit_value is not None:
        unit_value_text = decimal_text(fallback_unit_value)
    valuation_source = allocation_valuation_source(allocation_lines) or fallback_unit_value_source
    allocation_source_rows = unique_texts(line.get("source_row", "") for line in allocation_lines)
    allocation_import_declarations = unique_texts(line.get("import_declaration_no", "") for line in allocation_lines)
    allocation_import_lines = unique_texts(line.get("import_line_no", "") for line in allocation_lines)
    allocation_import_dates = unique_texts(line.get("import_declaration_date", "") for line in allocation_lines)
    available_qty = allocation_available_qty(allocation_lines, stock_candidates)
    currency = allocation_currency_summary(allocation_lines)
    if not currency:
        currency = stock.get("value_currency") or stock.get("currency") or material.get("value_currency") or material.get("currency", "")
    return {
        "source_row": ",".join(allocation_source_rows) or stock.get("source_row") or f"BOM:{row.get('source', '')}",
        "import_declaration_no": ", ".join(allocation_import_declarations) or stock.get("import_declaration_no", ""),
        "import_declaration_date": (
            ", ".join(allocation_import_dates)
            or stock.get("registration_date")
            or stock.get("declaration_date")
            or stock.get("import_declaration_date")
            or ""
        ),
        "import_line_no": ", ".join(allocation_import_lines) or stock.get("line_no", ""),
        "material_code": material_code,
        "material_sequence": str(material_sequence or ""),
        "customs_material_code": material.get("customs_code") or material_code,
        "internal_material_code": material.get("internal_code") or material_code,
        "material_description": material_description,
        "material_name_missing": not bool(material_description),
        "hs_code": hs_code,
        "origin_status": origin_status,
        "origin_status_label": origin_details["label"],
        "origin_status_source": origin_details["source"],
        "origin_status_note": origin_details["note"],
        "available_qty": available_qty,
        "consumed_qty": consumed_qty,
        "unit_value": unit_value_text,
        "currency": currency,
        "material_value": decimal_text(material_value) if material_value is not None else "",
        "material_value_native": decimal_text(material_value) if material_value is not None else "",
        "material_value_vnd": decimal_text(material_value_vnd) if material_value_vnd is not None else "",
        "non_origin_cif_value": decimal_text(vnm_value) if vnm_value is not None else "",
        "non_origin_cif_value_vnd": decimal_text(vnm_value_vnd) if vnm_value_vnd is not None else "",
        "exchange_rate_source": aggregated_fx_source,
        "unit_value_missing": unit_value_missing,
        "valuation_status": valuation_status,
        "valuation_status_label": valuation_status_label(valuation_status),
        "valuation_source": valuation_source,
        "valuation_source_label": valuation_source_label(valuation_source),
        "data_status_label": "Đủ evidence tính VNM" if valuation_status == "ready" else "Cần bổ sung evidence",
        "allocation_status": allocation_status,
        "allocation_shortage_qty": decimal_text(shortage_qty) if shortage_qty > 0 else "",
        "allocation_shortage_trace": shortage_trace,
        "allocation_lines": allocation_lines,
        "allocation_summary": allocation_summary(allocation_lines, allocation_status),
        "material_warnings": material_warnings,
        "material_warnings_text": " | ".join(material_warnings),
        "bom_qty_per": decimal_text(qty_per),
        "bom_scrap_rate": row.get("scrap_rate", ""),
        "bom_source": row.get("source", ""),
        "bom_row_class": row.get("row_class", ""),
        "uom": row.get("uom", ""),
        "source_document_ref": allocation_document_ref(allocation_lines) or row.get("source") or row.get("product_version_id", ""),
    }


def allocate_material_stock(
    material_code: str,
    required_qty: Decimal,
    stock_candidates: list[dict],
    bom_row: dict,
    material: dict,
    allocation_context: dict | None = None,
) -> tuple[list[dict], Decimal, str]:
    remaining_required = required_qty
    lines = []
    if remaining_required <= 0:
        return lines, Decimal("0"), ""
    allocation_context = allocation_context or {}
    for stock in stock_candidates:
        if not co_stock_is_usable(stock):
            continue
        available_qty = stock_allocation_remaining_qty(stock)
        if available_qty <= 0:
            continue
        allocated_qty = min(available_qty, remaining_required)
        if allocated_qty <= 0:
            continue
        line = stock_allocation_line(stock, allocated_qty, available_qty, bom_row, material, allocation_context)
        lines.append(line)
        if "_allocation_remaining_qty" in stock:
            stock["_allocation_remaining_qty"] = available_qty - allocated_qty
        stock.setdefault("_allocation_consumptions", []).append(stock_allocation_consumption(line, allocation_context))
        remaining_required -= allocated_qty
        if remaining_required <= 0:
            break
    shortage_qty = max(remaining_required, Decimal("0"))
    shortage_trace = stock_shortage_trace(stock_candidates, allocation_context) if shortage_qty > 0 else ""
    return lines, shortage_qty, shortage_trace


def stock_allocation_line(
    stock: dict,
    allocated_qty: Decimal,
    available_qty: Decimal,
    bom_row: dict,
    material: dict,
    allocation_context: dict | None = None,
) -> dict:
    allocation_context = allocation_context or {}
    unit_value, unit_value_source = first_decimal_source(
        ("bom", bom_row.get("unit_value")),
        ("bom", bom_row.get("unit_price")),
        ("co_stock", stock.get("unit_value")),
        ("co_stock", stock.get("unit_price")),
        ("co_stock", stock.get("taxable_unit_price")),
        ("material_catalog", material.get("unit_price")),
        ("material_catalog", material.get("taxable_unit_price")),
    )
    material_value = allocated_qty * unit_value if unit_value is not None else None
    fx_rate, fx_source = _allocation_line_fx(stock)
    if unit_value is not None and fx_rate is not None:
        unit_value_vnd = unit_value * fx_rate
        material_value_vnd = allocated_qty * unit_value_vnd
    else:
        unit_value_vnd = None
        material_value_vnd = None
    source_line_ids = stock.get("source_line_ids", [])
    if isinstance(source_line_ids, list):
        source_line_ids_text = ",".join(str(item) for item in source_line_ids if str(item).strip())
    else:
        source_line_ids_text = str(source_line_ids or "")
    return {
        "source_row": stock.get("source_row", ""),
        "source_line_ids": source_line_ids_text,
        "import_declaration_no": stock.get("import_declaration_no", ""),
        "import_declaration_date": (
            stock.get("registration_date")
            or stock.get("declaration_date")
            or stock.get("import_declaration_date")
            or ""
        ),
        "import_line_no": stock.get("line_no", ""),
        "customs_material_code": stock.get("customs_item_code", ""),
        "allocation_code": stock.get("allocation_code", ""),
        "product_sequence": allocation_context.get("product_sequence", ""),
        "product_code": allocation_context.get("product_code", ""),
        "material_sequence": allocation_context.get("material_sequence", ""),
        "opening_qty": decimal_text(available_qty),
        "available_qty": decimal_text(available_qty),
        "remaining_qty": decimal_text(available_qty - allocated_qty),
        "allocated_qty": decimal_text(allocated_qty),
        "unit_value": decimal_text(unit_value) if unit_value is not None else "",
        "unit_value_native": decimal_text(unit_value) if unit_value is not None else "",
        "unit_value_vnd": decimal_text(unit_value_vnd) if unit_value_vnd is not None else "",
        "currency": stock.get("value_currency") or stock.get("currency") or material.get("value_currency") or material.get("currency", ""),
        "material_value": decimal_text(material_value) if material_value is not None else "",
        "material_value_native": decimal_text(material_value) if material_value is not None else "",
        "material_value_vnd": decimal_text(material_value_vnd) if material_value_vnd is not None else "",
        "exchange_rate_to_vnd": decimal_text(fx_rate) if fx_rate is not None else "",
        "exchange_rate_source": fx_source,
        "valuation_source": unit_value_source,
        "valuation_source_label": valuation_source_label(unit_value_source),
        "material_description": stock.get("material_description", ""),
        "hs_code": stock.get("hs_code", ""),
    }


def _allocation_line_fx(stock: dict) -> tuple[Decimal | None, str]:
    """Return (rate, source) parsed from a co_stock row's FX payload fields.

    Materializer writes exchange_rate_to_vnd + exchange_rate_source (phase 1).
    Old snapshots predating phase 1 lack these fields — treat as 'missing'.
    """
    source = (stock.get("exchange_rate_source") or "").strip() or "missing"
    raw = (stock.get("exchange_rate_to_vnd") or "").strip()
    if not raw:
        return None, source
    try:
        return Decimal(raw), source
    except (InvalidOperation, ValueError):
        return None, source


def stock_allocation_consumption(line: dict, allocation_context: dict) -> dict:
    return {
        "product_sequence": allocation_context.get("product_sequence", ""),
        "product_code": allocation_context.get("product_code", ""),
        "product_name": allocation_context.get("product_name", ""),
        "material_sequence": allocation_context.get("material_sequence", ""),
        "material_code": allocation_context.get("material_code", ""),
        "material_uom": allocation_context.get("material_uom", ""),
        "allocated_qty": line.get("allocated_qty", ""),
        "source_row": line.get("source_row", ""),
        "import_declaration_no": line.get("import_declaration_no", ""),
        "import_line_no": line.get("import_line_no", ""),
    }


def stock_shortage_trace(stock_candidates: list[dict], allocation_context: dict) -> str:
    trace = []
    seen = set()
    for stock in stock_candidates:
        for consumption in stock.get("_allocation_consumptions", []):
            if not stock_consumption_is_before(consumption, allocation_context):
                continue
            marker = (
                consumption.get("product_sequence", ""),
                consumption.get("product_code", ""),
                consumption.get("material_sequence", ""),
                consumption.get("material_code", ""),
                consumption.get("source_row", ""),
                consumption.get("allocated_qty", ""),
            )
            if marker in seen:
                continue
            seen.add(marker)
            trace.append(stock_consumption_label(consumption))
    return "; ".join(trace)


def stock_consumption_is_before(consumption: dict, allocation_context: dict) -> bool:
    current_sequence = numeric_sequence(allocation_context.get("product_sequence", ""))
    consumed_sequence = numeric_sequence(consumption.get("product_sequence", ""))
    if current_sequence is None or consumed_sequence is None:
        return False
    return consumed_sequence < current_sequence


def numeric_sequence(value) -> int | None:
    try:
        return int(str(value or "").strip())
    except ValueError:
        return None


def stock_consumption_label(consumption: dict) -> str:
    step = consumption.get("product_sequence", "")
    product = consumption.get("product_code", "")
    qty = consumption.get("allocated_qty", "")
    uom = consumption.get("material_uom", "")
    source = consumption.get("import_declaration_no", "") or consumption.get("source_row", "")
    line_no = consumption.get("import_line_no", "")
    source_ref = f"{source}/{line_no}" if source and line_no else source
    prefix = f"Bước {step} {product}".strip()
    detail = f"{prefix} dùng {qty} {uom}".strip()
    return f"{detail} từ {source_ref}" if source_ref else detail


def allocation_available_qty(allocation_lines: list[dict], stock_candidates: list[dict]) -> Decimal:
    if allocation_lines:
        return sum((decimal_value(line.get("available_qty")) for line in allocation_lines), Decimal("0"))
    return sum(
        (stock_allocation_remaining_qty(row) for row in stock_candidates if co_stock_is_usable(row)),
        Decimal("0"),
    )


def allocation_unit_value_summary(allocation_lines: list[dict]) -> str:
    unit_values = unique_texts(line.get("unit_value", "") for line in allocation_lines)
    if len(unit_values) == 1:
        return unit_values[0]
    if len(unit_values) > 1:
        return "Nhiều đơn giá"
    return ""


def allocation_currency_summary(allocation_lines: list[dict]) -> str:
    currencies = unique_texts(line.get("currency", "") for line in allocation_lines)
    if len(currencies) == 1:
        return currencies[0]
    if len(currencies) > 1:
        return "Nhiều tiền tệ"
    return ""


def allocation_valuation_source(allocation_lines: list[dict]) -> str:
    if not allocation_lines:
        return ""
    sources = unique_texts(line.get("valuation_source", "") for line in allocation_lines)
    if len(allocation_lines) > 1 and sources == ["co_stock"]:
        return "co_stock_allocation"
    if len(sources) == 1:
        return sources[0]
    return "mixed_allocation"


def allocation_summary(allocation_lines: list[dict], allocation_status: str) -> str:
    if not allocation_lines:
        return "Thiếu tồn CO" if allocation_status == "shortage" else ""
    suffix = " + thiếu tồn" if allocation_status == "shortage" else ""
    return f"{len(allocation_lines)} dòng tồn{suffix}"


def allocation_document_ref(allocation_lines: list[dict]) -> str:
    refs = []
    for line in allocation_lines:
        declaration = line.get("import_declaration_no", "")
        line_no = line.get("import_line_no", "")
        if declaration and line_no:
            refs.append(f"{declaration}/{line_no}")
        elif declaration:
            refs.append(declaration)
    return "; ".join(unique_texts(refs))


def valuation_status_label(status: str) -> str:
    return {
        "ready": "Đủ giá trị",
        "missing_unit_value": "Thiếu đơn giá NVL",
        "partial_allocation": "Thiếu tồn CO",
        "partial_valuation": "Tạm tính trị giá",
    }.get(status, "Cần bổ sung evidence")


def origin_status_from_material(material: dict) -> str:
    return origin_status_details_from_material(material)["status"]


def origin_status_details_from_material(material: dict) -> dict:
    value = str(material.get("origin_default") or material.get("origin_status") or "").lower()
    if "không" in value or "khong" in value or value == "non_origin":
        return {
            "status": "non_origin",
            "label": "Không xuất xứ",
            "source": "material_catalog",
            "note": "Theo phân loại xuất xứ NVL hiện có.",
        }
    if "có" in value or value == "origin":
        return {
            "status": "origin",
            "label": "Có xuất xứ",
            "source": "material_catalog",
            "note": "Theo phân loại xuất xứ NVL hiện có.",
        }
    return {
        "status": "non_origin",
        "label": "Không xuất xứ",
        "source": "default_conservative",
        "note": "Chưa có phân loại xuất xứ, tạm tính bảo thủ vào VNM.",
    }


def attach_origin_readiness(case: dict) -> dict:
    enriched = dict(case)
    products = [
        enrich_origin_product({**product, "allocation_sequence": product.get("allocation_sequence") or str(index)})
        for index, product in enumerate(enriched.get("products", []), start=1)
    ]
    enriched["products"] = products
    snapshot = dict(enriched.get("origin_snapshot", {}))
    issue_count = sum(len(product.get("origin_warnings", [])) for product in products)
    statuses = [product.get("origin_readiness_status", "review") for product in products]
    if not products:
        readiness_status = "empty"
        readiness_label = "Chưa có dữ liệu xuất xứ"
    elif "blocked" in statuses:
        readiness_status = "blocked"
        readiness_label = "Cần bổ sung evidence"
    elif "fail" in statuses:
        readiness_status = "fail"
        readiness_label = "Có TP không đạt"
    elif "review" in statuses:
        readiness_status = "review"
        readiness_label = "Cần review tiêu chí"
    else:
        readiness_status = "ready"
        readiness_label = "Đủ điều kiện build-down"
    snapshot.update({
        "calculation_method": "build_down_lvc",
        "calculation_method_label": "Build-down LVC/RVC",
        "formula": "(FOB - VNM) / FOB x 100",
        "readiness_status": readiness_status,
        "readiness_label": readiness_label,
        "issue_count": issue_count,
        "product_count": len(products),
    })
    enriched["origin_snapshot"] = snapshot
    return enriched


def enrich_origin_product(product: dict) -> dict:
    enriched = dict(product)
    enriched["allocation_sequence"] = str(enriched.get("allocation_sequence") or "")
    materials = [enrich_origin_material(material) for material in enriched.get("materials", [])]
    enriched["materials"] = materials
    criterion = str(enriched.get("documented_result") or enriched.get("rule") or "")
    missing_unit_material_count = sum(1 for material in materials if material.get("valuation_status") == "missing_unit_value")
    shortage_material_count = sum(1 for material in materials if material.get("allocation_status") == "shortage")
    incomplete_material_count = missing_unit_material_count + shortage_material_count
    lvc = normalized_lvc_result(enriched, materials, criterion, incomplete_material_count > 0)
    enriched["lvc_percentage"] = lvc["percentage"]
    enriched["lvc_status"] = lvc["status"]
    enriched["lvc_status_label"] = lvc["status_label"]
    enriched["lvc_quality_warning_text"] = (
        f"Thiếu đơn giá {missing_unit_material_count} dòng NVL; LVC đang tạm tính từ các dòng đã có đơn giá."
        if missing_unit_material_count and lvc["percentage"]
        else f"Thiếu tồn CO {shortage_material_count} dòng NVL; LVC đang tạm tính từ phần đã phân bổ."
        if shortage_material_count and lvc["percentage"]
        else ""
    )
    ctc_rule = tariff_shift_rule_from_criterion(criterion)
    lvc_status = str(enriched.get("lvc_status") or "")
    warnings = []
    if not materials:
        warnings.append(f"{enriched.get('code', 'TP')}: chưa có BOM/NVL để tính xuất xứ.")
    if not enriched.get("fob"):
        warnings.append(f"{enriched.get('code', 'TP')}: thiếu FOB/trị giá TP.")
    for material in materials:
        warnings.extend(material.get("material_warnings", []))
    warnings = unique_texts(warnings)
    tariff_shift_status = ""
    tariff_shift_status_label = ""
    tariff_shift_note = ""
    if ctc_rule:
        non_origin_hs = [
            str(material.get("hs_code") or "")
            for material in materials
            if material.get("origin_status") == "non_origin"
        ]
        tariff_shift = evaluate_tariff_shift(str(enriched.get("finished_hs") or ""), non_origin_hs, ctc_rule)
        tariff_shift_status = "skipped" if tariff_shift.skipped else "pass" if tariff_shift.passed else "fail"
        if tariff_shift.skipped:
            tariff_shift_status_label = f"Thiếu HS cho {ctc_rule} preview"
        elif tariff_shift.passed:
            tariff_shift_status_label = f"Đạt {ctc_rule} preview"
        else:
            tariff_shift_status_label = f"Không đạt {ctc_rule} preview"
        tariff_shift_note = f"{ctc_rule} preview chỉ so HS TP với HS NVL không xuất xứ; chưa thay thế PSR engine/legal review."
    if lvc_status in {"missing_value", "missing_bom"} or any(
        material.get("valuation_status") in {"missing_unit_value", "partial_allocation", "partial_valuation"}
        or material.get("allocation_status") == "shortage"
        for material in materials
    ):
        readiness_status = "blocked"
        readiness_label = "Cần bổ sung evidence"
    elif lvc_status == "fail":
        readiness_status = "fail"
        readiness_label = "Không đạt build-down"
    elif ctc_rule:
        readiness_status = "review"
        readiness_label = "Cần review CTC"
    elif lvc_status == "pass":
        readiness_status = "ready"
        readiness_label = "Đủ điều kiện build-down"
    else:
        readiness_status = "review"
        readiness_label = "Cần review tiêu chí"
    enriched.update({
        "origin_method": "build_down_lvc",
        "origin_method_label": "Build-down LVC/RVC",
        "origin_formula": "(FOB - VNM) / FOB x 100",
        "origin_criterion_mode": criterion_mode(criterion),
        "origin_readiness_status": readiness_status,
        "origin_readiness_label": readiness_label,
        "origin_warnings": warnings,
        "origin_warning_summary": origin_warning_summary(enriched, materials, warnings),
        "origin_warnings_text": " | ".join(warnings),
        "tariff_shift_rule": ctc_rule,
        "tariff_shift_status": tariff_shift_status,
        "tariff_shift_status_label": tariff_shift_status_label,
        "tariff_shift_note": tariff_shift_note,
    })
    return enriched


def enrich_origin_material(material: dict) -> dict:
    enriched = dict(material)
    enriched["material_name_missing"] = not bool(str(enriched.get("material_description") or "").strip())
    if not enriched.get("origin_status_label"):
        enriched["origin_status_label"] = "Có xuất xứ" if enriched.get("origin_status") == "origin" else "Không xuất xứ"
    unit_missing = not str(enriched.get("unit_value") or "").strip()
    enriched["valuation_status"] = enriched.get("valuation_status") or ("missing_unit_value" if unit_missing else "ready")
    enriched["valuation_status_label"] = enriched.get("valuation_status_label") or (
        valuation_status_label(enriched["valuation_status"])
    )
    enriched["valuation_source_label"] = enriched.get("valuation_source_label") or valuation_source_label(enriched.get("valuation_source", ""))
    enriched["data_status_label"] = enriched.get("data_status_label") or (
        "Cần bổ sung evidence" if unit_missing else "Đủ evidence tính VNM"
    )
    enriched["allocation_lines"] = list(enriched.get("allocation_lines") or [])
    for allocation in enriched["allocation_lines"]:
        if not allocation.get("opening_qty"):
            allocation["opening_qty"] = allocation.get("available_qty", "")
    enriched["allocation_status"] = enriched.get("allocation_status") or ("covered" if enriched["allocation_lines"] else "")
    enriched["allocation_summary"] = enriched.get("allocation_summary") or allocation_summary(
        enriched["allocation_lines"],
        enriched["allocation_status"],
    )
    warnings = text_list(enriched.get("material_warnings") or enriched.get("material_warnings_text"))
    if unit_missing and not warnings:
        code = enriched.get("material_code") or enriched.get("internal_material_code") or "NVL"
        warnings.append(f"{code}: thiếu đơn giá để tính trị giá NVL/VNM.")
    if enriched["material_name_missing"]:
        code = enriched.get("material_code") or enriched.get("internal_material_code") or "NVL"
        warnings.append(f"{code}: thiếu tên NVL từ BOM, danh mục NVL và BCCT nhập.")
    if enriched.get("allocation_shortage_trace") and not any("đã dùng ở bước trước" in warning for warning in warnings):
        code = enriched.get("material_code") or enriched.get("internal_material_code") or "NVL"
        warnings.append(f"{code}: tồn CO đã dùng ở bước trước: {enriched['allocation_shortage_trace']}.")
    warnings = unique_texts(warnings)
    enriched["material_warnings"] = warnings
    enriched["material_warnings_text"] = " | ".join(warnings)
    return enriched


def origin_warning_summary(product: dict, materials: list[dict], warnings: list[str]) -> list[dict]:
    summary = []
    if not materials:
        summary.append({
            "kind": "missing_bom",
            "label": "Chưa có BOM/NVL",
            "count": 1,
            "detail": "Không kết luận LVC cho tới khi chọn BOM snapshot có dòng NVL.",
            "examples": product.get("code", ""),
        })
    if not product.get("fob"):
        summary.append({
            "kind": "missing_fob",
            "label": "Thiếu FOB",
            "count": 1,
            "detail": "Cần trị giá TP để tính build-down LVC/RVC.",
            "examples": product.get("code", ""),
        })
    summary.extend(material_issue_summary(materials, "missing_unit_value", "valuation_status", "Thiếu đơn giá NVL", "LVC đang tạm tính từ các dòng đã có đơn giá."))
    shortage_materials = [material for material in materials if material.get("allocation_status") == "shortage"]
    if shortage_materials:
        summary.append(material_summary_row(
            shortage_materials,
            "allocation_shortage",
            "Thiếu tồn CO",
            "LVC đang tạm tính từ phần tồn CO đã phân bổ được.",
        ))
    summary.extend(material_issue_summary(materials, "default_conservative", "origin_status_source", "Chưa phân loại xuất xứ", "Đang tạm tính bảo thủ là không xuất xứ."))
    missing_name_materials = [material for material in materials if material.get("material_name_missing")]
    if missing_name_materials:
        summary.append(material_summary_row(
            missing_name_materials,
            "missing_material_name",
            "Thiếu tên NVL",
            "Không tìm thấy tên trong BOM, danh mục NVL hoặc BCCT nhập.",
        ))
    if summary:
        return summary
    return []


def material_issue_summary(materials: list[dict], value: str, field: str, label: str, detail: str) -> list[dict]:
    rows = [material for material in materials if material.get(field) == value]
    return [material_summary_row(rows, value, label, detail)] if rows else []


def material_summary_row(materials: list[dict], kind: str, label: str, detail: str) -> dict:
    codes = unique_texts(
        material.get("internal_material_code") or material.get("material_code") or "NVL"
        for material in materials
    )
    return {
        "kind": kind,
        "label": label,
        "count": len(materials),
        "detail": detail,
        "examples": ", ".join(codes[:6]),
    }


def tariff_shift_rule_from_criterion(criterion: str) -> str:
    text = criterion.upper()
    for rule in ["CTSH", "CTH", "CC"]:
        if re.search(rf"\b{rule}\b", text):
            return rule
    return ""


def criterion_mode(criterion: str) -> str:
    has_value_content = bool(re.search(r"\b(?:LVC|RVC|AIFTA)\b", criterion, flags=re.IGNORECASE))
    has_tariff_shift = bool(tariff_shift_rule_from_criterion(criterion))
    if has_value_content and has_tariff_shift:
        return "compound"
    if has_value_content:
        return "value_content"
    if has_tariff_shift:
        return "tariff_shift"
    return "manual_review"


def valuation_source_label(source: str) -> str:
    return {
        "bom": "BOM",
        "co_stock": "BCCT nhập/tồn CO",
        "co_stock_allocation": "Tồn CO nhiều lô",
        "mixed_allocation": "Nhiều nguồn giá",
        "material_catalog": "Danh mục NVL",
    }.get(str(source or ""), "Chưa có")


def text_list(value) -> list[str]:
    if isinstance(value, list):
        return [str(item) for item in value if str(item).strip()]
    return [item.strip() for item in str(value or "").split("|") if item.strip()]


def unique_texts(values) -> list[str]:
    output = []
    seen = set()
    for value in values or []:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        output.append(text)
        seen.add(text)
    return output


def lvc_threshold_from_criterion(criterion: str) -> Decimal | None:
    if not criterion:
        return None
    match = re.search(r"(?:LVC|RVC|AIFTA)[^\d]*(\d+(?:[.,]\d+)?)\s*%", criterion, flags=re.IGNORECASE)
    if not match:
        match = re.search(r"(\d+(?:[.,]\d+)?)\s*%\s*(?:FOB|LVC|RVC)", criterion, flags=re.IGNORECASE)
    return decimal_value(match.group(1)) if match else None


def normalized_lvc_result(product: dict, materials: list[dict], criterion: str, missing_material_values: bool) -> dict:
    fob = decimal_value(product.get("fob")) if product.get("fob") not in (None, "") else None
    vnm_source = first_non_empty([product.get("vnm_value"), product.get("non_origin_value")])
    if vnm_source:
        vnm = decimal_value(vnm_source)
    else:
        vnm = sum(
            decimal_value(material.get("non_origin_cif_value"))
            for material in materials
            if material.get("origin_status") == "non_origin"
        )
    threshold_source = first_non_empty([product.get("lvc_threshold"), product.get("rvc_threshold")])
    threshold = decimal_value(threshold_source) if threshold_source else lvc_threshold_from_criterion(criterion)
    return calculate_lvc_result(fob, vnm, threshold, missing_material_values, missing_bom_materials=not materials)


def calculate_lvc_result(
    fob: Decimal | None,
    vnm: Decimal,
    threshold: Decimal | None,
    missing_material_values: bool,
    *,
    missing_bom_materials: bool = False,
) -> dict:
    if fob is None or fob <= 0:
        return {"percentage": "", "status": "missing_value", "status_label": "Thiếu FOB"}
    if missing_bom_materials:
        return {"percentage": "", "status": "missing_bom", "status_label": "Thiếu BOM/NVL"}
    percentage = ((fob - vnm) / fob * Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    percentage_text = f"{percentage:.2f}"
    if threshold is None:
        if missing_material_values:
            return {"percentage": percentage_text, "status": "partial_review", "status_label": "Tạm tính LVC"}
        return {"percentage": percentage_text, "status": "review", "status_label": "Thiếu ngưỡng"}
    if percentage >= threshold:
        if missing_material_values:
            return {"percentage": percentage_text, "status": "partial_pass", "status_label": "Tạm đạt LVC"}
        return {"percentage": percentage_text, "status": "pass", "status_label": "Đạt LVC"}
    if missing_material_values:
        return {"percentage": percentage_text, "status": "partial_fail", "status_label": "Tạm không đạt LVC"}
    return {"percentage": percentage_text, "status": "fail", "status_label": "Không đạt LVC"}


def first_non_empty(values) -> str:
    for value in values:
        if value not in (None, ""):
            return str(value)
    return ""


def first_decimal_value(*values) -> Decimal | None:
    for value in values:
        if value not in (None, ""):
            return decimal_value(value)
    return None


def first_decimal_source(*values: tuple[str, object]) -> tuple[Decimal | None, str]:
    for source, value in values:
        if value not in (None, ""):
            return decimal_value(value), source
    return None, ""


def decimal_value(value) -> Decimal:
    try:
        return Decimal(str(value or "0").replace(",", "").strip() or "0")
    except (InvalidOperation, ValueError):
        return Decimal("0")


def decimal_text(value: Decimal | str) -> str:
    if isinstance(value, str):
        return value
    if value == value.to_integral():
        return str(value.quantize(Decimal("1")))
    return str(value)


def invoice_match_criteria_rows(invoice_matches: list[dict], form_lane: dict) -> list[dict]:
    if not form_lane:
        return []
    rows = []
    seen = set()
    for row in invoice_matches:
        hs_code = str(row.get("hs_code", "")).strip()
        product_code = str(row.get("item_code", "")).strip()
        key = (product_code, hs_code, str(row.get("declaration_no", "")), str(row.get("line_no", "")))
        if not hs_code or key in seen:
            continue
        seen.add(key)
        preview = criteria_preview_for_hs(form_lane["form_code"], hs_code)
        rows.append({
            "product_code": product_code,
            "product_name": row.get("description", ""),
            "finished_hs": hs_code,
            "form": form_lane["display_name"],
            "agreement": form_lane["agreement"],
            "instrument": form_lane["instrument"],
            "rule": preview["criteria"],
            "rule_note": preview["note"],
            "source_reference": preview["source_reference"],
            "rvc_percentage": "",
            "tariff_shift_status": "Chờ BOM",
            "material_code": "",
            "material_name": "",
            "material_hs": "",
            "origin_status": "BCCT invoice",
            "non_origin_cif_value": "",
            "declaration_no": row.get("declaration_no", ""),
            "line_no": row.get("line_no", ""),
            "quantity": row.get("quantity", ""),
            "unit": row.get("unit", ""),
            "invoice_ref": row.get("invoice_ref", ""),
        })
    return rows


def enrich_client_with_source_summary(client: dict, source_summary: dict) -> dict:
    client["counts"] = {
        **client.get("counts", {}),
        "materials": source_summary["material_catalog"]["published_row_count"],
        "products": source_summary["product_catalog"]["published_row_count"],
        "bcct": source_summary["bcct"]["published_row_count"],
        "co_stock": source_summary["co_stock_row_count"],
    }
    return client


def attach_case_source_summary_snapshot(case: dict, source_summary: dict) -> dict:
    material = source_summary["material_catalog"].get("latest_version") or {}
    product = source_summary["product_catalog"].get("latest_version") or {}
    bcct = source_summary["bcct"].get("latest_version") or {}
    case["source_snapshot"] = {
        "material_catalog_version_id": material.get("version_id", ""),
        "material_catalog_version_no": material.get("version_no", ""),
        "product_catalog_version_id": product.get("version_id", ""),
        "product_catalog_version_no": product.get("version_no", ""),
        "bcct_version_id": bcct.get("version_id", ""),
        "bcct_version_no": bcct.get("version_no", ""),
        "bcct_reviewed_row_count": source_summary["bcct"].get("reviewed_row_count", 0),
        "correction_candidate_count": source_summary["bcct"].get("correction_candidate_count", 0),
        "client_config_version": source_summary["client_config"].get("config_version", ""),
        "client_config_hash": source_summary["client_config"].get("config_hash", ""),
    }
    return case


def minimal_bom_workspace() -> dict:
    return {
        "versions": [],
        "product_versions": [],
        "product_version_options_by_code": {},
        "latest_version": {},
    }


def _data_hub_overview_context(
    client_id: str,
    active: str,
    *,
    dh_path: str,
) -> dict | None:
    """Lean context for Data Hub-backed source views (catalog / bom / bcct).

    Returns None when CO is not in Data Hub mode — caller should fall through
    to the full `client_context`. When in DH mode, skips the expensive
    `source_workspace` pagination (which pulls full materials / products /
    BCCT rows over HTTP) and returns only the metadata + counts needed to
    render a summary card + link-out to Data Hub. Pattern mirrors
    `_co_stock_lean_client_context`.

    `dh_path` is the path segment on the Data Hub side
    (catalog / bom / bcct) — composed into a target URL the template can
    render as a "Mở trên Data Hub" button.
    """
    client = resolve_client(client_id)
    try:
        source_summary, source_backend = portfolio_service.source_summary(client)
    except AttributeError:
        # Test shims (FakePortfolioService / FakeSourceIndexStore) may not
        # expose the lean summary call; fall through to the legacy full-
        # workspace path which they do support.
        return None
    if source_backend != "data-hub":
        return None
    client_config = source_summary.get("client_config") or portfolio_service.get_client_config(client)
    bcct_summary = source_summary.get("bcct", {}) or {}
    material_summary = source_summary.get("material_catalog", {}) or {}
    product_summary = source_summary.get("product_catalog", {}) or {}
    bom_summary = source_summary.get("bom", {}) or {}
    counts = {
        **client.get("counts", {}),
        "materials": material_summary.get("published_row_count", 0),
        "products": product_summary.get("published_row_count", 0),
        "bcct": bcct_summary.get("published_row_count", 0),
        "bom_lines": bom_summary.get("published_row_count", client.get("counts", {}).get("bom_lines", 0)),
        "co_stock": source_summary.get("co_stock_row_count", 0),
    }
    client = {**client, "counts": counts}
    dh_base = data_hub_link_settings().data_hub_base_url
    context = {
        "client": client,
        "case": client_case(client),
        "active": active,
        "source_backend": source_backend,
        "client_config": client_config,
        "source_summary": source_summary,
        "data_hub_base_url": dh_base,
    }
    # Only emit a target URL when there is a real DH page to deep-link to.
    # /config passes dh_path="" (no DH page) → omit the field so a "Mở DH"
    # button rendered by a future config-tab template can't point at an
    # empty / bare-client URL.
    if dh_path:
        context["data_hub_target_url"] = (
            f"{dh_base.rstrip('/')}/clients/{client_id}/{dh_path}"
        )
    return context


def catalog_table_context(request: Request, client_id: str, view_name: str, **extra) -> dict:
    view = CATALOG_VIEWS[view_name]
    lean = _data_hub_overview_context(client_id, "catalog", dh_path="catalog")
    if lean is not None and not extra.get("catalog_result"):
        # DH mode and no upload result to surface → skip the table build entirely.
        lean["catalog_view"] = {**view, "name": view_name}
        return lean
    context = client_context(client_id, "catalog", **extra)
    rows = context["source_workspace"][view["module"]]["published_rows"]
    if context["source_backend"] == "data-hub":
        view = {
            **view,
            "columns": [
                {**column, "key": "uom"} if column.get("key") == "unit" else column
                for column in view["columns"]
            ],
            "filters": [
                {**filter_row, "field": "uom"} if filter_row.get("field") == "unit" else filter_row
                for filter_row in view["filters"]
            ],
            "summary_fields": [
                {**summary_row, "field": "uom"} if summary_row.get("field") == "unit" else summary_row
                for summary_row in view["summary_fields"]
            ],
        }
    context["catalog_view"] = {**view, "name": view_name}
    context["source_table"] = build_table_view(
        rows,
        columns=view["columns"],
        query=request.query_params,
        filters=view["filters"],
        summary_fields=view["summary_fields"],
        default_sort=view["default_sort"],
    )
    return context


def bcct_table_context(request: Request, client_id: str, direction: str | None = None, **extra) -> dict:
    lean = _data_hub_overview_context(client_id, "bcct", dh_path="bcct")
    if lean is not None and not extra.get("bcct_result"):
        lean["bcct_view"] = direction or "all"
        title_by_direction = {"import": "BCCT nhập khẩu", "export": "BCCT xuất khẩu"}
        lean["bcct_title"] = title_by_direction.get(direction, "BCCT nhập khẩu / xuất khẩu")
        return lean
    context = client_context(client_id, "bcct", **extra)
    rows = [bcct_table_row(row) for row in context["source_workspace"]["bcct"]["published_rows"]]
    if direction:
        rows = [row for row in rows if row.get("direction") == direction]
    if direction == "export":
        relevant_types = set(context["client_config"]["bcct"].get("relevant_export_declaration_types", []))
        if relevant_types:
            rows = [row for row in rows if row.get("declaration_type") in relevant_types]
    context["bcct_view"] = direction or "all"
    title_by_direction = {"import": "BCCT nhập khẩu", "export": "BCCT xuất khẩu"}
    context["bcct_title"] = title_by_direction.get(direction, "BCCT nhập khẩu / xuất khẩu")
    context["source_table"] = build_table_view(
        rows,
        columns=BCCT_COLUMNS,
        query=request.query_params,
        filters=[
            {
                "name": "direction",
                "field": "direction",
                "label": "Luồng",
                "options": [
                    {"value": "import", "label": "Nhập khẩu"},
                    {"value": "export", "label": "Xuất khẩu"},
                ],
            },
            {"name": "type", "field": "declaration_type", "label": "Loại hình"},
            {"name": "hs", "field": "hs_code", "label": "HS"},
            {"name": "origin", "field": "origin_country", "label": "Xuất xứ"},
        ],
        summary_fields=[
            {"field": "direction_label", "label": "Luồng"},
            {"field": "declaration_type", "label": "Loại hình"},
        ],
        default_sort="declaration_no",
    )
    return context


def _co_stock_lean_client_context(client_id: str, co_stock_row_count: int | None = None) -> dict:
    """Lean context for /co-stock — skips full source_workspace pagination.

    Standard client_context calls source_workspace_for_client which paginates
    full BCCT from Data Hub (~10s on Johnson). /co-stock only needs client
    meta + client_config + counts + materialized co_stock_rows, so we build
    those directly.

    Also computes sync_status by comparing current BCCT row count against
    the snapshot's recorded count (see co_stock_materializer.compute_sync_status).
    """
    client = resolve_client(client_id)
    source_summary, source_backend = portfolio_service.source_summary(client)
    client_config = source_summary.get("client_config") or portfolio_service.get_client_config(client)
    bcct_row_count = source_summary.get("bcct", {}).get("published_row_count", 0)
    if co_stock_row_count is None:
        co_stock_row_count = co_stock_materializer.row_count(client["id"])
    counts = {
        **client.get("counts", {}),
        "materials": source_summary.get("material_catalog", {}).get("published_row_count", 0),
        "products": source_summary.get("product_catalog", {}).get("published_row_count", 0),
        "bcct": bcct_row_count,
        "co_stock": co_stock_row_count,
    }
    client = {**client, "counts": counts}
    sync_status = co_stock_materializer.compute_sync_status(client["id"], bcct_row_count)
    return {
        "client": client,
        "active": "co-stock",
        "source_backend": source_backend,
        "client_config": client_config,
        "source_workspace": {
            "client_config": client_config,
            "bcct": {"latest_version": source_summary.get("bcct", {}).get("latest_version") or {},
                     "correction_candidates": []},
            "material_catalog": {"latest_version": source_summary.get("material_catalog", {}).get("latest_version") or {}},
            "product_catalog": {"latest_version": source_summary.get("product_catalog", {}).get("latest_version") or {}},
        },
        "co_stock_last_refresh_at": co_stock_materializer.last_refresh_at(client["id"]),
        "co_stock_sync_status": sync_status,
    }


_CO_STOCK_STATUS_OPTIONS = [
    {"value": "available", "label": "Khả dụng"},
    {"value": "review_required", "label": "Cần review"},
    {"value": "inactive", "label": "Không dùng"},
    {"value": "depleted", "label": "Hết tồn"},
]


def co_stock_table_context(request: Request, client_id: str) -> dict:
    """SQL-paginated Tồn CO context. The DB does WHERE/ORDER/LIMIT so the
    page returns ~50 rows × few-ms even on 60k-row clients (Johnson).
    Ledger + adjustments apply only to the visible slice.

    Falls back to in-memory build_table_view for clients whose stock pool
    hasn't been materialized into `co_stock_rows` yet (file-mode demo
    fixtures + first-time-ever loads when no refresh has run).
    """
    from app.table_view import (
        DEFAULT_PAGE_SIZE,
        PAGE_SIZES,
        normalize_column,
        normalize_query,
        page_query,
        parse_int,
        table_field_names,
    )

    total_co_stock_rows = co_stock_materializer.row_count(client_id)
    # File-mode (no DB or no Data Hub configured) → legacy in-memory path so
    # Growatt-style fixtures keep working. Data Hub mode always uses the
    # materialized table — when empty, the UI shows a "Chưa có snapshot"
    # state that prompts the operator to click Refresh.
    use_legacy_path = total_co_stock_rows == 0 and not data_hub_link_settings().source_enabled
    if use_legacy_path:
        context = client_context(client_id, "co-stock")
        rows = [co_stock_table_row(row) for row in context["client"]["co_stock"]]
        context["co_stock_last_refresh_at"] = co_stock_materializer.last_refresh_at(client_id)
        context["co_stock_empty_needs_refresh"] = not context["client"]["co_stock"]
        context["co_stock_sync_status"] = {"status": "no_snapshot", "snapshot_rows": 0,
                                           "snapshot_bcct_rows": 0, "bcct_now_rows": 0,
                                           "refreshed_at": "", "delta": 0}
        context["source_table"] = build_table_view(
            rows,
            columns=CO_STOCK_COLUMNS,
            query=request.query_params,
            filters=[{"name": "status", "field": "status", "label": "Trạng thái",
                      "options": _CO_STOCK_STATUS_OPTIONS}],
            summary_fields=[
                {"field": "status_label", "label": "Trạng thái"},
                {"field": "declaration_type", "label": "Loại hình"},
            ],
            default_sort="import_declaration_no",
        )
        return context

    context = _co_stock_lean_client_context(client_id, co_stock_row_count=total_co_stock_rows)
    client = context["client"]
    context["co_stock_empty_needs_refresh"] = total_co_stock_rows == 0

    query_values = normalize_query(request.query_params)
    field_names = table_field_names("")
    q = query_values.get(field_names["q"], "").strip()
    status_value = query_values.get("status", "").strip()
    sort_key = query_values.get(field_names["sort"]) or "import_declaration_no"
    direction = "desc" if query_values.get(field_names["dir"]) == "desc" else "asc"
    per_page = min(max(1, parse_int(query_values.get(field_names["per_page"]), DEFAULT_PAGE_SIZE)), 500)
    page = max(1, parse_int(query_values.get(field_names["page"]), 1))

    page_rows, filtered_count = co_stock_materializer.read_co_stock_page(
        client["id"],
        q=q,
        status=status_value,
        sort=sort_key,
        direction=direction,
        offset=(page - 1) * per_page,
        limit=per_page,
    )
    # Apply ledger + adjustments only to the visible page.
    if page_rows:
        used_by_lot = co_stock_ledger.used_qty_by_lot(client["id"])
        if used_by_lot:
            page_rows = co_stock_ledger.apply_used_qty(page_rows, used_by_lot)
        adjustments = co_stock_adjustments_store.aggregate_by_lookup_key(client["id"])
        if adjustments:
            page_rows = co_stock_adjustments_store.apply_adjustments(page_rows, adjustments)
    rows = [co_stock_table_row(row) for row in page_rows]

    column_defs = [normalize_column(col) for col in CO_STOCK_COLUMNS]
    page_count = max(1, (filtered_count + per_page - 1) // per_page)
    page = min(page, page_count)

    prepared_query = {k: v for k, v in query_values.items() if k != field_names["page"]}
    owned_names = set(field_names.values()) | {"status"}
    reset_query = {k: v for k, v in query_values.items() if k not in owned_names}

    for col in column_defs:
        col["sort_active"] = col["key"] == sort_key
        col["sort_dir"] = direction if col["sort_active"] else ""
        col["sort_query"] = page_query(
            prepared_query,
            **{
                field_names["sort"]: col["key"],
                field_names["dir"]: "desc" if col["sort_active"] and direction == "asc" else "asc",
                field_names["page"]: 1,
            },
        )

    filter_defs = [
        {
            "name": "status",
            "query_name": "status",
            "label": "Trạng thái",
            "field": "status",
            "options": _CO_STOCK_STATUS_OPTIONS,
            "value": status_value,
        }
    ]

    context["source_table"] = {
        "rows": rows,
        "columns": column_defs,
        "filters": filter_defs,
        "summary_chips": [],
        "query": query_values,
        "field_names": field_names,
        "passthrough_params": reset_query,
        "param_prefix": "",
        "q": q,
        "sort": sort_key,
        "dir": direction,
        "page": page,
        "per_page": per_page,
        "page_sizes": PAGE_SIZES,
        "total_pages": page_count,
        "total_count": total_co_stock_rows,
        "filtered_count": filtered_count,
        "start_index": (page - 1) * per_page + 1 if rows else 0,
        "end_index": min(page * per_page, filtered_count),
        "has_previous": page > 1,
        "has_next": page < page_count,
        "previous_query": page_query(prepared_query, **{field_names["page"]: page - 1}),
        "next_query": page_query(prepared_query, **{field_names["page"]: page + 1}),
        "first_query": page_query(prepared_query, **{field_names["page"]: 1}),
        "last_query": page_query(prepared_query, **{field_names["page"]: page_count}),
        "reset_query": page_query(reset_query),
    }
    return context


def customs_exchange_rate_context(request: Request, **extra) -> dict:
    context = dict(extra)
    store = get_customs_fx_store()
    rows = store.rows(CUSTOMS_FX_CLIENT_ID)
    query = dict(request.query_params)
    if "sort" not in query:
        query["sort"] = "effective_date"
        query["dir"] = "desc"
    context["customs_fx_scope"] = CUSTOMS_FX_CLIENT_ID
    context["customs_fx_summary"] = store.summary(CUSTOMS_FX_CLIENT_ID)
    context["source_table"] = build_table_view(
        rows,
        columns=CUSTOMS_FX_COLUMNS,
        query=query,
        filters=[
            {"name": "currency", "field": "currency_code", "label": "Nguyên tệ"},
            {"name": "endpoint", "field": "source_endpoint", "label": "Nguồn API"},
        ],
        summary_fields=[
            {"field": "currency_code", "label": "Nguyên tệ"},
            {"field": "source_endpoint", "label": "Nguồn API"},
        ],
        default_sort="effective_date",
    )
    return context


def bom_context(request: Request, client_id: str, **extra) -> dict:
    lean = _data_hub_overview_context(client_id, "bom", dh_path="bom")
    if lean is not None and not extra.get("message") and not extra.get("error"):
        # DH mode: BOM is read-only; CO renders summary + link rather than
        # paginating the full bom_workspace (which fetches every product
        # version + line over HTTP from Data Hub).
        return lean
    context = client_context(client_id, "bom", **extra)
    workspace = context["bom_workspace"]
    selected_product = selected_bom_product(workspace, request.query_params.get("product", ""))
    product_rows = bom_product_table_rows(workspace, client_id, selected_product)
    line_rows = [
        bom_line_table_row(row)
        for row in workspace.get("latest_rows", [])
        if not selected_product or row.get("product_code") == selected_product
    ]
    product_table = build_table_view(
        product_rows,
        columns=BOM_PRODUCT_COLUMNS,
        query=request.query_params,
        filters=[{"name": "status", "field": "status", "label": "Trạng thái"}],
        summary_fields=[{"field": "status", "label": "Trạng thái"}],
        default_sort="product_code",
        default_per_page=25,
        param_prefix="tp_",
    )
    product_table["search_placeholder"] = "Mã thành phẩm, version, hash..."
    line_table = build_table_view(
        line_rows,
        columns=BOM_LINE_COLUMNS,
        query=request.query_params,
        filters=[
            {"name": "uom", "field": "uom", "label": "ĐVT"},
            {"name": "source", "field": "source", "label": "Nguồn"},
            {"name": "status", "field": "row_class", "label": "Trạng thái"},
        ],
        summary_fields=[
            {"field": "row_class", "label": "Trạng thái"},
            {"field": "uom", "label": "ĐVT"},
        ],
        default_sort="material_code",
        default_per_page=50,
        param_prefix="line_",
    )
    line_table["search_placeholder"] = "Mã NVL, tên NVL, trạng thái..."
    context["selected_bom_product"] = selected_product
    context["selected_bom_product_summary"] = next(
        (row for row in product_rows if row["product_code"] == selected_product),
        {},
    )
    context["bom_product_table"] = product_table
    context["bom_line_table"] = line_table
    return context


def selected_bom_product(workspace: dict, requested_product: str = "") -> str:
    codes = sorted({str(row.get("product_code", "")) for row in workspace.get("product_composition", []) if row.get("product_code")})
    if not codes:
        codes = sorted({str(row.get("product_code", "")) for row in workspace.get("latest_rows", []) if row.get("product_code")})
    requested = str(requested_product or "").strip()
    if requested in codes:
        return requested
    return codes[0] if codes else ""


def bom_product_table_rows(workspace: dict, client_id: str, selected_product: str) -> list[dict]:
    version_index = {
        version.get("product_version_id"): version
        for version in workspace.get("product_versions", [])
        if version.get("product_version_id")
    }
    rows = []
    composition = workspace.get("product_composition", [])
    if not composition:
        composition = [
            {
                "product_code": version.get("product_code", ""),
                "product_version_id": version.get("product_version_id", ""),
                "product_version_no": version.get("product_version_no", ""),
                "row_count": version.get("row_count", 0),
                "status": version.get("status", ""),
                "version_hash": version.get("version_hash", ""),
            }
            for version in workspace.get("product_versions", [])
        ]
    for row in composition:
        version = version_index.get(row.get("product_version_id"), {})
        product_code = str(row.get("product_code", ""))
        version_hash = str(row.get("version_hash") or version.get("version_hash") or "")
        rows.append({
            "product_code": product_code,
            "product_version_no": row.get("product_version_no", version.get("product_version_no", "")),
            "row_count": row.get("row_count", version.get("row_count", 0)),
            "status": row.get("status", version.get("status", "")),
            "version_hash": version_hash,
            "version_hash_short": version_hash[:10],
            "selected": "Đang xem" if product_code == selected_product else "",
            "view_href": f"/clients/{client_id}/bom?product={quote(product_code, safe='')}#bom-lines",
        })
    return rows


def bom_line_table_row(row: dict) -> dict:
    return {
        **row,
        "material_name": row.get("material_name", ""),
        "scrap_rate": row.get("scrap_rate", ""),
        "source": row.get("source", ""),
        "row_class": row.get("row_class", ""),
    }


def co_case_context(client_id: str, case_id: str = "", current_step: str = "index", **extra) -> dict:
    client = resolve_client(client_id)
    case_was_supplied = "case" in extra
    case = extra.pop("case", None)
    effective_case_id = case_id or (case or {}).get("persisted_case_id", "")
    workspace = get_case_workspace(client, effective_case_id)
    record = get_case_record(client, effective_case_id) if effective_case_id else None
    if case is None:
        case = client_case(client)
        if record:
            case = case_from_record(case, client, record)
    elif record:
        case.setdefault("persisted_case_id", record["case_id"])
        case["supporting_files"] = [dict(file_row) for file_row in record.get("supporting_files", [])]
        for key in ["bom_snapshot", "origin_snapshot", "source_snapshot", "source_invoice_matches"]:
            if key in record and key not in case:
                case[key] = json_safe(record.get(key))
    case.setdefault("persisted_case_id", "")
    case.setdefault("shipment", {"invoice_no": "", "bill_of_lading_no": ""})
    case["shipment"].setdefault("invoice_no", "")
    case["shipment"]["export_declaration_nos"] = declaration_refs(case["shipment"].get("export_declaration_nos"))
    case["shipment"].setdefault("bill_of_lading_no", "")
    case["shipment_reference_label"] = primary_shipment_reference(case["shipment"])
    if not isinstance(case.get("origin_snapshot"), dict):
        case["origin_snapshot"] = {}
    if not isinstance(case.get("bom_snapshot"), dict):
        case["bom_snapshot"] = {"composition": []}
    case["bom_snapshot"].setdefault("composition", [])
    case.setdefault("supporting_files", [])
    if case.get("products") and current_step != "origin":
        case = attach_results(case)
    if case_was_supplied and current_step == "origin":
        extra.setdefault("preserve_origin_products", True)
    force_source_refresh = bool(extra.pop("force_source_refresh", False))
    if not force_source_refresh and case.get("source_snapshot"):
        extra.setdefault("cached_case_context", True)
    extra["force_source_refresh"] = force_source_refresh
    origin_lock = active_origin_calculation_lock(client)
    for dossier in workspace["cases"]:
        dossier["delete_block_reason"] = co_case_delete_block_reason(dossier, origin_lock)
        try:
            dossier["delete_claims_summary"] = co_stock_ledger.claims_summary_for_case(
                client_id, dossier.get("case_id", "")
            )
        except Exception:  # noqa: BLE001
            dossier["delete_claims_summary"] = {"count": 0, "lots": 0}
    current_case_id = case.get("persisted_case_id") or effective_case_id
    origin_lock_owned = bool(origin_lock and current_case_id and origin_lock.get("case_id") == current_case_id)
    origin_lock_blocked = bool(origin_lock and current_case_id and origin_lock.get("case_id") != current_case_id)
    if current_step == "origin" and origin_lock_blocked:
        extra["origin_calculation_blocked"] = True
        extra.setdefault(
            "error",
            f"Khách hàng này đang có hồ sơ {origin_lock.get('case_code') or origin_lock.get('case_id')} giữ phiên tính tồn.",
        )
    form_candidates = form_candidates_for_market(case.get("destination_market", ""))
    criteria_rows = build_case_criteria_rows(case, form_candidates)
    context = co_case_light_context(
        client_id,
        case=case,
        current_step=current_step,
        case_workspace=workspace,
        form_candidates=form_candidates,
        criteria_rows=criteria_rows,
        **extra,
    )
    context["origin_calculation_lock"] = origin_lock
    context["origin_calculation_lock_owned"] = origin_lock_owned
    context["origin_calculation_lock_blocked"] = origin_lock_blocked
    return context


def co_case_workflow_steps(
    client_id: str,
    case: dict,
    current_step: str,
    invoice_matches: list[dict] | None = None,
    criteria_rows: list[dict] | None = None,
    origin_demo_active: bool = False,
    tkx_tkn_summary: dict | None = None,
) -> list[dict]:
    case_id = case.get("persisted_case_id", "")
    base_url = f"/clients/{client_id}/co-case/{case_id}" if case_id else ""
    steps = []
    for step in CO_CASE_WORKFLOW_STEPS:
        href = base_url if step["key"] == "shipment" else f"{base_url}/{step['key']}"
        status = co_case_step_status(
            case,
            step["key"],
            invoice_matches=invoice_matches or [],
            criteria_rows=criteria_rows or [],
            origin_demo_active=origin_demo_active,
            tkx_tkn_summary=tkx_tkn_summary,
        )
        steps.append({
            **step,
            "href": href,
            "active": current_step == step["key"],
            "status": status,
            "status_label": CO_CASE_STEP_STATUS_LABELS.get(status, status),
        })
    return steps


def co_case_step_status(
    case: dict,
    step_key: str,
    invoice_matches: list[dict] | None = None,
    criteria_rows: list[dict] | None = None,
    origin_demo_active: bool = False,
    tkx_tkn_summary: dict | None = None,
) -> str:
    invoice_matches = invoice_matches or []
    criteria_rows = criteria_rows or []
    shipment = case.get("shipment", {})
    has_reference = has_shipment_reference(shipment)
    has_market = bool(case.get("destination_market") and case.get("destination_market") != "Chưa nhập")
    has_products = bool(case.get("products") or criteria_rows)
    has_bom_snapshot = bool(case.get("bom_snapshot", {}).get("composition"))
    if step_key == "shipment":
        # Partial = invoice OR market but not both. Avoids misleading "thiếu"
        # when operator has filled the invoice but hasn't yet set the market.
        if has_reference and has_market:
            return "ready"
        if has_reference or has_market:
            return "review"
        return "todo"
    if step_key == "documents":
        return "ready" if case.get("supporting_files") else "todo"
    if step_key == "exports":
        if not has_reference:
            return "todo"
        if not invoice_matches:
            return "review"
        # "ready" only when actual declaration files are uploaded — matching
        # the inner page's truth instead of just BCCT row presence. Falls
        # back to "review" when summary isn't available (caller didn't pass).
        if tkx_tkn_summary is not None:
            missing_tkx = tkx_tkn_summary.get("missing_tkx") or []
            missing_tkn = tkx_tkn_summary.get("missing_tkn") or []
            if missing_tkx or missing_tkn:
                return "review"
            return "ready"
        return "review"
    if step_key == "origin":
        if origin_demo_active:
            return "preview"
        if invoice_matches and has_products and has_bom_snapshot:
            return "review"
        if has_products or invoice_matches:
            return "preview"
        return "todo"
    if step_key == "review":
        if has_reference and invoice_matches and has_products:
            return "ready"
        return "preview" if has_products else "todo"
    return "todo"


def config_context(client_id: str, **extra) -> dict:
    # /config only renders client identity + client_config knobs. It does NOT
    # need source_workspace / bom_workspace, so skip the full pagination that
    # client_context triggers (Johnson: ~65k BCCT rows over HTTP per render).
    lean = _data_hub_overview_context(client_id, "config", dh_path="")
    if lean is not None:
        lean.update(extra)
        return lean
    return client_context(client_id, "config", **extra)


def bcct_table_row(row: dict) -> dict:
    direction = row.get("direction", "")
    return {
        **row,
        "direction_label": "Nhập khẩu" if direction == "import" else "Xuất khẩu",
    }


def co_stock_table_row(row: dict) -> dict:
    remaining_qty = str(row.get("remaining_qty", ""))
    if row.get("eligibility_status") == "inactive":
        status = "inactive"
    elif row.get("allocation_code_status") != "resolved":
        status = "review_required"
    else:
        status = "depleted" if remaining_qty in {"", "0", "0.0", "0.00"} else "available"
    return {
        **row,
        "status": status,
        "status_label": {
            "available": "Khả dụng",
            "depleted": "Hết tồn",
            "inactive": "Không dùng",
            "review_required": "Cần review",
        }[status],
        "stock_reason_label": stock_reason_label(row, status),
    }


def stock_reason_label(row: dict, status: str) -> str:
    if status == "inactive" and row.get("eligibility_reason") == "excluded_by_declaration_type_config":
        return "Loại hình không active trong config"
    if status == "review_required":
        return row.get("allocation_code_reason") or "Cần review mã phân bổ"
    return row.get("eligibility_reason", "")


@app.get("/", response_class=HTMLResponse)
@app.get("/clients", response_class=HTMLResponse)
async def clients(request: Request):
    clients = portfolio_service.clients()
    if co_auth.auth_required():
        clients = co_auth.filter_visible_clients(clients, co_auth.current_user(request))
    return templates.TemplateResponse(
        request=request,
        name="clients.html",
        context={"clients": clients},
    )


@app.get("/clients/{client_id}", response_class=HTMLResponse)
async def workspace(request: Request, client_id: str):
    # Workspace overview only renders client.counts tiles. Avoid the full
    # source_workspace + bom_service.workspace pagination here — those would
    # paginate every BCCT/material/BOM row from Data Hub on each render.
    return templates.TemplateResponse(
        request=request,
        name="workspace.html",
        context=client_overview_context(client_id),
    )


def client_overview_context(client_id: str) -> dict:
    client = resolve_client(client_id)
    try:
        source_summary, source_backend = portfolio_service.source_summary(client)
    except Exception:  # noqa: BLE001
        source_summary, source_backend = {
            "material_catalog": {"published_row_count": 0},
            "product_catalog": {"published_row_count": 0},
            "bcct": {"published_row_count": 0},
            "co_stock_row_count": 0,
        }, "n/a"
    client = enrich_client_with_source_summary(client, source_summary)
    # Workspace template references client.counts.bom_lines too; surface a
    # zero so the tile renders rather than crashes.
    client["counts"]["bom_lines"] = client["counts"].get("bom_lines", 0)
    return {
        "client": client,
        "case": client_case(client),
        "active": "overview",
        "source_backend": source_backend,
    }


@app.get("/clients/{client_id}/catalog", response_class=HTMLResponse)
async def catalog(request: Request, client_id: str):
    context = _data_hub_overview_context(client_id, "catalog", dh_path="catalog") \
        or client_context(client_id, "catalog")
    return templates.TemplateResponse(
        request=request,
        name="catalog.html",
        context=context,
    )


@app.get("/clients/{client_id}/catalog/materials", response_class=HTMLResponse)
async def material_catalog(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="catalog_table.html",
        context=catalog_table_context(request, client_id, "materials"),
    )


@app.get("/clients/{client_id}/catalog/products", response_class=HTMLResponse)
async def product_catalog(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="catalog_table.html",
        context=catalog_table_context(request, client_id, "products"),
    )


@app.get("/clients/{client_id}/catalog/material-template.xlsx")
async def download_material_catalog_template(client_id: str):
    require_local_source_writes()
    content = portfolio_service.material_catalog_template(resolve_client(client_id))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{client_id}-ds-nvl-template.xlsx"'},
    )


@app.get("/clients/{client_id}/catalog/product-template.xlsx")
async def download_product_catalog_template(client_id: str):
    require_local_source_writes()
    content = portfolio_service.product_catalog_template(resolve_client(client_id))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{client_id}-ds-sp-template.xlsx"'},
    )


@app.post("/clients/{client_id}/catalog/upload", response_class=HTMLResponse)
async def upload_catalog_workbook(
    request: Request,
    client_id: str,
    file: UploadFile = File(...),
    catalog_type: str = Form("material"),
    upload_scope: str = Form("full_catalog"),
):
    require_local_source_writes()
    client = resolve_client(client_id)
    result = portfolio_service.process_catalog_upload(
        client,
        catalog_type,
        await file.read(),
        file.filename or "catalog.xlsx",
        upload_scope,
    )
    status_code = 400 if result["status"] == "failed" else 200
    view_name = "products" if catalog_type == "product" else "materials"
    return templates.TemplateResponse(
        request=request,
        name="catalog_table.html",
        status_code=status_code,
        context=catalog_table_context(
            request,
            client_id,
            view_name,
            catalog_result=result,
            message=result["message"] if status_code == 200 else "",
            error=result["message"] if status_code == 400 else "",
        ),
    )


@app.get("/clients/{client_id}/bom", response_class=HTMLResponse)
async def bom(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="bom.html",
        context=bom_context(request, client_id),
    )


@app.post("/clients/{client_id}/bom/config", response_class=HTMLResponse)
async def save_bom_config(request: Request, client_id: str):
    require_local_source_writes()
    client = resolve_client(client_id)
    form = await request.form()
    bom_service.update_config(client, {key: str(value) for key, value in form.items()})
    return templates.TemplateResponse(
        request=request,
        name="bom.html",
        context=bom_context(request, client_id, message="Đã lưu cấu hình BOM cho công ty này."),
    )


@app.get("/clients/{client_id}/config", response_class=HTMLResponse)
async def client_config(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="client_config.html",
        context=config_context(client_id),
    )


@app.post("/clients/{client_id}/config", response_class=HTMLResponse)
async def save_client_config_route(request: Request, client_id: str):
    client = resolve_client(client_id)
    form = await request.form()
    # Identity fields (legal_name / tax_code) are CO-side render metadata, not
    # source data — editable even when Data Hub source-mode is enabled. The
    # `co_stock_min_days_before_export` knob is also CO-side: it controls a
    # local CO eligibility predicate, not anything DH owns, so it persists to
    # the same local overlay without going through `require_local_source_writes`.
    if "legal_name" in form or "tax_code" in form or "co_stock_min_days_before_export" in form:
        client = dict(client)
        if "legal_name" in form:
            client["legal_name"] = str(form.get("legal_name") or "").strip()
        if "tax_code" in form:
            client["tax_code"] = str(form.get("tax_code") or "").strip()
        if "co_stock_min_days_before_export" in form:
            raw = str(form.get("co_stock_min_days_before_export") or "").strip()
            overrides = dict(client.get("co_stock_overrides") or {})
            if raw == "":
                overrides.pop("min_days_before_export", None)
            else:
                try:
                    n = int(raw)
                    if n < 0:
                        n = co_stock_eligibility.DEFAULT_MIN_GAP_DAYS
                except ValueError:
                    n = co_stock_eligibility.DEFAULT_MIN_GAP_DAYS
                overrides["min_days_before_export"] = n
            client["co_stock_overrides"] = overrides
        store = get_app_state_store()
        if store:
            store.upsert_client(client)
    require_local_source_writes()
    config = portfolio_service.get_client_config(client)
    config["co_stock"]["lot_policy"] = str(form.get("co_stock_lot_policy", "line_level"))
    config["allocation_code"]["strategy"] = str(form.get("allocation_code_strategy", "same_as_customs_code"))
    config["allocation_code"]["description_regex"] = str(form.get("description_regex", ""))
    config["allocation_code"]["fallback"] = str(form.get("allocation_code_fallback", "same_as_customs_code"))
    try:
        portfolio_service.save_client_config(client, config)
    except ValueError as exc:
        return templates.TemplateResponse(
            request=request,
            name="client_config.html",
            status_code=400,
            context=config_context(client_id, error=str(exc)),
        )
    portfolio_service.refresh_client_indexes(client)
    return templates.TemplateResponse(
        request=request,
        name="client_config.html",
        context=config_context(client_id, message="Đã lưu cấu hình công ty."),
    )


# ---------- Cost allocation ratios (LVC/RVC cost-buildup auto-fill) ----------


def _decimal_str(value: Decimal) -> str:
    """Render Decimal for the admin form: strip trailing zeros but keep at
    least one digit. Empty for zero (so the placeholder shows)."""
    if value is None or value == 0:
        return ""
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def _row_for_template(row) -> dict:
    return {
        "product_code": row.product_code,
        "coef_wages_str": _decimal_str(row.coef_wages),
        "coef_welfare_str": _decimal_str(row.coef_welfare),
        "coef_rent_str": _decimal_str(row.coef_rent),
        "coef_depreciation_str": _decimal_str(row.coef_depreciation),
        "coef_other_mfg_str": _decimal_str(row.coef_other_mfg),
        "coef_transport_storage_str": _decimal_str(row.coef_transport_storage),
        "note": row.note,
        "has_value": any([
            row.coef_wages, row.coef_welfare, row.coef_rent,
            row.coef_depreciation, row.coef_other_mfg, row.coef_transport_storage,
            row.note,
        ]),
    }


def _cost_allocation_context(client_id: str, **extra) -> dict:
    """Lightweight context for the cost-allocation admin page.

    Deliberately avoids `client_context()` because that helper pulls the full
    source workspace (BCCT scan ~2.6s on Growatt) which this page does not
    need. We render the nav with no counts; the rest of the template only
    needs client identity + the ratio rows.
    """
    from app import cost_allocation_store
    from app.cost_allocation_store import CostAllocationRow
    client = resolve_client(client_id)
    rows = cost_allocation_store.list_ratios(client_id)
    mode_b = cost_allocation_store.get_mode_b_default(client_id) or CostAllocationRow(product_code="")
    return {
        "client": client,
        "active": "cost-allocation",
        "rows": [_row_for_template(r) for r in sorted(rows, key=lambda r: r.product_code)],
        "mode_b": _row_for_template(mode_b),
        **extra,
    }


def _coef_from_form(form, key: str) -> Decimal:
    raw = str(form.get(key) or "").strip().replace(",", ".")
    if not raw:
        return Decimal(0)
    try:
        value = Decimal(raw)
    except (InvalidOperation, ValueError):
        return Decimal(0)
    return value if value >= 0 else Decimal(0)


@app.get("/clients/{client_id}/cost-allocation", response_class=HTMLResponse)
async def cost_allocation_page(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="cost_allocation.html",
        context=_cost_allocation_context(client_id),
    )


@app.post("/clients/{client_id}/cost-allocation/mode-b", response_class=HTMLResponse)
async def cost_allocation_save_mode_b(request: Request, client_id: str):
    resolve_client(client_id)  # validates client exists
    from app import cost_allocation_store
    form = await request.form()
    row = cost_allocation_store.CostAllocationRow(
        product_code="",
        coef_wages=_coef_from_form(form, "coef_wages"),
        coef_welfare=_coef_from_form(form, "coef_welfare"),
        coef_rent=_coef_from_form(form, "coef_rent"),
        coef_depreciation=_coef_from_form(form, "coef_depreciation"),
        coef_other_mfg=_coef_from_form(form, "coef_other_mfg"),
        coef_transport_storage=_coef_from_form(form, "coef_transport_storage"),
        note=str(form.get("note", "") or "").strip(),
    )
    cost_allocation_store.upsert_ratio(client_id, row)
    return templates.TemplateResponse(
        request=request,
        name="cost_allocation.html",
        context=_cost_allocation_context(client_id, message="Đã lưu hệ số mặc định Mode B."),
    )


@app.post("/clients/{client_id}/cost-allocation/mode-b/delete", response_class=HTMLResponse)
async def cost_allocation_delete_mode_b(request: Request, client_id: str):
    resolve_client(client_id)
    from app import cost_allocation_store
    cost_allocation_store.delete_ratio(client_id, "")
    return templates.TemplateResponse(
        request=request,
        name="cost_allocation.html",
        context=_cost_allocation_context(client_id, message="Đã xóa hệ số mặc định Mode B."),
    )


@app.post("/clients/{client_id}/cost-allocation/row/delete", response_class=HTMLResponse)
async def cost_allocation_delete_row(request: Request, client_id: str, product_code: str = Form(...)):
    resolve_client(client_id)
    from app import cost_allocation_store
    cost_allocation_store.delete_ratio(client_id, product_code.strip())
    return templates.TemplateResponse(
        request=request,
        name="cost_allocation.html",
        context=_cost_allocation_context(client_id, message=f"Đã xóa hệ số cho {product_code}."),
    )


@app.post("/clients/{client_id}/cost-allocation/upload", response_class=HTMLResponse)
async def cost_allocation_upload(request: Request, client_id: str, file: UploadFile = File(...)):
    resolve_client(client_id)
    from app import cost_allocation_store, cost_allocation_importer
    try:
        parsed = cost_allocation_importer.parse_excel(await file.read())
    except Exception as exc:  # noqa: BLE001
        return templates.TemplateResponse(
            request=request,
            name="cost_allocation.html",
            status_code=400,
            context=_cost_allocation_context(client_id, error=f"Không đọc được file: {exc}"),
        )
    diff = cost_allocation_store.replace_all(client_id, parsed)
    return templates.TemplateResponse(
        request=request,
        name="cost_allocation.html",
        context=_cost_allocation_context(
            client_id,
            message=f"Đã import {len(parsed)} dòng từ {file.filename}.",
            upload_diff=diff,
        ),
    )


@app.get("/clients/{client_id}/cost-allocation/template.xlsx")
async def cost_allocation_template(client_id: str):
    resolve_client(client_id)
    from openpyxl import Workbook
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet3"
    ws["A1"] = "BẢNG PHÂN BỔ TỶ LỆ CHI PHÍ"
    ws["A2"] = "STT"
    ws["B2"] = "Mã SP"
    ws["C2"] = "Lương, thưởng"
    ws["D2"] = "Phúc lợi y tế"
    ws["E2"] = "Phí thuê nhà xưởng"
    ws["F2"] = "Phí khấu hao, BH, BD"
    ws["G2"] = "CP SX chung khác"
    ws["H2"] = "Lợi nhuận (bỏ qua khi import)"
    ws["I2"] = "Vận chuyển, lưu kho, dịch vụ"
    ws["J2"] = "Ghi chú"
    # Row 4 onward = data area.
    ws["A4"] = 1
    buf = BytesIO()
    wb.save(buf)
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="cost-allocation-{client_id}-template.xlsx"'},
    )


@app.get("/clients/{client_id}/cost-allocation/resolve")
async def cost_allocation_resolve(client_id: str, product_code: str, fob: str = "0"):
    """JSON endpoint for the 'Áp hệ số' button on the origin product panel.

    Returns the resolved coefficient (Mode A → Mode B fallback) plus the
    multiplied detail values. `found` is false when neither mode matches.
    """
    resolve_client(client_id)
    from app import cost_allocation_store, cost_allocation_importer
    row = cost_allocation_store.get_ratio(client_id, product_code.strip())
    if row is None:
        return JSONResponse({"found": False, "product_code": product_code})
    try:
        fob_dec = Decimal(str(fob).replace(",", "."))
    except (InvalidOperation, ValueError):
        fob_dec = Decimal(0)
    detail = cost_allocation_importer.apply_to_fob(row, fob_dec)
    return JSONResponse({
        "found": True,
        "product_code": product_code,
        "matched_mode": "A" if row.product_code else "B",
        "matched_product_code": row.product_code,
        "fob": str(fob_dec),
        "coefficients": {
            "wages": str(row.coef_wages),
            "welfare": str(row.coef_welfare),
            "rent": str(row.coef_rent),
            "depreciation": str(row.coef_depreciation),
            "other_mfg": str(row.coef_other_mfg),
            "transport_storage": str(row.coef_transport_storage),
        },
        "details": {k: str(v) for k, v in detail.items()},
        "note": row.note,
    })


@app.post("/clients/{client_id}/bom/upload", response_class=HTMLResponse)
async def upload_bom_workbook(
    request: Request,
    client_id: str,
    file: UploadFile = File(...),
    upload_mode: str = Form("direct_bom"),
    upload_scope: str = Form(""),
    accept_review_required: str = Form(""),
):
    require_local_source_writes()
    client = resolve_client(client_id)
    result = bom_service.process_upload(
        client,
        await file.read(),
        file.filename or "bom.xlsx",
        upload_mode,
        upload_scope or None,
        accept_review_required == "on",
    )
    status_code = 400 if result["status"] == "failed" else 200
    return templates.TemplateResponse(
        request=request,
        name="bom.html",
        status_code=status_code,
        context=bom_context(
            request,
            client_id,
            bom_result=result,
            message=result["message"] if status_code == 200 else "",
            error=result["message"] if status_code == 400 else "",
        ),
    )


@app.get("/clients/{client_id}/bom/template.xlsx")
async def download_bom_template(client_id: str):
    try:
        content = bom_service.template(resolve_client(client_id))
    except RuntimeError as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{client_id}-bom-template.xlsx"'},
    )


@app.get("/clients/{client_id}/co-stock", response_class=HTMLResponse)
async def co_stock(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="co_stock.html",
        context=co_stock_table_context(request, client_id),
    )


@app.get("/clients/{client_id}/bcct", response_class=HTMLResponse)
async def bcct(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="bcct.html",
        context=bcct_table_context(request, client_id),
    )


@app.get("/clients/{client_id}/bcct/imports", response_class=HTMLResponse)
async def bcct_imports(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="bcct.html",
        context=bcct_table_context(request, client_id, "import"),
    )


@app.get("/clients/{client_id}/bcct/exports", response_class=HTMLResponse)
async def bcct_exports(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="bcct.html",
        context=bcct_table_context(request, client_id, "export"),
    )


@app.get("/customs-exchange-rates", response_class=HTMLResponse)
async def customs_exchange_rates(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="customs_exchange_rates.html",
        context=customs_exchange_rate_context(request),
    )


@app.post("/customs-exchange-rates/refresh", response_class=HTMLResponse)
async def refresh_customs_exchange_rates_route(request: Request):
    try:
        result = refresh_customs_exchange_rates(client_id=CUSTOMS_FX_CLIENT_ID)
    except Exception as exc:
        return templates.TemplateResponse(
            request=request,
            name="customs_exchange_rates.html",
            status_code=502,
            context=customs_exchange_rate_context(
                request,
                error=f"Không cập nhật được tỷ giá hải quan: {exc}",
            ),
        )
    return templates.TemplateResponse(
        request=request,
        name="customs_exchange_rates.html",
        context=customs_exchange_rate_context(
            request,
            customs_fx_result=result,
            message=(
                f"Đã cập nhật {result['fetched_row_count']} dòng tỷ giá hải quan; "
                f"đang lưu {result['saved_row_count']} dòng."
            ),
        ),
    )


@app.get("/clients/{client_id}/customs-exchange-rates")
async def client_customs_exchange_rates_redirect(client_id: str):
    return RedirectResponse("/customs-exchange-rates", status_code=303)


@app.get("/clients/{client_id}/bcct/template.xlsx")
async def download_bcct_template(client_id: str):
    require_local_source_writes()
    content = portfolio_service.bcct_template(resolve_client(client_id))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{client_id}-bcct-template.xlsx"'},
    )


@app.post("/clients/{client_id}/bcct/upload", response_class=HTMLResponse)
async def upload_bcct_workbook(request: Request, client_id: str, file: UploadFile = File(...)):
    require_local_source_writes()
    client = resolve_client(client_id)
    result = portfolio_service.process_bcct_upload(client, await file.read(), file.filename or "bcct.xlsx")
    status_code = 400 if result["status"] == "failed" else 200
    return templates.TemplateResponse(
        request=request,
        name="bcct.html",
        status_code=status_code,
        context=bcct_table_context(
            request,
            client_id,
            bcct_result=result,
            message=result["message"] if status_code == 200 else "",
            error=result["message"] if status_code == 400 else "",
        ),
    )


@app.post("/clients/{client_id}/co-stock/import")
async def import_co_stock_workbook(client_id: str, file: UploadFile = File(...)):
    """Upload a standard CO stock template xlsx. Overwrites prior snapshot
    rows for the same (declaration_no, line_no, customs_code) keys; preserves
    rows untouched by this upload (so a partial upload only updates what it
    covers).

    Run scripts/convert_co_stock.py first if uploading from the agency
    `tru-lui-co-template.xlsm` workbook.
    """
    client = resolve_client(client_id)
    content = await file.read()
    try:
        rows, parse_errors = read_standard_co_stock(content)
    except CoStockTemplateError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    batch_id = "batch_" + hashlib.sha1(content).hexdigest()[:16]
    summary = co_stock_adjustments_store.upsert_batch(
        client["id"],
        rows,
        batch_id=batch_id,
        source_file_ref=file.filename or "co_stock.xlsx",
    )
    # Invalidate the cached source context so the next case page reflects the
    # new adjustments. The cache is process-local so this is cheap.
    _CO_CASE_SOURCE_CACHE.clear()
    return JSONResponse({
        "ok": not summary.get("errors"),
        "client_id": client["id"],
        "batch_id": batch_id,
        "filename": file.filename or "",
        "parsed_rows": len(rows),
        "parse_errors": parse_errors,
        "upsert": summary,
    })


@app.post("/clients/{client_id}/co-stock/refresh")
async def refresh_co_stock_endpoint(client_id: str):
    """Materialize the derived stock pool for one client into CO's
    co_stock_rows table.

    Tries the Data Hub delta path first (`list_bcct_with_envelope` with
    `since` + `include_tombstones`) when we have a stored
    `last_bcct_server_time` AND the response carries a fresh `server_time`.
    Falls back to full-pull when either is missing — the very first refresh
    of a client, or when Data Hub is on an older contract.
    """
    client = resolve_client(client_id)
    summary = _refresh_co_stock_delta_or_full(client)
    # Invalidate the case-source cache so the substitute modal / sheet calc
    # paths see the same fresh data.
    _CO_CASE_SOURCE_CACHE.clear()
    return JSONResponse({"ok": not summary.get("errors"), **summary})


def _refresh_co_stock_delta_or_full(client: dict) -> dict:
    """Pick delta vs full refresh and run it. Records refresh state so the
    next call can decide again. Errors fall back to full on the spot so a
    transient Data Hub issue doesn't strand the operator on an old snapshot."""
    state = co_stock_materializer.read_refresh_state(client["id"]) or {}
    last_server_time = state.get("last_bcct_server_time", "") if state else ""
    data_hub = getattr(portfolio_service, "data_hub", None)
    if last_server_time and data_hub is not None and hasattr(data_hub, "list_bcct_with_envelope"):
        delta_summary = _try_delta_refresh(client, data_hub, last_server_time)
        if delta_summary is not None:
            return delta_summary
    return _full_refresh(client)


def _try_delta_refresh(client: dict, data_hub, last_server_time: str) -> dict | None:
    """Returns a summary on success, or None if delta path can't be taken
    (e.g. response missing `server_time`, indicating Data Hub doesn't yet
    support the contract on this deployment)."""
    try:
        envelope = data_hub.list_bcct_with_envelope(
            client["id"], since=last_server_time, include_tombstones=True,
        )
    except Exception as exc:  # noqa: BLE001 — log + fall back to full
        logging.getLogger(__name__).warning(
            "co_stock delta refresh pull failed for %s: %s", client["id"], exc
        )
        return None
    server_time = envelope.get("server_time") or ""
    if not server_time:
        return None  # Data Hub on old contract — caller falls back to full.
    delta_items = envelope.get("items") or []
    tombstones = envelope.get("tombstones") or []
    tombstone_source_rows = [
        f"import-row-{hashlib.sha1(str(t.get('transaction_key') or '').encode('utf-8')).hexdigest()[:16]}"
        for t in tombstones
        if isinstance(t, dict) and t.get("transaction_key")
    ]
    client_config = portfolio_service.get_client_config(client)
    from app.source_store import _safe_customs_fx_rows, co_stock_rows_from_bcct
    from app.data_hub_client import normalize_bcct_row
    delta_rows = co_stock_rows_from_bcct(
        [normalize_bcct_row(row) for row in delta_items],
        client_config,
        customs_fx_rows=_safe_customs_fx_rows(),
    )
    summary = co_stock_materializer.refresh_co_stock_for_client(
        client,
        lambda: delta_rows,
        mode="delta",
        tombstone_source_rows=tombstone_source_rows,
    )
    source_summary, _ = portfolio_service.source_summary(client)
    co_stock_materializer.record_refresh_state(
        client["id"],
        snapshot_row_count=co_stock_materializer.row_count(client["id"]),
        bcct_row_count_at_refresh=source_summary.get("bcct", {}).get("published_row_count", 0),
        last_bcct_server_time=server_time,
    )
    summary["server_time"] = server_time
    summary["tombstones_received"] = len(tombstones)
    return summary


def _full_refresh(client: dict) -> dict:
    workspace, _backend = portfolio_service.source_workspace(client)
    summary = co_stock_materializer.refresh_co_stock_for_client(
        client, lambda: workspace.get("co_stock_rows") or [],
    )
    source_summary, _ = portfolio_service.source_summary(client)
    # Best-effort: probe a quick server_time so subsequent refreshes can go
    # delta. If the deployment doesn't carry server_time yet, leave it blank
    # — _refresh_co_stock_delta_or_full will keep trying full each time.
    server_time = _probe_server_time(client)
    co_stock_materializer.record_refresh_state(
        client["id"],
        snapshot_row_count=summary.get("rows_persisted", 0),
        bcct_row_count_at_refresh=source_summary.get("bcct", {}).get("published_row_count", 0),
        last_bcct_server_time=server_time,
    )
    if server_time:
        summary["server_time"] = server_time
    return summary


def _probe_server_time(client: dict) -> str:
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is None or not hasattr(data_hub, "list_bcct_with_envelope"):
        return ""
    try:
        # Empty `since` returns full set + server_time. We discard items here
        # because the full path already pulled them via source_workspace; the
        # only goal is to capture the high-water mark.
        envelope = data_hub.list_bcct_with_envelope(client["id"], since="", include_tombstones=False)
        return envelope.get("server_time") or ""
    except Exception:  # noqa: BLE001
        return ""


@app.get("/clients/{client_id}/co-stock/lot-history")
async def co_stock_lot_history(
    client_id: str,
    declaration_no: str = "",
    line_no: str = "",
    customs_code: str = "",
    limit: int = 200,
):
    """Return chronological audit log for one lot: claim_lock/release events
    from sheet allocations + adjustment_import_insert/update from manual
    workbook uploads. Newest first.
    """
    client = resolve_client(client_id)
    if not (declaration_no.strip() and line_no.strip() and customs_code.strip()):
        raise HTTPException(status_code=400, detail="declaration_no, line_no, customs_code required")
    events = co_stock_events_store.events_for_lot(
        client["id"],
        declaration_no.strip(),
        line_no.strip(),
        customs_code.strip(),
        limit=max(1, min(int(limit), 500)),
    )
    return JSONResponse({
        "ok": True,
        "client_id": client["id"],
        "lot": {"declaration_no": declaration_no, "line_no": line_no, "customs_code": customs_code},
        "events": events,
        "count": len(events),
    })


@app.get("/clients/{client_id}/co-stock/export.xlsx")
async def export_co_stock_workbook(client_id: str):
    """Dump effective ton CO state (BCCT opening + ledger + adjustments) into
    a standard template xlsx. Heavy for big clients (full BCCT pagination
    on Data Hub mode); intended for on-demand download.
    """
    client = resolve_client(client_id)
    workspace, _backend = portfolio_service.source_workspace(client)
    stock_rows = [dict(row) for row in workspace.get("co_stock_rows") or []]
    client_id_value = client.get("id", "")
    used_by_lot = co_stock_ledger.used_qty_by_lot(client_id_value)
    if used_by_lot:
        stock_rows = co_stock_ledger.apply_used_qty(stock_rows, used_by_lot)
    adjustments = co_stock_adjustments_store.aggregate_by_lookup_key(client_id_value)
    if adjustments:
        stock_rows = co_stock_adjustments_store.apply_adjustments(stock_rows, adjustments)
    rows_for_template = []
    for row in stock_rows:
        rows_for_template.append({
            "declaration_no": row.get("import_declaration_no", ""),
            "registration_date": row.get("registration_date") or row.get("declaration_date") or row.get("import_declaration_date") or "",
            "declaration_type": row.get("declaration_type", ""),
            "line_no": row.get("line_no", ""),
            "customs_code": row.get("customs_item_code", ""),
            "hs_code": row.get("hs_code", ""),
            "goods_name": row.get("material_description") or row.get("goods_name", ""),
            "origin_country": row.get("origin_country", ""),
            "unit_price": row.get("unit_value") or row.get("unit_price", ""),
            "taxable_unit_price": row.get("taxable_unit_price", ""),
            "opening_qty": row.get("available_qty", ""),
            "unit": row.get("unit", ""),
            "partner": row.get("partner", ""),
            "invoice_no": row.get("invoice_no") or row.get("invoice_ref", ""),
            "invoice_date": row.get("invoice_date", ""),
            "exchange_rate": row.get("exchange_rate", ""),
            "used_qty": row.get("used_qty", ""),
            "source_co_no": "",  # Per-CO attribution requires reading co_stock_claims.case_id; out of scope here.
            "transaction_key": row.get("source_transaction_key", ""),
        })
    xlsx = write_standard_co_stock(rows_for_template, use_labels=True)
    filename = f"{client_id_value}-co-stock-{date.today().isoformat()}.xlsx"
    return Response(
        content=xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/clients/{client_id}/co-case", response_class=HTMLResponse)
async def co_case(request: Request, client_id: str):
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(client_id),
    )


@app.get("/clients/{client_id}/co-case/invoice-preview")
async def co_case_invoice_preview(client_id: str, invoice_no: str = "", q: str = "", export_declaration_nos: str = ""):
    client = resolve_client(client_id)
    return invoice_lookup_payload(client, invoice_no, q, export_declaration_nos)


@app.post("/clients/{client_id}/co-case/create")
async def create_co_case(request: Request, client_id: str):
    client = resolve_client(client_id)
    form = {key: str(value) for key, value in (await request.form()).items()}
    resolved = resolve_shipment_reference(client, form.get("invoice_no", ""), form.get("export_declaration_nos", ""))
    form["invoice_no"] = resolved["invoice_no"]
    form["export_declaration_nos"] = ", ".join(resolved["export_declaration_nos"])
    record = create_case_record(client, form)
    # Set a short-lived cookie so the case detail page can surface a one-time
    # toast confirming the dossier was created (without changing the redirect
    # URL — many tests + back-references rely on the canonical path).
    response = RedirectResponse(
        f"/clients/{client_id}/co-case/{record['case_id']}",
        status_code=303,
    )
    response.set_cookie(
        "co_case_just_created",
        record["case_id"],
        max_age=60,
        path=f"/clients/{client_id}/co-case/{record['case_id']}",
        httponly=False,
        samesite="lax",
    )
    return response


@app.post("/clients/{client_id}/co-case/{case_id}/delete", response_class=HTMLResponse)
async def delete_co_case(request: Request, client_id: str, case_id: str):
    if not co_auth.can_delete_co_cases(co_auth.current_user(request)):
        raise HTTPException(status_code=403, detail="Không có quyền xoá hồ sơ C/O.")
    client = resolve_client(client_id)
    form = await request.form()
    release_claims = str(form.get("confirm_release_claims") or "").strip() == "1"
    try:
        result = delete_case_record(client, case_id, release_claims=release_claims)
    except KeyError:
        raise HTTPException(status_code=404) from None
    except CaseHasActiveClaimsError as exc:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(client_id, error=str(exc)),
        )
    except ValueError as exc:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(client_id, error=str(exc)),
        )
    released = int(result.get("claims_released") or 0)
    flash = (
        f"Đã xoá hồ sơ và nhả {released} dòng tồn về kho."
        if released
        else "Đã xoá hồ sơ."
    )
    redirect = RedirectResponse(f"/clients/{client_id}/co-case", status_code=303)
    # Cookies are latin-1 only; URL-encode the Vietnamese flash text and
    # decode in the template (request.cookies.get(...) | urldecode).
    redirect.set_cookie(
        "co_flash",
        quote(flash, safe=""),
        max_age=15,
        path=f"/clients/{client_id}/co-case",
    )
    return redirect


@app.post("/clients/{client_id}/co-case/{case_id}/shipment")
async def update_co_case_shipment(request: Request, client_id: str, case_id: str):
    form = {key: str(value) for key, value in (await request.form()).items()}
    client = resolve_client(client_id)
    resolved = resolve_shipment_reference(client, form.get("invoice_no", ""), form.get("export_declaration_nos", ""))
    update_case_record(
        client,
        {
            **form,
            "id": case_id,
            "persisted_case_id": case_id,
            "shipment": {
                "invoice_no": resolved["invoice_no"],
                "export_declaration_nos": resolved["export_declaration_nos"],
                "bill_of_lading_no": form.get("bill_of_lading_no", ""),
            },
        },
    )
    return RedirectResponse(f"/clients/{client_id}/co-case/{case_id}", status_code=303)


@app.get("/clients/{client_id}/co-case/{case_id}", response_class=HTMLResponse)
async def co_case_detail(request: Request, client_id: str, case_id: str):
    # Run preload in a thread so the case detail page returns immediately.
    # Preload fetches source_context (BCCT pagination) and persists it to the
    # case record; the next /origin click reads the cached snapshot instead of
    # hitting Data Hub again. Background is fine because the shipment tab
    # doesn't need origin context, and /origin has its own fallback if preload
    # hasn't finished yet.
    asyncio.get_event_loop().run_in_executor(
        None, preload_co_case_origin_context, client_id, case_id
    )
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(client_id, case_id, "shipment"),
    )


# Specific GET routes that share the /clients/{client_id}/co-case/{case_id}/...
# prefix MUST be defined before the catch-all {step} route below — FastAPI's
# router is order-sensitive, and {step} would otherwise swallow any single-
# segment GET (e.g. /export-bang-ke) and 404 it for not being a workflow key.


@app.get("/clients/{client_id}/co-case/{case_id}/export-bang-ke")
async def export_co_case_bang_ke_workbook_get(request: Request, client_id: str, case_id: str):
    return await export_co_case_bang_ke_workbook(request, client_id, case_id)


@app.get("/clients/{client_id}/co-case/{case_id}/{step}", response_class=HTMLResponse)
async def co_case_step(request: Request, client_id: str, case_id: str, step: str):
    if step not in CO_CASE_WORKFLOW_STEP_KEYS:
        raise HTTPException(status_code=404)
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(client_id, case_id, step),
    )


@app.get("/clients/{client_id}/co-case/{case_id}/origin/calculation-payload")
async def co_case_origin_calculation_payload(client_id: str, case_id: str):
    context = co_case_context(
        client_id,
        case_id,
        current_step="origin",
        origin_demo_allowed=False,
    )
    case = context["case"]
    source_context = context.get("origin_source_context", {})
    return {
        "case_id": case.get("persisted_case_id") or case_id,
        "case_code": case.get("case_code", ""),
        "revision": origin_case_revision(case),
        "source_snapshot": json_safe(case.get("source_snapshot", {})),
        "bom_snapshot": json_safe(case.get("bom_snapshot", {})),
        "origin_snapshot": json_safe(case.get("origin_snapshot", {})),
        "origin_product_order": origin_product_order(case),
        "origin_sheet_states": json_safe(case.get("origin_sheet_states", {})),
        "products": json_safe(case.get("products", [])),
        "form_lane": json_safe(context.get("recommended_form_lane", {})),
        "source": {
            "backend": source_context.get("source_backend", ""),
            "summary": json_safe(source_context.get("source_summary", {})),
            "invoice_matches": json_safe(source_context.get("invoice_matches", [])),
            "material_rows": json_safe(source_context.get("material_rows", [])),
            "stock_rows": json_safe(source_context.get("stock_rows", [])),
        },
    }


@app.post("/clients/{client_id}/co-case/{case_id}/supporting-files")
async def upload_co_case_supporting_file(
    request: Request,
    client_id: str,
    case_id: str,
    file: UploadFile = File(...),
    document_slot: str = Form("other"),
    invoice_no: str = Form(""),
    bill_of_lading_no: str = Form(""),
):
    client = resolve_client(client_id)
    # save_supporting_file bypasses update_case_record's close-state gate;
    # check it explicitly here so closed cases also reject uploads.
    record = get_case_record(client, case_id)
    if record and co_case_is_completed(record):
        raise HTTPException(
            status_code=409,
            detail="Hồ sơ đã đóng — bấm 'Mở lại hồ sơ' ở tab Review & Xuất trước khi upload chứng từ.",
        )
    content = await file.read(MAX_SUPPORTING_FILE_BYTES + 1)
    try:
        save_supporting_file(
            client,
            case_id,
            content,
            file.filename or "supporting-file",
            document_slot,
            invoice_no,
            bill_of_lading_no,
        )
    except ValueError as exc:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=400,
            context=co_case_context(client_id, case_id, "documents", error=str(exc)),
        )
    return RedirectResponse(f"/clients/{client_id}/co-case/{case_id}/documents", status_code=303)


@app.get("/clients/{client_id}/co-case/{case_id}/supporting-files/{upload_id}")
async def download_co_case_supporting_file(client_id: str, case_id: str, upload_id: str):
    try:
        file_row, path = get_supporting_file(resolve_client(client_id), case_id, upload_id)
    except (KeyError, FileNotFoundError):
        raise HTTPException(status_code=404) from None
    return FileResponse(
        path,
        media_type=file_row.get("mime_type") or file_row.get("content_type") or "application/octet-stream",
        filename=file_row.get("original_filename") or file_row.get("filename") or "supporting-file",
    )


@app.post("/clients/{client_id}/co-case/{case_id}/export")
async def export_co_case_workbook(request: Request, client_id: str, case_id: str):
    content_type = request.headers.get("content-type", "")
    posted_case = None
    if "application/x-www-form-urlencoded" in content_type or "multipart/form-data" in content_type:
        form = await large_request_form(request)
        if form:
            posted_case = update_products_from_form({key: str(value) for key, value in form.items()})
            posted_case["persisted_case_id"] = posted_case.get("persisted_case_id") or case_id
    client = resolve_client(client_id)
    lock_result = acquire_origin_calculation_lock(client, case_id, origin_lock_actor(request))
    if not lock_result["acquired"]:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=posted_case,
                origin_demo_allowed=False,
                origin_calculation_blocked=True,
                error=f"Chưa thể export: hồ sơ {lock_result['lock'].get('case_code') or lock_result['lock'].get('case_id')} đang giữ phiên tính tồn cho khách hàng này.",
            ),
        )
    context = co_case_context(
        client_id,
        case_id,
        current_step="origin",
        case=posted_case,
        origin_demo_allowed=False,
    )
    blockers = origin_sheet_export_blockers(context["case"])
    should_enforce_sheet_state = bool(posted_case)
    if should_enforce_sheet_state and blockers:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context={
                **context,
                "error": f"Chưa thể export: bảng kê {', '.join(blockers[:5])} cần tính lại hoặc chốt trước.",
            },
        )
    if context["case"].get("persisted_case_id") and not context.get("origin_demo_active"):
        try:
            update_case_record(client, context["case"])
        except KeyError:
            pass
    content = create_case_workbook(
        context["case"],
        context["form_candidates"],
        context["invoice_matches"],
        context["criteria_rows"],
    )
    filename = safe_filename(f"{context['case']['case_code'] or 'co-case'}-dossier.xlsx")
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/clients/{client_id}/co-case/{case_id}/export-bang-ke")
async def export_co_case_bang_ke_workbook(request: Request, client_id: str, case_id: str):
    content_type = request.headers.get("content-type", "")
    posted_case = None
    if "application/x-www-form-urlencoded" in content_type or "multipart/form-data" in content_type:
        form = await large_request_form(request)
        if form:
            posted_case = update_products_from_form({key: str(value) for key, value in form.items()})
            posted_case["persisted_case_id"] = posted_case.get("persisted_case_id") or case_id
    client = resolve_client(client_id)
    case = posted_case or persisted_origin_case(client, case_id)
    # The form-rebuilt case has empty origin_sheet_states (case_from_form starts
    # with {}). Re-hydrate from the persisted DB row so per-sheet overrides
    # (form / criteria / threshold / currency_mode) AND material_overrides
    # (delete / substitute / norm-edit / added rows) are honored by the
    # renderer. Without this, the export ships every material — including
    # ones the operator deleted via Substitute — and always picks the LVC
    # template because effective_criteria is unresolved.
    if posted_case and case_id:
        try:
            persisted = persisted_origin_case(client, case_id)
        except KeyError:
            persisted = None
        if persisted:
            case["origin_sheet_states"] = persisted.get("origin_sheet_states") or {}
    case = attach_origin_sheet_states(case)
    _hydrate_material_dates_from_stock(case, client)
    _hydrate_product_export_declaration_dates(case, client)
    blockers = origin_sheet_export_blockers(case)
    if blockers:
        raise HTTPException(
            status_code=409,
            detail=f"Chưa thể xuất bảng kê: bảng kê {', '.join(blockers[:5])} cần tính lại hoặc chốt trước.",
        )
    if case.get("persisted_case_id"):
        try:
            update_case_record(client, case)
        except KeyError:
            pass
    content = create_hq_bang_ke_workbook(case)
    filename = safe_filename(f"{case['case_code'] or 'co-case'}-bang-ke-hq.xlsx")
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/clients/{client_id}/co-case/{case_id}/export-dossier-zip")
async def export_co_case_dossier_zip(client_id: str, case_id: str):
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    case = attach_origin_sheet_states(case)
    # Hard gate: case must be closed. Close itself already requires every sheet
    # locked, so this implicitly guarantees the TKN summary is complete (it
    # filters to locked sheets — see case_tkx_tkn_summary). Without this gate
    # an operator could ship a dossier whose TKN list silently omits the
    # declarations referenced by half-finished sheets.
    if not co_case_is_completed(case):
        unlocked = [
            str(p.get("code") or "?")
            for p in case.get("products") or []
            if str(p.get("origin_sheet_status") or "").strip() != "locked"
        ]
        if unlocked:
            detail = (
                f"Còn {len(unlocked)} bảng kê chưa chốt: "
                f"{', '.join(unlocked[:5])}{'…' if len(unlocked) > 5 else ''}. "
                "Chốt hết các bảng kê rồi bấm 'Đóng hồ sơ' trước khi xuất file tổng hợp."
            )
        else:
            detail = "Đóng hồ sơ trước khi xuất file tổng hợp (cần khoá để chốt danh sách TKX/TKN)."
        raise HTTPException(status_code=409, detail=detail)
    source_context = co_case_source_context(client, case)
    invoice_matches = source_context.get("invoice_matches") or []
    stock_rows = source_context.get("stock_rows") or []
    summary = case_tkx_tkn_summary(
        case,
        invoice_matches,
        stock_rows,
        source_context.get("declaration_file_counts") or {},
    )
    supporting_files: list[dict] = []
    for file_row in case.get("supporting_files", []):
        upload_id = file_row.get("upload_id") or ""
        if not upload_id:
            continue
        try:
            row, path = get_supporting_file(client, case_id, upload_id)
        except (KeyError, FileNotFoundError):
            continue
        supporting_files.append({
            "slot": row.get("slot", "other"),
            "filename": row.get("filename", "supporting.bin"),
            "content": path.read_bytes(),
        })
    declaration_archives = _try_fetch_declaration_archives(client, case, summary)
    content = create_dossier_zip(
        case,
        supporting_files,
        summary,
        data_hub_base_url=data_hub_link_settings().data_hub_base_url,
        declaration_archives=declaration_archives,
    )
    filename = safe_filename(f"{case.get('case_code') or 'co-case'}-dossier.zip")
    return StreamingResponse(
        iter([content]),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _try_fetch_declaration_archives(client: dict, case: dict, tkx_tkn_summary: dict) -> dict[str, bytes]:
    """Probe Data Hub's Bearer-aware download.zip endpoint.

    Until DH ships the endpoint per
    `.ai/api-requests/2026-05-28-bcct-declarations-download-bearer.md`,
    any error (404 / 401 / network) silently falls back to manifest-only
    mode. The dossier ZIP renderer detects the empty dict and only writes
    the README + MANIFEST entries; once DH deploys, this function returns
    populated bytes and create_dossier_zip embeds them directly.

    Keyed by archive path inside the dossier ZIP:
        `TKX/<filename>.zip` and `TKN/<filename>.zip`.
    """
    data_hub = getattr(portfolio_service, "data_hub", None)
    if data_hub is None or not hasattr(data_hub, "download_declarations_zip"):
        return {}
    case_code = (case.get("case_code") or "co-case").strip() or "co-case"
    archives: dict[str, bytes] = {}

    def _fetch(direction: str, entries: list[dict], archive_label: str) -> None:
        nos = sorted({
            str(entry.get("declaration_no") or "").strip()
            for entry in (entries or [])
            if str(entry.get("declaration_no") or "").strip()
        })
        if not nos:
            return
        filename = safe_filename(f"{archive_label}_{case_code}.zip")
        try:
            blob = data_hub.download_declarations_zip(
                client["id"], direction=direction, declaration_nos=nos, filename=filename,
            )
        except Exception:  # noqa: BLE001 — fall back to manifest-only on any failure
            return
        if isinstance(blob, (bytes, bytearray)) and blob:
            archives[f"{archive_label}/{filename}"] = bytes(blob)

    _fetch("export", tkx_tkn_summary.get("tkx") or [], "TKX")
    _fetch("import", tkx_tkn_summary.get("tkn") or [], "TKN")
    return archives


@app.post("/clients/{client_id}/co-case/{case_id}/close")
async def close_co_case(request: Request, client_id: str, case_id: str):
    """Mark the case as completed. Pre-conditions:

    - Every product's origin sheet must be in `locked` status. A case with
      a half-finished bảng kê isn't ready to be filed; we refuse rather
      than silently freezing edits on top of incomplete data.
    - The case must currently be open. (Re-closing a closed case is a
      no-op; the route is idempotent in spirit, but `update_case_record`
      treats it as a normal mutation, so no-op early.)

    After close: every mutating endpoint refuses via CaseClosedError.
    Also releases any origin-calculation lock the case holds."""
    client = resolve_client(client_id)
    try:
        record = get_case_record(client, case_id)
    except KeyError:
        raise HTTPException(status_code=404) from None
    if record is None:
        raise HTTPException(status_code=404)
    case = case_from_record(default_client_case(client), client, record)
    # Idempotent: re-clicking close on a closed case redirects without write.
    if co_case_is_completed(case):
        return RedirectResponse(
            f"/clients/{client_id}/co-case/{case_id}/review",
            status_code=303,
        )
    case = attach_origin_sheet_states(case)
    products = case.get("products") or []
    if not products:
        raise HTTPException(
            status_code=409,
            detail="Chưa có bảng kê nào — không thể đóng hồ sơ rỗng.",
        )
    unlocked = [
        p.get("code") or "?"
        for p in products
        if str(p.get("origin_sheet_status") or "").strip() != "locked"
    ]
    if unlocked:
        raise HTTPException(
            status_code=409,
            detail=(
                f"Còn {len(unlocked)} bảng kê chưa chốt: "
                f"{', '.join(unlocked[:5])}{'…' if len(unlocked) > 5 else ''}. "
                "Chốt tất cả ở tab Bảng kê C/O trước khi đóng hồ sơ."
            ),
        )
    try:
        update_case_record(client, {"id": case_id, "persisted_case_id": case_id, "status": "completed"})
    except KeyError:
        raise HTTPException(status_code=404) from None
    # Best-effort lock release: don't fail the close if no lock is held.
    try:
        existing_lock = active_origin_calculation_lock(client)
        if existing_lock and existing_lock.get("case_id") == case_id:
            release_origin_calculation_lock(client, case_id)
    except Exception:  # noqa: BLE001 — lock release is housekeeping
        pass
    return RedirectResponse(
        f"/clients/{client_id}/co-case/{case_id}/review",
        status_code=303,
    )


@app.post("/clients/{client_id}/co-case/{case_id}/reopen-case")
async def reopen_co_case(request: Request, client_id: str, case_id: str):
    """Re-open a completed case for further edits."""
    client = resolve_client(client_id)
    try:
        update_case_record(client, {"id": case_id, "persisted_case_id": case_id, "status": "open"})
    except KeyError:
        raise HTTPException(status_code=404) from None
    return RedirectResponse(
        f"/clients/{client_id}/co-case/{case_id}/review",
        status_code=303,
    )


@app.post("/clients/{client_id}/co-case/{case_id}/origin-lock/release")
async def release_co_case_origin_lock(client_id: str, case_id: str, next_url: str = Form("")):
    client = resolve_client(client_id)
    try:
        record = get_case_record(client, case_id)
        case = case_from_record(default_client_case(client), client, record)
        case = mark_origin_sheets_stale(case, 0)
        update_case_record(client, case)
    except KeyError:
        pass
    release_origin_calculation_lock(client, case_id)
    redirect_url = next_url if next_url.startswith(f"/clients/{client_id}/co-case") else f"/clients/{client_id}/co-case/{case_id}/origin"
    return RedirectResponse(redirect_url, status_code=303)


@app.post("/clients/{client_id}/co-case/{case_id}/origin/save")
async def save_co_case_origin(request: Request, client_id: str, case_id: str):
    client = resolve_client(client_id)
    case, payload = await origin_case_from_request(request, client, case_id)
    try:
        stale_from_index = int(str(payload.get("stale_from_index", "0") or "0"))
    except ValueError:
        stale_from_index = 0
    if payload.get("mark_stale", True):
        case = mark_origin_sheets_stale(case, stale_from_index)
    else:
        case = attach_origin_sheet_states(case)
    try:
        update_case_record(client, case)
    except KeyError:
        raise HTTPException(status_code=404) from None
    return {
        "status": "ok",
        "revision": origin_case_revision(case),
        "stale_from_index": stale_from_index,
        "origin_product_order": origin_product_order(case),
        "origin_sheet_states": json_safe(case.get("origin_sheet_states", {})),
    }


@app.post("/clients/{client_id}/co-case/{case_id}/origin/autosave")
async def autosave_co_case_origin(request: Request, client_id: str, case_id: str):
    client = resolve_client(client_id)
    case, payload = await origin_case_from_request(request, client, case_id)
    try:
        stale_from_index = int(str(payload.get("stale_from_index", "0") or "0"))
    except ValueError:
        stale_from_index = 0
    case = mark_origin_sheets_stale(case, stale_from_index)
    try:
        update_case_record(client, case)
    except KeyError:
        raise HTTPException(status_code=404) from None
    return {
        "status": "ok",
        "revision": origin_case_revision(case),
        "stale_from_index": stale_from_index,
        "origin_product_order": origin_product_order(case),
        "origin_sheet_states": json_safe(case.get("origin_sheet_states", {})),
    }


CALCULATE_SNAPSHOT_FRESHNESS_SECONDS = 30


def _calculate_stock_rows_from_snapshot(client: dict) -> list[dict] | None:
    """Returns stock rows ready for `prepare_case_origin_sheet`, decorated
    with current ledger used/remaining qty, or None when the snapshot
    isn't usable (no DB, empty for this client, or delta refresh failed).

    Skips the DH delta round trip when the materialized snapshot was
    refreshed within `CALCULATE_SNAPSHOT_FRESHNESS_SECONDS` — operators
    clicking Load BOM repeatedly within a 30-second window don't pay
    the ~4s DH delta cost per click. The explicit /refresh-co-stock
    endpoint bypasses this TTL when the operator wants to force a
    pull (e.g. right after importing fresh BCCT in Data Hub).

    The caller falls back to the legacy full-pull path on None — the
    operator never silently calculates against an empty snapshot.
    """
    client_id = str(client.get("id", "")) if isinstance(client, dict) else ""
    if not client_id:
        return None
    state = co_stock_materializer.read_refresh_state(client_id) or {}
    refreshed_at = str(state.get("refreshed_at") or "")
    is_fresh = False
    if refreshed_at:
        try:
            last = datetime.fromisoformat(refreshed_at)
            age = (datetime.now(last.tzinfo or timezone.utc) - last).total_seconds()
            is_fresh = age < CALCULATE_SNAPSHOT_FRESHNESS_SECONDS
        except ValueError:
            is_fresh = False
    if not is_fresh:
        try:
            summary = _refresh_co_stock_delta_or_full(client)
        except Exception as exc:  # noqa: BLE001 — fall back, never block /calculate
            logging.getLogger(__name__).warning(
                "co_stock delta-refresh failed for %s; using legacy full pull: %s", client_id, exc
            )
            return None
        if summary.get("errors"):
            return None
    rows = co_stock_materializer.read_co_stock_rows_cached(client_id)
    if not rows:
        return None
    # apply_used_qty mutates the rows in place to attach used/remaining,
    # so copy the cached payloads first — the cache must stay clean.
    rows = [dict(r) for r in rows]
    used_by_lot = co_stock_ledger.used_qty_by_lot(client_id)
    return co_stock_ledger.apply_used_qty(rows, used_by_lot)


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/calculate", response_class=HTMLResponse)
async def calculate_co_case_origin_sheet(request: Request, client_id: str, case_id: str, product_code: str):
    client = resolve_client(client_id)
    case, _payload = await origin_case_from_request(request, client, case_id)
    try:
        persisted = get_case_record(client, case_id)
        if persisted.get("origin_sheet_states"):
            case["origin_sheet_states"] = dict(persisted.get("origin_sheet_states") or {})
    except KeyError:
        pass
    action_error = origin_sheet_action_error(case, product_code, "calculate")
    if action_error:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=action_error,
                preserve_origin_products=True,
                fast_origin_context=True,
            ),
        )
    lock_result = acquire_origin_calculation_lock(client, case_id, origin_lock_actor(request))
    if not lock_result["acquired"]:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=f"Chưa thể load BOM vào bảng kê: hồ sơ {lock_result['lock'].get('case_code') or lock_result['lock'].get('case_id')} đang giữ phiên tính tồn cho khách hàng này.",
                origin_calculation_blocked=True,
                preserve_origin_products=True,
            ),
        )
    # Fast path: delta-refresh the materialized stock snapshot (1-2s when
    # the upstream BCCT is quiet, vs 30-45s for a full DH pull every time)
    # and read stock rows from co_stock_rows directly. invoice_matches +
    # material_rows are pulled through the cached snapshot the shipment
    # step already populated. Falls back to the legacy full-pull path when
    # the snapshot is empty (fresh client) or delta refresh errored, so we
    # never silently calculate against stale data.
    snapshot_stock_rows = _calculate_stock_rows_from_snapshot(client)
    if snapshot_stock_rows is not None:
        context = co_case_context(
            client_id,
            case_id,
            current_step="origin",
            case=case,
            message=f"Đã load BOM vào bảng kê {product_code}.",
            preserve_origin_products=True,
            cached_case_context=True,
        )
        source_context = context.get("origin_source_context", {})
        stock_rows = snapshot_stock_rows
    else:
        context = co_case_context(
            client_id,
            case_id,
            current_step="origin",
            case=case,
            message=f"Đã load BOM vào bảng kê {product_code}.",
            preserve_origin_products=True,
            force_source_refresh=True,
        )
        source_context = context.get("origin_source_context", {})
        stock_rows = source_context.get("stock_rows", [])
    try:
        client_config_for_rule = (
            portfolio_service.get_client_config(client)
            if hasattr(portfolio_service, "get_client_config") else {}
        )
    except Exception:  # noqa: BLE001
        client_config_for_rule = {}
    min_gap_days = effective_min_gap_days(client, client_config_for_rule)
    context["case"] = prepare_case_origin_sheet(
        context["case"],
        product_code,
        source_context.get("invoice_matches", []),
        context.get("bom_workspace", minimal_bom_workspace()),
        context.get("recommended_form_lane", {}),
        source_context.get("material_rows", []),
        stock_rows,
        min_gap_days=min_gap_days,
    )
    context["case"] = attach_case_bom_snapshot(context["case"], context.get("bom_workspace", minimal_bom_workspace()))
    context["case"] = attach_origin_bom_product_codes(
        context["case"],
        context.get("bom_workspace", minimal_bom_workspace()),
    )
    context["case"] = attach_origin_readiness(context["case"])
    context["case"] = attach_results(context["case"])
    context["case"] = attach_origin_sheet_states(context["case"])
    context["case"] = set_origin_sheet_status(context["case"], product_code, "calculated")
    target_index = next(
        (
            index
            for index, product in enumerate(context["case"].get("products", []))
            if str(product.get("code") or "").strip() == product_code
        ),
        -1,
    )
    if target_index >= 0:
        context["case"] = mark_origin_sheets_stale(context["case"], target_index + 1)
        context["case"] = set_origin_sheet_status(context["case"], product_code, "calculated")
        states = dict(context["case"].get("origin_sheet_states") or {})
        previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
        if previous.get("material_overrides"):
            states[product_code] = {**previous, "material_overrides": {}}
            context["case"]["origin_sheet_states"] = states
            context["case"] = attach_origin_sheet_states(context["case"])
    context["criteria_rows"] = build_case_criteria_rows(context["case"], context.get("form_candidates", []))
    if context["case"].get("persisted_case_id") and not context.get("origin_demo_active"):
        update_case_record(client, context["case"])
    return templates.TemplateResponse(request=request, name="co_case.html", context=context)


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/lock", response_class=HTMLResponse)
async def lock_co_case_origin_sheet(request: Request, client_id: str, case_id: str, product_code: str):
    client = resolve_client(client_id)
    case, _payload = await origin_case_from_request(request, client, case_id)
    action_error = origin_sheet_action_error(case, product_code, "lock")
    if action_error:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=action_error,
                preserve_origin_products=True,
                fast_origin_context=True,
            ),
        )
    # Capture allocations from the PERSISTED case BEFORE update_case_record
    # rewrites the disk record. The form-rebuilt `case` may have stripped
    # materials/allocation_lines if the AJAX submitter only sent metadata.
    # The ledger writes claims in a single transaction with an availability
    # pre-check; on any failure we must NOT update case state, otherwise the
    # sheet appears locked while no claim was recorded (the ghost-claim bug
    # this code path used to suffer from silent exception swallowing).
    try:
        record_sheet_lock_claims(client_id, case_id, product_code, case)
    except co_stock_ledger.StockOverclaimError as exc:
        detail_lines = [
            f"{v['source_row']}: cần {v['claimed']}, còn {v['available']}"
            f" (gốc {v['bcct_remaining']} - case khác {v['other_claims']})"
            for v in exc.violations[:5]
        ]
        message = (
            f"Không chốt được bảng kê {product_code} vì vượt tồn ở "
            f"{len(exc.violations)} lot: " + "; ".join(detail_lines)
            + ". Hãy tính lại bảng kê để cập nhật phân bổ theo tồn hiện tại."
        )
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=message,
                preserve_origin_products=True,
                fast_origin_context=True,
            ),
        )
    case = set_origin_sheet_status(case, product_code, "locked")
    update_case_record(client, case)
    invalidate_co_case_source_cache(client_id, case_id)
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(
            client_id,
            case_id,
            current_step="origin",
            case=case,
            message=f"Đã chốt bảng kê {product_code}.",
            preserve_origin_products=True,
            fast_origin_context=True,
        ),
    )


@app.get("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/substitute-candidates")
async def co_case_origin_sheet_substitute_candidates(
    client_id: str,
    case_id: str,
    product_code: str,
    material_code: str = "",
    row_index: int = -1,
    search: str = "",
    seed_hs: str = "",
    limit: int = 20,
):
    if not material_code and not search:
        raise HTTPException(status_code=400, detail="material_code or search query required")
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    target = next((p for p in case.get("products", []) if str(p.get("code") or "").strip() == product_code), None)
    if target is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    sheet_state = (case.get("origin_sheet_states") or {}).get(product_code, {}) or {}
    optimization_mode = sheet_state.get("optimization_mode") or "max_lvc"

    # Empty stock summary placeholder. Stock data is fetched lazily by a
    # separate /substitute-stock endpoint so the recommendation list can
    # render in ~300ms (one Data Hub call) instead of waiting for full
    # BCCT pagination (~30s for Johnson). JS merges stock async.
    def empty_stock_summary() -> dict:
        return {
            "lot_count": 0,
            "usable_lot_count": 0,
            "total_remaining_qty": "0",
            "unit_price_min": "",
            "unit_price_max": "",
            "lots": [],
            "pending": True,
        }

    candidates: list[dict] = []
    error_detail = ""
    candidates_source = "data_hub"
    if material_code:
        try:
            raw, source = portfolio_service.list_material_substitutes(
                client_id, material_code, min_score=0.5, limit=min(limit, 50)
            )
        except Exception as exc:  # noqa: BLE001
            raw, source = [], "error"
            error_detail = str(exc)
        candidates_source = source
        for row in raw:
            code = str(row.get("material_b_code") or row.get("material_code") or "").strip()
            if not code:
                continue
            candidates.append({
                "material_code": code,
                "name": row.get("name", ""),
                "category": row.get("category", ""),
                "hs_code": row.get("hs_code", ""),
                "score": float(row.get("combined_score") or row.get("score") or 0.0),
                "raw_scores": row.get("raw_scores") or {},
                "sources": row.get("sources") or [],
                "confirmed": bool(row.get("confirmed")),
                "stock": empty_stock_summary(),
                "kind": "recommended",
            })
        # Heuristic fallback ONLY when Data Hub had nothing: this still needs
        # the materials catalog (one Data Hub list_materials pagination, but
        # cached). Caller can opt out via ?skip_heuristic=1 to keep first call
        # fast even on substitutes-empty.
        if not candidates:
            try:
                cached_ctx = co_case_source_context_cached(client, case)
                material_rows = cached_ctx.get("material_rows") or []
            except Exception:  # noqa: BLE001
                material_rows = []
            heuristic, hs_seed = compute_substitute_heuristic_candidates(
                client_id, material_code, material_rows, lambda _code: empty_stock_summary(),
                fallback_hs=seed_hs,
            )
            candidates_source = "co_heuristic"
            if source == "data_hub_unauthorized":
                error_detail = (
                    "Data Hub trả 401/403 (token thiếu scope hub:read?) — fallback HS-prefix "
                    f"heuristic từ {len(material_rows)} NVL trong catalog."
                )
            else:
                error_detail = (
                    f"Data Hub không có substitute precomputed cho {material_code or '(no code)'}. "
                    f"Fallback heuristic theo HS={hs_seed or 'n/a'} — {len(heuristic)} ứng viên."
                )
            candidates = heuristic
            candidates.sort(key=lambda item: -item.get("score", 0.0))

    search_results: list[dict] = []
    if search:
        raw_search: list[dict] = []
        search_error = ""
        try:
            raw_search = portfolio_service.search_materials(client_id, search, limit=min(limit, 50))
        except Exception as exc:  # noqa: BLE001
            search_error = str(exc)
        fallback_rows = search_case_material_rows(case, search, limit=min(limit, 50))
        seen_search_codes: set[str] = set()
        for row in [*raw_search, *fallback_rows]:
            code = str(row.get("material_code") or row.get("internal_code") or "").strip()
            if not code or code in seen_search_codes:
                continue
            seen_search_codes.add(code)
            search_results.append({
                "material_code": code,
                "name": row.get("name") or row.get("material_description") or "",
                "category": row.get("category", ""),
                "hs_code": row.get("hs_code", ""),
                "score": 0.0,
                "stock": empty_stock_summary(),
                "kind": "search",
            })
            if len(search_results) >= max(1, min(limit, 50)):
                break
        if not raw_search and fallback_rows and not error_detail:
            error_detail = (
                "Data Hub material catalog search unavailable or empty; showing matching NVL "
                "already present in this dossier."
            )
        elif search_error and not error_detail:
            error_detail = search_error

    # Initial sort by score only — re-sorted client-side once stock arrives.
    candidates.sort(key=lambda item: -item.get("score", 0.0))
    return JSONResponse({
        "ok": True,
        "product_code": product_code,
        "material_code": material_code,
        "row_index": row_index,
        "optimization_mode": optimization_mode,
        "candidates": candidates,
        "candidates_source": candidates_source,
        "search_results": search_results,
        "stock_pending": True,
        "stock_url": (
            f"/clients/{client_id}/co-case/{case_id}/origin/sheet/{quote(product_code, safe='')}/substitute-stock"
        ),
        "error": error_detail,
    })


def search_case_material_rows(case: dict, query: str, limit: int = 20) -> list[dict]:
    if not (query or "").strip():
        return []
    max_rows = max(1, min(limit, 100))
    scored: list[tuple[float, int, dict]] = []
    seen: set[str] = set()
    position = 0
    for product in case.get("products", []) or []:
        for material in product.get("materials", []) or []:
            code = str(
                material.get("material_code")
                or material.get("internal_material_code")
                or material.get("internal_code")
                or ""
            ).strip()
            if not code or code in seen:
                continue
            score = material_search.match_score(query, [
                code,
                material.get("internal_code"),
                material.get("internal_material_code"),
                material.get("material_description"),
                material.get("name"),
                material.get("hs_code") or material.get("import_hs"),
            ])
            if score is None:
                continue
            seen.add(code)
            scored.append((score, position, {
                "material_code": code,
                "internal_code": material.get("internal_code") or material.get("internal_material_code") or code,
                "name": material.get("name") or material.get("material_description") or "",
                "material_description": material.get("material_description") or material.get("name") or "",
                "category": material.get("category", ""),
                "hs_code": material.get("hs_code") or material.get("import_hs") or "",
            }))
            position += 1
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [row for _, _, row in scored[:max_rows]]


@app.get("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/substitute-stock")
async def co_case_origin_sheet_substitute_stock(
    client_id: str,
    case_id: str,
    product_code: str,
    codes: str = "",
):
    """Lazy stock-summary endpoint. Returns stock pool entries for the given
    comma-separated material codes. Prefers Data Hub `bcct/by-codes` for a
    narrow lookup (~500ms); falls back to TTL-cached source_context for the
    file-based service.
    """
    requested = [code.strip() for code in (codes or "").split(",") if code.strip()]
    if not requested:
        return JSONResponse({"ok": True, "stock": {}})
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    sheet_state = (case.get("origin_sheet_states") or {}).get(product_code, {}) or {}
    optimization_mode = sheet_state.get("optimization_mode") or "max_lvc"
    cached_matches = case.get("source_invoice_matches") if isinstance(case.get("source_invoice_matches"), list) else []
    client_config = portfolio_service.get_client_config(client) if hasattr(portfolio_service, "get_client_config") else {}
    min_gap = effective_min_gap_days(client, client_config)
    # Substitute feasibility only needs CO stock (tồn) lots per candidate, not
    # raw BCCT — read solely from the materialized CO-stock snapshot. This is
    # the single source of truth: a flaky Data Hub BCCT call can never blank
    # out the suggestions, and we never silently fall back to file-store or a
    # heavy live DH pull (consistent with the no-silent-local-fallback rule).
    # A candidate code absent from the snapshot simply has no tồn. The snapshot
    # may lag BCCT; the response carries `stock_refreshed_at` so the modal shows
    # how fresh the tồn is (empty when the client has never been materialized).
    requested_set = set(requested)
    snapshot_rows = co_stock_materializer.read_co_stock_rows_cached(client_id)
    candidate_rows = [
        row for row in snapshot_rows
        if any(key in requested_set for key in co_stock_key_candidates(row))
    ]
    stock_pool = case_allocation_pool(case, cached_matches, candidate_rows, min_gap_days=min_gap)
    out: dict[str, dict] = {}
    for code in requested:
        lots = stock_pool.get(code, [])
        usable = [lot for lot in lots if co_stock_is_usable(lot)]
        total_remaining = sum(decimal_value(lot.get("remaining_qty") or lot.get("available_qty") or "0") for lot in lots)
        prices = []
        for lot in lots:
            price = decimal_value(lot.get("unit_value") or lot.get("unit_price") or "0")
            if price > 0:
                prices.append(price)
        unit_price_min = min(prices) if prices else None
        unit_price_max = max(prices) if prices else None
        out[code] = {
            "lot_count": len(lots),
            "usable_lot_count": len(usable),
            "total_remaining_qty": str(total_remaining),
            "unit_price_min": str(unit_price_min) if unit_price_min is not None else "",
            "unit_price_max": str(unit_price_max) if unit_price_max is not None else "",
            "lots": [
                {
                    "source_row": lot.get("source_row", ""),
                    "import_declaration_no": lot.get("import_declaration_no", ""),
                    "line_no": lot.get("line_no", ""),
                    "registration_date": str(lot.get("registration_date") or lot.get("declaration_date") or ""),
                    "remaining_qty": str(lot.get("remaining_qty") or lot.get("available_qty") or "0"),
                    "available_qty": str(lot.get("available_qty") or lot.get("remaining_qty") or "0"),
                    "unit_value": str(lot.get("unit_value") or lot.get("unit_price") or ""),
                    "currency": lot.get("currency", ""),
                    "material_description": lot.get("material_description", ""),
                    "hs_code": lot.get("hs_code", ""),
                    "uom": lot.get("uom", ""),
                    "allocation_code": lot.get("allocation_code", ""),
                    "eligibility_ok": bool(lot.get("_eligibility_ok", True)),
                    "eligibility_reason": str(lot.get("_eligibility_reason", "ok")),
                    "eligibility_label": co_stock_eligibility.REJECTION_LABELS.get(
                        str(lot.get("_eligibility_reason", "ok")), ""
                    ),
                }
                for lot in lots[:50]
            ],
        }
    return JSONResponse({
        "ok": True,
        "optimization_mode": optimization_mode,
        "stock": out,
        "stock_refreshed_at": co_stock_materializer.last_refresh_at(client_id),
    })


def compute_substitute_heuristic_candidates(
    client_id: str,
    seed_material_code: str,
    material_rows: list[dict],
    stock_summary,
    *,
    fallback_hs: str = "",
) -> tuple[list[dict], str]:
    """HS-prefix heuristic fallback when Data Hub substitutes endpoint returns nothing.

    Score = 0.5 base for sharing the 4-digit HS prefix, +0.2 for sharing 6-digit,
    +0.2 if the candidate has any usable CO stock lot. Tagged with source="co_heuristic"
    so the UI shows the explanation banner.

    `fallback_hs` is used when the seed material is not in Data Hub catalog (common
    when BOM uses an internal code that hasn't been catalog-resolved). Pass the
    BOM row's HS code so heuristic still has a search seed.
    """
    seed = next(
        (row for row in material_rows if str(row.get("material_code") or "").strip() == seed_material_code),
        None,
    )
    if not seed:
        try:
            seed = portfolio_service.get_material(client_id, seed_material_code)
        except Exception:  # noqa: BLE001
            seed = {}
    seed_hs = re.sub(r"\D+", "", str(seed.get("hs_code") or fallback_hs or ""))[:6]
    if not seed_hs:
        return [], ""
    seed_category = str(seed.get("category") or "").strip().lower()
    output: list[dict] = []
    for row in material_rows:
        code = str(row.get("material_code") or "").strip()
        if not code or code == seed_material_code:
            continue
        candidate_hs = re.sub(r"\D+", "", str(row.get("hs_code") or ""))[:6]
        if not candidate_hs or candidate_hs[:4] != seed_hs[:4]:
            continue
        if seed_category and str(row.get("category") or "").strip().lower() not in {seed_category, ""}:
            continue
        score = 0.5
        if candidate_hs[:6] == seed_hs[:6]:
            score = 0.7
        stock = stock_summary(code)
        if stock.get("usable_lot_count", 0) > 0:
            score += 0.2
        output.append({
            "material_code": code,
            "name": row.get("name", ""),
            "category": row.get("category", ""),
            "hs_code": row.get("hs_code", ""),
            "score": round(score, 4),
            "raw_scores": {"hs_prefix": score},
            "sources": ["co_heuristic_hs_prefix"],
            "confirmed": False,
            "stock": stock,
            "kind": "heuristic",
        })
    return output[:30], seed_hs


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/substitute-row")
async def co_case_origin_sheet_substitute_row(
    request: Request, client_id: str, case_id: str, product_code: str
):
    payload: dict = {}
    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        try:
            payload = await request.json()
        except json.JSONDecodeError:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
    else:
        form = await large_request_form(request)
        payload = {key: str(value) for key, value in form.items()}
    row_index = payload.get("row_index")
    new_material_code = str(payload.get("new_material_code") or "").strip()
    new_norm = str(payload.get("new_norm_per_unit") or "").strip()
    new_name = str(payload.get("new_name") or "").strip()
    delete = bool(payload.get("delete"))
    if row_index is None or str(row_index).strip() == "":
        raise HTTPException(status_code=400, detail="row_index required")
    raw_key = str(row_index).strip()
    is_added_key = raw_key.startswith("added_")
    if is_added_key:
        key = raw_key
        row_index_int = -1
    else:
        try:
            row_index_int = int(raw_key)
        except (TypeError, ValueError) as exc:
            raise HTTPException(status_code=400, detail="row_index must be integer or added_<n>") from exc
        key = str(row_index_int)
    if not delete and not new_material_code:
        raise HTTPException(status_code=400, detail="new_material_code required when not deleting")
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    products = case.get("products", [])
    target_index = next(
        (i for i, p in enumerate(products) if str(p.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    reject_if_sheet_locked(case, product_code)
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    overrides = dict(previous.get("material_overrides") or {})
    if delete:
        if is_added_key:
            # added rows aren't part of product.materials; "deletion" = drop the override
            # so the row disappears completely from both UI and export.
            overrides.pop(key, None)
        else:
            overrides[key] = {"deleted": True}
    elif is_added_key:
        existing = overrides.get(key) if isinstance(overrides.get(key), dict) else {}
        overrides[key] = {
            **existing,
            "added": True,
            "material_code": new_material_code,
            "norm_per_unit": new_norm,
            "name": new_name,
        }
    else:
        overrides[key] = {
            "material_code": new_material_code,
            "norm_per_unit": new_norm,
            "name": new_name,
        }
    states[product_code] = {**previous, "material_overrides": overrides, "status": "stale", "status_label": ORIGIN_SHEET_STATUS_LABELS["stale"]}
    case["origin_sheet_states"] = states
    case = mark_origin_sheets_stale(case, target_index)
    update_case_record(client, case)
    return JSONResponse({
        "ok": True,
        "product_code": product_code,
        "row_index": row_index_int if not is_added_key else key,
        "applied_override": overrides.get(key, {"deleted": True}),
        "sheet_status": "stale",
    })


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/propose-bom")
async def co_case_origin_sheet_propose_bom(
    request: Request, client_id: str, case_id: str, product_code: str
):
    payload = await read_json_or_form(request)
    actor = str(payload.get("actor") or "co_system").strip() or "co_system"
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    target = next((p for p in case.get("products", []) if str(p.get("code") or "").strip() == product_code), None)
    if target is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    state = (case.get("origin_sheet_states") or {}).get(product_code, {}) or {}
    if state.get("status") != "locked":
        raise HTTPException(status_code=409, detail="Chỉ propose được BOM mới sau khi đã chốt sheet.")
    overrides = state.get("material_overrides") or {}
    if not overrides:
        raise HTTPException(status_code=409, detail="Không có thay đổi BOM so với artifact gốc; không cần propose.")
    parent_artifact_id = str(target.get("bom_product_artifact_id") or target.get("bom_product_version_id") or "").strip()
    if not parent_artifact_id:
        raise HTTPException(status_code=409, detail="Sheet chưa gắn BOM artifact gốc; không thể propose BOM mới.")
    bom_product_code = str(target.get("bom_product_code") or target.get("code") or "").strip()
    rows = build_bom_proposal_rows(target, overrides)
    if not rows:
        raise HTTPException(status_code=409, detail="Không có dòng NVL nào để propose.")
    try:
        result = portfolio_service.submit_bom_proposal(
            client_id,
            bom_product_code,
            parent_artifact_id=parent_artifact_id,
            rows=rows,
            context={
                "case_id": case_id,
                "case_code": case.get("case_code", ""),
                "sheet_product_code": product_code,
                "diff_summary": {
                    "added": state.get("material_diff_added", 0),
                    "removed": state.get("material_diff_removed", 0),
                    "replaced": state.get("material_diff_replaced", 0),
                    "norm_only": state.get("material_diff_norm_only", 0),
                },
            },
            actor=actor,
        )
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=502, detail=f"Data Hub propose failed: {exc}") from exc
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    states[product_code] = {
        **previous,
        "proposed_artifact_id": str(result.get("artifact_id") or result.get("proposal_id") or ""),
        "proposed_proposal_id": str(result.get("proposal_id") or ""),
        "proposed_status": str(result.get("status") or "submitted"),
    }
    case["origin_sheet_states"] = states
    update_case_record(client, case)
    return JSONResponse({
        "ok": True, "product_code": product_code,
        "proposal": {
            "artifact_id": result.get("artifact_id"),
            "proposal_id": result.get("proposal_id"),
            "status": result.get("status"),
        },
    })


def build_bom_proposal_rows(product: dict, overrides: dict) -> list[dict]:
    """Shape rows for the Data Hub BOM proposal submission.

    Field names match Data Hub's BOM row contract: qty_per_unit + uom are
    typed columns; everything else lands in the row payload jsonb. CO-internal
    overrides store the qty under `norm_per_unit` (operator-facing "định mức")
    — translate that to `qty_per_unit` at the boundary, never inside DH's
    payload. Sending `norm_per_unit` makes DH read qty as 0, which auto-rejects
    every proposal via `qty_delta_exceeds_tolerance` and leaves the qty cell
    blank in the reviewer UI.
    """
    materials = product.get("materials") or []
    output: list[dict] = []
    for index, material in enumerate(materials):
        override = overrides.get(str(index)) if isinstance(overrides.get(str(index)), dict) else {}
        if override.get("deleted"):
            continue
        material_code = override.get("material_code") or material.get("material_code") or material.get("internal_material_code")
        qty = override.get("norm_per_unit") or material.get("bom_qty_per") or "0"
        output.append({
            "material_code": str(material_code or "").strip(),
            "qty_per_unit": str(qty),
            "scrap_rate": str(material.get("bom_scrap_rate") or "0"),
            "uom": str(material.get("uom") or override.get("uom") or ""),
            "name": str(override.get("name") or material.get("material_description") or ""),
            "hs_code": str(material.get("hs_code") or override.get("hs_code") or ""),
            "source_row_index": index,
        })
    for key, value in overrides.items():
        if not key.startswith("added_") or not isinstance(value, dict):
            continue
        output.append({
            "material_code": str(value.get("material_code") or "").strip(),
            "qty_per_unit": str(value.get("norm_per_unit") or "0"),
            "scrap_rate": "0",
            "uom": str(value.get("uom") or ""),
            "name": str(value.get("name") or ""),
            "hs_code": str(value.get("hs_code") or ""),
            "source_row_index": None,
            "added": True,
        })
    return [row for row in output if row["material_code"]]


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/edit-row")
async def co_case_origin_sheet_edit_row(
    request: Request, client_id: str, case_id: str, product_code: str
):
    payload = await read_json_or_form(request)
    row_index = payload.get("row_index")
    new_norm = str(payload.get("new_norm_per_unit") or "").strip()
    if row_index is None or str(row_index).strip() == "" or not new_norm:
        raise HTTPException(status_code=400, detail="row_index and new_norm_per_unit required")
    try:
        row_index_int = int(row_index)
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="row_index must be integer") from exc
    try:
        Decimal(new_norm)
    except (InvalidOperation, ValueError) as exc:
        raise HTTPException(status_code=400, detail="new_norm_per_unit must be numeric") from exc
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    target_index = next(
        (i for i, p in enumerate(case.get("products", [])) if str(p.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    reject_if_sheet_locked(case, product_code)
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    overrides = dict(previous.get("material_overrides") or {})
    key = str(row_index_int)
    existing = overrides.get(key) if isinstance(overrides.get(key), dict) else {}
    overrides[key] = {**existing, "norm_per_unit": new_norm, "norm_edit_only": not existing.get("material_code")}
    states[product_code] = {**previous, "material_overrides": overrides, "status": "stale", "status_label": ORIGIN_SHEET_STATUS_LABELS["stale"]}
    case["origin_sheet_states"] = states
    case = mark_origin_sheets_stale(case, target_index)
    update_case_record(client, case)
    return JSONResponse({
        "ok": True, "product_code": product_code,
        "row_index": row_index_int, "applied_override": overrides[key], "sheet_status": "stale",
    })


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/add-row")
async def co_case_origin_sheet_add_row(
    request: Request, client_id: str, case_id: str, product_code: str
):
    payload = await read_json_or_form(request)
    new_material_code = str(payload.get("new_material_code") or "").strip()
    new_norm = str(payload.get("new_norm_per_unit") or "").strip()
    new_name = str(payload.get("new_name") or "").strip()
    new_uom = str(payload.get("new_uom") or "").strip()
    new_hs = str(payload.get("new_hs_code") or "").strip()
    if not new_material_code:
        raise HTTPException(status_code=400, detail="new_material_code required")
    if new_norm:
        try:
            Decimal(new_norm)
        except (InvalidOperation, ValueError) as exc:
            raise HTTPException(status_code=400, detail="new_norm_per_unit must be numeric") from exc
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    target_index = next(
        (i for i, p in enumerate(case.get("products", [])) if str(p.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    reject_if_sheet_locked(case, product_code)
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    overrides = dict(previous.get("material_overrides") or {})
    next_added = 1 + max(
        [int(k.split("_", 1)[1]) for k in overrides if k.startswith("added_") and k.split("_", 1)[1].isdigit()] + [-1]
    )
    key = f"added_{next_added}"
    overrides[key] = {
        "added": True,
        "material_code": new_material_code,
        "norm_per_unit": new_norm,
        "name": new_name,
        "uom": new_uom,
        "hs_code": new_hs,
    }
    states[product_code] = {**previous, "material_overrides": overrides, "status": "stale", "status_label": ORIGIN_SHEET_STATUS_LABELS["stale"]}
    case["origin_sheet_states"] = states
    case = mark_origin_sheets_stale(case, target_index)
    update_case_record(client, case)
    return JSONResponse({
        "ok": True, "product_code": product_code, "added_key": key,
        "applied_override": overrides[key], "sheet_status": "stale",
    })


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/save")
async def co_case_origin_sheet_save(
    request: Request, client_id: str, case_id: str, product_code: str
):
    """Batched persistence of client-side bảng kê edits.

    Accepts a JSON payload with four lists/maps:
      replaces: {row_index: {new_material_code, new_norm_per_unit, new_name, new_hs_code}}
      adds: [{key, new_material_code, new_norm_per_unit, new_name, new_uom, new_hs_code}]
      deletes: {row_index: true}
      norm_edits: {row_index: new_norm}

    Each maps to material_overrides entries the existing render path already
    consumes; the sheet is flipped to "stale" so the next /calculate (now
    "Load BOM vào Bảng Kê") re-runs allocation with these overrides.
    """
    payload = await read_json_or_form(request)
    replaces = payload.get("replaces") or {}
    adds = payload.get("adds") or []
    deletes = payload.get("deletes") or {}
    norm_edits = payload.get("norm_edits") or {}
    if not (replaces or adds or deletes or norm_edits):
        raise HTTPException(status_code=400, detail="empty payload")
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    expected_revision = str(payload.get("expected_revision") or "").strip()
    if expected_revision and expected_revision != origin_case_revision(case):
        raise HTTPException(status_code=409, detail="Origin case state changed; reload before saving.")
    case = merge_origin_action_payload(case, payload)
    target_index = next(
        (i for i, p in enumerate(case.get("products", [])) if str(p.get("code") or "").strip() == product_code),
        None,
    )
    if target_index is None:
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    reject_if_sheet_locked(case, product_code)
    states = dict(case.get("origin_sheet_states") or {})
    previous = states.get(product_code) if isinstance(states.get(product_code), dict) else {}
    overrides = dict(previous.get("material_overrides") or {})

    def _parse_row_index(value) -> int | None:
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    counts = {"replaces": 0, "adds": 0, "deletes": 0, "norm_edits": 0}

    if isinstance(replaces, dict):
        for raw_index, info in replaces.items():
            row_index = _parse_row_index(raw_index)
            if row_index is None or not isinstance(info, dict):
                continue
            new_material_code = str(info.get("new_material_code") or "").strip()
            if not new_material_code:
                continue
            norm = str(info.get("new_norm_per_unit") or "").strip()
            if norm:
                try:
                    Decimal(norm)
                except (InvalidOperation, ValueError):
                    raise HTTPException(status_code=400, detail=f"replaces[{row_index}].new_norm_per_unit must be numeric")
            overrides[str(row_index)] = {
                "material_code": new_material_code,
                "norm_per_unit": norm,
                "name": str(info.get("new_name") or "").strip(),
                "hs_code": str(info.get("new_hs_code") or "").strip(),
                "uom": str(info.get("new_uom") or "").strip(),
            }
            counts["replaces"] += 1

    if isinstance(norm_edits, dict):
        for raw_index, raw_norm in norm_edits.items():
            row_index = _parse_row_index(raw_index)
            if row_index is None:
                continue
            norm = str(raw_norm or "").strip()
            if not norm:
                continue
            try:
                Decimal(norm)
            except (InvalidOperation, ValueError):
                raise HTTPException(status_code=400, detail=f"norm_edits[{row_index}] must be numeric")
            key = str(row_index)
            existing = overrides.get(key) if isinstance(overrides.get(key), dict) else {}
            overrides[key] = {**existing, "norm_per_unit": norm, "norm_edit_only": not existing.get("material_code")}
            counts["norm_edits"] += 1

    if isinstance(deletes, dict):
        for raw_index, flag in deletes.items():
            if not flag:
                continue
            row_index = _parse_row_index(raw_index)
            if row_index is None:
                continue
            overrides[str(row_index)] = {"deleted": True}
            counts["deletes"] += 1

    used_added_ids = [
        int(k.split("_", 1)[1])
        for k in overrides
        if k.startswith("added_") and k.split("_", 1)[1].isdigit()
    ]
    next_added = (max(used_added_ids) + 1) if used_added_ids else 0
    if isinstance(adds, list):
        for entry in adds:
            if not isinstance(entry, dict):
                continue
            new_material_code = str(entry.get("new_material_code") or "").strip()
            if not new_material_code:
                continue
            norm = str(entry.get("new_norm_per_unit") or "").strip()
            if norm:
                try:
                    Decimal(norm)
                except (InvalidOperation, ValueError):
                    raise HTTPException(status_code=400, detail=f"adds[{new_material_code}].new_norm_per_unit must be numeric")
            key = f"added_{next_added}"
            next_added += 1
            overrides[key] = {
                "added": True,
                "material_code": new_material_code,
                "norm_per_unit": norm,
                "name": str(entry.get("new_name") or "").strip(),
                "uom": str(entry.get("new_uom") or "").strip(),
                "hs_code": str(entry.get("new_hs_code") or "").strip(),
            }
            counts["adds"] += 1

    states[product_code] = {
        **previous,
        "material_overrides": overrides,
        "status": "calculated",
        "status_label": ORIGIN_SHEET_STATUS_LABELS["calculated"],
    }
    case["origin_sheet_states"] = states
    try:
        client_config_for_rule = (
            portfolio_service.get_client_config(client)
            if hasattr(portfolio_service, "get_client_config") else {}
        )
    except Exception:  # noqa: BLE001
        client_config_for_rule = {}
    case = recalculate_origin_sheet_edits(
        client, case, product_code,
        min_gap_days=effective_min_gap_days(client, client_config_for_rule),
    )
    case = set_origin_sheet_status(case, product_code, "calculated")
    case = mark_origin_sheets_stale(case, target_index + 1)
    update_case_record(client, case)
    return JSONResponse({
        "ok": True,
        "product_code": product_code,
        "operations": counts,
        "override_count": len(overrides),
        "sheet_status": "calculated",
        "revision": origin_case_revision(case),
        "origin_product_order": origin_product_order(case),
        "origin_sheet_states": json_safe(case.get("origin_sheet_states", {})),
    })


async def read_json_or_form(request: Request) -> dict:
    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        try:
            payload = await request.json()
        except json.JSONDecodeError:
            payload = {}
        return payload if isinstance(payload, dict) else {}
    form = await large_request_form(request)
    return {key: str(value) for key, value in form.items()}


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/recommendation-override")
async def co_case_origin_sheet_recommendation_override(
    request: Request, client_id: str, case_id: str, product_code: str
):
    payload: dict = {}
    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        try:
            payload = await request.json()
        except json.JSONDecodeError:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
    else:
        form = await large_request_form(request)
        payload = {key: str(value) for key, value in form.items()}
    overrides: dict = {}
    if "form_override" in payload:
        form_override = str(payload.get("form_override") or "").strip()
        if form_override and form_override not in {row["form_code"] for row in load_co_form_config().get("forms", [])}:
            raise HTTPException(status_code=400, detail=f"Unknown form_code {form_override}")
        overrides["form_override"] = form_override
    if "criteria_override" in payload:
        overrides["criteria_override"] = str(payload.get("criteria_override") or "").strip()
    if "lvc_threshold_override" in payload:
        overrides["lvc_threshold_override"] = payload.get("lvc_threshold_override")
    if "rvc_threshold_override" in payload:
        overrides["rvc_threshold_override"] = payload.get("rvc_threshold_override")
    if "currency_mode" in payload:
        mode = str(payload.get("currency_mode") or "").strip().lower()
        if mode and mode not in SHEET_CURRENCY_MODES:
            raise HTTPException(status_code=400, detail=f"Unknown currency_mode {mode}")
        overrides["currency_mode"] = mode or "native"
    if "optimization_mode" in payload:
        mode = str(payload.get("optimization_mode") or "").strip().lower()
        if mode and mode not in SHEET_OPTIMIZATION_MODES:
            raise HTTPException(status_code=400, detail=f"Unknown optimization_mode {mode}")
        overrides["optimization_mode"] = mode or "max_lvc"
    client = resolve_client(client_id)
    case = persisted_origin_case(client, case_id)
    expected_revision = str(payload.get("expected_revision") or "").strip()
    if expected_revision and expected_revision != origin_case_revision(case):
        raise HTTPException(status_code=409, detail="Origin case state changed; reload before saving.")
    case = merge_origin_action_payload(case, payload)
    if not any(str(p.get("code") or "").strip() == product_code for p in case.get("products", [])):
        raise HTTPException(status_code=404, detail=f"Sheet {product_code} not found in case")
    case = set_origin_sheet_config_override(case, product_code, overrides)
    update_case_record(client, case)
    state = (case.get("origin_sheet_states") or {}).get(product_code, {})
    return JSONResponse({
        "ok": True,
        "product_code": product_code,
        "state": {
            "form_override": state.get("form_override", ""),
            "criteria_override": state.get("criteria_override", ""),
            "lvc_threshold_override": state.get("lvc_threshold_override", ""),
            "rvc_threshold_override": state.get("rvc_threshold_override", ""),
            "currency_mode": state.get("currency_mode", "native"),
            "optimization_mode": state.get("optimization_mode", "max_lvc"),
            "effective_form_code": state.get("effective_form_code", ""),
            "effective_criteria_text": state.get("effective_criteria_text", ""),
            "effective_lvc_threshold": state.get("effective_lvc_threshold", ""),
            "effective_rvc_threshold": state.get("effective_rvc_threshold", ""),
        },
    })


@app.post("/clients/{client_id}/co-case/{case_id}/origin/sheet/{product_code}/reopen", response_class=HTMLResponse)
async def reopen_co_case_origin_sheet(request: Request, client_id: str, case_id: str, product_code: str):
    client = resolve_client(client_id)
    case, _payload = await origin_case_from_request(request, client, case_id)
    action_error = origin_sheet_action_error(case, product_code, "reopen")
    if action_error:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=action_error,
                preserve_origin_products=True,
                fast_origin_context=True,
            ),
        )
    # Release ledger claims BEFORE updating case state so a DB failure leaves
    # the sheet in a consistent locked state (claims still held, sheet still
    # locked). The previous order (case state first, then release) meant a
    # release failure silently leaked the claim while the UI showed unlocked.
    try:
        co_stock_ledger.record_sheet_release(client_id, case_id, product_code)
    except Exception as exc:  # noqa: BLE001
        logging.getLogger(__name__).warning(
            "reopen failed for %s/%s/%s: %s", client_id, case_id, product_code, exc
        )
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=409,
            context=co_case_context(
                client_id,
                case_id,
                current_step="origin",
                case=case,
                error=f"Không mở chốt được bảng kê {product_code}: ledger lỗi ({exc}). Hãy thử lại.",
                preserve_origin_products=True,
                fast_origin_context=True,
            ),
        )
    case = set_origin_sheet_status(case, product_code, "calculated")
    update_case_record(client, case)
    invalidate_co_case_source_cache(client_id, case_id)
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(
            client_id,
            case_id,
            current_step="origin",
            case=case,
            message=f"Đã mở chốt bảng kê {product_code}.",
            preserve_origin_products=True,
            fast_origin_context=True,
        ),
    )


@app.post("/clients/{client_id}/evaluate", response_class=HTMLResponse)
async def evaluate(request: Request, client_id: str):
    form = await large_request_form(request)
    case = update_products_from_form({key: str(value) for key, value in form.items()})
    case_id = case.get("persisted_case_id", "")
    if case_id:
        client = resolve_client(client_id)
        lock_result = acquire_origin_calculation_lock(client, case_id, origin_lock_actor(request))
        if not lock_result["acquired"]:
            return templates.TemplateResponse(
                request=request,
                name="co_case.html",
                status_code=409,
                context=co_case_context(
                    client_id,
                    case=case,
                    current_step="origin",
                    error=f"Chưa thể tính lại: hồ sơ {lock_result['lock'].get('case_code') or lock_result['lock'].get('case_id')} đang giữ phiên tính tồn cho khách hàng này.",
                    origin_calculation_blocked=True,
                    preserve_origin_products=True,
                ),
            )
    context = co_case_context(
        client_id,
        case=case,
        current_step="origin",
        message="Đã tính lại theo dữ liệu đang sửa.",
        preserve_origin_products=False,
    )
    if context["case"].get("persisted_case_id"):
        try:
            update_case_record(
                resolve_client(client_id),
                case if context.get("origin_demo_active") else context["case"],
            )
        except KeyError:
            pass
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=context,
    )


@app.post("/clients/{client_id}/upload", response_class=HTMLResponse)
async def upload_workbook(request: Request, client_id: str, file: UploadFile = File(...)):
    client = resolve_client(client_id)
    content = await file.read()
    try:
        case = parse_input_workbook(content, source_label=f"Upload: {file.filename}")
    except WorkbookParseError as exc:
        return templates.TemplateResponse(
            request=request,
            name="co_case.html",
            status_code=400,
            context=co_case_context(client_id, error=str(exc)),
        )
    case["customer"] = client.get("legal_name") or client["name"]
    case["customer_legal_name"] = client.get("legal_name", "")
    case["customer_tax_code"] = client.get("tax_code", "")
    return templates.TemplateResponse(
        request=request,
        name="co_case.html",
        context=co_case_context(client_id, case=case, current_step="origin", message=f"Đã parse {file.filename}."),
    )


@app.get("/clients/{client_id}/demo-input.xlsx")
async def download_demo_input(client_id: str):
    content = create_input_workbook(client_case(resolve_client(client_id)))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{client_id}-demo-input.xlsx"'},
    )


@app.post("/clients/{client_id}/export")
async def export_evidence(request: Request, client_id: str):
    resolve_client(client_id)
    form = await large_request_form(request)
    case = update_products_from_form({key: str(value) for key, value in form.items()})
    content = create_evidence_workbook(case)
    filename = f"{case['case_code'] or 'co-case'}-evidence.xlsx"
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
