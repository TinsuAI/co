from __future__ import annotations

import inspect
from pathlib import Path
from typing import Any

from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.app_state_store import get_app_state_store
from app.client_config_store import get_client_config as load_client_config
from app.client_config_store import save_client_config as persist_client_config
from app.co_case_store import match_case_bcct_exports
from app.data_hub_client import DataHubPortfolioService, current_data_hub_token, data_hub_client_from_env
from app.data_hub_settings import data_hub_link_settings
from app.material_search import rank_matches
from app.demo_data import get_client as seed_get_client
from app.demo_data import get_clients as seed_get_clients
from app.source_index_store import get_source_index_store, rebuild_source_index_if_configured
from app.source_postgres_store import get_source_write_store
from app.source_store import (
    _safe_customs_fx_rows,
    co_stock_rows_from_bcct,
    create_bcct_template_workbook,
    create_material_catalog_template_workbook,
    create_product_catalog_template_workbook,
    get_source_workspace,
    load_module_state,
    process_bcct_upload,
    process_catalog_upload,
    source_summary_from_states,
)


ROOT = Path(__file__).resolve().parent
THEME_COOKIE = "co_theme"


def theme_context(request: Request) -> dict[str, str]:
    theme = request.cookies.get(THEME_COOKIE)
    theme = theme if theme in {"light", "dark"} else "light"
    return {"theme": theme, "next_theme": "light" if theme == "dark" else "dark"}


portfolio_templates = Jinja2Templates(directory=ROOT / "templates", context_processors=[theme_context])


class PortfolioService:
    def clients(self) -> list[dict]:
        store = get_app_state_store()
        if store and store.has_clients():
            return [self.client_summary(row["id"]) for row in store.clients()]
        return [self.client_summary(row["id"]) for row in seed_get_clients()]

    def client(self, client_id: str) -> dict:
        store = get_app_state_store()
        if store and store.has_clients():
            try:
                return store.client(client_id)
            except KeyError as exc:
                raise HTTPException(status_code=404, detail=f"Unknown client: {client_id}") from exc
        try:
            return seed_get_client(client_id)
        except KeyError as exc:
            raise HTTPException(status_code=404, detail=f"Unknown client: {client_id}") from exc

    def client_summary(self, client_id: str) -> dict:
        client = self.client(client_id)
        source_summary, source_backend = self.source_summary(client)
        summary = {
            key: value
            for key, value in client.items()
            if key not in {"material_catalog", "product_catalog", "bcct_rows", "co_stock", "bom_rows"}
        }
        summary["counts"] = {
            **client.get("counts", {}),
            "materials": source_summary["material_catalog"]["published_row_count"],
            "products": source_summary["product_catalog"]["published_row_count"],
            "bcct": source_summary["bcct"]["published_row_count"],
            "co_stock": source_summary["co_stock_row_count"],
        }
        summary["source_backend"] = source_backend
        summary["source_versions"] = {
            "material_catalog": source_summary["material_catalog"].get("latest_version") or {},
            "product_catalog": source_summary["product_catalog"].get("latest_version") or {},
            "bcct": source_summary["bcct"].get("latest_version") or {},
        }
        return summary

    def get_client_config(self, client: dict) -> dict:
        store = get_app_state_store()
        if store:
            return store.get_client_config(client)
        return load_client_config(client)

    def save_client_config(self, client: dict, config: dict) -> dict:
        store = get_app_state_store()
        if store:
            return store.save_client_config(client, config)
        return persist_client_config(client, config)

    def refresh_client_indexes(self, client: dict) -> dict | None:
        return rebuild_source_index_if_configured(client)

    def source_summary(self, client: dict) -> tuple[dict, str]:
        client_config = self.get_client_config(client)
        store = get_source_index_store()
        if store and store.has_client(client["id"]):
            return store.source_summary(client["id"], client_config), "postgres"

        material = load_module_state(client, "material_catalog")
        product = load_module_state(client, "product_catalog")
        bcct = load_module_state(client, "bcct")
        return source_summary_from_states(material, product, bcct, client_config), "files"

    def source_workspace(self, client: dict) -> tuple[dict, str]:
        client_config = self.get_client_config(client)
        store = get_source_index_store()
        if store and store.has_client(client["id"]):
            return store.source_workspace(client["id"], client_config), "postgres"
        return get_source_workspace(client), "files"

    def co_case_source_context(self, client: dict, case: dict, *, skip_heavy_context: bool = False) -> dict:
        # Local store context is in-memory and cheap; skip_heavy_context (a Data Hub
        # pagination optimization) is a no-op here.
        client_config = self.get_client_config(client)
        store = get_source_index_store()
        shipment = case.get("shipment", {})
        invoice_no = shipment.get("invoice_no", "")
        export_declaration_nos = shipment.get("export_declaration_nos", [])
        if store and store.has_client(client["id"]):
            relevant_types = client_config.get("bcct", {}).get("relevant_export_declaration_types", [])
            if source_index_accepts_declaration_refs(store.match_bcct_exports):
                invoice_matches = store.match_bcct_exports(client["id"], invoice_no, relevant_types, export_declaration_nos)
            else:
                invoice_matches = store.match_bcct_exports(client["id"], invoice_no, relevant_types)
            return {
                "source_backend": "postgres",
                "source_summary": store.source_summary(client["id"], client_config),
                "invoice_matches": invoice_matches,
                "material_rows": store.catalog_rows(client["id"], "material_catalog") if hasattr(store, "catalog_rows") else [],
                "stock_rows": store.co_stock_rows(client["id"]) if hasattr(store, "co_stock_rows") else [],
            }

        material = load_module_state(client, "material_catalog")
        product = load_module_state(client, "product_catalog")
        bcct = load_module_state(client, "bcct")
        return {
            "source_backend": "files",
            "source_summary": source_summary_from_states(material, product, bcct, client_config),
            "invoice_matches": match_case_bcct_exports(
                case,
                {"bcct": {"published_rows": bcct["published_rows"]}},
                client_config,
            ),
            "material_rows": material["published_rows"],
            "stock_rows": co_stock_rows_from_bcct(
                bcct["published_rows"],
                client_config,
                customs_fx_rows=_safe_customs_fx_rows(),
            ),
        }

    def process_catalog_upload(self, client: dict, catalog_type: str, content: bytes, filename: str, upload_scope: str) -> dict:
        store = get_source_write_store()
        if store and store.has_client(client["id"]):
            return store.process_catalog_upload(client, catalog_type, content, filename, upload_scope, self.get_client_config(client))
        return process_catalog_upload(client, catalog_type, content, filename, upload_scope)

    def process_bcct_upload(self, client: dict, content: bytes, filename: str) -> dict:
        store = get_source_write_store()
        if store and store.has_client(client["id"]):
            return store.process_bcct_upload(client, content, filename, self.get_client_config(client))
        return process_bcct_upload(client, content, filename)

    def list_material_substitutes(self, client_id: str, material_code: str, **_query) -> tuple[list[dict], str]:
        return [], "no_data_hub"

    def list_bcct_by_codes(self, client_id: str, codes: list[str], **_query) -> list[dict]:
        return []

    def get_material(self, client_id: str, material_code: str) -> dict:
        return {}

    def submit_bom_proposal(self, client_id: str, product_code: str, **_kwargs) -> dict:
        raise HTTPException(
            status_code=503,
            detail="BOM proposals require Data Hub. Enable Data Hub to propose modified BOM artifacts.",
        )

    def search_materials(self, client_id: str, query: str, limit: int = 20) -> list[dict]:
        try:
            client = self.client(client_id)
        except HTTPException:
            return []
        catalog = client.get("material_catalog")
        if isinstance(catalog, list):
            rows = catalog
        elif isinstance(catalog, dict):
            rows = catalog.get("published_rows") or catalog.get("rows") or []
        else:
            rows = []
        return rank_matches(
            query,
            rows,
            lambda row: [
                row.get("material_code"),
                row.get("internal_code"),
                row.get("name"),
                row.get("hs_code"),
            ],
            limit=limit,
        )

    def material_catalog_template(self, client: dict) -> bytes:
        return create_material_catalog_template_workbook(client)

    def product_catalog_template(self, client: dict) -> bytes:
        return create_product_catalog_template_workbook(client)

    def bcct_template(self, client: dict) -> bytes:
        return create_bcct_template_workbook(client)


def source_index_accepts_declaration_refs(match_func) -> bool:
    try:
        return "export_declaration_nos" in inspect.signature(match_func).parameters
    except (TypeError, ValueError):
        return False


class SourceBackendUnavailable(RuntimeError):
    """Data Hub is the source of truth but is not configured/reachable, and the
    local file-store fallback is not explicitly allowed. Surfaced as 503 so the
    operator sees a clear error instead of a silently-empty page built from
    stale local backup data."""


def get_portfolio_service() -> PortfolioService | DataHubPortfolioService:
    data_hub_client = data_hub_client_from_env(token_provider=current_data_hub_token)
    if data_hub_client:
        return DataHubPortfolioService(data_hub_client)
    # No Data Hub client means DATA_HUB_ENABLED is off. Only fall back to the
    # local file-store backend when explicitly allowed (tests / offline dev via
    # CO_ALLOW_LOCAL_SOURCE=1). In a real deployment this raises so the request
    # fails loudly rather than quietly serving local backup data.
    if data_hub_link_settings().allow_local_source:
        return PortfolioService()
    raise SourceBackendUnavailable(
        "Data Hub chưa được cấu hình (DATA_HUB_ENABLED). CO không dùng dữ liệu "
        "local backup; hãy bật Data Hub hoặc đặt CO_ALLOW_LOCAL_SOURCE=1 cho môi "
        "trường dev/test."
    )


_PORTFOLIO_SERVICE_CACHE: tuple[tuple, PortfolioService | DataHubPortfolioService] | None = None


def current_portfolio_service() -> PortfolioService | DataHubPortfolioService:
    global _PORTFOLIO_SERVICE_CACHE
    settings = data_hub_link_settings()
    cache_key = (
        settings.source_enabled,
        settings.allow_local_source,
        settings.data_hub_api_base_url,
        settings.api_token,
        settings.request_timeout_seconds,
    )
    if _PORTFOLIO_SERVICE_CACHE and _PORTFOLIO_SERVICE_CACHE[0] == cache_key:
        return _PORTFOLIO_SERVICE_CACHE[1]
    service = get_portfolio_service()
    _PORTFOLIO_SERVICE_CACHE = (cache_key, service)
    return service


class PortfolioServiceProxy:
    def __getattr__(self, name: str):
        return getattr(current_portfolio_service(), name)


portfolio_service = PortfolioServiceProxy()
portfolio_app = FastAPI(title="Barry Source Portfolio")


@portfolio_app.get("/", response_class=HTMLResponse)
async def portfolio_dashboard(request: Request):
    return portfolio_templates.TemplateResponse(
        request=request,
        name="portfolio.html",
        context={"clients": portfolio_service.clients()},
    )


@portfolio_app.get("/api/clients")
async def portfolio_clients() -> dict[str, list[dict]]:
    return {"clients": portfolio_service.clients()}


@portfolio_app.get("/api/clients/{client_id}")
async def portfolio_client(client_id: str) -> dict[str, dict]:
    return {"client": portfolio_service.client_summary(client_id)}


@portfolio_app.get("/api/clients/{client_id}/source-summary")
async def portfolio_source_summary(client_id: str) -> dict[str, Any]:
    client = portfolio_service.client(client_id)
    source_summary, source_backend = portfolio_service.source_summary(client)
    return {
        "client": {
            "id": client["id"],
            "name": client["name"],
            "code": client["code"],
            "tax_code": client.get("tax_code", ""),
        },
        "source_backend": source_backend,
        "source_summary": source_summary,
    }


@portfolio_app.get("/api/clients/{client_id}/source-workspace")
async def portfolio_source_workspace(client_id: str) -> dict[str, Any]:
    client = portfolio_service.client(client_id)
    source_workspace, source_backend = portfolio_service.source_workspace(client)
    return {
        "client": {
            "id": client["id"],
            "name": client["name"],
            "code": client["code"],
            "tax_code": client.get("tax_code", ""),
        },
        "source_backend": source_backend,
        "source_workspace": source_workspace,
    }


@portfolio_app.get("/api/clients/{client_id}/config")
async def portfolio_client_config(client_id: str) -> dict[str, Any]:
    client = portfolio_service.client(client_id)
    return {"client_id": client["id"], "client_config": portfolio_service.get_client_config(client)}


@portfolio_app.put("/api/clients/{client_id}/config")
async def portfolio_save_client_config(client_id: str, config: dict = Body(...)) -> dict[str, Any]:
    client = portfolio_service.client(client_id)
    saved = portfolio_service.save_client_config(client, config)
    portfolio_service.refresh_client_indexes(client)
    return {"client_id": client["id"], "client_config": saved}
