import hashlib
import html
import json
import os
import re
from copy import deepcopy
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import httpx
import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.bom_store import get_bom_workspace
from app.co_case_store import MAX_SUPPORTING_FILE_BYTES, get_case_record, match_case_bcct_exports, update_case_record
from app.client_config_store import (
    get_client_config,
    resolve_allocation_code,
    save_client_config,
)
from app.demo_data import get_client, get_client_case, get_clients, get_demo_case, update_products_from_form
from app.main import app
from app.origin import calculate_rvc, evaluate_tariff_shift
from app.source_store import (
    create_bcct_template_workbook,
    create_material_catalog_template_workbook,
    create_product_catalog_template_workbook,
    get_source_summary,
    get_source_workspace,
    parse_bcct_workbook,
    parse_catalog_workbook,
    process_bcct_upload,
    process_catalog_upload,
)
from app.source_index_store import (
    build_bcct_index_records,
    build_catalog_index_records,
    build_source_state_records,
    PostgresSourceIndexStore,
    source_state_from_workspace,
)
from app.material_search import fold_text, match_score, rank_matches
from app.table_view import build_table_view
from app.workbook_io import create_evidence_workbook, create_input_workbook, parse_input_workbook


@pytest.fixture(autouse=True)
def isolate_bom_store(tmp_path, monkeypatch):
    monkeypatch.setenv("BOM_STORE_ROOT", str(tmp_path / "bom-store"))
    monkeypatch.setenv("SOURCE_STORE_ROOT", str(tmp_path / "source-store"))
    monkeypatch.setenv("CLIENT_CONFIG_ROOT", str(tmp_path / "client-config"))
    monkeypatch.setenv("CO_CASE_STORE_ROOT", str(tmp_path / "co-case-store"))
    monkeypatch.setenv("CUSTOMS_FX_STORE_ROOT", str(tmp_path / "customs-fx-store"))
    monkeypatch.setenv("CO_FORM_CONFIG_PATH", str(tmp_path / "co-form-index.json"))


def workbook_bytes(workbook: Workbook) -> bytes:
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def hidden_form_data(markup: str) -> dict[str, str]:
    return {
        name: html.unescape(value)
        for name, value in re.findall(r'<input type="hidden" name="([^"]+)" value="([^"]*)"', markup)
    }


def co_form_settings_payload(config: dict) -> dict[str, str]:
    data = {
        "source_note": config["source_note"],
        "form_priority": ", ".join(config["form_priority"]),
        "form_count": str(len(config["forms"])),
        "market_count": str(len(config["market_presets"])),
        "psr_count": str(len(config.get("psr_rules", []))),
    }
    for index, form in enumerate(config["forms"]):
        prefix = f"form_{index}"
        data[f"{prefix}_enabled"] = "1" if form.get("enabled") else ""
        for key in [
            "form_code",
            "display_name",
            "agreement",
            "instrument",
            "instrument_note",
            "source_label",
            "source_url",
            "verification_status",
        ]:
            data[f"{prefix}_{key}"] = form.get(key, "")
    for index, market in enumerate(config["market_presets"]):
        prefix = f"market_{index}"
        data[f"{prefix}_enabled"] = "1" if market.get("enabled") else ""
        data[f"{prefix}_show_in_picker"] = "1" if market.get("show_in_picker") else ""
        data[f"{prefix}_market"] = market.get("market", "")
        data[f"{prefix}_label"] = market.get("label", "")
        data[f"{prefix}_form_code"] = market.get("form_code", "")
        data[f"{prefix}_aliases"] = ", ".join(market.get("aliases", []))
        data[f"{prefix}_selection_reason"] = market.get("selection_reason", "")
        data[f"{prefix}_source_label"] = market.get("source_label", "")
    for index, rule in enumerate(config.get("psr_rules", [])):
        prefix = f"psr_{index}"
        data[f"{prefix}_enabled"] = "1" if rule.get("enabled") else ""
        data[f"{prefix}_form_code"] = rule.get("form_code", "")
        data[f"{prefix}_hs_scope"] = rule.get("hs_scope", "")
        data[f"{prefix}_criteria"] = rule.get("criteria", "")
        data[f"{prefix}_source_reference"] = rule.get("source_reference", "")
        data[f"{prefix}_note"] = rule.get("note", "")
        data[f"{prefix}_status"] = rule.get("status", "")
    return data


def bcct_workbook(rows: list[dict]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "BCCT"
    worksheet.append([
        "coverage_period",
        "direction",
        "declaration_no",
        "declaration_date",
        "customs_office",
        "declaration_type",
        "line_no",
        "item_code",
        "description",
        "hs_code",
        "quantity",
        "unit",
        "customs_value",
        "currency",
        "invoice_ref",
    ])
    for row in rows:
        worksheet.append([
            row.get("coverage_period", ""),
            row["direction"],
            row["declaration_no"],
            row.get("declaration_date", ""),
            row.get("customs_office", ""),
            row.get("declaration_type", ""),
            row["line_no"],
            row["item_code"],
            row.get("description", ""),
            row.get("hs_code", ""),
            row["quantity"],
            row["unit"],
            row.get("customs_value", ""),
            row.get("currency", ""),
            row.get("invoice_ref", ""),
        ])
    return workbook_bytes(workbook)


def material_catalog_value_workbook(rows: list[dict]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "NVL"
    worksheet.append(["customs_code", "name", "unit", "hs_code", "unit_price", "origin_default", "status"])
    for row in rows:
        worksheet.append([
            row["customs_code"],
            row.get("name", ""),
            row.get("unit", ""),
            row.get("hs_code", ""),
            row.get("unit_price", ""),
            row.get("origin_default", ""),
            row.get("status", "active"),
        ])
    return workbook_bytes(workbook)


def customs_zip_entry(filename: str) -> bytes:
    archive_path = Path("temp/drive-download-20260428T145419Z-3-001.zip")
    if not archive_path.exists():
        pytest.skip("Customs sample ZIP is local-only.")
    with ZipFile(archive_path) as archive:
        return archive.read(filename)


def test_calculates_growatt_documented_rvc_seed_from_material_rows():
    case = get_demo_case()

    assert case["products"][0]["result"].rvc.percentage == Decimal("35.88")
    assert case["products"][1]["result"].rvc.percentage == Decimal("40.19")
    assert case["summary"]["passed_products"] == 2


def test_evaluates_ctsh_by_six_digit_subheading():
    result = evaluate_tariff_shift("8504.40", ["8542.39", "8536.90"], "CTSH")

    assert result.finished_key == "850440"
    assert result.input_keys == ("854239", "853690")
    assert result.passed is True


def test_form_update_recomputes_rvc_from_edited_material_value():
    case = update_products_from_form(
        {
            "customer": "Growatt",
            "case_code": "TEST",
            "product_count": "1",
            "product_0_code": "PV00.0048500",
            "product_0_name": "Model inverter PV00.0048500",
            "product_0_finished_hs": "8504.40",
            "product_0_fob": "100000",
            "product_0_rvc_threshold": "35",
            "product_0_material_count": "1",
            "product_0_material_0_source_row": "IMP-001/1",
            "product_0_material_0_hs_code": "8542.39",
            "product_0_material_0_origin_status": "non_origin",
            "product_0_material_0_available_qty": "520",
            "product_0_material_0_consumed_qty": "120",
            "product_0_material_0_non_origin_cif_value": "70000",
        }
    )

    assert case["products"][0]["result"].rvc.percentage == Decimal("30.00")
    assert case["products"][0]["result"].passed is False


def test_seed_workbook_round_trips_through_parser():
    parsed = parse_input_workbook(create_input_workbook(get_demo_case()), "roundtrip")

    assert parsed["source_label"] == "roundtrip"
    assert parsed["products"][0]["code"] == "PV00.0048500"
    assert parsed["products"][0]["materials"][0]["customs_material_code"] == "DEMO-NPL-001"
    assert parsed["products"][0]["materials"][0]["internal_material_code"] == "DEMO-NPL-001"
    assert parsed["products"][0]["result"].rvc.percentage == Decimal("35.88")


def test_evidence_workbook_is_generated():
    content = create_evidence_workbook(get_demo_case())

    assert content.startswith(b"PK")


def test_workspace_renders_interactive_controls():
    client = TestClient(app)

    response = client.get("/clients/growatt")

    assert response.status_code == 200
    assert "Growatt" in response.text
    assert "/clients/growatt/catalog" in response.text
    assert "/clients/growatt/bom" in response.text
    assert "/clients/growatt/co-stock" in response.text
    assert "/clients/growatt/bcct" in response.text
    assert "/clients/growatt/co-case" in response.text
    assert "Workspace theo công ty" in response.text


def test_theme_toggle_persists_dark_theme_cookie():
    client = TestClient(app)

    initial = client.get("/clients")
    assert initial.status_code == 200
    assert 'data-theme="light"' in initial.text

    response = client.post(
        "/settings/theme",
        data={"theme": "dark", "next_url": "/clients/growatt"},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == "/clients/growatt"
    assert "co_theme=dark" in response.headers["set-cookie"]

    themed = client.get("/clients/growatt")
    assert themed.status_code == 200
    assert 'data-theme="dark"' in themed.text
    assert 'name="theme" value="light"' in themed.text


def test_client_navigation_hides_data_modules_inside_co_workflow():
    client = TestClient(app)

    response = client.get("/clients/growatt/co-case")

    assert response.status_code == 200
    assert 'aria-label="Luồng làm C/O"' in response.text
    assert 'aria-label="Dữ liệu nền công ty"' not in response.text
    assert "Làm hồ sơ C/O" in response.text
    assert 'href="/clients/growatt/co-case">Hồ sơ C/O</a>' not in response.text
    assert "Overview" not in response.text
    assert "Danh mục mã hàng" not in response.text
    assert "/clients/growatt/catalog" not in response.text
    assert "/clients/growatt/bom" not in response.text
    assert "/clients/growatt/co-stock" not in response.text
    assert "/clients/growatt/bcct" not in response.text

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Hidden data nav case", "case_code": "NAV-HIDE", "destination_market": "EU"},
        follow_redirects=False,
    )
    detail = client.get(created.headers["location"])

    assert detail.status_code == 200
    assert 'aria-label="Dữ liệu nền công ty"' not in detail.text
    assert "Overview" not in detail.text
    assert "Danh mục mã hàng" not in detail.text

    catalog = client.get("/clients/growatt/catalog")

    assert catalog.status_code == 200
    assert 'aria-label="Dữ liệu nền công ty"' in catalog.text
    assert "Danh mục mã hàng" in catalog.text


def test_table_view_filters_searches_sorts_and_paginates():
    rows = [
        {"code": "MAT-001", "name": "Capacitor", "status": "inactive"},
        {"code": "MAT-002", "name": "Diode bridge", "status": "active"},
        {"code": "MAT-003", "name": "Small diode", "status": "active"},
    ]

    table = build_table_view(
        rows,
        columns=[
            {"key": "code", "label": "Code"},
            {"key": "name", "label": "Name"},
            {"key": "status", "label": "Status"},
        ],
        query={"q": "diode", "status": "active", "sort": "code", "dir": "desc", "page": "2", "per_page": "1"},
        filters=[{"name": "status", "field": "status", "label": "Status"}],
    )

    assert table["total_count"] == 3
    assert table["filtered_count"] == 2
    assert table["page"] == 2
    assert table["total_pages"] == 2
    assert [row["code"] for row in table["rows"]] == ["MAT-002"]


def test_table_view_clamps_page_to_available_results():
    table = build_table_view(
        [{"code": "A"}, {"code": "B"}, {"code": "C"}],
        columns=[{"key": "code", "label": "Code"}],
        query={"page": "99", "per_page": "2"},
    )

    assert table["page"] == 2
    assert table["total_pages"] == 2
    assert [row["code"] for row in table["rows"]] == ["C"]


def test_material_search_multi_token_matches_across_fields():
    # "res blue" must hit code (RES…) AND name (…blue), in any field, any order.
    assert match_score("res blue", ["RES-100", "Blue resistor"]) is not None
    assert match_score("blue res", ["RES-100", "Blue resistor"]) is not None
    # A token that matches no field excludes the row entirely.
    assert match_score("res green", ["RES-100", "Blue resistor"]) is None


def test_material_search_is_accent_insensitive():
    assert match_score("dien tro", ["RES-1", "Điện trở 1k"]) is not None
    assert fold_text("Điện trở") == "dien tro"


def test_material_search_ranks_code_match_above_incidental_name_match():
    rows = [
        {"material_code": "CAP-001", "name": "Resin coated"},   # 'res' only in name
        {"material_code": "RES-001", "name": "Standard part"},  # 'res' as code prefix
    ]
    ranked = rank_matches("res", rows, lambda r: [r["material_code"], r["name"]])
    assert [r["material_code"] for r in ranked] == ["RES-001", "CAP-001"]


def test_material_search_subsequence_tolerates_abbreviation():
    # 'dintro' is a subsequence of code 'DIENTRO.CHIP' (missing 'e') — matches.
    assert match_score("dintro", ["DIENTRO.CHIP", "Chip điện trở"]) is not None
    # Unrelated token does not.
    assert match_score("zzzq", ["DIENTRO.CHIP", "Chip điện trở"]) is None


def test_material_search_subsequence_does_not_scatter_match_long_descriptions():
    # 'aptomat' must NOT subsequence-match a long spaced IC description.
    ic = ["007.0061200", "Mạch tích hợp IC, đơn vị điều khiển. Hàng mới 100%"]
    assert match_score("aptomat", ic) is None


def test_material_search_numeric_token_requires_real_digit_run_not_subsequence():
    # '5000' scatter-matches HS '85044090' (5,0,0,0) as a subsequence — must be
    # rejected; only a real "5000" substring (in code/name/HS) should match.
    no_5000 = ["BIENTAN.05", "Thiết bị biến tần model MIN 10000TL", "85044090"]
    assert match_score("bientan 5000", no_5000) is None
    has_5000 = ["BIENTAN.18", "Thiết bị biến tần model MIN 5000TL-X2", "85044090"]
    assert match_score("bientan 5000", has_5000) is not None


def test_catalog_bom_stock_bcct_are_data_views_and_co_case_is_workflow_entry():
    client = TestClient(app)

    catalog_response = client.get("/clients/growatt/catalog")
    material_catalog_response = client.get("/clients/growatt/catalog/materials")
    product_catalog_response = client.get("/clients/growatt/catalog/products")
    bom_response = client.get("/clients/growatt/bom")
    stock_response = client.get("/clients/growatt/co-stock")
    bcct_response = client.get("/clients/growatt/bcct")
    customs_fx_response = client.get("/customs-exchange-rates")
    co_case_response = client.get("/clients/growatt/co-case")

    assert catalog_response.status_code == 200
    assert material_catalog_response.status_code == 200
    assert product_catalog_response.status_code == 200
    assert bom_response.status_code == 200
    assert stock_response.status_code == 200
    assert bcct_response.status_code == 200
    assert customs_fx_response.status_code == 200
    assert co_case_response.status_code == 200
    assert "Upload danh mục" in catalog_response.text
    assert "/clients/growatt/catalog/materials" in catalog_response.text
    assert "/clients/growatt/catalog/products" in catalog_response.text
    assert "DS NVL DK HQ" in material_catalog_response.text
    assert "DEMO-NPL-001" in material_catalog_response.text
    assert "Mã nội bộ" not in material_catalog_response.text
    assert "DS SP DK HQ" in product_catalog_response.text
    assert "PV00.0048500" in product_catalog_response.text
    assert "DEMO-NPL-001" in bom_response.text
    assert "Tồn CO khác tồn kho vật lý" in stock_response.text
    assert "BCCT nhập khẩu / xuất khẩu" in bcct_response.text
    assert "107101950210" in bcct_response.text
    assert "Tỷ giá hải quan" in customs_fx_response.text
    assert "app-level" in customs_fx_response.text
    assert "Refresh tỷ giá" in customs_fx_response.text
    assert "Quy trình làm C/O" in co_case_response.text
    assert "Tạo hoặc mở hồ sơ" in co_case_response.text
    assert "Danh sách hồ sơ C/O" in co_case_response.text
    assert "Các bước xử lý" not in co_case_response.text
    assert "Dữ liệu nền đang sẵn sàng" not in co_case_response.text
    assert "Đánh giá RVC + CTSH" not in co_case_response.text
    assert "BCCT xuất khẩu theo invoice" not in co_case_response.text
    assert "BTP" not in catalog_response.text


def test_catalog_child_route_search_limits_material_rows():
    client = TestClient(app)

    response = client.get("/clients/growatt/catalog/materials?q=DEMO-NPL-002")

    assert response.status_code == 200
    assert "DEMO-NPL-002" in response.text
    assert "DEMO-NPL-001" not in response.text
    assert "1 / 3 dòng" in response.text


def test_catalog_product_route_search_limits_product_rows():
    client = TestClient(app)

    response = client.get("/clients/growatt/catalog/products?q=PV01.0117600")

    assert response.status_code == 200
    assert "PV01.0117600" in response.text
    assert "PV00.0048500" not in response.text
    assert "1 / 2 dòng" in response.text


def test_product_catalog_does_not_render_origin_rule():
    client = TestClient(app)

    response = client.get("/clients/growatt/catalog/products")

    assert response.status_code == 200
    assert "Quy tắc" not in response.text
    assert "RVC 35% + CTSH" not in response.text


def test_product_catalog_template_does_not_include_origin_rule_column():
    workbook = load_workbook(BytesIO(create_product_catalog_template_workbook(get_client("growatt"))))
    headers = [cell.value for cell in workbook.active[1]]

    assert "Quy tắc" not in headers
    assert "origin_rule" not in headers
    assert "rule" not in headers


def test_product_catalog_upload_ignores_legacy_origin_rule_column():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["Mã", "Tên", "Đơn vị tính", "Mã HS", "Quy tắc"])
    worksheet.append(["TP-001", "Finished product", "PCS", "85044090", "RVC 35% + CTSH"])

    result = process_catalog_upload(
        get_client("do-thanh"),
        "product",
        workbook_bytes(workbook),
        "legacy-ds-sp.xlsx",
        "full_catalog",
    )

    rows = get_source_workspace(get_client("do-thanh"))["product_catalog"]["published_rows"]
    assert result["status"] == "new_version"
    assert rows[0]["product_code"] == "TP-001"
    assert "rule" not in rows[0]
    assert rows[0]["raw_fields"]["Quy tắc"] == "RVC 35% + CTSH"


def test_bcct_route_paginates_rows_server_side():
    client = TestClient(app)

    response = client.get("/clients/growatt/bcct?per_page=1&page=2")

    assert response.status_code == 200
    assert "GIN01425L031" in response.text
    assert "107101950210" not in response.text
    assert "Trang 2 / 3" in response.text


def customs_fx_payloads() -> tuple[dict, dict]:
    return (
        {
            "d": [
                {"DONG_TIEN": "USD", "TEN_DONG_TIEN": "Đô-la Mỹ"},
                {"DONG_TIEN": "JPY", "TEN_DONG_TIEN": "Yên Nhật"},
            ]
        },
        {
            "d": [
                {
                    "LOAI_NGOAI_TE": "USD",
                    "TEN_NGOAI_TE": "Đô-la Mỹ",
                    "HIEU_LUC_TU_NGAY": "27/04/2026",
                    "TY_GIA": "26.130 VNĐ",
                },
                {
                    "LOAI_NGOAI_TE": "JPY",
                    "TEN_NGOAI_TE": "Yên Nhật",
                    "HIEU_LUC_TU_NGAY": "27/04/2026",
                    "TY_GIA": "177 VNĐ",
                },
                {
                    "LOAI_NGOAI_TE": "JPY",
                    "TEN_NGOAI_TE": "Yên Nhật",
                    "HIEU_LUC_TU_NGAY": "20/04/2026",
                    "TY_GIA": "178 VNĐ",
                },
            ]
        },
    )


def test_customs_fx_parser_preserves_vietnamese_rate_format():
    from app.customs_fx_store import lookup_exchange_rate, parse_customs_exchange_rate_payloads, parse_vnd_rate_text

    rows = parse_customs_exchange_rate_payloads(*customs_fx_payloads())

    usd = [row for row in rows if row["currency_code"] == "USD"][0]
    assert parse_vnd_rate_text("26.130 VNĐ") == Decimal("26130")
    assert usd["effective_date"] == "2026-04-27"
    assert usd["rate_vnd_per_unit"] == "26130"
    assert usd["rate_display"] == "26.130 VNĐ"
    assert lookup_exchange_rate(rows, "USD", "2026-05-01")["rate_vnd_per_unit"] == "26130"


def test_customs_fx_fetch_uses_customs_public_json_endpoints():
    from app.customs_fx_store import fetch_customs_exchange_rates

    currency_payload, history_payload = customs_fx_payloads()

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        if "GetListDongTienTyGia" in url:
            return httpx.Response(200, json=currency_payload)
        if "GetListRateByNameOrDate" in url:
            payload = json.loads(request.content.decode())
            assert payload["ten_ngoai_te"] == ""
            assert payload["hieu_luc_tu_ngay"] == "01-04-2026"
            assert payload["hieu_luc_den_ngay"] == "01-05-2026"
            assert payload["captcha"] == ""
            return httpx.Response(200, json=history_payload)
        return httpx.Response(404, json={})

    rows = fetch_customs_exchange_rates(
        history_start_date="2026-04-01",
        history_end_date="2026-05-01",
        transport=httpx.MockTransport(handler),
    )

    assert {row["currency_code"] for row in rows} == {"USD", "JPY"}
    assert len([row for row in rows if row["currency_code"] == "JPY"]) == 2
    assert rows[0]["source"] == "customs.gov.vn"
    assert rows[0]["source_endpoint"] == "GetListRateByNameOrDate"


def test_customs_fx_file_store_upserts_global_rate_rows():
    from app.customs_fx_store import CUSTOMS_FX_CLIENT_ID, FileCustomsFxStore, parse_customs_exchange_rate_payloads

    rows = parse_customs_exchange_rate_payloads(*customs_fx_payloads())
    store = FileCustomsFxStore()

    first = store.save_refresh(CUSTOMS_FX_CLIENT_ID, rows)
    second = store.save_refresh(CUSTOMS_FX_CLIENT_ID, rows)

    assert first["fetched_row_count"] == 3
    assert first["saved_row_count"] == 3
    assert second["upserted_row_count"] == 0
    assert store.summary()["currency_count"] == 2
    assert store.rows()[0]["effective_date"] == "2026-04-27"


def test_customs_fx_route_refreshes_and_filters_rows(monkeypatch):
    from app import customs_fx_store as customs_fx_module
    from app.customs_fx_store import parse_customs_exchange_rate_payloads

    rows = parse_customs_exchange_rate_payloads(*customs_fx_payloads())
    monkeypatch.setattr(customs_fx_module, "fetch_customs_exchange_rates", lambda **_kwargs: rows)

    client = TestClient(app)
    refresh = client.post("/customs-exchange-rates/refresh")
    filtered = client.get("/customs-exchange-rates?currency=JPY")

    assert refresh.status_code == 200
    assert "Đã cập nhật 3 dòng tỷ giá hải quan" in refresh.text
    assert "26.130 VNĐ" in refresh.text
    assert filtered.status_code == 200
    assert "JPY" in filtered.text
    assert "26.130 VNĐ" not in filtered.text


def test_customs_fx_client_route_redirects_to_app_level_surface():
    client = TestClient(app)

    response = client.get("/clients/growatt/customs-exchange-rates", follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == "/customs-exchange-rates"


def test_co_stock_route_search_and_status_filter():
    client = TestClient(app)

    response = client.get("/clients/growatt/co-stock?q=DEMO-NPL-001&status=available")

    assert response.status_code == 200
    assert "DEMO-NPL-001" in response.text
    assert "GIN01425L031" not in response.text
    assert "Hiển thị 1-1 / 1 dòng" in response.text


def test_clients_page_is_entry_point():
    client = TestClient(app)

    response = client.get("/clients")

    assert response.status_code == 200
    assert "Danh sách công ty" in response.text
    assert "Growatt" in response.text
    assert "Johnson" in response.text
    assert "Danh mục TP" in response.text
    assert "Danh mục NVL" in response.text
    assert "BOM" in response.text
    assert "Tồn CO" in response.text
    assert "BTP" not in response.text


def test_upload_seed_workbook_runs_parser():
    client = TestClient(app)
    content = create_input_workbook(get_demo_case())

    response = client.post(
        "/clients/growatt/upload",
        files={"file": ("growatt.xlsx", content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "Đã parse growatt.xlsx" in response.text
    assert "PV01.0117600" in response.text


def test_download_seed_input_and_export_routes_return_xlsx():
    client = TestClient(app)

    seed_response = client.get("/clients/growatt/demo-input.xlsx")
    assert seed_response.status_code == 200
    assert seed_response.content.startswith(b"PK")

    export_response = client.post(
        "/clients/growatt/export",
        data={
            "customer": "Growatt",
            "case_code": "TEST",
            "product_count": "1",
            "product_0_code": "PV00.0048500",
            "product_0_name": "Model inverter PV00.0048500",
            "product_0_finished_hs": "8504.40",
            "product_0_fob": "100000",
            "product_0_rvc_threshold": "35",
            "product_0_material_count": "1",
            "product_0_material_0_hs_code": "8542.39",
            "product_0_material_0_origin_status": "non_origin",
            "product_0_material_0_available_qty": "520",
            "product_0_material_0_consumed_qty": "120",
            "product_0_material_0_non_origin_cif_value": "64120",
        },
    )
    assert export_response.status_code == 200
    assert export_response.content.startswith(b"PK")


def test_low_level_rvc_calculation_still_matches_formula():
    result = calculate_rvc("100000", "64120", "35")

    assert result.percentage == Decimal("35.88")
    assert result.passed is True


def test_client_seed_has_required_company_modules():
    growatt = [client for client in get_clients() if client["id"] == "growatt"][0]
    johnson = [client for client in get_clients() if client["id"] == "johnson"][0]
    case = get_client_case("growatt")

    assert growatt["counts"]["materials"] > 0
    assert growatt["counts"]["products"] > 0
    assert growatt["counts"]["bom_lines"] > 0
    assert growatt["counts"]["co_stock"] > 0
    assert growatt["counts"]["bcct"] > 0
    assert "boms" not in growatt["counts"]
    assert johnson["counts"]["products"] > 0
    assert johnson["counts"]["materials"] > 0
    assert johnson["counts"]["bom_lines"] > 0
    assert case["customer"] == "Growatt"


def test_bom_view_surfaces_company_source_profiles():
    client = TestClient(app)

    growatt_response = client.get("/clients/growatt/bom")
    johnson_response = client.get("/clients/johnson/bom")

    assert growatt_response.status_code == 200
    assert johnson_response.status_code == 200
    assert "Growatt multi-workbook technical BOM graph" in growatt_response.text
    assert "Johnson SAP exploded BOM export" in johnson_response.text
    assert "SAP WebAS" in johnson_response.text


def test_bom_template_download_and_no_change_upload_keeps_current_version():
    client = TestClient(app)

    template = client.get("/clients/growatt/bom/template.xlsx")
    assert template.status_code == 200
    assert template.content.startswith(b"PK")

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom.xlsx", template.content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "Không có thay đổi BOM" in response.text
    workspace = get_bom_workspace(get_client("growatt"))
    assert workspace["latest_version"]["version_no"] == 1
    assert {row["product_version_no"] for row in workspace["product_composition"]} == {1}


def test_bom_changed_upload_creates_next_version():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    worksheet = workbook["BOM"]
    worksheet["F2"] = "2.00"
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-changed.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "Đã tạo BOM composition #2" in response.text
    assert "Composition #2" in response.text
    assert "Đổi <strong>1</strong>" in response.text
    workspace = get_bom_workspace(get_client("growatt"))
    composition = {row["product_code"]: row["product_version_no"] for row in workspace["product_composition"]}
    assert workspace["latest_version"]["version_no"] == 2
    assert composition["PV00.0048500"] == 2
    assert composition["PV01.0117600"] == 1
    assert [version["product_version_no"] for version in workspace["product_versions"] if version["product_code"] == "PV00.0048500"] == [2, 1]


def test_direct_bom_full_aggregate_retires_missing_products():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    worksheet = workbook["BOM"]
    for row_index in range(worksheet.max_row, 1, -1):
        if worksheet.cell(row_index, 1).value == "PV01.0117600":
            worksheet.delete_rows(row_index)
    worksheet["F2"] = "2.00"
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-full.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "PV01.0117600 #1 -&gt; retired" in response.text
    workspace = get_bom_workspace(get_client("growatt"))
    composition = {row["product_code"]: row["product_version_no"] for row in workspace["product_composition"]}
    assert composition == {"PV00.0048500": 2}


def test_direct_bom_partial_upload_preserves_missing_products():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    worksheet = workbook["BOM"]
    for row_index in range(worksheet.max_row, 1, -1):
        if worksheet.cell(row_index, 1).value == "PV01.0117600":
            worksheet.delete_rows(row_index)
    worksheet["F2"] = "2.00"
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom", "upload_scope": "partial_product"},
        files={"file": ("growatt-bom-partial.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    workspace = get_bom_workspace(get_client("growatt"))
    composition = {row["product_code"]: row["product_version_no"] for row in workspace["product_composition"]}
    assert composition["PV00.0048500"] == 2
    assert composition["PV01.0117600"] == 1


def test_duplicate_bom_rows_are_rejected_before_versioning():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    worksheet = workbook["BOM"]
    duplicate_values = [worksheet.cell(2, column_index).value for column_index in range(1, worksheet.max_column + 1)]
    worksheet.append(duplicate_values)
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-duplicate.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 400
    assert "BOM có dòng trùng khóa" in response.text
    assert get_bom_workspace(get_client("growatt"))["latest_version"]["version_no"] == 1


def test_reupload_historical_bom_can_create_new_product_version_against_current():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx").content
    client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom.xlsx", template, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    workbook = load_workbook(BytesIO(template))
    workbook["BOM"]["F2"] = "2.00"
    changed = BytesIO()
    workbook.save(changed)
    client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-changed.xlsx", changed.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom.xlsx", template, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "PV00.0048500 #2 -&gt; #3" in response.text
    workspace = get_bom_workspace(get_client("growatt"))
    composition = {row["product_code"]: row["product_version_no"] for row in workspace["product_composition"]}
    assert workspace["latest_version"]["version_no"] == 3
    assert composition["PV00.0048500"] == 3
    assert composition["PV01.0117600"] == 1
    assert [version["product_version_no"] for version in workspace["product_versions"] if version["product_code"] == "PV00.0048500"] == [3, 2, 1]


def test_bom_company_config_can_be_updated():
    client = TestClient(app)

    response = client.post(
        "/clients/growatt/bom/config",
        data={
            "bom_profile": "johnson_sap_exploded",
            "default_import_mode": "technical_bom",
            "code_system_mode": "single_code",
        },
    )

    assert response.status_code == 200
    assert "Đã lưu cấu hình BOM" in response.text
    assert '<option value="johnson_sap_exploded" selected>' in response.text
    assert '<option value="technical_bom" selected>' in response.text
    assert '<option value="single_code" selected>' in response.text


def test_growatt_technical_bom_upload_requires_flatten_review_before_publish():
    client = TestClient(app)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Growatt"
    worksheet.append(["成品物料", "组件物料", "组件物料描述", "标准用量", "单位"])
    worksheet.append(["PV00.0048500", "GW-NVL-001", "Growatt board", "1.5", "PCE"])
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "technical_bom"},
        files={"file": ("growatt-technical.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "cần review flatten" in response.text
    assert get_bom_workspace(get_client("growatt"))["latest_version"]["version_no"] == 1


def test_growatt_technical_bom_upload_can_be_accepted_as_flat_for_demo():
    client = TestClient(app)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Growatt"
    worksheet.append(["成品物料", "组件物料", "组件物料描述", "标准用量", "单位"])
    worksheet.append(["PV00.0048500", "GW-NVL-001", "Growatt board", "1.5", "PCE"])
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "technical_bom", "accept_review_required": "on"},
        files={"file": ("growatt-technical.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "Đã tạo BOM composition #2" in response.text
    assert "GW-NVL-001" in response.text
    assert "needs_graph_flatten_review" in response.text
    workspace = get_bom_workspace(get_client("growatt"))
    composition = {row["product_code"]: row["product_version_no"] for row in workspace["product_composition"]}
    assert composition["PV00.0048500"] == 2
    assert composition["PV01.0117600"] == 1


def test_repeated_uploads_get_distinct_upload_ids():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx").content

    for _ in range(2):
        response = client.post(
            "/clients/growatt/bom/upload",
            data={"upload_mode": "direct_bom"},
            files={"file": ("growatt-bom.xlsx", template, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200

    uploads = get_bom_workspace(get_client("growatt"))["uploads"]
    upload_ids = [upload["upload_id"] for upload in uploads]
    assert len(upload_ids) == len(set(upload_ids))


def test_bom_state_prefers_postgres_store_when_available(monkeypatch):
    from app import bom_store

    saved_states = []

    class FakeBomStateStore:
        def __init__(self):
            self.state = None

        def get_state(self, client_id: str) -> dict | None:
            assert client_id == "growatt"
            return self.state

        def save_state(self, client_id: str, state: dict) -> None:
            assert client_id == "growatt"
            self.state = dict(state)
            saved_states.append(dict(state))

    fake_store = FakeBomStateStore()
    monkeypatch.setattr(bom_store, "get_bom_state_store", lambda: fake_store)

    state = bom_store.load_state(get_client("growatt"))
    bom_store.update_bom_config(get_client("growatt"), {"bom_profile": "manual_flat"})

    assert state["client_id"] == "growatt"
    assert saved_states
    assert fake_store.state["config"]["bom_profile"] == "manual_flat"
    assert not bom_store.state_path("growatt").exists()


def test_co_case_can_select_aggregate_bom_version_snapshot():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    workbook["BOM"]["F2"] = "2.00"
    stream = BytesIO()
    workbook.save(stream)
    client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-changed.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    workspace = get_bom_workspace(get_client("growatt"))
    v1 = [version for version in workspace["versions"] if version["version_no"] == 1][0]

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "BOM snapshot case", "case_code": "CO-BOM", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    get_response = client.get(f"{created.headers['location']}/origin")
    assert get_response.status_code == 200
    assert "data-bom-version-select" in get_response.text
    assert "#2 · " in get_response.text

    post_response = client.post(
        "/clients/growatt/evaluate",
        data={
            "case_id": "TEST",
            "customer": "Growatt",
            "case_code": "TEST",
            "document_count": "0",
            "product_count": "1",
            "bom_artifact_id": v1["version_id"],
            "product_0_code": "PV00.0048500",
            "product_0_name": "Model inverter PV00.0048500",
            "product_0_finished_hs": "8504.40",
            "product_0_quantity": "demo",
            "product_0_fob": "100000",
            "product_0_rvc_threshold": "35",
            "product_0_documented_result": "demo",
            "product_0_material_count": "0",
        },
    )

    assert post_response.status_code == 200
    assert "#1 · 2 dòng" in post_response.text


def test_johnson_technical_bom_upload_keeps_sap_leaf_rows_only():
    client = TestClient(app)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet.append(["Level", "Explosion level", "Component number", "Object description", "Comp. Qty (CUn)", "Component unit"])
    worksheet.append([1, 1, "ASM-001", "Assembly parent", 1, "EA"])
    worksheet.append([2, 2, "004426-00", "Screw; Flat Head; M6x1.0Px12L", 2, "EA"])
    worksheet.append([2, 2, "1000461274", "Tube; Round; 20#", 1.306, "EA"])
    stream = BytesIO()
    workbook.save(stream)

    response = client.post(
        "/clients/johnson/bom/upload",
        data={"upload_mode": "technical_bom"},
        files={"file": ("MFW0502-39.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "Đã tạo BOM composition #2" in response.text
    assert "004426-00" in response.text
    assert "1000461274" in response.text
    assert "ASM-001" not in response.text
    workspace = get_bom_workspace(get_client("johnson"))
    product_versions = [version for version in workspace["product_versions"] if version["product_code"] == "MFW0502-39"]
    assert [version["product_version_no"] for version in product_versions] == [2, 1]


def test_material_catalog_full_upload_marks_omitted_code_inactive_pending_review():
    client = get_client("growatt")
    template = create_material_catalog_template_workbook(client)
    workbook = load_workbook(BytesIO(template))
    worksheet = workbook.active
    for row_index in range(worksheet.max_row, 1, -1):
        if worksheet.cell(row_index, 2).value == "DEMO-NPL-003":
            worksheet.delete_rows(row_index)

    result = process_catalog_upload(client, "material", workbook_bytes(workbook), "ds-nvl.xlsx", "full_catalog")

    assert result["status"] == "new_version"
    assert result["summary"]["inactive_pending_review"] == 1
    workspace = get_source_workspace(client)
    omitted_row = [row for row in workspace["material_catalog"]["published_rows"] if row["customs_code"] == "DEMO-NPL-003"][0]
    assert omitted_row["status"] == "inactive_pending_review"


def test_source_upload_metadata_uses_standard_file_contract():
    client = get_client("growatt")
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "NVL"
    worksheet.append(["customs_code", "internal_code", "name", "hs_code", "unit", "role", "origin_default", "status"])
    worksheet.append(["STD-MAT-001", "STD-MAT-001", "Standard material", "8542.39", "PCS", "NVL", "Không xuất xứ", "active"])

    content = workbook_bytes(workbook)
    result = process_catalog_upload(client, "material", content, "../DS NVL chuẩn.xlsx", "partial_update")

    assert result["status"] == "new_version"
    upload = result["upload"]
    assert upload["metadata_schema_version"] == 1
    assert upload["client_id"] == "growatt"
    assert upload["module"] == "material_catalog"
    assert upload["original_filename"] == "DS NVL chuẩn.xlsx"
    assert upload["stored_filename"].endswith(".xlsx")
    assert upload["stored_filename"] != upload["original_filename"]
    assert upload["stored_path"].startswith("clients/growatt/material-catalog/uploads/")
    assert upload["storage_backend"] == "filesystem"
    assert upload["content_sha256"] == hashlib.sha256(content).hexdigest()
    assert upload["size_bytes"] == len(content)
    assert upload["file_ext"] == ".xlsx"
    assert upload["upload_scope"] == "partial_update"
    assert upload["parse_status"] == "parsed"
    assert upload["parse_error"] == ""
    assert upload["row_count"] == 1
    assert upload["snapshot_id"].startswith("snapshot-")
    assert upload["snapshot_rows_hash"]
    assert upload["created_version_id"] == result["version"]["version_id"]
    assert upload["result"] == "new_version"
    assert upload["diff_summary"] == result["summary"]
    assert result["version"]["snapshot_id"] == upload["snapshot_id"]
    assert Path(os.environ["SOURCE_STORE_ROOT"], upload["stored_path"]).exists()


def test_postgres_source_state_records_include_standard_upload_versions_and_audit():
    client = get_client("growatt")
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "NVL"
    worksheet.append(["customs_code", "internal_code", "name", "hs_code", "unit", "role", "origin_default", "status"])
    worksheet.append(["PG-MAT-001", "PG-MAT-001", "Postgres metadata material", "8542.39", "PCS", "NVL", "Không xuất xứ", "active"])

    result = process_catalog_upload(client, "material", workbook_bytes(workbook), "metadata.xlsx", "partial_update")
    state = get_source_workspace(client)["material_catalog"]
    records = build_source_state_records(client["id"], "material_catalog", state)

    upload = records["uploads"][0]
    assert upload["upload_id"] == result["upload"]["upload_id"]
    assert upload["client_id"] == "growatt"
    assert upload["module"] == "material_catalog"
    assert upload["original_filename"] == "metadata.xlsx"
    assert upload["stored_path"] == result["upload"]["stored_path"]
    assert upload["content_sha256"] == result["upload"]["content_sha256"]
    assert upload["size_bytes"] == result["upload"]["size_bytes"]
    assert upload["parse_status"] == "parsed"
    assert upload["snapshot_id"] == result["upload"]["snapshot_id"]
    assert upload["created_version_id"] == result["version"]["version_id"]
    assert upload["diff_summary"] == result["summary"]
    assert records["raw_files"][0]["upload_id"] == upload["upload_id"]
    assert records["raw_files"][0]["storage_key"] == upload["stored_path"]
    assert records["snapshots"][0]["snapshot_id"] == result["upload"]["snapshot_id"]
    assert records["snapshots"][0]["rows_hash"] == result["upload"]["snapshot_rows_hash"]
    assert records["versions"][0]["version_id"] == result["version"]["version_id"]
    assert records["versions"][0]["snapshot_id"] == result["upload"]["snapshot_id"]
    snapshot_rows = [
        row for row in records["snapshot_rows"]
        if row["snapshot_id"] == result["upload"]["snapshot_id"]
    ]
    assert snapshot_rows[0]["row_key"] == "PG-MAT-001"
    assert snapshot_rows[0]["row_index"] == 1
    assert snapshot_rows[0]["customs_code"] == "PG-MAT-001"
    assert snapshot_rows[0]["payload"]["name"] == "Postgres metadata material"
    version_rows = [
        row for row in records["version_rows"]
        if row["version_id"] == result["version"]["version_id"] and row["row_key"] == "PG-MAT-001"
    ]
    assert version_rows[0]["customs_code"] == "PG-MAT-001"
    assert version_rows[0]["payload"]["name"] == "Postgres metadata material"
    assert records["audit_events"][0]["event"].startswith("material_catalog.")
    assert records["module_state"]["latest_version_id"] == result["version"]["version_id"]


def test_postgres_source_state_records_keep_direct_upload_history_rows_without_artifacts():
    workspace = {
        "module": "material_catalog",
        "published_rows": [{"customs_code": "PG-DIRECT-001", "name": "Direct row", "unit": "PCS"}],
        "latest_version": {"version_id": "material_catalog-v1-direct", "version_no": 1},
        "versions": [
            {
                "version_id": "material_catalog-v1-direct",
                "version_no": 1,
                "source_upload_id": "upload-direct",
                "snapshot_id": "snapshot-direct",
                "row_count": 1,
                "rows_hash": "hash-direct",
                "summary": {"added": 1},
            }
        ],
        "uploads": [
            {
                "upload_id": "upload-direct",
                "original_filename": "direct.xlsx",
                "stored_filename": "upload-direct-direct.xlsx",
                "stored_path": "clients/growatt/material-catalog/uploads/upload-direct/raw/upload-direct-direct.xlsx",
                "parse_status": "parsed",
                "snapshot_id": "snapshot-direct",
                "row_count": 1,
                "result": "new_version",
            }
        ],
        "snapshot_rows": {
            "snapshot-direct": [{"customs_code": "PG-DIRECT-001", "name": "Direct row", "unit": "PCS"}],
        },
        "version_rows": {
            "material_catalog-v1-direct": [{"customs_code": "PG-DIRECT-001", "name": "Direct row", "unit": "PCS"}],
        },
        "correction_candidates": [],
        "audit_events": [],
    }

    state = source_state_from_workspace("growatt", "material_catalog", workspace)
    records = build_source_state_records("growatt", "material_catalog", state)

    assert records["snapshot_rows"][0]["snapshot_id"] == "snapshot-direct"
    assert records["snapshot_rows"][0]["row_key"] == "PG-DIRECT-001"
    assert records["version_rows"][0]["version_id"] == "material_catalog-v1-direct"
    assert records["version_rows"][0]["payload"]["name"] == "Direct row"


def test_material_catalog_partial_update_does_not_deactivate_missing_codes():
    client = get_client("growatt")
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "NVL"
    worksheet.append(["customs_code", "internal_code", "name", "hs_code", "unit", "role", "origin_default", "status"])
    worksheet.append(["DEMO-NPL-001", "DEMO-NPL-001", "Main control board - edited", "8542.39", "PCS", "NVL", "Không xuất xứ", "active"])

    result = process_catalog_upload(client, "material", workbook_bytes(workbook), "ds-nvl-partial.xlsx", "partial_update")

    assert result["status"] == "new_version"
    workspace = get_source_workspace(client)
    retained_row = [row for row in workspace["material_catalog"]["published_rows"] if row["customs_code"] == "DEMO-NPL-003"][0]
    edited_row = [row for row in workspace["material_catalog"]["published_rows"] if row["customs_code"] == "DEMO-NPL-001"][0]
    assert retained_row["status"] != "inactive_pending_review"
    assert edited_row["name"] == "Main control board - edited"


def test_bcct_arbitrary_overlap_upload_adds_new_rows_without_deleting_missing_rows():
    client = get_client("do-thanh")
    first_upload = bcct_workbook([
        {
            "coverage_period": "2026-01",
            "direction": "import",
            "declaration_no": "TK-001",
            "line_no": "1",
            "item_code": "MAT-001",
            "description": "Material 1",
            "hs_code": "8542.39",
            "quantity": "100",
            "unit": "PCS",
            "customs_value": "1000",
        }
    ])
    second_upload = bcct_workbook([
        {
            "coverage_period": "2026-02",
            "direction": "import",
            "declaration_no": "TK-002",
            "line_no": "1",
            "item_code": "MAT-002",
            "description": "Material 2",
            "hs_code": "8536.90",
            "quantity": "50",
            "unit": "PCS",
            "customs_value": "500",
        }
    ])

    process_bcct_upload(client, first_upload, "bcct-jan.xlsx")
    result = process_bcct_upload(client, second_upload, "bcct-feb.xlsx")

    assert result["status"] == "new_version"
    assert result["summary"]["added_rows"] == 1
    workspace = get_source_workspace(client)
    assert {row["transaction_key"] for row in workspace["bcct"]["published_rows"]} == {
        "import||TK-001||1||MAT-001",
        "import||TK-002||1||MAT-002",
    }


def test_bcct_reupload_identical_row_with_unit_alias_is_noop():
    client = get_client("do-thanh")
    first_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"}
    ])
    alias_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCE"}
    ])

    process_bcct_upload(client, first_upload, "bcct-a.xlsx")
    result = process_bcct_upload(client, alias_upload, "bcct-alias.xlsx")

    assert result["status"] == "no_change"
    assert result["summary"]["unchanged_rows"] == 1
    workspace = get_source_workspace(client)
    assert workspace["bcct"]["published_rows"][0]["unit"] == "PCS"
    assert workspace["bcct"]["correction_candidates"] == []


def test_bcct_changed_unit_same_key_creates_correction_candidate():
    client = get_client("do-thanh")
    first_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"}
    ])
    changed_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "KG"}
    ])

    process_bcct_upload(client, first_upload, "bcct-a.xlsx")
    result = process_bcct_upload(client, changed_upload, "bcct-conflict.xlsx")

    assert result["status"] == "review_required"
    assert result["summary"]["correction_candidates"] == 1
    workspace = get_source_workspace(client)
    assert workspace["bcct"]["published_rows"][0]["unit"] == "PCS"
    assert workspace["bcct"]["correction_candidates"][0]["incoming_row"]["unit"] == "KG"


def test_bcct_parser_infers_configured_import_types_when_direction_is_blank():
    rows = parse_bcct_workbook(bcct_workbook([
        {"direction": "", "declaration_type": "E21", "declaration_no": "GC-001", "line_no": "1", "item_code": "MAT-GC-1", "quantity": "10", "unit": "PCS"},
        {"direction": "", "declaration_type": "E23", "declaration_no": "GC-002", "line_no": "1", "item_code": "MAT-GC-2", "quantity": "20", "unit": "PCS"},
        {"direction": "", "declaration_type": "E31", "declaration_no": "SXXK-001", "line_no": "1", "item_code": "MAT-SX-1", "quantity": "30", "unit": "PCS"},
        {"direction": "", "declaration_type": "E62", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-1", "quantity": "5", "unit": "PCS"},
    ]))

    assert {row["declaration_type"]: row["direction"] for row in rows} == {
        "E21": "import",
        "E23": "import",
        "E31": "import",
        "E62": "export",
    }


def test_growatt_client_config_defaults_to_dncx_and_description_allocation():
    config = get_client_config(get_client("growatt"))

    assert config["bcct"]["eligible_import_declaration_types"] == ["E11", "E15"]
    assert config["bcct"]["relevant_export_declaration_types"] == ["E42"]
    assert config["co_stock"]["lot_policy"] == "line_level"
    assert config["allocation_code"]["strategy"] == "description_regex"
    assert config["allocation_code"]["fallback"] == "same_as_customs_code"
    assert config["config_hash"]


def test_allocation_code_resolver_handles_regex_fallback_and_ambiguity():
    config = get_client_config(get_client("growatt"))

    resolved = resolve_allocation_code(
        {"item_code": "DIENTRO", "description": "DIENTRO#&Điện trở. Hàng mới 100% (001.0001400)"},
        config,
    )
    fallback = resolve_allocation_code(
        {"item_code": "DIENTRO", "description": "DIENTRO#&Điện trở không có mã trong ngoặc"},
        config,
    )
    ambiguous = resolve_allocation_code(
        {"item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400) hoặc (001.0001500)"},
        config,
    )

    assert resolved["allocation_code"] == "001.0001400"
    assert resolved["status"] == "resolved"
    assert resolved["source"] == "description_regex"
    assert fallback["allocation_code"] == "DIENTRO"
    assert fallback["source"] == "same_as_customs_code"
    assert ambiguous["allocation_code"] == ""
    assert ambiguous["status"] == "requires_review"
    assert ambiguous["reason"] == "multiple_regex_matches"


def test_allocation_code_resolver_prefers_data_hub_material_identity_internal_code():
    config = get_client_config(get_client("do-thanh"))

    resolved = resolve_allocation_code(
        {
            "item_code": "DAYTINHIEU",
            "description": "DAYTINHIEU#&Điện trở nhiệt NTSA3153/15Kohm.Hàng mới 100% (012.0002700)",
            "material_identity": {
                "resolution_status": "resolved",
                "resolved_code": "DAYTINHIEU",
                "customs_code": "DAYTINHIEU",
                "internal_code": "012.0002700",
            },
        },
        config,
    )

    assert resolved["allocation_code"] == "012.0002700"
    assert resolved["source"] == "material_identity.internal_code"
    assert resolved["status"] == "resolved"


def test_client_config_rejects_invalid_regex():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["allocation_code"]["strategy"] = "description_regex"
    config["allocation_code"]["description_regex"] = "("

    with pytest.raises(ValueError, match="Invalid allocation code regex"):
        save_client_config(client, config)


def test_bcct_import_rows_create_immutable_co_stock_source_rows():
    client = get_client("do-thanh")
    upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS", "customs_value": "1000"},
        {"direction": "export", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-001", "quantity": "10", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    workspace = get_source_workspace(client)
    import_rows = [row for row in workspace["bcct"]["published_rows"] if row["direction"] == "import"]
    assert len(import_rows) == 1
    assert import_rows[0]["import_row_id"].startswith("import-row-")
    assert len(workspace["co_stock_rows"]) == 1
    stock_row = workspace["co_stock_rows"][0]
    assert stock_row["source_row"] == import_rows[0]["import_row_id"]
    assert stock_row["source_line_ids"] == [import_rows[0]["import_row_id"]]
    assert stock_row["import_declaration_no"] == "TK-001"
    assert stock_row["line_no"] == "1"
    assert stock_row["customs_item_code"] == "MAT-001"
    assert stock_row["allocation_code"] == "MAT-001"
    assert stock_row["material_code"] == "MAT-001"
    assert stock_row["allocation_code_source"] == "same_as_customs_code"
    assert stock_row["allocation_code_status"] == "resolved"
    assert stock_row["available_qty"] == "100"
    assert stock_row["used_qty"] == "0"
    assert stock_row["remaining_qty"] == "100"
    assert stock_row["customs_value"] == "1000"
    assert stock_row["unit_value"] == "10"
    assert stock_row["unit_value_source"] == "bcct_customs_value_per_qty"


def test_co_stock_line_level_keeps_duplicate_codes_as_separate_lots():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["allocation_code"] = {
        "strategy": "description_regex",
        "description_regex": r"\(([A-Z0-9][A-Z0-9._/-]{3,})\)",
        "fallback": "same_as_customs_code",
    }
    config["co_stock"]["lot_policy"] = "line_level"
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400)", "quantity": "100", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "2", "item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400)", "quantity": "50", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_rows = get_source_workspace(client)["co_stock_rows"]
    assert len(stock_rows) == 2
    assert [row["line_no"] for row in stock_rows] == ["1", "2"]
    assert {row["customs_item_code"] for row in stock_rows} == {"DIENTRO"}
    assert {row["allocation_code"] for row in stock_rows} == {"001.0001400"}


def test_co_stock_can_aggregate_within_one_declaration_without_losing_source_lines():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["allocation_code"] = {
        "strategy": "description_regex",
        "description_regex": r"\(([A-Z0-9][A-Z0-9._/-]{3,})\)",
        "fallback": "same_as_customs_code",
    }
    config["co_stock"]["lot_policy"] = "aggregate_by_declaration_and_allocation_code"
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400)", "quantity": "100", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "2", "item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400)", "quantity": "50", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-002", "line_no": "1", "item_code": "DIENTRO", "description": "DIENTRO#&Điện trở (001.0001400)", "quantity": "25", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_rows = get_source_workspace(client)["co_stock_rows"]
    assert len(stock_rows) == 2
    grouped = [row for row in stock_rows if row["import_declaration_no"] == "TK-001"][0]
    assert grouped["line_no"] == "1,2"
    assert grouped["allocation_code"] == "001.0001400"
    assert grouped["available_qty"] == "150"
    assert len(grouped["source_line_ids"]) == 2


def test_co_stock_aggregation_handles_thousands_separators():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["co_stock"]["lot_policy"] = "aggregate_by_declaration_and_allocation_code"
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "1,000", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "2", "item_code": "MAT-001", "quantity": "250.5", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_rows = get_source_workspace(client)["co_stock_rows"]
    assert len(stock_rows) == 1
    assert stock_rows[0]["available_qty"] == "1250.5"


def test_client_config_declaration_types_marks_excluded_import_stock_inactive():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11", "E15"]
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E13", "declaration_no": "TK-002", "line_no": "1", "item_code": "TOOL-001", "quantity": "1", "unit": "PCS"},
        {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-001", "quantity": "10", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_rows = get_source_workspace(client)["co_stock_rows"]
    assert [row["customs_item_code"] for row in stock_rows] == ["MAT-001", "TOOL-001"]
    active_row = [row for row in stock_rows if row["customs_item_code"] == "MAT-001"][0]
    inactive_row = [row for row in stock_rows if row["customs_item_code"] == "TOOL-001"][0]
    assert active_row["eligibility_status"] == "active"
    assert active_row["remaining_qty"] == "100"
    assert inactive_row["eligibility_status"] == "inactive"
    assert inactive_row["eligibility_reason"] == "excluded_by_declaration_type_config"
    assert inactive_row["available_qty"] == "1"
    assert inactive_row["remaining_qty"] == "0"


def test_config_change_deactivates_and_reactivates_existing_stock_candidate():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11", "E15"]
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E15", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"},
    ])
    process_bcct_upload(client, upload, "bcct.xlsx")

    initial_stock = get_source_workspace(client)["co_stock_rows"][0]
    source_row = initial_stock["source_row"]
    assert initial_stock["eligibility_status"] == "active"

    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    save_client_config(client, config)
    deactivated_stock = get_source_workspace(client)["co_stock_rows"][0]
    assert deactivated_stock["source_row"] == source_row
    assert deactivated_stock["eligibility_status"] == "inactive"
    assert deactivated_stock["remaining_qty"] == "0"

    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11", "E15"]
    save_client_config(client, config)
    reactivated_stock = get_source_workspace(client)["co_stock_rows"][0]
    assert reactivated_stock["source_row"] == source_row
    assert reactivated_stock["eligibility_status"] == "active"
    assert reactivated_stock["remaining_qty"] == "100"


def test_co_stock_aggregation_keeps_active_and_inactive_candidates_separate():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["co_stock"]["lot_policy"] = "aggregate_by_declaration_and_allocation_code"
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"},
        {"direction": "import", "declaration_type": "E13", "declaration_no": "TK-001", "line_no": "2", "item_code": "MAT-001", "quantity": "50", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_rows = get_source_workspace(client)["co_stock_rows"]
    assert len(stock_rows) == 2
    assert {row["eligibility_status"] for row in stock_rows} == {"active", "inactive"}
    assert sorted(row["remaining_qty"] for row in stock_rows) == ["0", "100"]


def test_co_stock_declaration_type_filter_normalizes_uploaded_case():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "", "declaration_type": "e11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    workspace = get_source_workspace(client)
    assert workspace["bcct"]["published_rows"][0]["declaration_type"] == "E11"
    assert [row["customs_item_code"] for row in workspace["co_stock_rows"]] == ["MAT-001"]
    assert workspace["co_stock_rows"][0]["eligibility_status"] == "active"


def test_unresolved_allocation_code_is_review_only_stock():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["allocation_code"] = {
        "strategy": "description_regex",
        "description_regex": r"\(([A-Z0-9][A-Z0-9._/-]{3,})\)",
        "fallback": "requires_review",
    }
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "DIENTRO", "description": "DIENTRO không có mã nội bộ", "quantity": "100", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_row = get_source_workspace(client)["co_stock_rows"][0]
    assert stock_row["allocation_code_status"] == "requires_review"
    assert stock_row["allocation_code"] == ""
    assert stock_row["material_code"] == ""
    response = TestClient(app).get("/clients/do-thanh/co-stock")
    assert response.status_code == 200
    assert "Cần review" in response.text


def test_manual_review_lot_policy_keeps_stock_rows_unusable_until_review():
    client = get_client("do-thanh")
    config = get_client_config(client)
    config["bcct"]["eligible_import_declaration_types"] = ["E11"]
    config["co_stock"]["lot_policy"] = "manual_review"
    save_client_config(client, config)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"},
    ])

    process_bcct_upload(client, upload, "bcct.xlsx")

    stock_row = get_source_workspace(client)["co_stock_rows"][0]
    assert stock_row["allocation_code_status"] == "requires_review"
    assert stock_row["allocation_code_reason"] == "manual_stock_review"
    assert stock_row["material_code"] == ""


def test_export_bcct_view_filters_to_relevant_configured_types():
    client = TestClient(app)
    config = get_client_config(get_client("do-thanh"))
    config["bcct"]["relevant_export_declaration_types"] = ["E42"]
    save_client_config(get_client("do-thanh"), config)
    upload = bcct_workbook([
        {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-CO", "quantity": "10", "unit": "PCS"},
        {"direction": "export", "declaration_type": "B11", "declaration_no": "XK-002", "line_no": "1", "item_code": "TP-NORMAL", "quantity": "20", "unit": "PCS"},
    ])
    client.post(
        "/clients/do-thanh/bcct/upload",
        files={"file": ("bcct.xlsx", upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    response = client.get("/clients/do-thanh/bcct/exports")

    assert response.status_code == 200
    assert "TP-CO" in response.text
    assert "TP-NORMAL" not in response.text


def test_config_and_bcct_child_routes_render():
    client = TestClient(app)

    config_response = client.get("/clients/growatt/config")
    imports_response = client.get("/clients/growatt/bcct/imports")
    exports_response = client.get("/clients/growatt/bcct/exports")

    assert config_response.status_code == 200
    assert imports_response.status_code == 200
    assert exports_response.status_code == 200
    assert "Cấu hình công ty" in config_response.text
    assert "E11" in config_response.text
    assert "BCCT nhập khẩu" in imports_response.text
    assert "BCCT xuất khẩu" in exports_response.text


def test_co_case_snapshots_client_config_hash():
    client = TestClient(app)
    upload = bcct_workbook([
        {"direction": "import", "declaration_type": "E11", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"}
    ])
    client.post(
        "/clients/do-thanh/bcct/upload",
        files={"file": ("bcct.xlsx", upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    created = client.post(
        "/clients/do-thanh/co-case/create",
        data={"title": "Config snapshot case", "case_code": "CO-CONFIG", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    response = client.get(f"{created.headers['location']}/review")

    # 2026-05-28: review tab redesigned to operator-facing essentials only.
    # The config-hash + "Config snapshot" debug audit details aren't shown
    # on the review page anymore. The hash still lives in case persistence
    # for traceability; this regression just confirms the page renders.
    assert response.status_code == 200
    assert "Review &amp; Xuất hồ sơ" in response.text or "Review & Xuất hồ sơ" in response.text


def test_catalog_and_bcct_routes_offer_templates_and_upload_forms():
    client = TestClient(app)

    material_template = client.get("/clients/growatt/catalog/material-template.xlsx")
    bcct_template = client.get("/clients/growatt/bcct/template.xlsx")
    catalog_page = client.get("/clients/growatt/catalog")
    bcct_page = client.get("/clients/growatt/bcct")

    assert material_template.status_code == 200
    assert material_template.content.startswith(b"PK")
    assert bcct_template.status_code == 200
    assert bcct_template.content.startswith(b"PK")
    assert "Upload danh mục" in catalog_page.text
    assert "Upload BCCT" in bcct_page.text
    assert "append_or_review_by_transaction_key" in bcct_page.text


def test_catalog_upload_renders_selected_child_table():
    client = TestClient(app)
    content = create_material_catalog_template_workbook(get_client("growatt"))

    response = client.post(
        "/clients/do-thanh/catalog/upload",
        data={"catalog_type": "material", "upload_scope": "full_catalog"},
        files={"file": ("ds-nvl.xlsx", content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "DS NVL DK HQ" in response.text
    assert "DEMO-NPL-001" in response.text
    assert "Mã nội bộ" not in response.text


def test_real_customs_material_catalog_xls_preserves_hq_schema_fields():
    rows = parse_catalog_workbook(customs_zip_entry("DANH MUC NPL DK HQ MOI.xls"), "material_catalog")

    assert rows[0]["customs_code"] == "DIOT"
    assert rows[0]["name"] == "Đi ốt"
    assert rows[0]["unit"] == "PCS"
    assert rows[0]["source_schema"] == "customs_material_catalog"
    assert rows[0]["source_row_number"] == 2
    assert rows[0]["raw_fields"]["Mã"] == "DIOT"
    assert "Mã biểu thuế NK" in rows[0]["raw_fields"]
    assert "import_tariff_code" in rows[0]


def test_real_customs_product_catalog_xls_preserves_hq_schema_fields():
    rows = parse_catalog_workbook(customs_zip_entry("DANH MUC SP DK HQ MOI.xls"), "product_catalog")

    assert rows[0]["product_code"] == "BIENTAN.01"
    assert rows[0]["hs_code"] == "85044090"
    assert rows[0]["unit"] == "PCS"
    assert rows[0]["source_schema"] == "customs_product_catalog"
    assert rows[0]["raw_fields"]["Mã định danh của lệnh SX"] == ""
    assert "production_order_identifier" in rows[0]


def test_real_customs_bcct_xlsx_reads_header_row_10_and_preserves_hq_fields():
    rows = parse_bcct_workbook(customs_zip_entry("BaoCaoHangChiTiet 01.01.2025 - 31.12.2025 08.01 or.xlsx"))

    assert len(rows) == 19898
    assert rows[0]["declaration_no"] == "106865355330"
    assert rows[0]["declaration_date"] == "2025-01-07"
    assert rows[0]["declaration_type"] == "E15"
    assert rows[0]["direction"] == "import"
    assert rows[0]["line_no"] == "1"
    assert rows[0]["item_code"] == "LKN-VO"
    assert rows[0]["quantity"] == "2000"
    assert rows[0]["unit"] == "PCS"
    assert rows[0]["origin_country"] == "VIETNAM"
    assert rows[0]["customs_value"] == "1288209000"
    assert rows[0]["invoice_ref"] == "00000020"
    assert rows[0]["raw_fields"]["Số To Khai"] == "106865355330"
    assert rows[0]["source_header_row"] == 10
    assert any(row["direction"] == "export" and row["declaration_type"] == "E42" for row in rows)


def test_bcct_upload_route_surfaces_correction_candidates():
    client = TestClient(app)
    first_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"}
    ])
    changed_upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "KG"}
    ])
    client.post(
        "/clients/do-thanh/bcct/upload",
        files={"file": ("bcct.xlsx", first_upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    response = client.post(
        "/clients/do-thanh/bcct/upload",
        files={"file": ("bcct-conflict.xlsx", changed_upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 200
    assert "correction_candidate" in response.text
    assert "TK-001" in response.text


def test_co_case_snapshots_reviewed_source_versions_without_correction_candidates():
    client = TestClient(app)
    upload = bcct_workbook([
        {"direction": "import", "declaration_no": "TK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "100", "unit": "PCS"}
    ])
    client.post(
        "/clients/do-thanh/bcct/upload",
        files={"file": ("bcct.xlsx", upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    created = client.post(
        "/clients/do-thanh/co-case/create",
        data={"title": "Source snapshot case", "case_code": "CO-SOURCE", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    response = client.get(f"{created.headers['location']}/review")

    # 2026-05-28: review tab no longer renders the source-evidence snapshot
    # (operator-facing redesign). The data still feeds export-dossier-zip
    # via case_tkx_tkn_summary; only the visual block was removed.
    assert response.status_code == 200
    assert "correction_candidate" not in response.text


def test_co_case_can_create_persisted_dossier_and_select_it():
    client = TestClient(app)

    response = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "C/O GROWATT INV-77",
            "case_code": "CO-INV-77",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-77",
            "bill_of_lading_no": "BL-77",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    location = response.headers["location"]
    assert location.startswith("/clients/growatt/co-case/")
    detail = client.get(location)
    assert detail.status_code == 200
    assert "C/O GROWATT INV-77" in detail.text
    assert "INV-77" in detail.text
    assert "BL-77" in detail.text
    index = client.get("/clients/growatt/co-case")
    assert 'aria-label="Danh sách hồ sơ C/O"' in index.text
    assert "Danh sách hồ sơ C/O" in index.text
    assert 'data-case-filter' in index.text
    assert 'data-case-search' in index.text
    assert "Trạng thái" in index.text
    assert "Tên hồ sơ" in index.text
    assert "Invoice" in index.text
    assert "B/L" in index.text
    assert "Mở" in index.text
    assert "C/O GROWATT INV-77" in index.text
    assert "CO-INV-77" in index.text
    assert "INV-77" in index.text
    assert "BL-77" in index.text
    assert "Hồ sơ lưu local: CO-INV-77" not in index.text


def test_co_case_auto_generates_case_code_and_step_status_labels():
    client = TestClient(app)

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Auto code", "destination_market": "Canada", "invoice_no": "INV-AUTO"},
        follow_redirects=False,
    )

    assert created.status_code == 303
    detail = client.get(created.headers["location"])
    index = client.get("/clients/growatt/co-case")
    assert "CO-GROWATT-INV-AUTO-" in detail.text
    assert "CO-GROWATT-INV-AUTO-" in index.text
    assert "Chưa nhập · Auto code" not in index.text
    assert "Đủ" in detail.text
    assert "Cần soát" in detail.text


def test_co_case_detail_is_split_into_workflow_step_views():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Workflow dossier",
            "case_code": "CO-WORKFLOW",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-WORKFLOW",
            "bill_of_lading_no": "BL-WORKFLOW",
        },
        follow_redirects=False,
    )
    case_url = created.headers["location"]

    shipment = client.get(case_url)
    documents = client.get(f"{case_url}/documents")
    exports = client.get(f"{case_url}/exports")
    guidance = client.get(f"{case_url}/guidance")
    origin = client.get(f"{case_url}/origin")
    review = client.get(f"{case_url}/review")

    assert shipment.status_code == 200
    assert documents.status_code == 200
    assert exports.status_code == 200
    # `guidance` was the placeholder Form & PSR step; removed per user request
    # (2026-05-28). Re-add when a real PSR confirmation flow lands.
    assert guidance.status_code == 404
    assert origin.status_code == 200
    assert review.status_code == 200
    assert "Thông tin lô hàng" in shipment.text
    assert "Chứng từ hồ sơ" not in shipment.text
    assert "Chứng từ hồ sơ" in documents.text
    assert "Bắt buộc" in documents.text
    assert "Bổ sung" in documents.text
    assert "BCCT xuất khẩu theo tham chiếu hồ sơ" in exports.text
    assert "TKX / TKN" in exports.text
    assert "TKX_CO-WORKFLOW.zip" in exports.text
    assert "TKN_CO_CO-WORKFLOW.ZIP" in exports.text
    assert "Tờ khai xuất (TKX)" in exports.text
    assert "Tờ khai nhập (TKN)" in exports.text
    assert ">Load BOM<" in origin.text
    assert "origin-config-bar" in origin.text
    assert "data-origin-recommendation-optimization" in origin.text
    assert 'role="tablist" aria-label="Sheet sản phẩm trong bảng kê"' in origin.text
    assert 'data-origin-sheet-tab' in origin.text
    assert 'data-origin-sheet-panel' in origin.text
    assert 'replaceHistory: method === "GET"' in origin.text
    assert "activeOriginProductCode" in origin.text
    assert "__originSaveInFlight" in origin.text
    assert "currentOriginUrl" in origin.text
    assert 'class="table-input"' not in origin.text
    assert "Upload và parse" not in origin.text
    assert "Tải seed XLSX" not in origin.text
    assert "Xuất evidence XLSX" not in origin.text
    assert "Xuất bảng kê HQ" in origin.text
    assert "/export-bang-ke" in origin.text
    assert 'id="origin-export-bang-ke" hx-boost="false"' in origin.text
    assert 'form="origin-export-bang-ke" data-origin-export-action' in origin.text
    assert 'form.getAttribute("hx-boost") === "false"' in origin.text
    # Step 6 was redesigned (2026-05-28): single dossier ZIP + close-case
    # button. Legacy XLSX dossier export removed.
    assert "Xuất hồ sơ .zip" in review.text
    assert "Đóng hồ sơ" in review.text
    assert f"{case_url}/documents" in shipment.text
    assert f"{case_url}/origin" in shipment.text


def test_tkx_tkn_status_uses_declaration_files_not_bcct_rows():
    from app.main import case_tkx_tkn_summary

    case = {
        "shipment": {"export_declaration_nos": ["XK-001"]},
        "products": [
            {
                "code": "TP-001",
                "origin_sheet_status": "locked",
                "materials": [
                    {
                        "material_code": "MAT-001",
                        "allocation_lines": [
                            {"import_declaration_no": "NK-001", "import_line_no": "1", "allocated_qty": "5"}
                        ],
                    }
                ],
            }
        ],
    }
    invoice_matches = [{"declaration_no": "XK-001", "line_no": "1", "declaration_type": "E42"}]
    stock_rows = [{"import_declaration_no": "NK-001"}]

    summary = case_tkx_tkn_summary(case, invoice_matches, stock_rows)

    assert summary["tkx"][0]["in_data_hub"] is False
    assert summary["tkn"][0]["in_data_hub"] is False
    assert [entry["declaration_no"] for entry in summary["missing_tkx"]] == ["XK-001"]
    assert [entry["declaration_no"] for entry in summary["missing_tkn"]] == ["NK-001"]


def test_tkx_tkn_status_marks_present_when_declaration_file_exists():
    from app.main import case_tkx_tkn_summary

    case = {
        "shipment": {"export_declaration_nos": ["XK-001"]},
        "products": [
            {
                "code": "TP-001",
                "origin_sheet_status": "locked",
                "materials": [
                    {
                        "material_code": "MAT-001",
                        "allocation_lines": [
                            {"import_declaration_no": "NK-001", "import_line_no": "1", "allocated_qty": "5"}
                        ],
                    }
                ],
            }
        ],
    }
    invoice_matches = [{"declaration_no": "XK-001", "line_no": "1", "declaration_type": "E42"}]
    file_counts = {"export": {"XK-001": 1}, "import": {"NK-001": 2}}

    summary = case_tkx_tkn_summary(case, invoice_matches, [], file_counts)

    assert summary["tkx"][0]["in_data_hub"] is True
    assert summary["tkx"][0]["file_count"] == 1
    assert summary["tkn"][0]["in_data_hub"] is True
    assert summary["tkn"][0]["file_count"] == 2
    assert summary["missing_tkx"] == []
    assert summary["missing_tkn"] == []


def test_co_case_origin_step_renders_without_origin_snapshot():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "No snapshot",
            "case_code": "CO-NO-SNAPSHOT",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-NO-SNAPSHOT",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-NO-SNAPSHOT",
            "title": "No snapshot",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-NO-SNAPSHOT"},
            "products": [{
                "code": "TP-NO-SNAPSHOT",
                "name": "No snapshot product",
                "quantity": "1",
                "unit": "PCS",
                "fob": "100",
                "currency": "USD",
                "materials": [],
            }],
            "origin_product_order": ["TP-NO-SNAPSHOT"],
        },
    )

    origin = client.get(f"/clients/growatt/co-case/{case_id}/origin")

    assert origin.status_code == 200
    assert "Chưa tính" in origin.text


def test_co_case_shipment_step_updates_metadata_without_dropping_origin_view():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Shipment edit",
            "case_code": "CO-SHIP",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-OLD",
            "bill_of_lading_no": "BL-OLD",
        },
        follow_redirects=False,
    )
    case_url = created.headers["location"]

    response = client.post(
        f"{case_url}/shipment",
        data={
            "title": "Shipment edit",
            "case_code": "CO-SHIP-NEW",
            "destination_market": "Canada",
            "invoice_no": "INV-NEW",
            "bill_of_lading_no": "BL-NEW",
            "agreement": "CPTPP",
            "co_form_type": "Form CPTPP",
            "rule": "Cần tra cứu PSR theo HS",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == case_url
    shipment = client.get(case_url)
    origin = client.get(f"{case_url}/origin")
    assert "CO-SHIP-NEW" in shipment.text
    assert "INV-NEW" in shipment.text
    assert "BL-NEW" in shipment.text
    assert "PV00.0048500" in origin.text
    assert "origin-config-bar" in origin.text


def test_co_case_origin_preloads_demo_when_case_has_no_invoice_source_data():
    client = TestClient(app)
    created = client.post(
        "/clients/do-thanh/co-case/create",
        data={"title": "Origin demo", "case_code": "CO-DEMO", "destination_market": "Canada"},
        follow_redirects=False,
    )
    assert created.status_code == 303

    origin = client.get(f"{created.headers['location']}/origin")
    review = client.get(f"{created.headers['location']}/review")

    assert origin.status_code == 200
    assert "Demo tự nạp" in origin.text
    assert "2 TP mẫu" in origin.text
    assert "4 NVL mẫu" in origin.text
    assert "PV00.0048500" in origin.text
    assert "Upload và parse" not in origin.text
    assert "Demo tự nạp" not in review.text
    assert "2 TP mẫu" not in review.text


def test_co_case_origin_builds_and_persists_invoice_bom_snapshot():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-BOM-1",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "1000",
                        "currency": "VND",
                    },
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-BOM-2",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-002",
                        "description": "Connector set",
                        "hs_code": "8536.90",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "2000",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-BOM",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "3",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-BOM",
                    }
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Invoice BOM", "case_code": "CO-BOM-INV", "destination_market": "Ấn Độ", "invoice_no": "INV-BOM"},
        follow_redirects=False,
    )
    location = created.headers["location"]
    case_id = location.rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "shipment": {"invoice_no": "INV-BOM", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "STALE-TP",
                    "materials": [{"material_code": "STALE-MAT", "unit_value": ""}],
                }
            ],
            "origin_snapshot": {"source": "invoice_bcct_bom", "invoice_no": "INV-BOM"},
        },
    )

    origin = client.get(f"{location}/origin")

    assert origin.status_code == 200
    assert "Demo tự nạp" not in origin.text
    assert "STALE-MAT" not in origin.text
    assert "PV00.0048500" in origin.text
    form_data = hidden_form_data(origin.text)
    assert form_data["product_0_material_count"] == "0"
    assert form_data["product_0_origin_sheet_status"] == "draft"
    assert form_data["product_0_fob"] == "1000"
    assert form_data["product_0_currency"] == "VND"
    assert "1,000" in origin.text
    assert "VND" in origin.text
    assert 'name="product_0_bom_product_artifact_id"' in origin.text
    assert "DEMO-NPL-001" not in origin.text
    assert "<th>Tờ khai nhập</th>" not in origin.text
    assert "<th>Tồn CO</th>" not in origin.text
    assert "<th>Còn lại</th>" not in origin.text
    assert '<th data-origin-column="bom">Định mức</th>' in origin.text
    assert '<th data-origin-column="consumed">Lượng dùng</th>' in origin.text
    assert '<th data-origin-column="unit-value">Đơn giá</th>' in origin.text
    assert '<th data-origin-column="material-value">Trị giá NVL</th>' in origin.text
    assert '<th data-origin-column="non-origin">Trị giá KXX/VNM</th>' in origin.text

    calculated = client.post(
        f"{location}/origin/sheet/PV00.0048500/calculate",
        json={
            "origin_product_order": ["PV00.0048500"],
            "products": [
                {
                    "code": "PV00.0048500",
                    "bom_product_code": form_data["product_0_bom_product_code"],
                    "bom_product_artifact_id": form_data.get("product_0_bom_product_artifact_id", ""),
                    "name": form_data["product_0_name"],
                    "finished_hs": form_data["product_0_finished_hs"],
                    "quantity": form_data["product_0_quantity"],
                    "unit": form_data["product_0_unit"],
                    "currency": form_data["product_0_currency"],
                    "fob": form_data["product_0_fob"],
                    "rvc_threshold": form_data["product_0_rvc_threshold"],
                    "origin_sheet_status": form_data["product_0_origin_sheet_status"],
                    "origin_sheet_status_label": form_data["product_0_origin_sheet_status_label"],
                }
            ],
            "mark_stale": False,
        },
    )
    calculated_form = hidden_form_data(calculated.text)
    persisted = client.get(f"{location}/origin")

    assert calculated.status_code == 200
    assert "Đã load BOM vào bảng kê PV00.0048500" in calculated.text
    assert 'name="product_0_material_0_material_code" value="DEMO-NPL-001"' in calculated.text
    assert calculated_form["product_0_material_0_consumed_qty"] == "3"
    assert calculated_form["product_0_material_0_unit_value"] == "10"
    assert calculated_form["product_0_material_0_currency"] == "VND"
    assert calculated_form["product_0_material_0_material_value"] == "30"
    assert calculated_form["product_0_material_1_material_value"] == "105"
    assert calculated_form["product_0_vnm_value"] == "135"
    assert calculated_form["product_0_lvc_percentage"] == "86.50"
    assert "86.50%" in calculated.text
    assert "Đạt LVC" in calculated.text
    assert "DEMO-NPL-001" in persisted.text
    assert "#1 · 2 dòng" in persisted.text


def test_co_case_origin_switches_product_bom_version_from_dropdown():
    client = TestClient(app)
    template = client.get("/clients/growatt/bom/template.xlsx")
    workbook = load_workbook(BytesIO(template.content))
    workbook["BOM"]["F2"] = "2.00"
    stream = BytesIO()
    workbook.save(stream)
    client.post(
        "/clients/growatt/bom/upload",
        data={"upload_mode": "direct_bom"},
        files={"file": ("growatt-bom-v2.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    versions = get_bom_workspace(get_client("growatt"))["product_version_options_by_code"]["PV00.0048500"]
    v1 = [version for version in versions if version["product_version_no"] == 1][0]
    v2 = [version for version in versions if version["product_version_no"] == 2][0]
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {"direction": "import", "declaration_type": "E11", "declaration_no": "NK-BOM-SWITCH-1", "line_no": "1", "item_code": "DEMO-NPL-001", "description": "Main control board", "hs_code": "8542.39", "quantity": "100", "unit": "PCE", "customs_value": "1000", "currency": "VND"},
                    {"direction": "import", "declaration_type": "E11", "declaration_no": "NK-BOM-SWITCH-2", "line_no": "1", "item_code": "DEMO-NPL-002", "description": "Connector set", "hs_code": "8536.90", "quantity": "100", "unit": "PCE", "customs_value": "2000", "currency": "VND"},
                    {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-BOM-SWITCH", "line_no": "1", "item_code": "PV00.0048500", "description": "Growatt inverter", "hs_code": "850440", "quantity": "3", "unit": "PCS", "customs_value": "1000", "currency": "VND", "invoice_ref": "INV-BOM-SWITCH"},
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "BOM switch", "case_code": "CO-BOM-SWITCH", "destination_market": "Ấn Độ", "invoice_no": "INV-BOM-SWITCH"},
        follow_redirects=False,
    )

    origin = client.get(f"{created.headers['location']}/origin")

    assert origin.status_code == 200
    assert f'value="{v2["product_version_id"]}" selected' in origin.text
    assert "#1 · 2 dòng" in origin.text
    assert "#2 · 2 dòng" in origin.text

    form_data = hidden_form_data(origin.text)
    form_data["product_0_bom_product_artifact_id"] = v1["product_version_id"]
    switched = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        json={
            "origin_product_order": ["PV00.0048500"],
            "products": [
                {
                    "code": "PV00.0048500",
                    "bom_product_code": form_data["product_0_bom_product_code"],
                    "bom_product_artifact_id": v1["product_version_id"],
                    "name": form_data["product_0_name"],
                    "finished_hs": form_data["product_0_finished_hs"],
                    "quantity": form_data["product_0_quantity"],
                    "unit": form_data["product_0_unit"],
                    "currency": form_data["product_0_currency"],
                    "fob": form_data["product_0_fob"],
                    "rvc_threshold": form_data["product_0_rvc_threshold"],
                    "origin_sheet_status": form_data["product_0_origin_sheet_status"],
                    "origin_sheet_status_label": form_data["product_0_origin_sheet_status_label"],
                }
            ],
            "mark_stale": False,
        },
    )
    switched_data = hidden_form_data(switched.text)

    assert switched.status_code == 200
    assert f'value="{v1["product_version_id"]}" selected' in switched.text
    assert switched_data["product_0_material_0_consumed_qty"] == "3"
    assert switched_data["product_0_material_0_material_value"] == "30"


def test_cached_origin_context_loads_live_bom_artifact_options(monkeypatch):
    from app import main as main_module

    class FakeBomService:
        def workspace(self, client, product_codes=None, *, case_id=""):
            assert client["id"] == "growatt"
            assert product_codes == ["TP-BOM"]
            versions = [
                {
                    "product_code": "TP-BOM",
                    "product_artifact_id": "bom-artifact-1",
                    "product_version_id": "bom-artifact-1",
                    "product_artifact_no": 1,
                    "product_version_no": 1,
                    "row_count": 122,
                    "status": "published",
                    "flatten_status": "flattened",
                    "rows": [],
                },
                {
                    "product_code": "TP-BOM",
                    "product_artifact_id": "bom-artifact-2",
                    "product_version_id": "bom-artifact-2",
                    "product_artifact_no": 2,
                    "product_version_no": 2,
                    "row_count": 368,
                    "status": "current",
                    "flatten_status": "flattened",
                    "rows": [],
                },
            ]
            return {
                "versions": [],
                "product_versions": versions,
                "product_version_options_by_code": {"TP-BOM": versions},
                "latest_version": {},
                "latest_rows": [],
            }

    monkeypatch.setattr(main_module, "bom_service", FakeBomService())
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Cached BOM options", "case_code": "CO-CACHED-BOM", "destination_market": "Ấn Độ", "invoice_no": "INV-CACHED-BOM"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-CACHED-BOM",
            "title": "Cached BOM options",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-CACHED-BOM", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-CACHED",
                    "bom_product_code": "TP-BOM",
                    "bom_product_artifact_id": "bom-artifact-2",
                    "bom_product_artifact_no": 2,
                    "name": "Cached product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "materials": [],
                }
            ],
            "origin_product_order": ["TP-CACHED"],
            "origin_sheet_states": {"TP-CACHED": {"status": "calculated", "status_label": "Đã tính"}},
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {
                "aggregate_artifact_id": "bom-composition",
                "aggregate_artifact_no": 1,
                "composition": [
                    {
                        "product_code": "TP-BOM",
                        "product_artifact_id": "bom-artifact-2",
                        "product_artifact_no": 2,
                        "row_count": 368,
                    }
                ],
            },
            "source_snapshot": {"client_config_hash": "snapshot"},
            "source_invoice_matches": [
                {
                    "item_code": "TP-CACHED",
                    "quantity": "1",
                    "customs_value": "100",
                    "currency": "VND",
                    "material_identity": {"resolution_status": "resolved", "bom_product_code": "TP-BOM"},
                }
            ],
        },
    )

    origin = client.get(f"{created.headers['location']}/origin")

    assert origin.status_code == 200
    assert 'value="bom-artifact-2" selected' in origin.text
    assert "#1 · 122 dòng" in origin.text
    assert "#2 · 368 dòng" in origin.text


def test_co_case_origin_uses_data_hub_material_identity_for_bom_code():
    from app.main import co_case_bom_product_codes, prepare_case_origin_products, selected_bom_rows_by_product

    match = {
        "item_code": "BIENTAN.17",
        "quantity": "3",
        "customs_value": "100",
        "currency": "USD",
        "material_identity": {
            "resolution_status": "resolved",
            "product_kind": "tp",
            "bom_product_code": "PV01.0117500",
        },
    }
    case = {"products": [{"code": "BIENTAN.17", "bom_product_code": "PV01.0117500"}]}
    workspace = {
        "latest_version": {
            "version_id": "agg-1",
            "product_versions": [{"product_code": "PV01.0117500", "product_version_id": "bv-1"}],
        },
        "versions": [
            {
                "version_id": "agg-1",
                "rows": [
                    {
                        "product_code": "PV01.0117500",
                        "product_version_id": "bv-1",
                        "product_version_no": 2,
                        "material_code": "NVL-1",
                        "qty_per": "2",
                    }
                ],
                "product_versions": [{"product_code": "PV01.0117500", "product_version_id": "bv-1"}],
            }
        ],
        "latest_rows": [],
        "product_versions": [
            {
                "product_code": "PV01.0117500",
                "product_version_id": "bv-1",
                "product_version_no": 2,
                "rows": [{"product_code": "PV01.0117500", "material_code": "NVL-1", "qty_per": "2"}],
            }
        ],
        "product_version_options_by_code": {
            "PV01.0117500": [
                {
                    "product_code": "PV01.0117500",
                    "product_version_id": "bv-1",
                    "product_version_no": 2,
                    "rows": [{"product_code": "PV01.0117500", "material_code": "NVL-1", "qty_per": "2"}],
                }
            ]
        },
    }

    assert co_case_bom_product_codes({}, [match]) == ["PV01.0117500"]
    assert selected_bom_rows_by_product(case, workspace)["BIENTAN.17"][0]["material_code"] == "NVL-1"
    prepared = prepare_case_origin_products(
        {},
        [match],
        workspace,
        {},
        [],
        [],
    )
    assert prepared["products"][0]["bom_product_code"] == "PV01.0117500"
    assert prepared["products"][0]["materials"][0]["material_code"] == "NVL-1"
    assert prepared["products"][0]["materials"][0]["consumed_qty"] == Decimal("6")


def test_co_case_origin_does_not_auto_select_unresolved_material_identity():
    from app.main import co_case_bom_product_codes

    assert co_case_bom_product_codes({}, [
        {
            "item_code": "BIENTAN.17",
            "material_identity": {
                "resolution_status": "ambiguous",
                "product_kind": "tp",
                "bom_product_code": "",
                "selected_candidate_code": "PV01.0117500",
            },
        }
    ]) == ["BIENTAN.17"]


def test_co_case_origin_page_surfaces_method_readiness_and_evidence_gaps():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-ORIGIN-READY",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "1000",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-ORIGIN-READY",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "3",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-ORIGIN-READY",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Origin readiness",
            "case_code": "CO-ORIGIN-READY",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-ORIGIN-READY",
        },
        follow_redirects=False,
    )

    origin = client.get(f"{created.headers['location']}/origin")

    assert origin.status_code == 200
    assert "origin-config-bar" in origin.text
    assert "Cần bổ sung evidence" in origin.text
    assert "Thiếu đơn giá NVL" not in origin.text
    assert "DEMO-NPL-002: thiếu đơn giá để tính trị giá NVL/VNM." not in origin.text
    assert "CTSH preview" in origin.text
    assert "chưa thay thế PSR engine/legal review" in origin.text
    assert "Nguồn giá" in origin.text
    assert "Trạng thái dữ liệu" in origin.text

    form_data = hidden_form_data(origin.text)
    assert form_data["product_0_origin_method"] == "build_down_lvc"
    assert form_data["product_0_origin_readiness_status"] == "blocked"
    assert form_data["product_0_material_count"] == "0"

    calculated = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        json={
            "origin_product_order": ["PV00.0048500"],
            "products": [
                {
                    "code": "PV00.0048500",
                    "bom_product_code": form_data["product_0_bom_product_code"],
                    "bom_product_artifact_id": form_data.get("product_0_bom_product_artifact_id", ""),
                    "name": form_data["product_0_name"],
                    "finished_hs": form_data["product_0_finished_hs"],
                    "quantity": form_data["product_0_quantity"],
                    "unit": form_data["product_0_unit"],
                    "currency": form_data["product_0_currency"],
                    "fob": form_data["product_0_fob"],
                    "rvc_threshold": form_data["product_0_rvc_threshold"],
                    "origin_sheet_status": form_data["product_0_origin_sheet_status"],
                    "origin_sheet_status_label": form_data["product_0_origin_sheet_status_label"],
                }
            ],
            "mark_stale": False,
        },
    )
    calculated_data = hidden_form_data(calculated.text)

    assert calculated.status_code == 200
    assert "Thiếu đơn giá NVL" in calculated.text
    assert "DEMO-NPL-002: thiếu đơn giá để tính trị giá NVL/VNM." in calculated.text
    assert calculated_data["product_0_lvc_status"] == "partial_pass"
    assert calculated_data["product_0_lvc_percentage"] == "97.00"
    assert calculated_data["product_0_material_1_valuation_status"] == "missing_unit_value"
    assert "97.00%" in calculated.text
    assert "Tạm đạt LVC" in calculated.text
    assert "Thiếu đơn giá 1 dòng NVL; LVC đang tạm tính từ các dòng đã có đơn giá." in calculated.text


def test_co_case_origin_does_not_calculate_lvc_without_bom_materials():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-NO-BOM",
                        "line_no": "1",
                        "item_code": "NO-BOM-TP",
                        "description": "Finished good without BOM",
                        "hs_code": "850440",
                        "quantity": "5",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-NO-BOM",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "No BOM", "case_code": "CO-NO-BOM", "destination_market": "Ấn Độ", "invoice_no": "INV-NO-BOM"},
        follow_redirects=False,
    )

    origin = client.get(f"{created.headers['location']}/origin")
    form_data = hidden_form_data(origin.text)

    assert origin.status_code == 200
    assert form_data["product_0_lvc_status"] == "missing_bom"
    assert form_data["product_0_lvc_percentage"] == ""
    assert "Thiếu BOM/NVL" in origin.text
    assert "100.00%" not in origin.text


def test_origin_material_uses_stock_description_and_summarizes_repeated_warnings():
    from app.main import enrich_origin_product, origin_material_from_bom_row

    material = origin_material_from_bom_row(
        {"material_code": "MAT-001", "qty_per": "2", "uom": "PCS"},
        Decimal("3"),
        {},
        {
            "MAT-001": {
                "material_description": "Imported material name",
                "hs_code": "853690",
                "unit_value": "5",
                "currency": "USD",
                "remaining_qty": "100",
            }
        },
    )
    product = enrich_origin_product({
        "code": "TP-001",
        "fob": "100",
        "non_origin_value": "20",
        "rvc_threshold": "35",
        "lvc_status": "missing_value",
        "materials": [
            {"material_code": "MAT-001", "unit_value": "", "material_warnings": ["Repeated warning"]},
            {"material_code": "MAT-002", "unit_value": "", "material_warnings": ["Repeated warning"]},
        ],
    })

    assert material["material_description"] == "Imported material name"
    assert material["hs_code"] == "853690"
    assert product["lvc_percentage"] == "80.00"
    assert product["lvc_status"] == "partial_pass"
    assert product["lvc_status_label"] == "Tạm đạt LVC"
    assert product["origin_warnings"].count("Repeated warning") == 1
    assert any(row["label"] == "Thiếu đơn giá NVL" and row["count"] == 2 for row in product["origin_warning_summary"])


def test_origin_material_allocates_required_quantity_across_multiple_stock_lots():
    from app.main import co_stock_allocation_pool, origin_material_from_bom_row

    pool = co_stock_allocation_pool([
        {
            "source_row": "LOT-001",
            "import_declaration_no": "NK-LOT-1",
            "line_no": "1",
            "material_code": "MAT-LOT",
            "allocation_code": "MAT-LOT",
            "customs_item_code": "MAT-LOT",
            "remaining_qty": "2",
            "unit_value": "10",
            "currency": "VND",
            "value_currency": "VND",
            "eligibility_status": "active",
            "allocation_code_status": "resolved",
        },
        {
            "source_row": "LOT-002",
            "import_declaration_no": "NK-LOT-2",
            "line_no": "2",
            "material_code": "MAT-LOT",
            "allocation_code": "MAT-LOT",
            "customs_item_code": "MAT-LOT",
            "remaining_qty": "4",
            "unit_value": "20",
            "currency": "VND",
            "value_currency": "VND",
            "eligibility_status": "active",
            "allocation_code_status": "resolved",
        },
    ])

    material = origin_material_from_bom_row(
        {"material_code": "MAT-LOT", "qty_per": "3", "uom": "PCS"},
        Decimal("2"),
        {"MAT-LOT": {"origin_default": "Không xuất xứ"}},
        pool,
    )
    depleted = origin_material_from_bom_row(
        {"material_code": "MAT-LOT", "qty_per": "1", "uom": "PCS"},
        Decimal("1"),
        {"MAT-LOT": {"origin_default": "Không xuất xứ"}},
        pool,
    )

    assert material["consumed_qty"] == Decimal("6")
    assert material["source_row"] == "LOT-001,LOT-002"
    assert material["unit_value"] == "Nhiều đơn giá"
    assert material["material_value"] == "100"
    assert material["non_origin_cif_value"] == "100"
    assert material["allocation_status"] == "covered"
    assert [line["allocated_qty"] for line in material["allocation_lines"]] == ["2", "4"]
    assert [line["material_value"] for line in material["allocation_lines"]] == ["20", "80"]
    assert depleted["allocation_status"] == "shortage"
    assert depleted["allocation_shortage_qty"] == "1"
    assert depleted["allocation_lines"] == []


def test_origin_products_calculate_sequentially_against_case_stock_pool():
    from app.main import prepare_case_origin_products

    case = prepare_case_origin_products(
        {"shipment": {"invoice_no": "INV-SEQ"}},
        [
            {
                "item_code": "TP-A",
                "description": "Finished product A",
                "hs_code": "850440",
                "quantity": "4",
                "unit": "PCS",
                "customs_value": "1000",
                "value_currency": "VND",
                "invoice_ref": "INV-SEQ",
            },
            {
                "item_code": "TP-B",
                "description": "Finished product B",
                "hs_code": "850440",
                "quantity": "3",
                "unit": "PCS",
                "customs_value": "1000",
                "value_currency": "VND",
                "invoice_ref": "INV-SEQ",
            },
        ],
        {
            "latest_version": {
                "version_id": "bom-seq",
                "rows": [
                    {"product_code": "TP-A", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
                    {"product_code": "TP-B", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
                ],
            },
            "versions": [],
            "product_versions": [],
        },
        {"form_code": "B", "display_name": "C/O form B"},
        [{"customs_code": "MAT-SHARED", "origin_default": "Không xuất xứ"}],
        [
            {
                "source_row": "SEQ-STOCK-1",
                "import_declaration_no": "NK-SEQ",
                "line_no": "1",
                "material_code": "MAT-SHARED",
                "remaining_qty": "5",
                "unit_value": "10",
                "currency": "VND",
                "value_currency": "VND",
                "eligibility_status": "active",
                "allocation_code_status": "resolved",
            }
        ],
    )

    product_a, product_b = case["products"]
    material_a = product_a["materials"][0]
    material_b = product_b["materials"][0]

    assert [product_a["allocation_sequence"], product_b["allocation_sequence"]] == ["1", "2"]
    assert material_a["allocation_lines"][0]["opening_qty"] == "5"
    assert material_a["allocation_lines"][0]["allocated_qty"] == "4"
    assert material_a["allocation_lines"][0]["remaining_qty"] == "1"
    assert material_b["allocation_status"] == "shortage"
    assert material_b["allocation_shortage_qty"] == "2"
    assert material_b["allocation_lines"][0]["opening_qty"] == "1"
    assert material_b["allocation_lines"][0]["allocated_qty"] == "1"
    assert material_b["allocation_lines"][0]["product_sequence"] == "2"
    assert "Bước 1 TP-A dùng 4 PCS" in material_b["allocation_shortage_trace"]
    assert any("đã dùng ở bước trước" in warning for warning in material_b["material_warnings"])


def test_origin_sheet_calculate_updates_only_target_sheet():
    from app.main import prepare_case_origin_products, prepare_case_origin_sheet

    invoice_matches = [
        {
            "item_code": "TP-A",
            "description": "Finished product A",
            "hs_code": "850440",
            "quantity": "4",
            "unit": "PCS",
            "customs_value": "1000",
            "value_currency": "VND",
            "invoice_ref": "INV-SEQ",
        },
        {
            "item_code": "TP-B",
            "description": "Finished product B",
            "hs_code": "850440",
            "quantity": "3",
            "unit": "PCS",
            "customs_value": "1000",
            "value_currency": "VND",
            "invoice_ref": "INV-SEQ",
        },
    ]
    bom_workspace = {
        "latest_version": {
            "version_id": "bom-seq",
            "rows": [
                {"product_code": "TP-A", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
                {"product_code": "TP-B", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
            ],
        },
        "versions": [],
        "product_versions": [],
    }
    material_rows = [{"customs_code": "MAT-SHARED", "origin_default": "Không xuất xứ"}]
    stock_rows = [
        {
            "source_row": "SEQ-STOCK-1",
            "import_declaration_no": "NK-SEQ",
            "line_no": "1",
            "material_code": "MAT-SHARED",
            "remaining_qty": "5",
            "unit_value": "10",
            "currency": "VND",
            "value_currency": "VND",
            "eligibility_status": "active",
            "allocation_code_status": "resolved",
        }
    ]
    case = prepare_case_origin_products(
        {"shipment": {"invoice_no": "INV-SEQ"}},
        invoice_matches,
        bom_workspace,
        {"form_code": "B", "display_name": "C/O form B"},
        material_rows,
        stock_rows,
    )
    original_product_a = deepcopy(case["products"][0])

    recalculated = prepare_case_origin_sheet(
        case,
        "TP-B",
        invoice_matches,
        bom_workspace,
        {"form_code": "B", "display_name": "C/O form B"},
        material_rows,
        [{**stock_rows[0], "remaining_qty": "10"}],
    )

    product_a, product_b = recalculated["products"]
    material_b = product_b["materials"][0]

    assert product_a == original_product_a
    assert material_b["allocation_status"] == "covered"
    assert material_b["allocation_lines"][0]["opening_qty"] == "6"
    assert material_b["allocation_lines"][0]["allocated_qty"] == "3"
    assert material_b["allocation_lines"][0]["product_sequence"] == "2"


def test_origin_sheet_lock_uses_cached_case_context(monkeypatch):
    from app import main as main_module

    def fail_source_refresh(_client, _case):
        raise AssertionError("state-only origin actions must not refresh source context")

    monkeypatch.setattr(main_module, "co_case_source_context", fail_source_refresh)
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Fast lock", "case_code": "CO-FAST-LOCK", "destination_market": "Ấn Độ", "invoice_no": "INV-FAST"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-FAST-LOCK",
            "title": "Fast lock",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-FAST", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-FAST",
                    "name": "Fast product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "lvc_status": "pass",
                    "lvc_status_label": "Đạt LVC",
                    "materials": [],
                }
            ],
            "origin_sheet_states": {"TP-FAST": {"status": "calculated", "status_label": "Đã tính"}},
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {"aggregate_version_id": "bom-fast", "aggregate_version_no": 1, "composition": []},
            "source_snapshot": {"client_config_hash": "snapshot"},
        },
    )
    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-FAST/lock",
        data={
            "case_id": case_id,
            "persisted_case_id": case_id,
            "customer": "Growatt",
            "case_code": "CO-FAST-LOCK",
            "title": "Fast lock",
            "destination_market": "Ấn Độ",
            "agreement": "",
            "co_form_type": "",
            "rule": "",
            "invoice_no": "INV-FAST",
            "bill_of_lading_no": "",
            "mode": "Invoice + BCCT + BOM snapshot",
            "mode_note": "",
            "source_label": "",
            "document_count": "0",
            "product_count": "1",
            "origin_product_order": "TP-FAST",
            "product_0_code": "TP-FAST",
            "product_0_name": "Fast product",
            "product_0_quantity": "1",
            "product_0_unit": "PCS",
            "product_0_currency": "VND",
            "product_0_fob": "100",
            "product_0_lvc_status": "pass",
            "product_0_lvc_status_label": "Đạt LVC",
            "product_0_origin_sheet_status": "calculated",
            "product_0_origin_sheet_status_label": "Đã tính",
            "product_0_material_count": "0",
        },
    )

    assert response.status_code == 200
    assert "Đã chốt bảng kê TP-FAST" in response.text
    assert hidden_form_data(response.text)["product_0_origin_sheet_status"] == "locked"


def test_origin_sheet_recommendation_override_persists_per_sheet():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Override test",
            "case_code": "CO-OVERRIDE",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-OVERRIDE",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-OVERRIDE",
            "title": "Override test",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-OVERRIDE", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-OVR",
                    "name": "Override product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "materials": [],
                }
            ],
        },
    )

    initial = client.get(f"/clients/growatt/co-case/{case_id}/origin")
    assert initial.status_code == 200
    assert "origin-config-bar" in initial.text

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-OVR/recommendation-override",
        json={"form_override": "CPTPP", "criteria_override": "RVC 40% override"},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["state"]["form_override"] == "CPTPP"
    assert payload["state"]["criteria_override"] == "RVC 40% override"
    assert payload["state"]["effective_form_code"] == "CPTPP"

    saved = get_case_record(get_client("growatt"), case_id)
    assert saved["origin_sheet_states"]["TP-OVR"]["form_override"] == "CPTPP"
    assert saved["origin_sheet_states"]["TP-OVR"]["criteria_override"] == "RVC 40% override"

    rendered = client.get(f"/clients/growatt/co-case/{case_id}/origin")
    assert rendered.status_code == 200
    assert "RVC 40% override" in rendered.text

    rejected = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-OVR/recommendation-override",
        json={"form_override": "BOGUS", "criteria_override": ""},
    )
    assert rejected.status_code == 400


def test_origin_sheet_substitute_candidates_endpoint_returns_search_and_recommended():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Substitute test",
            "case_code": "CO-SUB",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-SUB",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SUB",
            "title": "Substitute test",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SUB", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-SUB",
                    "name": "Substitute product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "materials": [
                        {"material_code": "M-1", "material_description": "Material 1", "uom": "PCS", "bom_qty_per": "1"},
                        {"material_code": "M-2", "material_description": "Material 2", "uom": "PCS", "bom_qty_per": "2"},
                    ],
                }
            ],
        },
    )

    response = client.get(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-SUB/substitute-candidates",
        params={"material_code": "M-1", "row_index": "0", "search": "M-2"},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["product_code"] == "TP-SUB"
    assert payload["material_code"] == "M-1"
    assert payload["optimization_mode"] == "max_lvc"
    assert isinstance(payload["candidates"], list)
    assert isinstance(payload["search_results"], list)


def test_substitute_stock_reads_materialized_snapshot_not_bcct(monkeypatch):
    from app import main as main_module

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Substitute stock snapshot",
            "case_code": "CO-SUB-STOCK",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-SUB-STOCK",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SUB-STOCK",
            "title": "Substitute stock snapshot",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SUB-STOCK", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-STK",
                    "name": "Stock product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "materials": [
                        {"material_code": "STK-1", "material_description": "Stock material", "uom": "PCS", "bom_qty_per": "1"},
                    ],
                }
            ],
        },
    )

    snapshot_row = {
        "material_code": "STK-1",
        "customs_item_code": "STK-1",
        "allocation_code": "STK-1",
        "import_declaration_no": "IMP-9",
        "line_no": "1",
        "registration_date": "2026-01-02",
        "remaining_qty": "12",
        "available_qty": "12",
        "unit_value": "7.5",
        "currency": "USD",
        "material_description": "Stock material",
        "hs_code": "850440",
    }
    monkeypatch.setattr(
        main_module.co_stock_materializer, "read_co_stock_rows_cached",
        lambda _cid: [snapshot_row],
    )
    monkeypatch.setattr(
        main_module.co_stock_materializer, "last_refresh_at",
        lambda _cid: "2026-06-01T08:30:00+00:00",
    )

    def _bcct_must_not_be_called(*_args, **_kwargs):
        raise AssertionError("substitute-stock must not call Data Hub BCCT")

    monkeypatch.setattr(main_module.portfolio_service, "list_bcct_by_codes", _bcct_must_not_be_called)

    response = client.get(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-STK/substitute-stock",
        params={"codes": "STK-1"},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["stock_refreshed_at"] == "2026-06-01T08:30:00+00:00"
    entry = payload["stock"]["STK-1"]
    assert entry["lot_count"] == 1
    assert entry["total_remaining_qty"] not in ("", "0")


def test_local_material_search_supports_seed_catalog_lists():
    from app.portfolio import PortfolioService

    results = PortfolioService().search_materials("growatt", "DEMO-NPL-001")

    assert results
    assert results[0]["internal_code"] == "DEMO-NPL-001"


def test_origin_sheet_substitute_search_falls_back_to_case_materials_when_catalog_fails(monkeypatch):
    from app import main as main_module

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Substitute search fallback",
            "case_code": "CO-SUB-SEARCH",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-SUB-SEARCH",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SUB-SEARCH",
            "title": "Substitute search fallback",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SUB-SEARCH", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-SUB-SEARCH",
                    "name": "Substitute search product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "materials": [
                        {
                            "material_code": "MAT-FALLBACK-1",
                            "material_description": "Fallback material one",
                            "hs_code": "850490",
                            "uom": "PCS",
                            "bom_qty_per": "1",
                        }
                    ],
                }
            ],
        },
    )

    def broken_search(*_args, **_kwargs):
        raise RuntimeError("material catalog unavailable")

    monkeypatch.setattr(main_module.portfolio_service, "search_materials", broken_search, raising=False)

    response = client.get(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-SUB-SEARCH/substitute-candidates",
        params={"search": "MAT-FALLBACK", "row_index": "0"},
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["search_results"][0]["material_code"] == "MAT-FALLBACK-1"
    assert payload["search_results"][0]["name"] == "Fallback material one"
    assert "this dossier" in payload["error"]


def test_origin_sheet_substitute_row_persists_override_and_marks_stale():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Substitute apply",
            "case_code": "CO-SUB-APPLY",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-APPLY",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SUB-APPLY",
            "title": "Substitute apply",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-APPLY"},
            "products": [
                {
                    "code": "TP-APPLY",
                    "name": "Apply product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "materials": [
                        {"material_code": "M-OLD", "material_description": "Old material", "uom": "PCS", "bom_qty_per": "1"},
                    ],
                }
            ],
            "origin_sheet_states": {"TP-APPLY": {"status": "calculated", "status_label": "Đã tính"}},
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-APPLY/substitute-row",
        json={
            "row_index": 0,
            "new_material_code": "M-NEW",
            "new_norm_per_unit": "1.5",
            "new_name": "New material",
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["sheet_status"] == "stale"
    assert body["applied_override"]["material_code"] == "M-NEW"

    saved = get_case_record(get_client("growatt"), case_id)
    state = saved["origin_sheet_states"]["TP-APPLY"]
    assert state["status"] == "stale"
    assert state["material_overrides"]["0"]["material_code"] == "M-NEW"
    assert state["material_overrides"]["0"]["norm_per_unit"] == "1.5"

    deleted = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-APPLY/substitute-row",
        json={"row_index": 0, "delete": True},
    )
    assert deleted.status_code == 200
    assert deleted.json()["applied_override"] == {"deleted": True}

    rejected = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-APPLY/substitute-row",
        json={"row_index": 0},
    )
    assert rejected.status_code == 400


def test_case_lock_serializes_concurrent_save_state_on_postgres():
    """Two operators editing different cases of the same client must not
    clobber each other. After Phase 3.2 final, case_lock is file-only
    (fcntl.LOCK_EX), which still serializes within a single host;
    cross-host safety comes from per-case co_cases optimistic
    concurrency on `revision` and per-case co_supporting_files writes.
    This test pins the single-host invariant by running both writers in
    a single process — under the file lock, they take turns; both
    mutations persist.

    Skipped without a DB.
    """
    import threading
    import time
    import pytest

    from app.database import database_url
    from app.co_case_store import (
        case_lock,
        load_state,
        save_state,
        update_case_record,
    )

    if not database_url():
        pytest.skip("BARRY_DATABASE_URL not set; advisory lock requires Postgres")

    from app.co_case_store import _persist_case_row

    client = {"id": "lock-test"}
    # Seed two cases for this client via the per-case write path
    # (Phase 2.4: save_state no longer rewrites co_cases).
    with case_lock(client["id"]):
        state = load_state(client["id"])
        state.setdefault("cases", [])
        for seed in (
            {"case_id": "lock-a", "case_code": "LOCK-A", "title": "A",
             "destination_market": "", "shipment": {}, "products": [],
             "updated_at": "2026-05-28T00:00:00Z", "created_at": "2026-05-28T00:00:00Z",
             "supporting_files": []},
            {"case_id": "lock-b", "case_code": "LOCK-B", "title": "B",
             "destination_market": "", "shipment": {}, "products": [],
             "updated_at": "2026-05-28T00:00:00Z", "created_at": "2026-05-28T00:00:00Z",
             "supporting_files": []},
        ):
            state["cases"].append(seed)
            _persist_case_row(client["id"], seed, expected_revision=None)
        save_state(client["id"], state)

    timings: dict[str, float] = {}
    errors: list[Exception] = []

    def mutate(case_id: str, new_title: str, hold_seconds: float, key: str):
        try:
            with case_lock(client["id"]):
                timings[f"{key}_lock_acquired"] = time.monotonic()
                state = load_state(client["id"])
                for row in state["cases"]:
                    if row["case_id"] == case_id:
                        row["title"] = new_title
                        _persist_case_row(client["id"], row)
                if hold_seconds:
                    time.sleep(hold_seconds)
                save_state(client["id"], state)
                timings[f"{key}_saved"] = time.monotonic()
        except Exception as exc:  # noqa: BLE001
            errors.append(exc)

    t1 = threading.Thread(target=mutate, args=("lock-a", "A-edited", 0.5, "t1"))
    t2 = threading.Thread(target=mutate, args=("lock-b", "B-edited", 0.0, "t2"))
    t1.start()
    time.sleep(0.1)  # Make sure T1 holds the lock before T2 tries.
    t2.start()
    t1.join(timeout=10)
    t2.join(timeout=10)

    assert not errors, errors

    # T2 must not have acquired the lock until T1 saved (~0.4s after t1_lock).
    assert timings["t2_lock_acquired"] >= timings["t1_saved"] - 0.01, (
        "T2 acquired the lock before T1 saved — advisory lock is not blocking"
    )

    final = load_state(client["id"])
    titles = {row["case_id"]: row["title"] for row in final["cases"]}
    assert titles == {"lock-a": "A-edited", "lock-b": "B-edited"}, (
        f"clobber detected, titles={titles}"
    )

    # Cleanup.
    with case_lock(client["id"]):
        state = load_state(client["id"])
        state["cases"] = []
        save_state(client["id"], state)


def test_delete_case_requires_confirm_when_active_claims_and_then_releases(monkeypatch):
    """HIGH #2 fix: deleting a case that still holds locked co_stock_claims
    must (a) refuse without explicit confirm_release_claims=1 (modal alert
    path), and (b) release every claim when confirm is passed. Without this
    fix, deleting an open case with locked sheets used to orphan claims and
    silently leak Tồn CO.
    """
    from app import co_stock_ledger

    # Simulate the ledger reporting active claims even without a DB so the
    # business rule is exercised in unit-test mode.
    claims_state = {"count": 2, "lots": 2}
    released_calls: list[tuple[str, str]] = []

    def fake_summary(client_id: str, case_id: str) -> dict:
        return dict(claims_state)

    def fake_release(client_id: str, case_id: str) -> int:
        released_calls.append((client_id, case_id))
        n = claims_state["count"]
        claims_state["count"] = 0
        claims_state["lots"] = 0
        return n

    monkeypatch.setattr(co_stock_ledger, "claims_summary_for_case", fake_summary)
    monkeypatch.setattr(co_stock_ledger, "release_all_claims_for_case", fake_release)

    test_client = TestClient(app)
    created = test_client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Delete with claims",
            "case_code": "CO-DEL-CLAIMS",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-DEL-CLAIMS",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]

    # Without confirm → 409 with claim-count message, no release call.
    resp_no_confirm = test_client.post(f"/clients/growatt/co-case/{case_id}/delete")
    assert resp_no_confirm.status_code == 409
    assert "đang giữ" in resp_no_confirm.text.lower()
    assert released_calls == [], "release must NOT run when operator hasn't confirmed"
    assert claims_state["count"] == 2

    # With confirm → claims released, case removed, redirect with flash cookie.
    resp_confirm = test_client.post(
        f"/clients/growatt/co-case/{case_id}/delete",
        data={"confirm_release_claims": "1"},
        follow_redirects=False,
    )
    assert resp_confirm.status_code == 303
    assert released_calls == [("growatt", case_id)]
    assert claims_state["count"] == 0
    # Flash cookie is URL-encoded; substring check on the encoded form.
    flash = resp_confirm.cookies.get("co_flash") or ""
    assert "nh%E1%BA%A3" in flash or "nhả" in flash  # noqa: RUF001


def test_origin_sheet_locked_rejects_material_and_norm_mutations():
    """Locked sheet must refuse server-side POSTs from all material/norm
    mutation endpoints. The UI hides the edit buttons, but a dev-tools or
    curl bypass would otherwise quietly leave the ledger holding claims for
    materials no longer on the persisted sheet (Tồn CO leak).
    """
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Locked guard",
            "case_code": "CO-LOCKED-GUARD",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-LOCKED",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    seed_record = {
        "persisted_case_id": case_id,
        "case_code": "CO-LOCKED-GUARD",
        "title": "Locked guard",
        "destination_market": "Ấn Độ",
        "shipment": {"invoice_no": "INV-LOCKED"},
        "products": [
            {
                "code": "TP-LOCKED",
                "name": "Locked product",
                "finished_hs": "850440",
                "quantity": "1",
                "unit": "PCS",
                "fob": "100",
                "currency": "USD",
                "materials": [
                    {"material_code": "M-LOCKED", "material_description": "Locked material", "uom": "PCS", "bom_qty_per": "1"},
                ],
            }
        ],
        "origin_sheet_states": {"TP-LOCKED": {"status": "locked", "status_label": "Chốt"}},
    }
    update_case_record(get_client("growatt"), seed_record)

    base = f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-LOCKED"

    rejections = [
        ("substitute-row", {"row_index": 0, "new_material_code": "M-X", "new_norm_per_unit": "1"}),
        ("edit-row", {"row_index": 0, "new_norm_per_unit": "2.5"}),
        ("add-row", {"new_material_code": "M-NEW", "new_norm_per_unit": "1"}),
        ("save", {"replaces": {"0": {"new_material_code": "M-Y", "new_norm_per_unit": "1"}}}),
    ]
    for path, body in rejections:
        resp = client.post(f"{base}/{path}", json=body)
        assert resp.status_code == 409, f"{path} should reject when sheet locked, got {resp.status_code}: {resp.text}"
        assert "mở chốt" in resp.text.lower()

    saved = get_case_record(get_client("growatt"), case_id)
    state = saved["origin_sheet_states"]["TP-LOCKED"]
    assert state["status"] == "locked", "sheet status must stay locked after rejected mutations"
    assert "material_overrides" not in state or not state["material_overrides"], (
        "no material_overrides should be written when server rejected the mutation"
    )


def test_export_dossier_zip_bundles_chung_tu_tkx_tkn_and_hq_bang_ke(monkeypatch):
    import io
    import zipfile
    from openpyxl import load_workbook

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Zip dossier", "case_code": "CO-ZIP", "destination_market": "Ấn Độ", "invoice_no": "INV-ZIP"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-ZIP",
            "title": "Zip dossier",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-ZIP"},
            "products": [
                {
                    "code": "TP-ZIP",
                    "name": "Zip product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "documented_result": "LVC 30%",
                    "lvc_threshold": "30",
                    "materials": [
                        {"material_code": "M-Z", "material_description": "Zip mat", "uom": "PCS",
                         "bom_qty_per": "1", "unit_value": "10", "material_value": "10",
                         "origin_status": "non_origin", "consumed_qty": "1"},
                    ],
                }
            ],
            "origin_sheet_states": {"TP-ZIP": {"status": "locked", "status_label": "Chốt"}},
        },
    )
    # Dossier export requires case closed (sheets locked + case marked completed).
    update_case_record(get_client("growatt"), {"persisted_case_id": case_id, "status": "completed"})

    response = client.post(f"/clients/growatt/co-case/{case_id}/export-dossier-zip")
    assert response.status_code == 200, response.text
    assert response.headers["content-type"] == "application/zip"

    archive = zipfile.ZipFile(io.BytesIO(response.content))
    names = archive.namelist()
    # Layout redesigned 2026-05-28: numbered sections + MD readme/manifest.
    assert "00-README.md" in names
    bang_ke_entry = next(n for n in names if n.startswith("01-bang-ke/") and n.endswith(".xlsx"))
    assert "03-to-khai/MANIFEST.md" in names

    hq_workbook_bytes = archive.read(bang_ke_entry)
    hq_archive = zipfile.ZipFile(io.BytesIO(hq_workbook_bytes))
    hq_names = hq_archive.namelist()
    hq_workbook_xml = hq_archive.read("xl/workbook.xml")
    hq_content_types = hq_archive.read("[Content_Types].xml")
    assert not [name for name in hq_names if "external" in name.lower()]
    assert not [name for name in hq_names if "vba" in name.lower() or name.endswith(".bin")]
    assert b"externalReferences" not in hq_workbook_xml
    assert b"#REF!" not in hq_workbook_xml
    assert b"vnd.ms-office.vbaProject" not in hq_content_types
    wb = load_workbook(io.BytesIO(hq_workbook_bytes))
    # Template-based path renames each sheet to `<seq><product_code>`, e.g. "1TP-ZIP".
    # Shell-fallback path keeps the criterion sheet name (LVC). Accept either.
    template_named = any(name.endswith("TP-ZIP") for name in wb.sheetnames)
    shell_named = "LVC" in wb.sheetnames
    assert template_named or shell_named
    target_name = next((n for n in wb.sheetnames if n.endswith("TP-ZIP")), None) or "LVC"
    sheet = wb[target_name]
    if template_named:
        # 2026 FORM MAU template path (data/local/hq-templates/form-mau-combined.xlsx
        # present locally). CI may not have this gitignored agency file, so it
        # may exercise the shell fallback path instead — assertions below are
        # tied to the LVC form-mau compact layout.
        # Per FORM LVC.xlsx, body starts at row 16 and ends ~row 437.
        assert sheet.print_area == f"'{target_name}'!$A$1:$N$477"
        assert sheet["B12"].value == "Tên nguyên phụ liệu\n原材料的名称"
        assert sheet["C12"].value == "Mã nguyên vật liệu"
        assert sheet["D12"].value == "Mã HS\nHS CODE"
        assert sheet.cell(row=16, column=1).value == 1
        assert sheet.cell(row=16, column=2).value == "Zip mat"
        # LVC compact layout puts material_code at C, HS at D, UOM at E.
        assert sheet.cell(row=16, column=3).value == "M-Z"
        # LVC product header: name at L7, qty at L9, uom at N9, FOB at L10.
        assert sheet["L9"].value == Decimal("1")
        assert sheet["N9"].value == "PCS"
        assert sheet["L10"].value == Decimal("100")
    # Title row is produced by both paths.
    title = sheet["A3"].value or ""
    assert isinstance(title, str) and ("BẢNG KÊ" in title.upper() or "BẢNG TÍNH HÀM LƯỢNG" in title.upper())


def test_export_bang_ke_direct_route_works_on_open_case(monkeypatch):
    """Per-product Excel export is a WIP tool — must work without case close."""
    import io
    from openpyxl import load_workbook

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Direct bang ke", "case_code": "CO-DIRECT", "destination_market": "Ấn Độ", "invoice_no": "INV-DIRECT"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-DIRECT",
            "title": "Direct bang ke",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-DIRECT"},
            "products": [
                {"code": "TP-DIRECT", "name": "Direct", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD",
                 "documented_result": "LVC 30%", "lvc_threshold": "30",
                 "materials": [
                     {"material_code": "M-D", "uom": "PCS", "bom_qty_per": "1", "unit_value": "10", "material_value": "10",
                      "origin_status": "non_origin", "consumed_qty": "1"},
                 ]},
            ],
            "origin_sheet_states": {"TP-DIRECT": {"status": "locked", "status_label": "Chốt"}},
        },
    )

    quick_wb = Workbook()
    quick_wb.active.title = "1TP-DIRECT"
    quick_wb.active["A1"] = "quick bang ke"
    import app.main as main_module
    monkeypatch.setattr(main_module, "create_hq_bang_ke_workbook", lambda _case: workbook_bytes(quick_wb))

    direct = client.post(f"/clients/growatt/co-case/{case_id}/export-bang-ke")
    assert direct.status_code == 200, direct.text
    assert 'filename="CO-DIRECT-bang-ke-hq.xlsx"' in direct.headers["content-disposition"]
    direct_wb = load_workbook(io.BytesIO(direct.content))
    assert direct_wb["1TP-DIRECT"]["A1"].value == "quick bang ke"


def test_export_dossier_zip_blocks_when_sheet_stale_or_draft():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Block zip", "case_code": "CO-BLOCK", "destination_market": "Ấn Độ", "invoice_no": "INV-BLOCK"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-BLOCK",
            "title": "Block zip",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-BLOCK"},
            "products": [
                {"code": "TP-DRAFT", "name": "Draft", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []},
            ],
            "origin_sheet_states": {"TP-DRAFT": {"status": "stale", "status_label": "Cần tính lại"}},
        },
    )
    response = client.post(f"/clients/growatt/co-case/{case_id}/export-dossier-zip")
    assert response.status_code == 409
    assert "TP-DRAFT" in response.json()["detail"]


def test_export_dossier_zip_requires_case_closed_even_when_all_sheets_locked():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "All locked open", "case_code": "CO-ALL-LOCKED", "destination_market": "Ấn Độ", "invoice_no": "INV-AL"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-ALL-LOCKED",
            "title": "All locked open",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-AL"},
            "products": [
                {"code": "TP-A", "name": "A", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []},
            ],
            "origin_sheet_states": {"TP-A": {"status": "locked", "status_label": "Chốt"}},
        },
    )
    # Sheet locked but case still open → must close first.
    response = client.post(f"/clients/growatt/co-case/{case_id}/export-dossier-zip")
    assert response.status_code == 409
    assert "Đóng hồ sơ" in response.json()["detail"]


def test_export_dossier_zip_blocks_open_case_with_partial_locks():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Partial locks", "case_code": "CO-PARTIAL", "destination_market": "Ấn Độ", "invoice_no": "INV-PL"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-PARTIAL",
            "title": "Partial locks",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-PL"},
            "products": [
                {"code": "TP-A", "name": "A", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []},
                {"code": "TP-B", "name": "B", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []},
            ],
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính"},
            },
        },
    )
    response = client.post(f"/clients/growatt/co-case/{case_id}/export-dossier-zip")
    assert response.status_code == 409
    body = response.json()["detail"]
    assert "TP-B" in body
    assert "chưa chốt" in body


def test_origin_sheet_lock_records_cross_case_stock_ledger_claims():
    from app import co_stock_ledger
    from app.database import database_url

    if not database_url():
        pytest.skip("co_stock_ledger requires BARRY_DATABASE_URL — run with Postgres for ledger coverage")
    # Reset ledger for a clean test scope (ledger is global per client_id).
    # Also seed co_stock_rows for the synthetic lots this test claims against:
    # the lock pre-check rejects allocations whose source_row isn't in the
    # materialized snapshot (real lots that no longer exist in BCCT). Provide
    # generous remaining_qty so the test exercises the happy path, not the
    # over-claim path (that's covered by a dedicated test below).
    from psycopg.types.json import Jsonb
    try:
        with co_stock_ledger._connect() as conn, conn.cursor() as cur:
            cur.execute("delete from co_stock_claims where client_id = %s", ("growatt",))
            cur.execute(
                "delete from co_stock_rows where client_id = %s and source_row = any(%s)",
                ("growatt", ["ROW-A1", "ROW-A2", "ROW-B1"]),
            )
            for sr, qty in (("ROW-A1", "100"), ("ROW-A2", "100"), ("ROW-B1", "100")):
                cur.execute(
                    """insert into co_stock_rows (
                        client_id, source_row, remaining_qty, payload
                       ) values (%s, %s, %s, %s)""",
                    ("growatt", sr, qty, Jsonb({"source_row": sr, "remaining_qty": qty})),
                )
    except Exception:  # noqa: BLE001
        pytest.skip("co_stock_claims table missing — apply migrations first")
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Lock ledger", "case_code": "CO-LEDGER", "destination_market": "Ấn Độ", "invoice_no": "INV-LEDGER"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-LEDGER",
            "title": "Lock ledger",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-LEDGER"},
            "products": [
                {
                    "code": "TP-LEDGER",
                    "name": "Lock product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "lvc_status": "pass",
                    "lvc_status_label": "Đạt LVC",
                    "lvc_percentage": "70.00",
                    "materials": [
                        {
                            "material_code": "M-A",
                            "uom": "PCS",
                            "bom_qty_per": "1",
                            "allocation_lines": [
                                {"source_row": "ROW-A1", "allocated_qty": "5"},
                                {"source_row": "ROW-A2", "allocated_qty": "3"},
                            ],
                        },
                        {
                            "material_code": "M-B",
                            "uom": "PCS",
                            "bom_qty_per": "1",
                            "allocation_lines": [
                                {"source_row": "ROW-B1", "allocated_qty": "10"},
                            ],
                        },
                    ],
                }
            ],
            "origin_sheet_states": {"TP-LEDGER": {"status": "calculated", "status_label": "Đã tính"}},
        },
    )

    lock_response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-LEDGER/lock",
        data={
            "case_id": case_id,
            "persisted_case_id": case_id,
            "case_code": "CO-LEDGER",
            "title": "Lock ledger",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-LEDGER",
            "product_count": "1",
            "origin_product_order": "TP-LEDGER",
            "product_0_code": "TP-LEDGER",
            "product_0_name": "Lock product",
            "product_0_quantity": "1",
            "product_0_unit": "PCS",
            "product_0_fob": "100",
            "product_0_currency": "VND",
            "product_0_lvc_status": "pass",
            "product_0_lvc_status_label": "Đạt LVC",
            "product_0_lvc_percentage": "70.00",
            "product_0_origin_sheet_status": "calculated",
            "product_0_material_count": "0",
        },
    )
    assert lock_response.status_code == 200, lock_response.text

    from decimal import Decimal
    used = co_stock_ledger.used_qty_by_lot("growatt")
    assert used.get("ROW-A1") == Decimal("5")
    assert used.get("ROW-A2") == Decimal("3")
    assert used.get("ROW-B1") == Decimal("10")

    release_response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-LEDGER/reopen",
        data={
            "case_id": case_id,
            "persisted_case_id": case_id,
            "case_code": "CO-LEDGER",
            "title": "Lock ledger",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-LEDGER",
            "product_count": "1",
            "origin_product_order": "TP-LEDGER",
            "product_0_code": "TP-LEDGER",
            "product_0_name": "Lock product",
            "product_0_quantity": "1",
            "product_0_unit": "PCS",
            "product_0_fob": "100",
            "product_0_currency": "VND",
            "product_0_lvc_status": "pass",
            "product_0_origin_sheet_status": "locked",
            "product_0_material_count": "0",
        },
    )
    assert release_response.status_code == 200, release_response.text
    used_after = co_stock_ledger.used_qty_by_lot("growatt")
    assert used_after == {}, f"expected ledger to release all locks, got {used_after}"


def test_origin_sheet_lock_rejects_overclaim_against_materialized_snapshot():
    """When the requested allocation exceeds (snapshot remaining - other-case
    claims) for any lot, the lock must fail with 409 and leave the sheet in
    its previous 'calculated' state — no claim should be written, no case
    state mutation."""
    from app import co_stock_ledger
    from app.database import database_url
    from psycopg.types.json import Jsonb

    if not database_url():
        pytest.skip("co_stock_ledger requires BARRY_DATABASE_URL")
    try:
        with co_stock_ledger._connect() as conn, conn.cursor() as cur:
            cur.execute("delete from co_stock_claims where client_id = %s", ("growatt",))
            cur.execute(
                "delete from co_stock_rows where client_id = %s and source_row = any(%s)",
                ("growatt", ["ROW-LIMIT"]),
            )
            # Seed a snapshot lot with only 4 units left.
            cur.execute(
                """insert into co_stock_rows (client_id, source_row, remaining_qty, payload)
                   values (%s, %s, %s, %s)""",
                ("growatt", "ROW-LIMIT", "4", Jsonb({"source_row": "ROW-LIMIT", "remaining_qty": "4"})),
            )
    except Exception:  # noqa: BLE001
        pytest.skip("co_stock_rows table missing — apply migrations first")

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Overclaim", "case_code": "CO-OVERCLAIM", "destination_market": "Ấn Độ", "invoice_no": "INV-OC"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-OVERCLAIM",
            "title": "Overclaim",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-OC"},
            "products": [
                {
                    "code": "TP-OC",
                    "name": "Overclaim product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "lvc_status": "pass",
                    "lvc_status_label": "Đạt LVC",
                    "lvc_percentage": "70.00",
                    "materials": [
                        {
                            "material_code": "M-OC",
                            "uom": "PCS",
                            "bom_qty_per": "1",
                            # Claim 10 against a lot with only 4 remaining.
                            "allocation_lines": [{"source_row": "ROW-LIMIT", "allocated_qty": "10"}],
                        }
                    ],
                }
            ],
            "origin_sheet_states": {"TP-OC": {"status": "calculated", "status_label": "Đã tính"}},
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-OC/lock",
        data={
            "case_id": case_id, "persisted_case_id": case_id,
            "case_code": "CO-OVERCLAIM", "title": "Overclaim",
            "destination_market": "Ấn Độ", "invoice_no": "INV-OC",
            "product_count": "1", "origin_product_order": "TP-OC",
            "product_0_code": "TP-OC", "product_0_name": "Overclaim product",
            "product_0_quantity": "1", "product_0_unit": "PCS",
            "product_0_fob": "100", "product_0_currency": "VND",
            "product_0_lvc_status": "pass", "product_0_lvc_status_label": "Đạt LVC",
            "product_0_lvc_percentage": "70.00",
            "product_0_origin_sheet_status": "calculated", "product_0_material_count": "0",
        },
    )
    assert response.status_code == 409, response.text
    assert "ROW-LIMIT" in response.text and "vượt tồn" in response.text.lower()

    # Verify no claim was written.
    assert co_stock_ledger.used_qty_by_lot("growatt").get("ROW-LIMIT") is None

    # Verify the case sheet stayed 'calculated' (the lock did not mutate state).
    persisted = get_case_record(get_client("growatt"), case_id)
    sheet_state = (persisted.get("origin_sheet_states") or {}).get("TP-OC", {})
    assert sheet_state.get("status") == "calculated", sheet_state


def test_origin_sheet_propose_bom_requires_lock_and_overrides(monkeypatch):
    from app import main as main_module

    captured = {}

    def fake_submit(client_id, product_code, **kwargs):
        captured["client_id"] = client_id
        captured["product_code"] = product_code
        captured["rows"] = kwargs["rows"]
        captured["context"] = kwargs["context"]
        captured["parent_artifact_id"] = kwargs["parent_artifact_id"]
        return {"proposal_id": "prop-1", "artifact_id": "bv_NEW", "status": "submitted"}

    monkeypatch.setattr(main_module.portfolio_service, "submit_bom_proposal", fake_submit, raising=False)

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Propose", "case_code": "CO-PROP", "destination_market": "Ấn Độ", "invoice_no": "INV-PROP"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-PROP",
            "title": "Propose",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-PROP"},
            "products": [
                {
                    "code": "TP-PROP",
                    "name": "Propose product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "bom_product_code": "TP-PROP",
                    "bom_product_artifact_id": "bv_OLD",
                    "materials": [
                        {"material_code": "M-A", "uom": "PCS", "bom_qty_per": "1", "material_description": "Mat A"},
                        {"material_code": "M-B", "uom": "PCS", "bom_qty_per": "1", "material_description": "Mat B"},
                    ],
                }
            ],
        },
    )

    # Without lock should 409
    not_locked = client.post(f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-PROP/propose-bom", json={})
    assert not_locked.status_code == 409

    # Mark locked but no overrides → 409
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "origin_sheet_states": {"TP-PROP": {"status": "locked", "status_label": "Chốt"}},
        },
    )
    no_overrides = client.post(f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-PROP/propose-bom", json={})
    assert no_overrides.status_code == 409

    # Add override + lock → success
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "origin_sheet_states": {
                "TP-PROP": {
                    "status": "locked", "status_label": "Chốt",
                    "material_overrides": {"0": {"material_code": "M-NEW", "norm_per_unit": "2"}},
                }
            },
        },
    )
    success = client.post(f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-PROP/propose-bom", json={})
    assert success.status_code == 200, success.text
    assert success.json()["proposal"]["artifact_id"] == "bv_NEW"
    assert captured["parent_artifact_id"] == "bv_OLD"
    assert captured["product_code"] == "TP-PROP"
    # Field-name contract: Data Hub's bom-row contract uses qty_per_unit, not
    # norm_per_unit. Sending the wrong key makes DH read qty as 0 and reject
    # via qty_delta_exceeds_tolerance (auto mode) or display blank định mức
    # in the reviewer UI (manual mode).
    rows_by_code = {row["material_code"]: row for row in captured["rows"]}
    assert rows_by_code["M-NEW"]["qty_per_unit"] == "2"
    assert "norm_per_unit" not in rows_by_code["M-NEW"]
    assert rows_by_code["M-B"]["qty_per_unit"] == "1"
    assert captured["context"]["case_id"] == case_id

    saved = get_case_record(get_client("growatt"), case_id)
    assert saved["origin_sheet_states"]["TP-PROP"]["proposed_artifact_id"] == "bv_NEW"


def test_origin_sheet_edit_row_persists_norm_only_override():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Edit row", "case_code": "CO-EDIT", "destination_market": "Ấn Độ", "invoice_no": "INV-EDIT"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-EDIT",
            "title": "Edit row",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-EDIT"},
            "products": [
                {"code": "TP-EDIT", "name": "Edit prod", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD",
                 "materials": [{"material_code": "M-1", "uom": "PCS", "bom_qty_per": "1"}]},
            ],
            "origin_sheet_states": {"TP-EDIT": {"status": "calculated", "status_label": "Đã tính"}},
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-EDIT/edit-row",
        json={"row_index": 0, "new_norm_per_unit": "2.5"},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["applied_override"]["norm_per_unit"] == "2.5"
    assert body["applied_override"]["norm_edit_only"] is True
    assert body["sheet_status"] == "stale"

    saved = get_case_record(get_client("growatt"), case_id)
    assert saved["origin_sheet_states"]["TP-EDIT"]["material_overrides"]["0"]["norm_per_unit"] == "2.5"

    bad = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-EDIT/edit-row",
        json={"row_index": 0, "new_norm_per_unit": "abc"},
    )
    assert bad.status_code == 400


def test_origin_sheet_renders_effective_norm_as_live_recompute_baseline():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-NORM",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "10",
                        "unit": "PCE",
                        "customs_value": "100",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-NORM",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "1",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-NORM",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Norm baseline", "case_code": "CO-NORM", "destination_market": "Ấn Độ", "invoice_no": "INV-NORM"},
        follow_redirects=False,
    )
    origin = client.get(f"{created.headers['location']}/origin")
    calculated = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        data=hidden_form_data(origin.text),
    )
    assert calculated.status_code == 200
    edited = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/edit-row",
        json={"row_index": 0, "new_norm_per_unit": "2.5"},
    )
    assert edited.status_code == 200

    response = client.get(f"{created.headers['location']}/origin")
    assert response.status_code == 200
    assert 'data-row-original-norm="2.5"' in response.text
    assert 'data-current-norm="2.5"' in response.text
    assert 'data-original-value="2.5"' in response.text
    assert 'value="2.5"' in response.text
    assert 'data-row-original-norm="1"' not in response.text


def test_origin_sheet_save_recomputes_replaced_material_snapshot():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-SWAP",
                        "line_no": "1",
                        "item_code": "M-NEW",
                        "description": "New material from stock",
                        "hs_code": "8542.39",
                        "quantity": "10",
                        "unit": "PCE",
                        "customs_value": "80",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-SWAP",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "1",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-SWAP",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Swap material", "case_code": "CO-SWAP", "destination_market": "Ấn Độ", "invoice_no": "INV-SWAP"},
        follow_redirects=False,
    )
    origin = client.get(f"{created.headers['location']}/origin")
    calculated = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        data=hidden_form_data(origin.text),
    )
    assert calculated.status_code == 200
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    before = get_case_record(get_client("growatt"), case_id)
    assert before["products"][0]["materials"][0]["material_code"] != "M-NEW"

    saved_response = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/save",
        json={
            "replaces": {
                "0": {
                    "new_material_code": "M-NEW",
                    "new_norm_per_unit": "2",
                    "new_name": "New material override",
                    "new_hs_code": "854239",
                }
            }
        },
    )
    assert saved_response.status_code == 200
    assert saved_response.json()["sheet_status"] == "calculated"

    saved = get_case_record(get_client("growatt"), case_id)
    material = saved["products"][0]["materials"][0]
    assert material["material_code"] == "M-NEW"
    assert material["internal_material_code"] == "M-NEW"
    assert material["material_description"] == "New material override"
    assert material["hs_code"] == "854239"
    assert str(material["bom_qty_per"]) == "2"
    assert str(material["consumed_qty"]) == "2"
    assert saved["origin_sheet_states"]["PV00.0048500"]["status"] == "calculated"

    reloaded = client.get(f"{created.headers['location']}/origin")
    assert reloaded.status_code == 200
    assert 'data-material-code="M-NEW"' in reloaded.text
    assert "→ M-NEW" not in reloaded.text


def test_origin_sheet_add_row_appends_added_override():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Add row", "case_code": "CO-ADD", "destination_market": "Ấn Độ", "invoice_no": "INV-ADD"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-ADD",
            "title": "Add row",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-ADD"},
            "products": [{"code": "TP-ADD", "name": "Add prod", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []}],
        },
    )

    first = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-ADD/add-row",
        json={"new_material_code": "M-NEW", "new_norm_per_unit": "1.25", "new_name": "New mat", "new_uom": "PCS", "new_hs_code": "850440"},
    )
    assert first.status_code == 200
    second = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-ADD/add-row",
        json={"new_material_code": "M-NEW2", "new_norm_per_unit": "0.5"},
    )
    assert second.status_code == 200
    assert first.json()["added_key"].startswith("added_")
    assert first.json()["added_key"] != second.json()["added_key"]

    saved = get_case_record(get_client("growatt"), case_id)
    overrides = saved["origin_sheet_states"]["TP-ADD"]["material_overrides"]
    assert len(overrides) == 2
    assert all(v.get("added") for v in overrides.values())

    rejected = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-ADD/add-row",
        json={"new_norm_per_unit": "0.1"},
    )
    assert rejected.status_code == 400


def test_co_stock_import_upserts_adjustments_and_clears_cache():
    from app import co_stock_adjustments_store
    from app.co_stock_template import write_standard_co_stock
    from app.database import database_url
    from decimal import Decimal

    if not database_url():
        pytest.skip("co_stock_adjustments requires BARRY_DATABASE_URL")
    # Clean slate for this test client.
    try:
        with co_stock_adjustments_store.connect() as conn, conn.cursor() as cur:
            cur.execute("delete from co_stock_adjustments where client_id = %s", ("growatt",))
    except Exception:  # noqa: BLE001
        pytest.skip("co_stock_adjustments table missing — apply migrations first")

    xlsx = write_standard_co_stock([
        {"declaration_no": "D1", "line_no": "1", "customs_code": "M-A", "opening_qty": Decimal("500"), "used_qty": Decimal("11.489"), "source_co_no": "VNG-001"},
        {"declaration_no": "D1", "line_no": "2", "customs_code": "M-B", "opening_qty": Decimal("200"), "used_qty": Decimal("80")},
    ])
    test_client = TestClient(app)
    response = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("test.xlsx", xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["parsed_rows"] == 2
    assert body["upsert"]["inserted"] == 2
    assert body["upsert"]["updated"] == 0
    assert body["batch_id"].startswith("batch_")

    # Re-upload same content → all rows updated (not inserted).
    response2 = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("test.xlsx", xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response2.status_code == 200
    body2 = response2.json()
    assert body2["upsert"]["inserted"] == 0
    assert body2["upsert"]["updated"] == 2
    assert body2["batch_id"] == body["batch_id"]  # identical content → identical batch_id

    # Adjustments are visible via the aggregation helper.
    agg = co_stock_adjustments_store.aggregate_by_lookup_key("growatt")
    assert ("D1", "1", "M-A") in agg
    assert agg[("D1", "1", "M-A")]["used_qty"] == Decimal("11.489")


def test_co_stock_import_emits_audit_events_with_diff():
    from app import co_stock_adjustments_store, co_stock_events_store
    from app.co_stock_template import write_standard_co_stock
    from app.database import database_url
    from decimal import Decimal

    if not database_url():
        pytest.skip("co_stock_adjustments requires BARRY_DATABASE_URL")
    # Clean slate.
    try:
        with co_stock_adjustments_store.connect() as conn, conn.cursor() as cur:
            cur.execute("delete from co_stock_adjustments where client_id = %s", ("growatt",))
            cur.execute("delete from co_stock_events where client_id = %s", ("growatt",))
    except Exception:
        pytest.skip("co_stock_adjustments / co_stock_events table missing")
    test_client = TestClient(app)

    # First upload: insert events expected.
    xlsx_v1 = write_standard_co_stock([
        {"declaration_no": "D-EVT", "line_no": "1", "customs_code": "M-EVT", "opening_qty": Decimal("100"), "used_qty": Decimal("20")},
    ])
    r1 = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("v1.xlsx", xlsx_v1, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r1.status_code == 200
    assert r1.json()["upsert"]["events"] == 1

    # Re-upload same values: no events (no diff).
    r2 = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("v1-again.xlsx", xlsx_v1, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r2.status_code == 200
    assert r2.json()["upsert"]["events"] == 0

    # Upload with bumped used_qty: one update event with positive delta.
    xlsx_v2 = write_standard_co_stock([
        {"declaration_no": "D-EVT", "line_no": "1", "customs_code": "M-EVT", "opening_qty": Decimal("100"), "used_qty": Decimal("55")},
    ])
    r3 = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("v2.xlsx", xlsx_v2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r3.status_code == 200
    assert r3.json()["upsert"]["events"] == 1

    events = co_stock_events_store.events_for_lot("growatt", "D-EVT", "1", "M-EVT")
    assert len(events) == 2
    # Newest first.
    assert events[0]["event_type"] == "adjustment_import_update"
    assert Decimal(events[0]["qty_delta"]) == Decimal("35")
    assert Decimal(events[0]["qty_before"]) == Decimal("20")
    assert Decimal(events[0]["qty_after"]) == Decimal("55")
    assert events[1]["event_type"] == "adjustment_import_insert"
    assert Decimal(events[1]["qty_after"]) == Decimal("20")


def test_co_stock_lot_history_endpoint_returns_chronological_events():
    from app import co_stock_adjustments_store, co_stock_events_store
    from app.database import database_url

    if not database_url():
        pytest.skip("requires BARRY_DATABASE_URL")
    try:
        with co_stock_adjustments_store.connect() as conn, conn.cursor() as cur:
            cur.execute("delete from co_stock_events where client_id = %s", ("growatt",))
    except Exception:
        pytest.skip("co_stock_events table missing")
    # Seed two events manually.
    co_stock_events_store.record_event(
        client_id="growatt", declaration_no="D-H", line_no="2", customs_code="M-H",
        event_type="adjustment_import_insert", qty_delta=10, qty_after=10, opening_qty_after=100,
        batch_id="b1", source_file_ref="seed.xlsx", actor="seed",
    )
    co_stock_events_store.record_event(
        client_id="growatt", declaration_no="D-H", line_no="2", customs_code="M-H",
        event_type="claim_lock", qty_delta=5, case_id="case-X", sheet_product_code="TP-1",
        actor="ledger:lock",
    )
    test_client = TestClient(app)
    r = test_client.get("/clients/growatt/co-stock/lot-history", params={
        "declaration_no": "D-H", "line_no": "2", "customs_code": "M-H",
    })
    assert r.status_code == 200
    body = r.json()
    assert body["count"] == 2
    types = [e["event_type"] for e in body["events"]]
    # Newest first: claim_lock then adjustment_import_insert
    assert types == ["claim_lock", "adjustment_import_insert"]

    # Missing key params → 400
    r_bad = test_client.get("/clients/growatt/co-stock/lot-history", params={"declaration_no": "X"})
    assert r_bad.status_code == 400


def test_co_stock_import_rejects_malformed_workbook():
    test_client = TestClient(app)
    response = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("garbage.xlsx", b"this is not a workbook", "application/octet-stream")},
    )
    assert response.status_code == 400
    assert "workbook" in response.json()["detail"].lower()


def test_co_stock_import_rejects_missing_key_columns():
    from openpyxl import Workbook
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "co_stock"
    ws.append(["declaration_no", "line_no"])  # missing customs_code
    ws.append(["D1", "1"])
    payload = BytesIO()
    wb.save(payload)
    test_client = TestClient(app)
    response = test_client.post(
        "/clients/growatt/co-stock/import",
        files={"file": ("bad.xlsx", payload.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 400
    assert "customs_code" in response.json()["detail"]


def test_co_stock_export_returns_xlsx_with_standard_headers():
    from app.co_stock_template import read_standard_co_stock

    test_client = TestClient(app)
    response = test_client.get("/clients/growatt/co-stock/export.xlsx")
    assert response.status_code == 200
    assert "spreadsheetml" in response.headers["content-type"]
    assert "co-stock-" in response.headers["content-disposition"]
    # Body must be parseable by our own loader (Vietnamese-label header round-trip).
    rows, errors = read_standard_co_stock(response.content)
    assert errors == []
    # Growatt demo seeds at least one BCCT import row; export should reflect ≥1 row.
    assert isinstance(rows, list)


def test_origin_sheet_save_batches_replaces_adds_deletes_and_norm_edits():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Save batch", "case_code": "CO-SAVE", "destination_market": "Ấn Độ", "invoice_no": "INV-SAVE"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SAVE",
            "title": "Save batch",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SAVE"},
            "products": [{
                "code": "TP-SAVE",
                "name": "Save prod",
                "quantity": "1",
                "unit": "PCS",
                "fob": "100",
                "currency": "USD",
                "materials": [],
            }],
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-SAVE/save",
        json={
            "replaces": {"0": {"new_material_code": "M-SWAP", "new_norm_per_unit": "1.5", "new_name": "Swap"}},
            "norm_edits": {"1": "2.25"},
            "deletes": {"2": True},
            "adds": [
                {"new_material_code": "M-ADD", "new_norm_per_unit": "0.5", "new_name": "Added"},
                {"new_material_code": "M-ADD2", "new_norm_per_unit": "0.1"},
            ],
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["operations"] == {"replaces": 1, "adds": 2, "deletes": 1, "norm_edits": 1}
    assert body["sheet_status"] == "calculated"

    saved = get_case_record(get_client("growatt"), case_id)
    overrides = saved["origin_sheet_states"]["TP-SAVE"]["material_overrides"]
    assert overrides["0"]["material_code"] == "M-SWAP"
    assert overrides["0"]["norm_per_unit"] == "1.5"
    assert overrides["1"]["norm_per_unit"] == "2.25"
    assert overrides["1"]["norm_edit_only"] is True
    assert overrides["2"]["deleted"] is True
    added = [k for k in overrides if k.startswith("added_")]
    assert len(added) == 2
    assert {overrides[k]["material_code"] for k in added} == {"M-ADD", "M-ADD2"}


def test_origin_sheet_save_merges_full_workbook_state_before_recompute():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Sheet save workbook", "case_code": "CO-SHEET-WB", "destination_market": "Ấn Độ", "invoice_no": "INV-SHEET-WB"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SHEET-WB",
            "title": "Sheet save workbook",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SHEET-WB"},
            "products": [
                {"code": "TP-A", "name": "A", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []},
                {
                    "code": "TP-B",
                    "name": "B",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "materials": [
                        {"material_code": "M-OLD", "material_description": "Old", "uom": "PCS", "bom_qty_per": "1"}
                    ],
                },
            ],
            "origin_product_order": ["TP-A", "TP-B"],
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt", "criteria_override": "preserve A"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính", "currency_mode": "native"},
            },
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-B/save",
        json={
            "origin_product_order": ["TP-A", "TP-B"],
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt", "criteria_override": "preserve A"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính", "currency_mode": "vnd"},
            },
            "products": [
                {"code": "TP-A", "origin_sheet_status": "locked", "origin_sheet_status_label": "Chốt"},
                {"code": "TP-B", "origin_sheet_status": "calculated", "origin_sheet_status_label": "Đã tính"},
            ],
            "norm_edits": {"0": "2"},
        },
    )

    saved = get_case_record(get_client("growatt"), case_id)
    assert response.status_code == 200
    assert response.json()["origin_product_order"] == ["TP-A", "TP-B"]
    assert saved["origin_product_order"] == ["TP-A", "TP-B"]
    assert saved["origin_sheet_states"]["TP-A"]["status"] == "locked"
    assert saved["origin_sheet_states"]["TP-A"]["criteria_override"] == "preserve A"
    assert saved["origin_sheet_states"]["TP-B"]["status"] == "calculated"
    assert saved["origin_sheet_states"]["TP-B"]["currency_mode"] == "vnd"
    assert saved["origin_sheet_states"]["TP-B"]["material_overrides"]["0"]["norm_per_unit"] == "2"


def test_origin_sheet_save_rejects_empty_payload_and_invalid_norm():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Save invalid", "case_code": "CO-SAVEX", "destination_market": "Ấn Độ", "invoice_no": "INV-SAVEX"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SAVEX",
            "title": "Save invalid",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SAVEX"},
            "products": [{"code": "TP-SAVEX", "name": "X", "quantity": "1", "unit": "PCS", "fob": "100", "currency": "USD", "materials": []}],
        },
    )
    empty = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-SAVEX/save",
        json={},
    )
    assert empty.status_code == 400

    bad_norm = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-SAVEX/save",
        json={"norm_edits": {"0": "not-a-number"}},
    )
    assert bad_norm.status_code == 400


def test_origin_sheet_threshold_currency_optimization_overrides_persist():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Threshold test",
            "case_code": "CO-THRESHOLD",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-THRESHOLD",
        },
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-THRESHOLD",
            "title": "Threshold test",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-THRESHOLD", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-THR",
                    "name": "Threshold product",
                    "finished_hs": "850440",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "USD",
                    "lvc_threshold": "30",
                    "materials": [],
                }
            ],
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-THR/recommendation-override",
        json={
            "lvc_threshold_override": "45",
            "rvc_threshold_override": "55",
            "currency_mode": "vnd",
            "optimization_mode": "min_lvc",
        },
    )
    assert response.status_code == 200
    state = response.json()["state"]
    assert state["lvc_threshold_override"] == "45"
    assert state["rvc_threshold_override"] == "55"
    assert state["currency_mode"] == "vnd"
    assert state["optimization_mode"] == "min_lvc"
    assert state["effective_lvc_threshold"] == "45"

    saved = get_case_record(get_client("growatt"), case_id)
    assert saved["origin_sheet_states"]["TP-THR"]["lvc_threshold_override"] == "45"
    assert saved["origin_sheet_states"]["TP-THR"]["currency_mode"] == "vnd"

    # invalid threshold should be rejected as empty (out of range)
    bad_threshold = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-THR/recommendation-override",
        json={"lvc_threshold_override": "150"},
    )
    assert bad_threshold.status_code == 200
    assert bad_threshold.json()["state"]["lvc_threshold_override"] == ""

    bad_currency = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-THR/recommendation-override",
        json={"currency_mode": "btc"},
    )
    assert bad_currency.status_code == 400


def test_origin_sheet_actions_accept_large_ajax_forms():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Large origin form", "case_code": "CO-LARGE-FORM", "destination_market": "Ấn Độ", "invoice_no": "INV-LARGE"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-LARGE-FORM",
            "title": "Large origin form",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-LARGE", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-LARGE",
                    "name": "Large product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "lvc_status": "pass",
                    "lvc_status_label": "Đạt LVC",
                    "materials": [],
                }
            ],
            "origin_sheet_states": {"TP-LARGE": {"status": "calculated", "status_label": "Đã tính"}},
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {"aggregate_version_id": "bom-large", "aggregate_version_no": 1, "composition": []},
            "source_snapshot": {"client_config_hash": "snapshot"},
        },
    )
    data = {
        "case_id": case_id,
        "persisted_case_id": case_id,
        "customer": "Growatt",
        "case_code": "CO-LARGE-FORM",
        "title": "Large origin form",
        "destination_market": "Ấn Độ",
        "agreement": "",
        "co_form_type": "",
        "rule": "",
        "invoice_no": "INV-LARGE",
        "bill_of_lading_no": "",
        "mode": "Invoice + BCCT + BOM snapshot",
        "mode_note": "",
        "source_label": "",
        "document_count": "0",
        "product_count": "1",
        "origin_product_order": "TP-LARGE",
        "product_0_code": "TP-LARGE",
        "product_0_name": "Large product",
        "product_0_quantity": "1",
        "product_0_unit": "PCS",
        "product_0_currency": "VND",
        "product_0_fob": "100",
        "product_0_lvc_status": "pass",
        "product_0_lvc_status_label": "Đạt LVC",
        "product_0_origin_sheet_status": "calculated",
        "product_0_origin_sheet_status_label": "Đã tính",
        "product_0_material_count": "0",
        **{f"extra_field_{index}": str(index) for index in range(21000)},
    }

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-LARGE/lock",
        data=data,
    )
    legacy_multipart_response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-LARGE/lock",
        data=data,
        files={"_multipart_marker": ("marker.txt", b"x", "text/plain")},
    )

    assert response.status_code == 200
    assert "Đã chốt bảng kê TP-LARGE" in response.text
    assert legacy_multipart_response.status_code == 200


def test_origin_calculation_payload_returns_case_json():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Origin payload", "case_code": "CO-PAYLOAD", "destination_market": "Ấn Độ", "invoice_no": "INV-PAYLOAD"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-PAYLOAD",
            "title": "Origin payload",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-PAYLOAD", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-PAYLOAD",
                    "name": "Payload product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "materials": [{"material_code": "MAT-PAYLOAD", "consumed_qty": "1"}],
                }
            ],
            "origin_product_order": ["TP-PAYLOAD"],
            "origin_sheet_states": {"TP-PAYLOAD": {"status": "calculated", "status_label": "Đã tính"}},
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {"aggregate_artifact_id": "bom-payload", "aggregate_artifact_no": 1, "composition": []},
            "source_snapshot": {"client_config_hash": "snapshot"},
            "source_invoice_matches": [{"item_code": "TP-PAYLOAD", "quantity": "1"}],
        },
    )

    response = client.get(f"/clients/growatt/co-case/{case_id}/origin/calculation-payload")
    payload = response.json()

    assert response.status_code == 200
    assert payload["case_id"] == case_id
    assert payload["revision"]
    assert payload["origin_product_order"] == ["TP-PAYLOAD"]
    assert payload["origin_sheet_states"]["TP-PAYLOAD"]["status"] == "calculated"
    assert payload["products"][0]["materials"][0]["material_code"] == "MAT-PAYLOAD"
    assert payload["source"]["invoice_matches"][0]["item_code"] == "TP-PAYLOAD"


def test_origin_save_accepts_compact_json_and_marks_stale():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Origin save", "case_code": "CO-SAVE", "destination_market": "Ấn Độ", "invoice_no": "INV-SAVE"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-SAVE",
            "title": "Origin save",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-SAVE", "bill_of_lading_no": ""},
            "products": [{"code": "TP-A", "materials": []}, {"code": "TP-B", "materials": []}],
            "origin_product_order": ["TP-A", "TP-B"],
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính"},
            },
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {"aggregate_artifact_id": "bom-save", "aggregate_artifact_no": 1, "composition": []},
            "source_snapshot": {"client_config_hash": "snapshot"},
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/save",
        json={"origin_product_order": ["TP-B", "TP-A"], "stale_from_index": 0},
    )
    record = get_case_record(get_client("growatt"), case_id)

    assert response.status_code == 200
    assert response.json()["status"] == "ok"
    assert record["origin_product_order"] == ["TP-B", "TP-A"]
    assert record["origin_sheet_states"]["TP-B"]["status"] == "stale"
    assert record["origin_sheet_states"]["TP-A"]["status"] == "stale"


def test_origin_autosave_persists_full_workbook_state_without_dropping_sheets():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Workbook state", "case_code": "CO-WB", "destination_market": "Ấn Độ", "invoice_no": "INV-WB"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-WB",
            "title": "Workbook state",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-WB", "bill_of_lading_no": ""},
            "products": [{"code": "TP-A", "materials": []}, {"code": "TP-B", "materials": []}, {"code": "TP-C", "materials": []}],
            "origin_product_order": ["TP-A", "TP-B", "TP-C"],
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt", "currency_mode": "native"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính", "currency_mode": "native"},
                "TP-C": {"status": "draft", "status_label": "Chưa tính", "criteria_override": "keep me"},
            },
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/autosave",
        json={
            "origin_product_order": ["TP-B", "TP-A", "TP-C"],
            "stale_from_index": 3,
            "origin_sheet_states": {
                "TP-A": {"status": "locked", "status_label": "Chốt", "currency_mode": "native"},
                "TP-B": {"status": "calculated", "status_label": "Đã tính", "currency_mode": "vnd"},
            },
            "products": [
                {"code": "TP-A", "origin_sheet_status": "locked", "origin_sheet_status_label": "Chốt"},
                {"code": "TP-B", "origin_sheet_status": "calculated", "origin_sheet_status_label": "Đã tính"},
                {"code": "TP-C", "origin_sheet_status": "draft", "origin_sheet_status_label": "Chưa tính"},
            ],
        },
    )

    record = get_case_record(get_client("growatt"), case_id)
    assert response.status_code == 200
    assert response.json()["revision"]
    assert record["origin_product_order"] == ["TP-B", "TP-A", "TP-C"]
    assert record["origin_sheet_states"]["TP-A"]["status"] == "locked"
    assert record["origin_sheet_states"]["TP-B"]["status"] == "calculated"
    assert record["origin_sheet_states"]["TP-B"]["currency_mode"] == "vnd"
    assert record["origin_sheet_states"]["TP-C"]["criteria_override"] == "keep me"


def test_origin_sheet_lock_accepts_compact_json_without_source_refresh(monkeypatch):
    from app import main as main_module

    def fail_source_refresh(_client, _case):
        raise AssertionError("compact state-only origin actions must not refresh source context")

    monkeypatch.setattr(main_module, "co_case_source_context", fail_source_refresh)
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Compact lock", "case_code": "CO-COMPACT-LOCK", "destination_market": "Ấn Độ", "invoice_no": "INV-COMPACT"},
        follow_redirects=False,
    )
    case_id = created.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": case_id,
            "case_code": "CO-COMPACT-LOCK",
            "title": "Compact lock",
            "destination_market": "Ấn Độ",
            "shipment": {"invoice_no": "INV-COMPACT", "bill_of_lading_no": ""},
            "products": [
                {
                    "code": "TP-COMPACT",
                    "name": "Compact product",
                    "quantity": "1",
                    "unit": "PCS",
                    "fob": "100",
                    "currency": "VND",
                    "lvc_status": "pass",
                    "lvc_status_label": "Đạt LVC",
                    "materials": [],
                }
            ],
            "origin_product_order": ["TP-COMPACT"],
            "origin_sheet_states": {"TP-COMPACT": {"status": "calculated", "status_label": "Đã tính"}},
            "origin_snapshot": {"source": "invoice_bcct_bom", "readiness_label": "Sẵn sàng"},
            "bom_snapshot": {"aggregate_artifact_id": "bom-compact", "aggregate_artifact_no": 1, "composition": []},
            "source_snapshot": {"client_config_hash": "snapshot"},
        },
    )

    response = client.post(
        f"/clients/growatt/co-case/{case_id}/origin/sheet/TP-COMPACT/lock",
        json={"origin_product_order": ["TP-COMPACT"], "mark_stale": False},
    )

    assert response.status_code == 200
    assert "Đã chốt bảng kê TP-COMPACT" in response.text
    assert hidden_form_data(response.text)["product_0_origin_sheet_status"] == "locked"


def test_origin_sheet_calculate_accepts_compact_json_from_fresh_origin_page():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-COMPACT-CALC",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "10",
                        "unit": "PCE",
                        "customs_value": "100",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-COMPACT-CALC",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "1",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-COMPACT-CALC",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Compact calculate", "case_code": "CO-COMPACT-CALC", "destination_market": "Ấn Độ", "invoice_no": "INV-COMPACT-CALC"},
        follow_redirects=False,
    )
    origin = client.get(f"{created.headers['location']}/origin")
    form_data = hidden_form_data(origin.text)
    case_id = created.headers["location"].rstrip("/").split("/")[-1]

    response = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        json={
            "origin_product_order": ["PV00.0048500"],
            "products": [
                {
                    "code": "PV00.0048500",
                    "bom_product_code": form_data["product_0_bom_product_code"],
                    "bom_product_artifact_id": form_data.get("product_0_bom_product_artifact_id", ""),
                    "name": form_data["product_0_name"],
                    "finished_hs": form_data["product_0_finished_hs"],
                    "quantity": form_data["product_0_quantity"],
                    "unit": form_data["product_0_unit"],
                    "currency": form_data["product_0_currency"],
                    "fob": form_data["product_0_fob"],
                    "rvc_threshold": form_data["product_0_rvc_threshold"],
                    "origin_sheet_status": form_data["product_0_origin_sheet_status"],
                    "origin_sheet_status_label": form_data["product_0_origin_sheet_status_label"],
                }
            ],
            "mark_stale": False,
        },
    )
    record = get_case_record(get_client("growatt"), case_id)

    assert response.status_code == 200
    assert "Đã load BOM vào bảng kê PV00.0048500" in response.text
    assert record["products"][0]["materials"]
    assert record["origin_sheet_states"]["PV00.0048500"]["status"] == "calculated"


def test_origin_sheet_actions_follow_sequential_locking_rules():
    from app.main import (
        attach_origin_sheet_states,
        mark_origin_sheets_stale,
        origin_sheet_action_error,
        set_origin_sheet_status,
    )

    case = {
        "products": [{"code": "TP-1"}, {"code": "TP-2"}, {"code": "TP-3"}, {"code": "TP-4"}, {"code": "TP-5"}],
        "origin_sheet_states": {
            "TP-1": {"status": "locked"},
            "TP-2": {"status": "calculated"},
            "TP-3": {"status": "draft"},
            "TP-4": {"status": "draft"},
            "TP-5": {"status": "draft"},
        },
    }
    guarded = attach_origin_sheet_states(case)

    assert guarded["products"][2]["origin_can_calculate"] is False
    assert "TP-2" in origin_sheet_action_error(guarded, "TP-3", "calculate")

    guarded = set_origin_sheet_status(guarded, "TP-2", "locked")
    guarded = set_origin_sheet_status(guarded, "TP-3", "locked")
    guarded = set_origin_sheet_status(guarded, "TP-4", "calculated")

    assert origin_sheet_action_error(guarded, "TP-4", "lock") == ""
    guarded = set_origin_sheet_status(guarded, "TP-4", "locked")
    assert guarded["products"][2]["origin_can_reopen"] is False
    assert "TP-4" in origin_sheet_action_error(guarded, "TP-3", "reopen")
    assert origin_sheet_action_error(guarded, "TP-4", "reopen") == ""

    released = mark_origin_sheets_stale(guarded, 0)
    assert [product["origin_sheet_status"] for product in released["products"]] == ["stale", "stale", "stale", "stale", "draft"]


def test_origin_product_order_override_changes_sequential_allocation():
    from app.main import prepare_case_origin_products

    case = prepare_case_origin_products(
        {"shipment": {"invoice_no": "INV-SEQ"}, "origin_product_order": ["TP-B", "TP-A"]},
        [
            {
                "item_code": "TP-A",
                "description": "Finished product A",
                "hs_code": "850440",
                "quantity": "4",
                "unit": "PCS",
                "customs_value": "1000",
                "value_currency": "VND",
            },
            {
                "item_code": "TP-B",
                "description": "Finished product B",
                "hs_code": "850440",
                "quantity": "3",
                "unit": "PCS",
                "customs_value": "1000",
                "value_currency": "VND",
            },
        ],
        {
            "latest_version": {
                "version_id": "bom-seq",
                "rows": [
                    {"product_code": "TP-A", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
                    {"product_code": "TP-B", "material_code": "MAT-SHARED", "qty_per": "1", "uom": "PCS"},
                ],
            },
            "versions": [],
            "product_versions": [],
        },
        {"form_code": "B", "display_name": "C/O form B"},
        [{"customs_code": "MAT-SHARED", "origin_default": "Không xuất xứ"}],
        [
            {
                "source_row": "SEQ-STOCK-1",
                "import_declaration_no": "NK-SEQ",
                "line_no": "1",
                "material_code": "MAT-SHARED",
                "remaining_qty": "5",
                "unit_value": "10",
                "currency": "VND",
                "value_currency": "VND",
                "eligibility_status": "active",
                "allocation_code_status": "resolved",
            }
        ],
    )

    product_b, product_a = case["products"]
    material_b = product_b["materials"][0]
    material_a = product_a["materials"][0]

    assert [product["code"] for product in case["products"]] == ["TP-B", "TP-A"]
    assert case["origin_snapshot"]["product_order"] == ["TP-B", "TP-A"]
    assert material_b["allocation_status"] == "covered"
    assert material_b["allocation_lines"][0]["allocated_qty"] == "3"
    assert material_a["allocation_status"] == "shortage"
    assert material_a["allocation_shortage_qty"] == "2"
    assert "Bước 1 TP-B dùng 3 PCS" in material_a["allocation_shortage_trace"]


def test_origin_material_blocks_mixed_currency_allocation_value():
    from app.main import co_stock_allocation_pool, origin_material_from_bom_row

    pool = co_stock_allocation_pool([
        {
            "source_row": "LOT-USD",
            "import_declaration_no": "NK-CUR-1",
            "line_no": "1",
            "material_code": "MAT-CUR",
            "remaining_qty": "1",
            "unit_value": "10",
            "currency": "USD",
            "value_currency": "USD",
            "eligibility_status": "active",
            "allocation_code_status": "resolved",
        },
        {
            "source_row": "LOT-VND",
            "import_declaration_no": "NK-CUR-2",
            "line_no": "2",
            "material_code": "MAT-CUR",
            "remaining_qty": "1",
            "unit_value": "200000",
            "currency": "VND",
            "value_currency": "VND",
            "eligibility_status": "active",
            "allocation_code_status": "resolved",
        },
    ])

    material = origin_material_from_bom_row(
        {"material_code": "MAT-CUR", "qty_per": "2", "uom": "PCS"},
        Decimal("1"),
        {"MAT-CUR": {"origin_default": "Không xuất xứ"}},
        pool,
    )

    assert material["allocation_status"] == "covered"
    assert material["valuation_status"] == "partial_valuation"
    assert material["currency"] == "Nhiều tiền tệ"
    assert material["material_value"] == ""
    assert material["non_origin_cif_value"] == ""
    assert "MAT-CUR: nhiều tiền tệ trong các dòng tồn, chưa cộng VNM tự động." in material["material_warnings"]


def test_co_case_export_workbook_contains_bom_snapshot_rows_from_origin_form():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-BOM-XLSX-1",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "1000",
                    },
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-BOM-XLSX-2",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-002",
                        "description": "Connector set",
                        "hs_code": "8536.90",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "2000",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-BOM-XLSX",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "2",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "invoice_ref": "INV-BOM-XLSX",
                    }
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "BOM export", "case_code": "CO-BOM-XLSX", "destination_market": "Ấn Độ", "invoice_no": "INV-BOM-XLSX"},
        follow_redirects=False,
    )
    origin = client.get(f"{created.headers['location']}/origin")

    # Sheets default to "Chưa tính"; calling /calculate flips them to "Đã tính"
    # which is required for export. Mirrors the manual UI gate (staff clicks
    # "Tính bảng kê").
    form_data = hidden_form_data(origin.text)
    client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        data=form_data,
    )

    origin = client.get(f"{created.headers['location']}/origin")
    response = client.post(f"{created.headers['location']}/export", data=hidden_form_data(origin.text))

    assert response.status_code == 200
    workbook = load_workbook(BytesIO(response.content))
    assert "LVC Statement" in workbook.sheetnames
    values = [cell.value for row in workbook["LVC Statement"].iter_rows(values_only=False) for cell in row]
    assert "PV00.0048500" in values
    assert "DEMO-NPL-001" in values
    assert "91.00" in values
    assert "90" in values


def test_co_case_export_workbook_contains_origin_snapshot_metadata_from_web():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {"direction": "import", "declaration_type": "E11", "declaration_no": "NK-ORIGIN-XLSX-1", "line_no": "1", "item_code": "DEMO-NPL-001", "description": "Main control board", "hs_code": "8542.39", "quantity": "100", "unit": "PCE", "customs_value": "1000", "currency": "VND"},
                    {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-ORIGIN-XLSX", "line_no": "1", "item_code": "PV00.0048500", "description": "Growatt inverter", "hs_code": "850440", "quantity": "3", "unit": "PCS", "customs_value": "1000", "currency": "VND", "invoice_ref": "INV-ORIGIN-XLSX"},
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Origin export", "case_code": "CO-ORIGIN-XLSX", "destination_market": "Ấn Độ", "invoice_no": "INV-ORIGIN-XLSX"},
        follow_redirects=False,
    )
    origin = client.get(f"{created.headers['location']}/origin")

    # Mark each sheet calculated (mirrors UI "Tính bảng kê" gate).
    form_data = hidden_form_data(origin.text)
    for product_code in ("PV00.0048500", "DEMO-NPL-002"):
        client.post(
            f"{created.headers['location']}/origin/sheet/{product_code}/calculate",
            data=form_data,
        )
    origin = client.get(f"{created.headers['location']}/origin")
    response = client.post(f"{created.headers['location']}/export", data=hidden_form_data(origin.text))

    assert response.status_code == 200
    workbook = load_workbook(BytesIO(response.content))
    assert "Origin Snapshot" in workbook.sheetnames
    values = [cell.value for row in workbook["Origin Snapshot"].iter_rows(values_only=False) for cell in row]
    assert "build_down_lvc" in values
    assert "blocked" in values
    assert "DEMO-NPL-002: thiếu đơn giá để tính trị giá NVL/VNM." in values
    assert "CTSH preview" in values


def test_co_case_origin_round_trips_multi_lot_allocation_to_export_workbook():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-ALLOC-1",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board dòng tồn 1",
                        "hs_code": "8542.39",
                        "quantity": "1",
                        "unit": "PCE",
                        "customs_value": "10",
                        "currency": "VND",
                    },
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-ALLOC-2",
                        "line_no": "2",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board dòng tồn 2",
                        "hs_code": "8542.39",
                        "quantity": "2",
                        "unit": "PCE",
                        "customs_value": "40",
                        "currency": "VND",
                    },
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-ALLOC-3",
                        "line_no": "3",
                        "item_code": "DEMO-NPL-002",
                        "description": "Connector set",
                        "hs_code": "8536.90",
                        "quantity": "100",
                        "unit": "PCE",
                        "customs_value": "2000",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-ALLOC",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "3",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-ALLOC",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Allocation export", "case_code": "CO-ALLOC", "destination_market": "Ấn Độ", "invoice_no": "INV-ALLOC"},
        follow_redirects=False,
    )

    origin = client.get(f"{created.headers['location']}/origin")
    form_data = hidden_form_data(origin.text)
    # Mark each sheet calculated (mirrors UI "Tính bảng kê" gate).
    for product_code in ("PV00.0048500", "DEMO-NPL-001"):
        client.post(
            f"{created.headers['location']}/origin/sheet/{product_code}/calculate",
            data=form_data,
        )
    origin = client.get(f"{created.headers['location']}/origin")
    form_data = hidden_form_data(origin.text)
    response = client.post(f"{created.headers['location']}/export", data=form_data)

    assert origin.status_code == 200
    assert 'data-origin-product-order' in origin.text
    assert 'data-origin-sheet-tab' in origin.text
    assert 'data-origin-sequence-move="left"' in origin.text
    assert 'data-origin-sequence-move="right"' in origin.text
    assert 'data-origin-tab-drag-handle' not in origin.text
    assert 'draggable="true"' not in origin.text
    assert 'data-origin-sequence-position' not in origin.text
    assert 'data-origin-sequence-move="up"' not in origin.text
    assert 'class="origin-sheet-toolbar"' in origin.text
    assert 'origin-sheet-status-pill' in origin.text
    assert 'data-origin-step-input' in origin.text
    assert 'data-origin-sheet-calculate' in origin.text
    assert 'data-origin-export-action' in origin.text
    assert "Bước 1" in origin.text
    assert form_data["product_0_material_0_allocation_line_count"] == "2"
    assert form_data["origin_product_order"] == "PV00.0048500"
    assert form_data["product_0_allocation_sequence"] == "1"
    assert form_data["product_0_material_0_material_sequence"] == "1"
    assert form_data["product_0_material_0_allocation_0_import_declaration_no"] == "NK-ALLOC-1"
    assert form_data["product_0_material_0_allocation_1_import_declaration_no"] == "NK-ALLOC-2"
    assert form_data["product_0_material_0_allocation_0_product_sequence"] == "1"
    assert form_data["product_0_material_0_allocation_0_opening_qty"] == "1"
    assert form_data["product_0_material_0_allocation_0_allocated_qty"] == "1"
    assert form_data["product_0_material_0_allocation_1_allocated_qty"] == "2"
    assert form_data["product_0_material_0_material_value"] == "50"
    assert "2 dòng tồn" in origin.text
    assert "Dòng tồn 1" in origin.text
    assert "tồn trước" in origin.text
    assert 'data-allocation-toggle' in origin.text
    assert 'aria-expanded="false"' in origin.text
    assert 'data-allocation-detail' in origin.text
    assert "const initOriginAllocationToggles" in origin.text
    assert "initOriginAllocationToggles(root);" in origin.text
    assert 'class="origin-allocation-row"' in origin.text
    assert "hidden" in origin.text
    assert "origin-material-name" in origin.text
    assert "origin-source-cell" in origin.text
    assert "source-chip" in origin.text
    assert 'data-origin-column-controls' in origin.text
    assert 'data-origin-column-toggle="hs"' in origin.text
    assert 'data-origin-column-toggle="source"' in origin.text
    assert 'data-origin-column="source"' in origin.text
    assert "NK-ALLOC-1 / line 1" in origin.text
    assert "NK-ALLOC-2 / line 2" in origin.text
    assert "NK-ALLOC-3 / line 3" in origin.text
    assert response.status_code == 200

    workbook = load_workbook(BytesIO(response.content))
    lvc_rows = list(workbook["LVC Statement"].iter_rows(values_only=True))
    snapshot_values = [value for row in workbook["Origin Snapshot"].iter_rows(values_only=True) for value in row]
    assert "Allocation source row" in lvc_rows[0]
    assert "Product sequence" in lvc_rows[0]
    assert "Allocation opening qty" in lvc_rows[0]
    declaration_index = lvc_rows[0].index("Allocation declaration")
    qty_index = lvc_rows[0].index("Allocated qty")
    opening_qty_index = lvc_rows[0].index("Allocation opening qty")
    allocation_rows = {
        row[declaration_index]: row
        for row in lvc_rows[1:]
        if row[declaration_index] in {"NK-ALLOC-1", "NK-ALLOC-2"}
    }
    assert allocation_rows["NK-ALLOC-1"][qty_index] == "1"
    assert allocation_rows["NK-ALLOC-1"][opening_qty_index] == "1"
    assert allocation_rows["NK-ALLOC-2"][qty_index] == "2"
    assert "allocation" in snapshot_values
    assert "NK-ALLOC-1 / line 1" in snapshot_values
    assert "NK-ALLOC-2 / line 2" in snapshot_values

    autosaved = client.post(f"{created.headers['location']}/origin/autosave", data={**form_data, "stale_from_index": "0"})
    stale_origin = client.get(f"{created.headers['location']}/origin")
    stale_form_data = hidden_form_data(stale_origin.text)
    blocked_export = client.post(f"{created.headers['location']}/export", data=stale_form_data)
    calculated = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/calculate",
        data=stale_form_data,
    )
    locked = client.post(
        f"{created.headers['location']}/origin/sheet/PV00.0048500/lock",
        data=hidden_form_data(calculated.text),
    )

    assert autosaved.status_code == 200
    assert autosaved.json()["status"] == "ok"
    assert "Cần tính lại" in stale_origin.text
    assert stale_form_data["product_0_origin_sheet_status"] == "stale"
    assert blocked_export.status_code == 409
    assert "Chưa thể export" in blocked_export.text
    assert calculated.status_code == 200
    assert "Đã load BOM vào bảng kê PV00.0048500" in calculated.text
    assert hidden_form_data(calculated.text)["product_0_origin_sheet_status"] == "calculated"
    assert locked.status_code == 200
    assert "Đã chốt bảng kê PV00.0048500" in locked.text
    assert hidden_form_data(locked.text)["product_0_origin_sheet_status"] == "locked"


def test_origin_calculation_lock_blocks_parallel_cases_for_same_client():
    client = TestClient(app)
    client.post(
        "/clients/growatt/bcct/upload",
        files={
            "file": (
                "bcct.xlsx",
                bcct_workbook([
                    {
                        "direction": "import",
                        "declaration_type": "E11",
                        "declaration_no": "NK-LOCK",
                        "line_no": "1",
                        "item_code": "DEMO-NPL-001",
                        "description": "Main control board",
                        "hs_code": "8542.39",
                        "quantity": "20",
                        "unit": "PCE",
                        "customs_value": "200",
                        "currency": "VND",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-LOCK-1",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "1",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-LOCK-1",
                    },
                    {
                        "direction": "export",
                        "declaration_type": "E42",
                        "declaration_no": "XK-LOCK-2",
                        "line_no": "1",
                        "item_code": "PV00.0048500",
                        "description": "Growatt inverter",
                        "hs_code": "850440",
                        "quantity": "1",
                        "unit": "PCS",
                        "customs_value": "1000",
                        "currency": "VND",
                        "invoice_ref": "INV-LOCK-2",
                    },
                ]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    first = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Lock one", "case_code": "CO-LOCK-1", "destination_market": "Ấn Độ", "invoice_no": "INV-LOCK-1"},
        follow_redirects=False,
    )
    second = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Lock two", "case_code": "CO-LOCK-2", "destination_market": "Ấn Độ", "invoice_no": "INV-LOCK-2"},
        follow_redirects=False,
    )
    first_origin = client.get(f"{first.headers['location']}/origin")
    second_origin = client.get(f"{second.headers['location']}/origin")

    first_recalculate = client.post("/clients/growatt/evaluate", data=hidden_form_data(first_origin.text))
    blocked_recalculate = client.post("/clients/growatt/evaluate", data=hidden_form_data(second_origin.text))
    blocked_export = client.post(f"{second.headers['location']}/export", data=hidden_form_data(second_origin.text))
    index = client.get("/clients/growatt/co-case")

    assert first_recalculate.status_code == 200
    assert "Hồ sơ này đang giữ phiên tính tồn" in first_recalculate.text
    assert blocked_recalculate.status_code == 409
    assert "Chưa thể tính lại" in blocked_recalculate.text
    assert "CO-LOCK-1" in blocked_recalculate.text
    assert blocked_export.status_code == 409
    assert "Chưa thể export" in blocked_export.text
    assert "Nhả phiên" in index.text
    assert "Đang giữ tồn" in index.text
    assert "Chỉ chuẩn bị" in index.text

    released = client.post(
        f"{first.headers['location']}/origin-lock/release",
        data={"next_url": "/clients/growatt/co-case"},
        follow_redirects=False,
    )
    second_recalculate = client.post("/clients/growatt/evaluate", data=hidden_form_data(second_origin.text))

    assert released.status_code == 303
    assert released.headers["location"] == "/clients/growatt/co-case"
    assert second_recalculate.status_code == 200
    assert "Hồ sơ này đang giữ phiên tính tồn" in second_recalculate.text


def test_co_case_delete_only_allows_draft_unlocked_cases():
    client = TestClient(app)
    draft = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Delete draft", "case_code": "CO-DELETE-DRAFT", "destination_market": "Ấn Độ", "invoice_no": "INV-DELETE-DRAFT"},
        follow_redirects=False,
    )
    locked = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Delete locked", "case_code": "CO-DELETE-LOCK", "destination_market": "Ấn Độ", "invoice_no": "INV-DELETE-LOCK"},
        follow_redirects=False,
    )
    completed = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Delete completed", "case_code": "CO-DELETE-DONE", "destination_market": "Ấn Độ", "invoice_no": "INV-DELETE-DONE"},
        follow_redirects=False,
    )
    completed_case_id = completed.headers["location"].rstrip("/").split("/")[-1]
    update_case_record(
        get_client("growatt"),
        {
            "persisted_case_id": completed_case_id,
            "status": "completed",
            "shipment": {"invoice_no": "INV-DELETE-DONE", "bill_of_lading_no": ""},
        },
    )
    locked_origin = client.get(f"{locked.headers['location']}/origin")
    client.post("/clients/growatt/evaluate", data=hidden_form_data(locked_origin.text))
    index = client.get("/clients/growatt/co-case")

    blocked_locked = client.post(f"{locked.headers['location']}/delete")
    blocked_completed = client.post(f"{completed.headers['location']}/delete")
    deleted_draft = client.post(f"{draft.headers['location']}/delete", follow_redirects=False)
    after_delete = client.get("/clients/growatt/co-case")

    assert index.status_code == 200
    assert "Xoá" in index.text
    assert "data-delete-case-modal" in index.text
    assert "return confirm(" not in index.text
    assert "Hồ sơ đang giữ phiên tính tồn" in index.text
    assert "Hồ sơ đã hoàn tất" in index.text
    assert blocked_locked.status_code == 409
    assert "Hồ sơ đang giữ phiên tính tồn" in blocked_locked.text
    assert blocked_completed.status_code == 409
    assert "Hồ sơ đã hoàn tất" in blocked_completed.text
    assert deleted_draft.status_code == 303
    assert deleted_draft.headers["location"] == "/clients/growatt/co-case"
    assert "CO-DELETE-DRAFT" not in after_delete.text


def test_co_case_supporting_upload_saves_invoice_metadata_and_matches_bcct_exports():
    client = TestClient(app)
    upload = bcct_workbook([
        {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-001", "description": "Finished good", "hs_code": "850440", "quantity": "10", "unit": "PCS", "invoice_ref": "INV-42"},
        {"direction": "import", "declaration_type": "E11", "declaration_no": "NK-001", "line_no": "1", "item_code": "MAT-001", "description": "Material", "hs_code": "853690", "quantity": "100", "unit": "PCS", "invoice_ref": "INV-42"},
        {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-002", "line_no": "1", "item_code": "TP-OTHER", "description": "Other finished good", "hs_code": "850440", "quantity": "5", "unit": "PCS", "invoice_ref": "INV-99"},
    ])
    client.post(
        "/clients/growatt/bcct/upload",
        files={"file": ("bcct.xlsx", upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Invoice lookup", "case_code": "CO-INV-42", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    location = created.headers["location"]

    response = client.post(
        f"{location}/supporting-files",
        data={"document_slot": "invoice", "invoice_no": "INV-42", "bill_of_lading_no": "BL-42"},
        files={"file": ("invoice-INV-42.pdf", b"%PDF-1.4 invoice", "application/pdf")},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == f"{location}/documents"
    documents = client.get(f"{location}/documents")
    exports = client.get(f"{location}/exports")
    assert "invoice-INV-42.pdf" in documents.text
    download_href = re.search(r'href="([^"]+/supporting-files/[^"]+)"', documents.text)
    assert download_href is not None
    downloaded = client.get(download_href.group(1))
    assert downloaded.status_code == 200
    assert downloaded.content == b"%PDF-1.4 invoice"
    assert "INV-42" in documents.text
    assert "BL-42" in documents.text
    assert "TP-001" in exports.text
    assert "XK-001" in exports.text
    assert "TKX_XK-001.zip" in exports.text
    assert "/clients/growatt/declarations/download.zip?direction=export" in exports.text
    assert "filename=TKX_XK-001.zip" in exports.text
    assert "MAT-001" not in exports.text
    assert "TP-OTHER" not in exports.text


def test_co_case_guidance_maps_invoice_bcct_products_to_form_instrument_and_hs_criteria():
    """The guidance workflow step was removed (2026-05-28). What remains is
    that the shipment page still surfaces the recommended form + instrument
    based on destination market — verified here. The criteria preview that
    used to live on /guidance is gone; if/when PSR confirmation comes back,
    re-add the assertions."""
    client = TestClient(app)
    upload = bcct_workbook([
        {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-PSR", "line_no": "1", "item_code": "PV00.0048500", "description": "Growatt inverter", "hs_code": "850440", "quantity": "12", "unit": "PCS", "invoice_ref": "INV-PSR"},
    ])
    client.post(
        "/clients/growatt/bcct/upload",
        files={"file": ("bcct.xlsx", upload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Criteria lookup", "case_code": "CO-PSR", "destination_market": "Canada", "invoice_no": "INV-PSR"},
        follow_redirects=False,
    )
    location = created.headers["location"]

    index = client.get("/clients/growatt/co-case")
    shipment = client.get(location)
    guidance = client.get(f"{location}/guidance")

    assert 'role="combobox"' in index.text
    assert "Form CPTPP" in shipment.text
    assert "03/2019/TT-BCT" in shipment.text
    assert "PV00.0048500" in shipment.text
    assert "850440" in shipment.text
    assert guidance.status_code == 404


def test_co_case_state_prefers_postgres_store_and_keeps_supporting_file_metadata(monkeypatch):
    from app import co_case_store

    saved_states = []

    class FakeCoCaseStateStore:
        def __init__(self):
            self.state = None

        def get_state(self, client_id: str) -> dict | None:
            assert client_id == "growatt"
            return self.state

        def save_state(self, client_id: str, state: dict) -> None:
            assert client_id == "growatt"
            self.state = dict(state)
            saved_states.append(dict(state))

    fake_store = FakeCoCaseStateStore()
    monkeypatch.setattr(co_case_store, "get_co_case_state_store", lambda: fake_store)
    client = get_client("growatt")
    record = co_case_store.create_case_record(
        client,
        {"title": "Postgres C/O", "case_code": "CO-PG", "destination_market": "Ấn Độ"},
    )

    file_row = co_case_store.save_supporting_file(
        client,
        record["case_id"],
        b"%PDF-1.4 invoice",
        "invoice-CO-PG.pdf",
        "invoice",
        "INV-PG",
        "BL-PG",
    )

    assert saved_states
    assert fake_store.state["cases"][0]["case_id"] == record["case_id"]
    assert file_row["original_filename"] == "invoice-CO-PG.pdf"
    assert file_row["stored_filename"].startswith(file_row["upload_id"])
    assert file_row["storage_backend"] == "filesystem"
    assert file_row["content_sha256"] == hashlib.sha256(b"%PDF-1.4 invoice").hexdigest()
    assert fake_store.state["cases"][0]["supporting_files"][0]["content_sha256"] == file_row["content_sha256"]
    assert not co_case_store.state_path("growatt").exists()


def test_co_case_evaluate_keeps_persisted_dossier_supporting_metadata():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Evaluate dossier", "case_code": "CO-EVAL", "destination_market": "Ấn Độ", "invoice_no": "INV-OLD"},
        follow_redirects=False,
    )
    location = created.headers["location"]
    case_id = location.rsplit("/", 1)[-1]
    client.post(
        f"{location}/supporting-files",
        data={"document_slot": "invoice", "invoice_no": "INV-OLD", "bill_of_lading_no": "BL-OLD"},
        files={"file": ("invoice-INV-OLD.pdf", b"%PDF-1.4 invoice", "application/pdf")},
        follow_redirects=False,
    )

    response = client.post(
        "/clients/growatt/evaluate",
        data={
            "case_id": case_id,
            "persisted_case_id": case_id,
            "customer": "Growatt",
            "case_code": "CO-EVAL",
            "title": "Evaluate dossier",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-NEW",
            "bill_of_lading_no": "BL-NEW",
            "agreement": "AIFTA",
            "co_form_type": "Form AI",
            "rule": "Cần tra cứu PSR theo HS",
            "document_count": "0",
            "product_count": "0",
        },
    )

    assert response.status_code == 200
    assert "INV-NEW" in response.text
    assert "BL-NEW" in response.text
    detail = client.get(f"{location}/documents")
    assert "invoice-INV-OLD.pdf" in detail.text
    assert "INV-NEW" in detail.text
    assert "BL-NEW" in detail.text


def test_invoice_matching_uses_only_reviewed_export_rows():
    matches = match_case_bcct_exports(
        {"shipment": {"invoice_no": "INV-REVIEW"}},
        {
            "bcct": {
                "published_rows": [
                    {"direction": "export", "review_status": "correction_candidate", "declaration_no": "XK-DRAFT", "line_no": "1", "declaration_type": "E42", "item_code": "TP-DRAFT", "quantity": "1", "unit": "PCS", "invoice_ref": "INV-REVIEW"},
                    {"direction": "export", "review_status": "reviewed", "declaration_no": "XK-OK", "line_no": "1", "declaration_type": "E42", "item_code": "TP-OK", "quantity": "1", "unit": "PCS", "invoice_ref": "INV-REVIEW"},
                    {"direction": "import", "review_status": "reviewed", "declaration_no": "NK-OK", "line_no": "1", "declaration_type": "E11", "item_code": "MAT-OK", "quantity": "1", "unit": "PCS", "invoice_ref": "INV-REVIEW"},
                ]
            }
        },
        {"bcct": {"relevant_export_declaration_types": ["E42"]}},
    )

    assert [row["item_code"] for row in matches] == ["TP-OK"]


def test_invoice_matching_prefers_export_declaration_when_invoice_is_missing():
    matches = match_case_bcct_exports(
        {"shipment": {"invoice_no": "", "export_declaration_nos": ["XK-NO-INV"]}},
        {
            "bcct": {
                "published_rows": [
                    {"direction": "export", "review_status": "reviewed", "declaration_no": "XK-NO-INV", "line_no": "1", "declaration_type": "E42", "item_code": "TP-DECL", "quantity": "1", "unit": "PCS", "invoice_ref": ""},
                ]
            }
        },
        {"bcct": {"relevant_export_declaration_types": ["E42"]}},
    )

    assert [row["item_code"] for row in matches] == ["TP-DECL"]
    assert matches[0]["match_source"] == "declaration"


def test_invoice_matching_prefers_export_declaration_and_warns_on_invoice_mismatch():
    matches = match_case_bcct_exports(
        {"shipment": {"invoice_no": "INV-WRONG", "export_declaration_nos": ["XK-RIGHT"]}},
        {
            "bcct": {
                "published_rows": [
                    {"direction": "export", "review_status": "reviewed", "declaration_no": "XK-RIGHT", "line_no": "1", "declaration_type": "E42", "item_code": "TP-RIGHT", "quantity": "1", "unit": "PCS", "invoice_ref": "INV-RIGHT"},
                    {"direction": "export", "review_status": "reviewed", "declaration_no": "XK-WRONG", "line_no": "1", "declaration_type": "E42", "item_code": "TP-WRONG", "quantity": "1", "unit": "PCS", "invoice_ref": "INV-WRONG"},
                ]
            }
        },
        {"bcct": {"relevant_export_declaration_types": ["E42"]}},
    )

    assert [row["item_code"] for row in matches] == ["TP-RIGHT"]
    assert matches[0]["invoice_mismatch"] is True
    assert "INV-WRONG" in matches[0]["reference_warning"]
    assert "invoice_ref INV-RIGHT" in matches[0]["reference_warning"]
    assert "tờ khai XK-RIGHT" in matches[0]["reference_warning"]


def test_invoice_lookup_options_accept_export_declaration_no():
    from app.main import invoice_search_options

    client = get_client("growatt")
    process_bcct_upload(
        client,
        bcct_workbook([
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-DECL-001", "line_no": "1", "item_code": "TP-001", "quantity": "2", "unit": "PCS", "invoice_ref": "INV-DECL-001"},
        ]),
        "bcct.xlsx",
    )

    options = invoice_search_options(client, "XK-DECL-001")

    assert options[0]["invoice_no"] == "INV-DECL-001"
    assert options[0]["declarations"] == ["XK-DECL-001"]
    preview = TestClient(app).get(
        "/clients/growatt/co-case/invoice-preview",
        params={"invoice_no": "XK-DECL-001"},
    ).json()
    assert preview["invoice_no"] == "INV-DECL-001"
    assert preview["source_reference"] == "XK-DECL-001"


def test_co_case_create_accepts_export_declaration_no():
    client_data = get_client("growatt")
    process_bcct_upload(
        client_data,
        bcct_workbook([
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-CREATE-001", "line_no": "1", "item_code": "TP-001", "quantity": "2", "unit": "PCS", "invoice_ref": "INV-CREATE-001"},
        ]),
        "bcct.xlsx",
    )
    client = TestClient(app)

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Declaration create", "case_code": "CO-DECL-CREATE", "destination_market": "Ấn Độ", "invoice_no": "XK-CREATE-001"},
        follow_redirects=False,
    )
    page = client.get(created.headers["location"])

    assert created.status_code == 303
    assert "Tờ khai XK-CREATE-001" in page.text
    assert "INV-CREATE-001" in page.text
    assert "CO-DECL-CREATE" in page.text


def test_co_case_create_accepts_export_declaration_without_invoice_ref():
    client_data = get_client("growatt")
    process_bcct_upload(
        client_data,
        bcct_workbook([
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-NO-INVOICE", "line_no": "1", "item_code": "TP-001", "quantity": "2", "unit": "PCS", "invoice_ref": ""},
        ]),
        "bcct.xlsx",
    )
    client = TestClient(app)

    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Declaration only", "case_code": "CO-DECL-ONLY", "destination_market": "Ấn Độ", "export_declaration_nos": "XK-NO-INVOICE"},
        follow_redirects=False,
    )
    page = client.get(f"{created.headers['location']}/exports")

    assert created.status_code == 303
    assert "Tờ khai XK-NO-INVOICE" in page.text
    assert "TP-001" in page.text
    assert "không có invoice_ref" in page.text


def test_invoice_preview_warns_when_invoice_and_export_declaration_disagree():
    client_data = get_client("growatt")
    process_bcct_upload(
        client_data,
        bcct_workbook([
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-RIGHT", "line_no": "1", "item_code": "TP-RIGHT", "quantity": "2", "unit": "PCS", "invoice_ref": "INV-RIGHT"},
        ]),
        "bcct.xlsx",
    )

    preview = TestClient(app).get(
        "/clients/growatt/co-case/invoice-preview",
        params={"invoice_no": "INV-WRONG", "export_declaration_nos": "XK-RIGHT"},
    ).json()

    assert preview["status"] == "found"
    assert preview["match_count"] == 1
    assert preview["reference_warnings"] == [
        "Invoice nhập INV-WRONG không khớp invoice_ref INV-RIGHT trên tờ khai XK-RIGHT."
    ]


def test_co_case_shipment_page_surfaces_invoice_declaration_warning():
    client_data = get_client("growatt")
    process_bcct_upload(
        client_data,
        bcct_workbook([
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-RIGHT", "line_no": "1", "item_code": "TP-RIGHT", "quantity": "2", "unit": "PCS", "invoice_ref": "INV-RIGHT"},
        ]),
        "bcct.xlsx",
    )
    client = TestClient(app)

    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Mismatch visible",
            "case_code": "CO-MISMATCH-VISIBLE",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-WRONG",
            "export_declaration_nos": "XK-RIGHT",
        },
        follow_redirects=False,
    )
    page = client.get(created.headers["location"])

    assert created.status_code == 303
    assert "Kiểm tra invoice / tờ khai xuất" in page.text
    assert "Invoice nhập INV-WRONG không khớp invoice_ref INV-RIGHT trên tờ khai XK-RIGHT." in page.text


def test_co_case_page_uses_lightweight_source_summary(monkeypatch):
    from app import portfolio as portfolio_module
    from app import source_store as source_store_module

    def fail_full_workspace_load(*_args, **_kwargs):
        raise AssertionError("C/O pages should not load the full source workspace.")

    monkeypatch.setattr(portfolio_module, "get_source_workspace", fail_full_workspace_load)
    monkeypatch.setattr(source_store_module, "get_source_workspace", fail_full_workspace_load)

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Indexed source case",
            "case_code": "CO-INDEX",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-INDEX",
            "bill_of_lading_no": "BL-INDEX",
        },
        follow_redirects=False,
    )

    response = client.get(created.headers["location"])

    assert response.status_code == 200
    assert "CO-INDEX" in response.text


def test_source_summary_exposes_counts_and_snapshot_metadata():
    client = get_client("do-thanh")
    process_bcct_upload(
        client,
        bcct_workbook([
            {"direction": "import", "declaration_no": "NK-001", "line_no": "1", "item_code": "MAT-001", "quantity": "10", "unit": "PCS"},
            {"direction": "export", "declaration_type": "E42", "declaration_no": "XK-001", "line_no": "1", "item_code": "TP-001", "quantity": "2", "unit": "PCS", "invoice_ref": "INV-001"},
        ]),
        "bcct.xlsx",
    )

    summary = get_source_summary(client)

    assert summary["bcct"]["published_row_count"] == 2
    assert summary["bcct"]["reviewed_row_count"] == 2
    assert summary["co_stock_row_count"] == 1
    assert summary["bcct"]["latest_version"]["version_no"] == 1


def test_co_case_page_uses_postgres_source_index_when_available(monkeypatch):
    from app import portfolio as portfolio_module

    class FakeSourceIndexStore:
        def has_client(self, client_id: str) -> bool:
            return client_id == "growatt"

        def source_summary(self, _client_id: str, client_config: dict) -> dict:
            return {
                "client_config": client_config,
                "material_catalog": {"published_row_count": 10, "latest_version": {"version_no": 2}},
                "product_catalog": {"published_row_count": 3, "latest_version": {"version_no": 4}},
                "bcct": {
                    "published_row_count": 20,
                    "reviewed_row_count": 9,
                    "correction_candidate_count": 0,
                    "latest_version": {"version_no": 5},
                },
                "co_stock_row_count": 8,
            }

        def match_bcct_exports(self, _client_id: str, invoice_no: str, _relevant_types: list[str]) -> list[dict]:
            assert invoice_no == "INV-PG"
            return [
                {
                    "declaration_no": "XK-PG",
                    "line_no": "1",
                    "declaration_type": "E42",
                    "item_code": "TP-PG",
                    "hs_code": "8504.40",
                    "quantity": "2",
                    "unit": "PCS",
                    "invoice_ref": "INV-PG",
                }
            ]

    def fail_file_source_load(*_args, **_kwargs):
        raise AssertionError("Postgres-indexed C/O pages should not load source JSON.")

    monkeypatch.setattr(portfolio_module, "get_source_index_store", lambda: FakeSourceIndexStore())
    monkeypatch.setattr(portfolio_module, "load_module_state", fail_file_source_load)

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Postgres indexed case",
            "case_code": "CO-PG",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-PG",
        },
        follow_redirects=False,
    )

    response = client.get(f"{created.headers['location']}/exports")

    assert response.status_code == 200
    assert "XK-PG" in response.text
    assert "TP-PG" in response.text
    assert "20 dòng BCCT" in response.text


def test_postgres_source_index_signature_check_does_not_swallow_type_errors(monkeypatch):
    from app import portfolio as portfolio_module

    class FakeSourceIndexStore:
        def has_client(self, client_id: str) -> bool:
            return client_id == "growatt"

        def source_summary(self, _client_id: str, client_config: dict) -> dict:
            return {
                "client_config": client_config,
                "material_catalog": {"published_row_count": 0, "latest_version": {}},
                "product_catalog": {"published_row_count": 0, "latest_version": {}},
                "bcct": {"published_row_count": 1, "reviewed_row_count": 1, "latest_version": {}},
                "co_stock_row_count": 0,
            }

        def match_bcct_exports(self, _client_id: str, _invoice_no: str, _relevant_types: list[str], export_declaration_nos=None) -> list[dict]:
            raise TypeError("internal matching bug")

    monkeypatch.setattr(portfolio_module, "get_source_index_store", lambda: FakeSourceIndexStore())

    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Postgres error case", "case_code": "CO-PG-ERR", "destination_market": "Ấn Độ", "invoice_no": "INV-PG-ERR"},
        follow_redirects=False,
    )

    with pytest.raises(TypeError, match="internal matching bug"):
        client.get(f"{created.headers['location']}/exports")


def test_postgres_catalog_index_records_keep_payload_and_keys():
    rows = [
        {
            "customs_code": "MAT-PG-001",
            "name": "Postgres material",
            "hs_code": "8504.40",
            "unit": "PCS",
            "status": "active",
        }
    ]

    records = build_catalog_index_records("growatt", "material_catalog", rows)

    assert records == [
        {
            "client_id": "growatt",
            "module": "material_catalog",
            "row_key": "MAT-PG-001",
            "customs_code": "MAT-PG-001",
            "product_code": "",
            "hs_code": "8504.40",
            "unit": "PCS",
            "status": "active",
            "payload": rows[0],
        }
    ]


def test_source_tables_use_postgres_workspace_when_available(monkeypatch):
    from app import portfolio as portfolio_module

    class FakeSourceIndexStore:
        def has_client(self, client_id: str) -> bool:
            return client_id == "growatt"

        def source_workspace(self, _client_id: str, client_config: dict) -> dict:
            return {
                "client_config": client_config,
                "material_catalog": {
                    "module": "material_catalog",
                    "published_rows": [
                        {
                            "customs_code": "MAT-PG-001",
                            "name": "Postgres material",
                            "hs_code": "8504.40",
                            "unit": "PCS",
                            "status": "active",
                        }
                    ],
                    "latest_version": {"version_no": 7},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "product_catalog": {
                    "module": "product_catalog",
                    "published_rows": [
                        {
                            "product_code": "TP-PG-001",
                            "name": "Postgres product",
                            "hs_code": "8504.40",
                            "unit": "PCS",
                            "status": "active",
                        }
                    ],
                    "latest_version": {"version_no": 8},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "bcct": {
                    "module": "bcct",
                    "published_rows": [
                        {
                            "transaction_key": "export||XK-PG||1||TP-PG-001",
                            "direction": "export",
                            "review_status": "reviewed",
                            "declaration_no": "XK-PG",
                            "line_no": "1",
                            "declaration_type": "E42",
                            "item_code": "TP-PG-001",
                            "hs_code": "8504.40",
                            "quantity": "2",
                            "unit": "PCS",
                            "invoice_ref": "INV-PG",
                        }
                    ],
                    "latest_version": {"version_no": 9},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "co_stock_rows": [
                    {
                        "source_row": "PG-STOCK-001",
                        "import_declaration_no": "NK-PG",
                        "line_no": "1",
                        "declaration_type": "E11",
                        "customs_item_code": "MAT-PG-001",
                        "allocation_code": "MAT-PG-001",
                        "allocation_code_status": "resolved",
                        "eligibility_status": "eligible",
                        "remaining_qty": "5",
                        "unit": "PCS",
                    }
                ],
            }

    monkeypatch.setattr(portfolio_module, "get_source_index_store", lambda: FakeSourceIndexStore())

    client = TestClient(app)

    materials = client.get("/clients/growatt/catalog/materials")
    bcct = client.get("/clients/growatt/bcct")
    stock = client.get("/clients/growatt/co-stock")

    assert materials.status_code == 200
    assert "MAT-PG-001" in materials.text
    assert "MAT-001" not in materials.text
    assert "XK-PG" in bcct.text
    assert "TP-PG-001" in bcct.text
    assert "NK-PG" in stock.text


def test_postgres_source_index_initializes_empty_client():
    class FakePostgresSourceIndexStore(PostgresSourceIndexStore):
        def __init__(self):
            super().__init__("postgresql:///unused")
            self.exists = False
            self.replaced = None

        def ensure_schema(self) -> None:
            return None

        def _client_index_exists(self, client_id: str) -> bool:
            assert client_id == "new-client"
            return self.exists

        def replace_client_indexes(
            self,
            client_id: str,
            catalog_records: list[dict],
            bcct_records: list[dict],
            invoice_records: list[dict],
            stock_records: list[dict],
            correction_records: list[dict],
            metadata_records: list[dict],
            source_state_records: list[dict] | None = None,
        ) -> None:
            self.exists = True
            self.replaced = {
                "client_id": client_id,
                "catalog_records": catalog_records,
                "bcct_records": bcct_records,
                "invoice_records": invoice_records,
                "stock_records": stock_records,
                "correction_records": correction_records,
                "metadata_records": metadata_records,
                "source_state_records": source_state_records or [],
            }

    store = FakePostgresSourceIndexStore()

    assert store.has_client("new-client") is True
    assert store.replaced["client_id"] == "new-client"
    assert {row["module"] for row in store.replaced["metadata_records"]} == {
        "material_catalog",
        "product_catalog",
        "bcct",
        "co_stock",
    }
    assert [records["module_state"]["module"] for records in store.replaced["source_state_records"]] == [
        "material_catalog",
        "product_catalog",
        "bcct",
    ]
    assert all(records["module_state"]["published_row_count"] == 0 for records in store.replaced["source_state_records"])


def test_portfolio_app_exposes_source_dashboard_and_summary_api():
    client = TestClient(app)

    dashboard = client.get("/portfolio")
    summary = client.get("/portfolio/api/clients/growatt/source-summary")

    assert dashboard.status_code == 200
    assert "Source Portfolio" in dashboard.text
    assert "Growatt" in dashboard.text
    assert summary.status_code == 200
    payload = summary.json()
    assert payload["client"]["id"] == "growatt"
    assert payload["source_backend"] in {"files", "postgres"}
    assert payload["source_summary"]["material_catalog"]["published_row_count"] >= 3
    assert payload["source_summary"]["product_catalog"]["published_row_count"] >= 2


def test_portfolio_service_prefers_postgres_clients_and_config(monkeypatch):
    from app import portfolio as portfolio_module

    class FakeAppStateStore:
        def has_clients(self) -> bool:
            return True

        def clients(self) -> list[dict]:
            return [
                {
                    "id": "pg-client",
                    "name": "Postgres Client",
                    "code": "PG",
                    "status": "active",
                    "tax_code": "010-PG",
                    "contact": "db",
                    "module_status": {},
                    "material_catalog": [],
                    "product_catalog": [],
                    "bom_rows": [],
                    "co_stock": [],
                    "bcct_rows": [],
                    "counts": {"materials": 0, "products": 0, "bom_lines": 0, "co_stock": 0, "bcct": 0},
                }
            ]

        def client(self, client_id: str) -> dict:
            assert client_id == "pg-client"
            return self.clients()[0]

        def get_client_config(self, client: dict) -> dict:
            return {
                "schema_version": 1,
                "client_id": client["id"],
                "config_version": 3,
                "config_hash": "pg-hash",
                "bcct": {"eligible_import_declaration_types": ["E11"], "relevant_export_declaration_types": ["E42"]},
                "co_stock": {"lot_policy": "line_level"},
                "allocation_code": {"strategy": "same_as_customs_code", "description_regex": "", "fallback": "same_as_customs_code"},
            }

    monkeypatch.setattr(portfolio_module, "get_app_state_store", lambda: FakeAppStateStore(), raising=False)

    service = portfolio_module.PortfolioService()

    assert service.client("pg-client")["name"] == "Postgres Client"
    assert service.clients()[0]["id"] == "pg-client"
    assert service.get_client_config({"id": "pg-client"})["config_hash"] == "pg-hash"


def test_portfolio_service_saves_client_config_to_postgres(monkeypatch):
    from app import portfolio as portfolio_module

    saved_configs = []

    class FakeAppStateStore:
        def save_client_config(self, client: dict, config: dict) -> dict:
            saved_configs.append((client["id"], config["config_hash"]))
            return {**config, "config_version": 9, "config_hash": "saved-pg-hash"}

    monkeypatch.setattr(portfolio_module, "get_app_state_store", lambda: FakeAppStateStore(), raising=False)

    service = portfolio_module.PortfolioService()
    saved = service.save_client_config({"id": "growatt"}, {"config_hash": "draft"})

    assert saved["config_version"] == 9
    assert saved["config_hash"] == "saved-pg-hash"
    assert saved_configs == [("growatt", "draft")]


def test_portfolio_service_uses_postgres_source_writer_for_uploads(monkeypatch):
    from app import portfolio as portfolio_module

    calls = []

    class FakeSourceWriteStore:
        def has_client(self, client_id: str) -> bool:
            return client_id == "growatt"

        def process_catalog_upload(self, client: dict, catalog_type: str, content: bytes, filename: str, upload_scope: str, client_config: dict) -> dict:
            calls.append(("catalog", client["id"], catalog_type, filename, upload_scope, client_config["config_hash"]))
            return {"status": "postgres_catalog", "upload": {"upload_id": "pg-catalog-upload"}}

        def process_bcct_upload(self, client: dict, content: bytes, filename: str, client_config: dict) -> dict:
            calls.append(("bcct", client["id"], filename, client_config["config_hash"]))
            return {"status": "postgres_bcct", "upload": {"upload_id": "pg-bcct-upload"}}

    def fail_file_catalog(*_args, **_kwargs):
        raise AssertionError("Postgres source uploads should not use JSON catalog writer.")

    def fail_file_bcct(*_args, **_kwargs):
        raise AssertionError("Postgres source uploads should not use JSON BCCT writer.")

    monkeypatch.setattr(portfolio_module, "get_source_write_store", lambda: FakeSourceWriteStore(), raising=False)
    monkeypatch.setattr(portfolio_module, "process_catalog_upload", fail_file_catalog)
    monkeypatch.setattr(portfolio_module, "process_bcct_upload", fail_file_bcct)

    service = portfolio_module.PortfolioService()
    monkeypatch.setattr(service, "get_client_config", lambda client: {"config_hash": "cfg-pg"})

    catalog = service.process_catalog_upload({"id": "growatt"}, "material", b"catalog", "catalog.xlsx", "full_catalog")
    bcct = service.process_bcct_upload({"id": "growatt"}, b"bcct", "bcct.xlsx")

    assert catalog["status"] == "postgres_catalog"
    assert bcct["status"] == "postgres_bcct"
    assert calls == [
        ("catalog", "growatt", "material", "catalog.xlsx", "full_catalog", "cfg-pg"),
        ("bcct", "growatt", "bcct.xlsx", "cfg-pg"),
    ]


def test_co_routes_use_portfolio_service_adapter(monkeypatch):
    from app import main as main_module

    class FakePortfolioService:
        def source_workspace(self, client: dict) -> tuple[dict, str]:
            return {
                "client_config": {"config_version": 1, "config_hash": "fake", "bcct": {"eligible_import_declaration_types": [], "relevant_export_declaration_types": []}},
                "material_catalog": {
                    "module": "material_catalog",
                    "published_rows": [
                        {
                            "customs_code": "PF-MAT-001",
                            "name": "Portfolio material",
                            "hs_code": "8504.40",
                            "unit": "PCS",
                            "status": "active",
                        }
                    ],
                    "latest_version": {"version_no": 1},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "product_catalog": {
                    "module": "product_catalog",
                    "published_rows": [],
                    "latest_version": {},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "bcct": {
                    "module": "bcct",
                    "published_rows": [],
                    "latest_version": {},
                    "versions": [],
                    "uploads": [],
                    "correction_candidates": [],
                    "audit_events": [],
                },
                "co_stock_rows": [],
            }, "portfolio-fake"

        def co_case_source_context(self, client: dict, case: dict, *, skip_heavy_context: bool = False) -> dict:
            return {
                "source_backend": "portfolio-fake",
                "source_summary": {
                    "client_config": {"config_version": 1, "config_hash": "fake"},
                    "material_catalog": {"published_row_count": 1, "latest_version": {"version_no": 1}},
                    "product_catalog": {"published_row_count": 0, "latest_version": {}},
                    "bcct": {
                        "published_row_count": 1,
                        "reviewed_row_count": 1,
                        "correction_candidate_count": 0,
                        "latest_version": {"version_no": 1},
                    },
                    "co_stock_row_count": 0,
                },
                "invoice_matches": [
                    {
                        "declaration_no": "PF-XK-001",
                        "line_no": "1",
                        "declaration_type": "E42",
                        "item_code": "PF-TP-001",
                        "hs_code": "8504.40",
                        "quantity": "1",
                        "unit": "PCS",
                        "invoice_ref": case.get("shipment", {}).get("invoice_no", ""),
                    }
                ],
            }

    monkeypatch.setattr(main_module, "portfolio_service", FakePortfolioService())

    client = TestClient(app)
    materials = client.get("/clients/growatt/catalog/materials")
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Portfolio case",
            "case_code": "CO-PF",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-PF",
        },
        follow_redirects=False,
    )
    exports = client.get(f"{created.headers['location']}/exports")

    assert materials.status_code == 200
    assert "PF-MAT-001" in materials.text
    assert exports.status_code == 200
    assert "PF-XK-001" in exports.text
    assert "PF-TP-001" in exports.text


def test_postgres_bcct_index_records_tokenize_invoice_refs():
    rows = [
        {
            "transaction_key": "export||XK-001||1||TP-001",
            "direction": "export",
            "review_status": "reviewed",
            "declaration_no": "XK-001",
            "line_no": "1",
            "declaration_type": "E42",
            "item_code": "TP-001",
            "hs_code": "8504.40",
            "quantity": "2",
            "unit": "PCS",
            "invoice_ref": "INV-001 / INV 002",
        }
    ]

    bcct_records, invoice_records = build_bcct_index_records("growatt", rows)

    assert bcct_records[0]["transaction_key"] == "export||XK-001||1||TP-001"
    assert bcct_records[0]["payload"]["item_code"] == "TP-001"
    assert {(row["invoice_key"], row["transaction_key"]) for row in invoice_records} == {
        ("INV001INV002", "export||XK-001||1||TP-001"),
        ("INV001", "export||XK-001||1||TP-001"),
        ("INV", "export||XK-001||1||TP-001"),
        ("002", "export||XK-001||1||TP-001"),
    }


def test_co_case_supporting_upload_rejects_unsupported_or_oversized_files():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Upload validation", "case_code": "CO-UP", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    location = created.headers["location"]

    unsupported = client.post(
        f"{location}/supporting-files",
        data={"document_slot": "invoice"},
        files={"file": ("invoice.exe", b"bad", "application/octet-stream")},
        follow_redirects=False,
    )
    oversized = client.post(
        f"{location}/supporting-files",
        data={"document_slot": "invoice"},
        files={"file": ("invoice.pdf", b"x" * (MAX_SUPPORTING_FILE_BYTES + 1), "application/pdf")},
        follow_redirects=False,
    )

    assert unsupported.status_code == 400
    assert "Không hỗ trợ định dạng file" in unsupported.text
    assert oversized.status_code == 400
    assert "File supporting vượt quá giới hạn" in oversized.text


def test_co_case_destination_market_shows_verified_form_candidates():
    client = TestClient(app)

    india = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "India shipment", "case_code": "CO-IN", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )
    france = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "France shipment", "case_code": "CO-FR", "destination_market": "Pháp"},
        follow_redirects=False,
    )
    canada = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Canada shipment", "case_code": "CO-CA", "destination_market": "Canada"},
        follow_redirects=False,
    )

    # The /guidance step was removed; the form-candidate hint still renders
    # on the shipment page (step 1) via the recommended_form_lane sidebar.
    india_page = client.get(india.headers["location"])
    france_page = client.get(france.headers["location"])
    canada_page = client.get(canada.headers["location"])

    assert "Form AI" in india_page.text
    assert "15/2010/TT-BCT" in india_page.text
    assert "Form EUR.1" in france_page.text
    assert "11/2020/TT-BCT" in france_page.text
    assert "Form CPTPP" in canada_page.text
    assert "03/2019/TT-BCT" in canada_page.text


def test_co_form_index_defaults_cover_initial_priority_forms():
    from app.co_forms import criteria_preview_for_hs, form_candidates_for_market, hs_scope_is_ex, prioritized_form_lanes, recommended_form_lane
    from app.co_form_config_store import default_co_form_config

    config = default_co_form_config()
    assert form_candidates_for_market("United States")[0]["form_code"] == "B"
    assert recommended_form_lane(prioritized_form_lanes("India"))["form_code"] == "AI"
    assert recommended_form_lane(prioritized_form_lanes("Canada"))["form_code"] == "CPTPP"
    assert recommended_form_lane(prioritized_form_lanes("Pháp"))["form_code"] == "EUR.1"
    assert len(config["psr_rules"]) > 6000
    assert all("Chương" not in rule["hs_scope"] for rule in config["psr_rules"])
    assert any(rule["hs_scope"] == "01" for rule in config["psr_rules"])
    assert criteria_preview_for_hs("AI", "850440")["criteria"] == "AIFTA 35% FOB + CTSH"
    assert "Product Specific Rules" in criteria_preview_for_hs("AI", "850440")["note"]
    assert criteria_preview_for_hs("CPTPP", "850440")["criteria"].startswith("CTH; hoặc RVC không thấp hơn")
    assert criteria_preview_for_hs("B", "999999")["criteria"] == "Tra PSR Form B theo Phụ lục I"
    assert hs_scope_is_ex("ex 0307")


def test_co_form_ex_hs_scope_requires_product_description_confirmation():
    from app.co_forms import criteria_preview_for_hs

    preview = criteria_preview_for_hs("EUR.1", "030600")

    assert preview["status"] == "requires_manual_lookup"
    assert preview["criteria"].startswith("Cần đối chiếu mô tả hàng hóa trước khi áp dụng")
    assert "Match by HS code alone is not enough" in preview["note"]


def test_co_form_settings_page_saves_market_alias_config():
    from app.co_form_config_store import default_co_form_config
    from app.co_forms import form_candidates_for_market

    config = default_co_form_config()
    data = co_form_settings_payload(config)
    new_index = len(config["market_presets"])
    data[f"market_{new_index}_enabled"] = "1"
    data[f"market_{new_index}_show_in_picker"] = "1"
    data[f"market_{new_index}_market"] = "Bharat"
    data[f"market_{new_index}_label"] = "Bharat / Form AI"
    data[f"market_{new_index}_form_code"] = "AI"
    data[f"market_{new_index}_aliases"] = "Bharat"
    data[f"market_{new_index}_selection_reason"] = "Test alias maps to Form AI."
    data[f"market_{new_index}_source_label"] = "test"

    client = TestClient(app)
    settings = client.get("/settings")
    saved = client.post("/settings/co-forms", data=data, follow_redirects=False)
    page = client.get("/settings/co-forms?saved=1")

    assert 'href="/settings/co-forms"' in settings.text
    assert saved.status_code == 303
    assert saved.headers["location"] == "/settings/co-forms?saved=1"
    assert "Đã lưu cấu hình form" in page.text
    assert form_candidates_for_market("Bharat")[0]["form_code"] == "AI"


def test_co_form_settings_page_saves_psr_rule_config():
    from app.co_form_config_store import default_co_form_config
    from app.co_forms import criteria_preview_for_hs

    config = default_co_form_config()
    data = co_form_settings_payload(config)
    new_index = len(config["psr_rules"])
    data[f"psr_{new_index}_enabled"] = "1"
    data[f"psr_{new_index}_form_code"] = "AI"
    data[f"psr_{new_index}_hs_scope"] = "850760"
    data[f"psr_{new_index}_criteria"] = "AIFTA custom battery rule"
    data[f"psr_{new_index}_source_reference"] = "Test source"
    data[f"psr_{new_index}_note"] = "Test note"
    data[f"psr_{new_index}_status"] = "confirmed_by_trong_tin"

    client = TestClient(app)
    saved = client.post("/settings/co-forms", data=data, follow_redirects=False)
    page = client.get("/settings/co-forms?saved=1")
    preview = criteria_preview_for_hs("AI", "85076039")

    assert saved.status_code == 303
    assert "HS Criteria" in page.text
    assert preview["criteria"] == "AIFTA custom battery rule"
    assert preview["status"] == "confirmed_by_trong_tin"


def test_co_form_settings_page_shows_readable_status_labels():
    client = TestClient(app)
    page = client.get("/settings/co-forms?tab=psr&psr_form=AI")
    forms_page = client.get("/settings/co-forms?tab=forms")

    assert page.status_code == 200
    assert forms_page.status_code == 200
    assert "Chờ Trọng Tín xác nhận" in page.text
    assert "Chờ Trọng Tín xác nhận" in forms_page.text
    assert 'placeholder="pending_trong_tin_confirmation"' not in page.text
    assert 'name="psr_0_status"' in page.text
    assert '<select name="psr_0_status">' in page.text


def test_co_form_settings_page_updates_filtered_psr_rule_without_reposting_all_rules():
    from app.co_form_config_store import default_co_form_config
    from app.co_forms import criteria_preview_for_hs

    config = default_co_form_config()
    original_index, rule = next(
        (index, row)
        for index, row in enumerate(config["psr_rules"])
        if row["form_code"] == "CPTPP" and row["hs_scope"] == "85.04"
    )
    data = {
        "active_tab": "psr",
        "psr_visible_count": "1",
        "psr_0_original_index": str(original_index),
        "psr_0_enabled": "1",
        "psr_0_form_code": rule["form_code"],
        "psr_0_hs_scope": rule["hs_scope"],
        "psr_0_criteria": "CPTPP updated filtered 8504 rule",
        "psr_0_source_reference": rule["source_reference"],
        "psr_0_note": rule["note"],
        "psr_0_status": "confirmed_by_trong_tin",
    }

    client = TestClient(app)
    filtered_page = client.get("/settings/co-forms?tab=psr&psr_form=CPTPP&psr_query=8504")
    saved = client.post("/settings/co-forms", data=data, follow_redirects=False)

    assert "Đang hiện" in filtered_page.text
    assert 'name="psr_0_original_index"' in filtered_page.text
    assert saved.status_code == 303
    assert saved.headers["location"] == "/settings/co-forms?saved=1&tab=psr"
    assert criteria_preview_for_hs("CPTPP", "850440")["criteria"] == "CPTPP updated filtered 8504 rule"
    assert criteria_preview_for_hs("B", "850440")["criteria"] == "LVC 30% hoặc CTH"


def test_co_case_create_explains_invoice_market_hint_without_auto_selecting(monkeypatch):
    from app import main as main_module

    class FakePortfolioService:
        def client(self, client_id: str) -> dict:
            return {"id": client_id, "name": "Growatt VN", "code": client_id, "counts": {}}

        def co_case_source_context(self, client: dict, case: dict, *, skip_heavy_context: bool = False) -> dict:
            return {
                "source_backend": "data-hub",
                "source_summary": {
                    "client_config": {"config_version": 1, "config_hash": "fake"},
                    "material_catalog": {"published_row_count": 0, "latest_version": {}},
                    "product_catalog": {"published_row_count": 0, "latest_version": {}},
                    "bcct": {"published_row_count": 1, "reviewed_row_count": 1, "latest_version": {}},
                    "co_stock_row_count": 0,
                },
                "invoice_matches": [
                    {
                        "declaration_no": "XK1",
                        "line_no": "1",
                        "item_code": "TP-US",
                        "invoice_ref": case.get("shipment", {}).get("invoice_no", ""),
                        "market_hint": {
                            "country_code": "US",
                            "country_name": "United States",
                            "source_field": "unloading_location",
                            "source_value": "USLAX - LOS ANGELES - CA",
                            "confidence": "high",
                        },
                    }
                ],
            }

    monkeypatch.setattr(main_module, "portfolio_service", FakePortfolioService())
    client = TestClient(app)

    created = client.post(
        "/clients/growatt-vn/co-case/create",
        data={"title": "Invoice hinted", "case_code": "CO-HINT", "invoice_no": "GUS28826A131-3F"},
        follow_redirects=False,
    )
    page = client.get(created.headers["location"])
    preview = client.get(
        "/clients/growatt-vn/co-case/invoice-preview",
        params={"invoice_no": "GUS28826A131-3F"},
    ).json()

    assert created.status_code == 303
    assert "Thị trường Chưa nhập" in page.text
    assert "Gợi ý thị trường" in page.text
    assert "United States" in page.text
    assert "Form B" in page.text
    assert 'data-market-value="United States"' in page.text
    assert preview["market_inference"]["destination_market"] == "United States"
    assert "unloading_location" in preview["market_inference"]["explanation"]
    assert "source_reference" not in preview
    assert preview["suggested_forms"][0]["form_code"] == "B"


def test_invoice_market_hint_requires_single_high_confidence_country():
    from app.co_market_hints import infer_market_from_invoice_matches

    conflict = infer_market_from_invoice_matches([
        {"market_hint": {"country_code": "US", "country_name": "United States", "confidence": "high"}},
        {"market_hint": {"country_code": "IN", "country_name": "India", "confidence": "high"}},
    ])
    low_confidence = infer_market_from_invoice_matches([
        {"market_hint": {"country_code": "US", "country_name": "United States", "confidence": "low"}}
    ])

    assert conflict["status"] == "conflict"
    assert conflict["destination_market"] == ""
    assert low_confidence["status"] == "missing"


def test_co_case_export_workbook_contains_dossier_sheets_and_criteria_rows():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={
            "title": "Export dossier",
            "case_code": "CO-XLSX",
            "destination_market": "Ấn Độ",
            "invoice_no": "INV-XLSX",
        },
        follow_redirects=False,
    )

    response = client.post(f"{created.headers['location']}/export")

    assert response.status_code == 200
    assert response.content.startswith(b"PK")
    workbook = load_workbook(BytesIO(response.content))
    assert set(["Case", "Supporting Files", "BCCT Invoice Matches", "Form Guidance", "Criteria"]).issubset(workbook.sheetnames)
    assert workbook["Case"]["B2"].value == "CO-XLSX"
    assert workbook["Form Guidance"]["A2"].value == "Form AI"
    criteria_values = [cell.value for row in workbook["Criteria"].iter_rows(values_only=False) for cell in row]
    assert "PV00.0048500" in criteria_values


def test_co_case_export_filename_is_sanitized():
    client = TestClient(app)
    created = client.post(
        "/clients/growatt/co-case/create",
        data={"title": "Unsafe filename", "case_code": "CO/../../bad", "destination_market": "Ấn Độ"},
        follow_redirects=False,
    )

    response = client.post(f"{created.headers['location']}/export")

    assert response.status_code == 200
    assert 'filename="bad-dossier.xlsx"' in response.headers["content-disposition"]
