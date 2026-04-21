#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from copy import deepcopy
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from itertools import product
from pathlib import Path
from zipfile import ZipFile

import openpyxl
import xlrd


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DEFAULT_WORKBOOK = Path(
    "data/extracted/Growatt-20260421/Growatt/"
    "tru lui CO final  SXXK - 2025 commercial-MAC - Huyền đúng.xlsm"
)
DEFAULT_XK_REPORT = Path(
    "data/extracted/Growatt-20260421/Growatt/BaoCaoHangChiTietXK GRW T3-T4.2026 moi.xls"
)
DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENTS = ("GIN01426B282", "GIN01426C171")
TARGET_RVC = 35.0
MODEL_RE = re.compile(r"\((PV[^)]+)\)")
VN_ORIGINS = {"VIETNAM", "VIỆT NAM", "VN"}
IMPORT_LEAD_DAYS = 2


@dataclass
class ExportLine:
    shipment_id: str
    declaration_no: str
    declaration_item_no: int
    export_row_no: int
    export_date: date
    internal_code: str
    model_code: str
    hs_code: str
    quantity: float
    unit_price_usd: float
    incoterm: str
    name: str


@dataclass
class BomLine:
    material_code: str
    qty_per_unit: float
    row_no: int
    ordinal_key: str


@dataclass
class BomVariant:
    variant_id: str
    model_code: str
    bom_code: str
    block_index: int
    start_row: int
    end_row: int
    line_count: int
    lines: list[BomLine]


@dataclass
class StockBucket:
    bucket_id: str
    sheet_row_no: int
    declaration_no: str
    declaration_item_no: int
    import_date: date | None
    material_code: str
    hs_code: str
    name: str
    origin: str
    unit_price_usd: float
    exchange_rate: float
    remaining_qty: float


@dataclass
class SourceAllocation:
    bucket_id: str
    declaration_no: str
    declaration_item_no: int
    import_date: date | None
    allocated_qty: float
    unit_price_usd: float
    exchange_rate: float
    origin: str


@dataclass
class MaterialAllocation:
    material_code: str
    need_qty: float
    allocated_qty: float
    unmet_qty: float
    blocked_by_date_qty: float
    source_bucket_count: int
    source_declarations: list[str]
    source_origins: list[str]
    unit_price_usd: float | None
    non_origin_value_usd: float
    stock_sufficient: bool
    source_allocations: list[SourceAllocation]


@dataclass
class ProductScenarioResult:
    shipment_id: str
    model_code: str
    bom_code: str
    variant_id: str
    declaration_no: str
    export_date: date
    export_qty: float
    export_unit_price_usd: float
    fob_value_usd: float
    non_origin_value_usd: float
    rvc_percent: float | None
    passes_rvc: bool
    stock_sufficient: bool
    unmet_material_count: int
    unmet_qty_total: float
    material_count: int
    multi_source_material_count: int
    materials: list[MaterialAllocation]


@dataclass
class ShipmentScenarioResult:
    shipment_id: str
    scenario_id: str
    valuation_mode: str
    product_results: list[ProductScenarioResult]

    @property
    def stock_sufficient(self) -> bool:
        return all(item.stock_sufficient for item in self.product_results)

    @property
    def passes_rvc(self) -> bool:
        return self.stock_sufficient and all(item.passes_rvc for item in self.product_results)

    @property
    def min_margin(self) -> float | None:
        margins = [item.rvc_percent - TARGET_RVC for item in self.product_results if item.rvc_percent is not None]
        return min(margins) if margins else None

    @property
    def total_unmet_qty(self) -> float:
        return sum(item.unmet_qty_total for item in self.product_results)


def normalize_origin(origin: str | None) -> str:
    if origin is None:
        return ""
    return str(origin).strip().upper()


def is_vietnam_origin(origin: str | None) -> bool:
    return normalize_origin(origin) in VN_ORIGINS


def parse_numeric(value: object) -> float:
    if value in (None, ""):
        return 0.0
    return float(value)


def to_date(value: object, datemode: int | None = None) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and datemode is not None:
        try:
            return datetime(*xlrd.xldate_as_tuple(value, datemode)).date()
        except Exception:
            return None
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def load_shared_strings(archive: ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall(f"{NS}si"):
        strings.append("".join(text.text or "" for text in si.iter(f"{NS}t")))
    return strings


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str | None:
    raw = cell.find(f"{NS}v")
    if raw is None:
        return None
    value = raw.text or ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(value)]
    return value


def load_dm_variants(workbook_path: Path) -> tuple[list[BomVariant], dict[str, list[BomVariant]]]:
    variants: list[BomVariant] = []
    variants_by_model: dict[str, list[BomVariant]] = defaultdict(list)

    with ZipFile(workbook_path) as archive:
        shared_strings = load_shared_strings(archive)
        rows: list[tuple[int, str, str, float]] = []
        for _, elem in ET.iterparse(archive.open("xl/worksheets/sheet4.xml"), events=("end",)):
            if elem.tag != f"{NS}row":
                continue
            row_no = int(elem.attrib["r"])
            if row_no < 7:
                elem.clear()
                continue
            values: dict[str, str | None] = {}
            for cell in elem.findall(f"{NS}c"):
                ref = cell.attrib["r"]
                col = "".join(ch for ch in ref if ch.isalpha())
                values[col] = cell_value(cell, shared_strings)
            model_code = values.get("A")
            ordinal_key = values.get("B")
            material_code = values.get("F")
            qty_raw = values.get("I")
            if model_code and material_code and qty_raw not in (None, ""):
                rows.append((row_no, str(model_code), str(ordinal_key or ""), str(material_code), float(qty_raw)))
            elem.clear()

    block_counts: dict[str, int] = defaultdict(int)
    current_model = None
    current_rows: list[tuple[int, str, str, float]] = []

    def flush_current() -> None:
        nonlocal current_rows, current_model
        if not current_model or not current_rows:
            current_rows = []
            current_model = None
            return
        block_counts[current_model] += 1
        block_index = block_counts[current_model]
        variant_id = f"{current_model}__block{block_index}"
        lines = [
            BomLine(material_code=material_code, qty_per_unit=qty_per_unit, row_no=row_no, ordinal_key=ordinal_key)
            for row_no, _, ordinal_key, material_code, qty_per_unit in current_rows
        ]
        variant = BomVariant(
            variant_id=variant_id,
            model_code=current_model,
            bom_code=current_model,
            block_index=block_index,
            start_row=current_rows[0][0],
            end_row=current_rows[-1][0],
            line_count=len(lines),
            lines=lines,
        )
        variants.append(variant)
        variants_by_model[current_model].append(variant)
        current_rows = []
        current_model = None

    for row_no, model_code, ordinal_key, material_code, qty_per_unit in rows:
        if model_code != current_model:
            flush_current()
            current_model = model_code
        current_rows.append((row_no, model_code, ordinal_key, material_code, qty_per_unit))
    flush_current()

    return variants, variants_by_model


def load_export_rows(report_path: Path, shipment_ids: set[str]) -> dict[str, list[ExportLine]]:
    book = xlrd.open_workbook(str(report_path))
    sheet = book.sheet_by_index(0)
    shipments: dict[str, list[ExportLine]] = defaultdict(list)

    for row_idx in range(10, sheet.nrows):
        row = sheet.row_values(row_idx)
        shipment_id = str(row[50]).strip()
        if shipment_id not in shipment_ids:
            continue
        name = str(row[22]).strip()
        match = MODEL_RE.search(name)
        if not match:
            continue
        shipments[shipment_id].append(
            ExportLine(
                shipment_id=shipment_id,
                declaration_no=str(int(row[1])),
                declaration_item_no=int(row[19]),
                export_row_no=row_idx + 1,
                export_date=to_date(row[2], book.datemode) or date.min,
                internal_code=str(row[20]).strip(),
                model_code=match.group(1),
                hs_code=str(row[21]).strip(),
                quantity=float(row[26]),
                unit_price_usd=float(row[24]),
                incoterm=str(row[17]).strip().upper(),
                name=name,
            )
        )

    for shipment_id in shipments:
        shipments[shipment_id].sort(key=lambda item: (item.declaration_no, item.declaration_item_no, item.export_row_no))
    return shipments


def load_nk2_stock(workbook_path: Path) -> dict[str, list[StockBucket]]:
    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
        keep_vba=True,
    )
    sheet = workbook["NK2"]
    stock_by_code: dict[str, list[StockBucket]] = defaultdict(list)
    for row_no, row in enumerate(sheet.iter_rows(min_row=5, values_only=True), start=5):
        material_code = row[4]
        if material_code is None:
            continue
        remaining_qty = parse_numeric(row[17])
        if remaining_qty <= 0:
            continue
        bucket = StockBucket(
            bucket_id=f"{int(row[0])}-{int(row[3])}-{str(material_code).strip()}-{row_no}",
            sheet_row_no=row_no,
            declaration_no=str(int(row[0])),
            declaration_item_no=int(row[3]),
            import_date=to_date(row[1]),
            material_code=str(material_code).strip(),
            hs_code=str(row[5]).strip() if row[5] is not None else "",
            name=str(row[6] or "").strip(),
            origin=str(row[7] or "").strip(),
            unit_price_usd=float(row[8] or 0),
            exchange_rate=float(row[15] or 0),
            remaining_qty=remaining_qty,
        )
        stock_by_code[bucket.material_code].append(bucket)
    return stock_by_code


def related_variants_for_model(model_code: str, variants_by_model: dict[str, list[BomVariant]]) -> list[BomVariant]:
    variants: list[BomVariant] = []
    for key, items in variants_by_model.items():
        if key == model_code or key.startswith(f"{model_code}-"):
            variants.extend(items)
    return variants


def is_bucket_eligible_for_export(bucket: StockBucket, export_line: ExportLine) -> bool:
    if bucket.import_date is None:
        return False
    return bucket.import_date <= export_line.export_date - timedelta(days=IMPORT_LEAD_DAYS)


def scenario_sort_key(result: ShipmentScenarioResult) -> tuple[float, float, float, float]:
    min_margin = result.min_margin if result.min_margin is not None else -9999.0
    return (
        0.0 if result.stock_sufficient else 1.0,
        0.0 if result.passes_rvc else 1.0,
        -min_margin,
        result.total_unmet_qty,
    )


def compute_material_unit_price(
    allocations: list[tuple[StockBucket, float]],
    need_qty: float,
    valuation_mode: str,
) -> float | None:
    if not allocations:
        return None
    if valuation_mode == "weighted":
        allocated_qty = sum(qty for _, qty in allocations)
        if allocated_qty <= 0:
            return None
        return sum(bucket.unit_price_usd * qty for bucket, qty in allocations) / allocated_qty
    return sum(bucket.unit_price_usd for bucket, _ in allocations) / len(allocations)


def evaluate_scenario(
    shipment_id: str,
    export_lines: list[ExportLine],
    scenario_variants: dict[str, BomVariant],
    stock_snapshot: dict[str, list[StockBucket]],
    valuation_mode: str,
) -> ShipmentScenarioResult:
    stock_state = {code: deepcopy(buckets) for code, buckets in stock_snapshot.items()}
    product_results: list[ProductScenarioResult] = []

    for export_line in export_lines:
        variant = scenario_variants[export_line.model_code]
        material_results: list[MaterialAllocation] = []
        non_origin_total = 0.0
        unmet_material_count = 0
        unmet_qty_total = 0.0
        multi_source_material_count = 0

        for bom_line in variant.lines:
            need_qty = bom_line.qty_per_unit * export_line.quantity
            remaining_need = need_qty
            allocations: list[tuple[StockBucket, float]] = []
            blocked_by_date_qty = 0.0

            for bucket in stock_state.get(bom_line.material_code, []):
                if remaining_need <= 1e-9:
                    break
                if bucket.remaining_qty <= 1e-9:
                    continue
                if not is_bucket_eligible_for_export(bucket, export_line):
                    blocked_by_date_qty += bucket.remaining_qty
                    continue
                take_qty = min(bucket.remaining_qty, remaining_need)
                if take_qty <= 0:
                    continue
                bucket.remaining_qty -= take_qty
                remaining_need -= take_qty
                allocations.append((bucket, take_qty))

            allocated_qty = need_qty - remaining_need
            stock_sufficient = remaining_need <= 1e-9
            if not stock_sufficient:
                unmet_material_count += 1
                unmet_qty_total += remaining_need
            if len(allocations) > 1:
                multi_source_material_count += 1

            unit_price_usd = compute_material_unit_price(allocations, need_qty, valuation_mode)
            source_origins = sorted({bucket.origin for bucket, _ in allocations if bucket.origin})
            source_declarations = sorted({bucket.declaration_no for bucket, _ in allocations})
            all_vn = bool(allocations) and all(is_vietnam_origin(bucket.origin) for bucket, _ in allocations)
            non_origin_value = 0.0
            if unit_price_usd is not None and not all_vn:
                non_origin_value = need_qty * unit_price_usd
            non_origin_total += non_origin_value
            source_allocations = [
                SourceAllocation(
                    bucket_id=bucket.bucket_id,
                    declaration_no=bucket.declaration_no,
                    declaration_item_no=bucket.declaration_item_no,
                    import_date=bucket.import_date,
                    allocated_qty=qty,
                    unit_price_usd=bucket.unit_price_usd,
                    exchange_rate=bucket.exchange_rate,
                    origin=bucket.origin,
                )
                for bucket, qty in allocations
            ]

            material_results.append(
                MaterialAllocation(
                    material_code=bom_line.material_code,
                    need_qty=need_qty,
                    allocated_qty=allocated_qty,
                    unmet_qty=remaining_need,
                    blocked_by_date_qty=blocked_by_date_qty,
                    source_bucket_count=len(allocations),
                    source_declarations=source_declarations,
                    source_origins=source_origins,
                    unit_price_usd=unit_price_usd,
                    non_origin_value_usd=non_origin_value,
                    stock_sufficient=stock_sufficient,
                    source_allocations=source_allocations,
                )
            )

        fob_value_usd = export_line.quantity * export_line.unit_price_usd
        rvc_percent = None
        if fob_value_usd > 0:
            rvc_percent = ((fob_value_usd - non_origin_total) / fob_value_usd) * 100.0
        stock_sufficient = unmet_material_count == 0
        passes_rvc = stock_sufficient and rvc_percent is not None and rvc_percent >= TARGET_RVC

        product_results.append(
            ProductScenarioResult(
                shipment_id=shipment_id,
                model_code=export_line.model_code,
                bom_code=variant.bom_code,
                variant_id=variant.variant_id,
                declaration_no=export_line.declaration_no,
                export_date=export_line.export_date,
                export_qty=export_line.quantity,
                export_unit_price_usd=export_line.unit_price_usd,
                fob_value_usd=fob_value_usd,
                non_origin_value_usd=non_origin_total,
                rvc_percent=rvc_percent,
                passes_rvc=passes_rvc,
                stock_sufficient=stock_sufficient,
                unmet_material_count=unmet_material_count,
                unmet_qty_total=unmet_qty_total,
                material_count=len(material_results),
                multi_source_material_count=multi_source_material_count,
                materials=material_results,
            )
        )

    scenario_bits = [f"{result.model_code}:{result.variant_id}" for result in product_results]
    return ShipmentScenarioResult(
        shipment_id=shipment_id,
        scenario_id=" | ".join(scenario_bits),
        valuation_mode=valuation_mode,
        product_results=product_results,
    )


def enumerate_shipment_scenarios(
    shipment_id: str,
    export_lines: list[ExportLine],
    variants_by_model: dict[str, list[BomVariant]],
    stock_snapshot: dict[str, list[StockBucket]],
    valuation_mode: str,
) -> tuple[list[ShipmentScenarioResult], dict[str, list[BomVariant]]]:
    variant_options: dict[str, list[BomVariant]] = {}
    for export_line in export_lines:
        options = related_variants_for_model(export_line.model_code, variants_by_model)
        if not options:
            raise SystemExit(f"No BOM variants found for {export_line.model_code}")
        variant_options[export_line.model_code] = options

    scenario_results: list[ShipmentScenarioResult] = []
    option_lists = [variant_options[export_line.model_code] for export_line in export_lines]
    for selection in product(*option_lists):
        scenario_variants = {
            export_lines[idx].model_code: variant
            for idx, variant in enumerate(selection)
        }
        scenario_results.append(
            evaluate_scenario(
                shipment_id=shipment_id,
                export_lines=export_lines,
                scenario_variants=scenario_variants,
                stock_snapshot=stock_snapshot,
                valuation_mode=valuation_mode,
            )
        )

    scenario_results.sort(key=scenario_sort_key)
    return scenario_results, variant_options


def print_shipment_summary(
    shipment_id: str,
    export_lines: list[ExportLine],
    variant_options: dict[str, list[BomVariant]],
    scenario_results: list[ShipmentScenarioResult],
    top_n: int,
) -> None:
    print(f"\nShipment {shipment_id}")
    print("Products:")
    for export_line in export_lines:
        options = variant_options[export_line.model_code]
        print(
            f"- {export_line.model_code} | qty={int(export_line.quantity)} | "
            f"decl={export_line.declaration_no} | export_date={export_line.export_date.isoformat()} | "
            f"variant_options={len(options)}"
        )
        for variant in options:
            print(
                f"  option {variant.variant_id} | bom_code={variant.bom_code} | dm_rows={variant.start_row}-{variant.end_row} | "
                f"line_count={variant.line_count}"
            )

    print(f"\nTop {min(top_n, len(scenario_results))} baseline scenarios:")
    for idx, result in enumerate(scenario_results[:top_n], start=1):
        min_margin = result.min_margin
        min_margin_text = f"{min_margin:.2f}" if min_margin is not None else "n/a"
        print(
            f"{idx}. stock={'ok' if result.stock_sufficient else 'fail'} | "
            f"rvc={'pass' if result.passes_rvc else 'fail'} | "
            f"min_margin={min_margin_text} | unmet_qty={result.total_unmet_qty:.4f}"
        )
        for product_result in result.product_results:
            rvc_text = f"{product_result.rvc_percent:.2f}%" if product_result.rvc_percent is not None else "n/a"
            print(
                f"   - {product_result.model_code} via {product_result.variant_id} | "
                f"RVC={rvc_text} | stock={'ok' if product_result.stock_sufficient else 'fail'} | "
                f"multi_source_materials={product_result.multi_source_material_count} | "
                f"unmet_materials={product_result.unmet_material_count}"
            )


def serialize_variants(variants_by_model: dict[str, list[BomVariant]]) -> dict[str, object]:
    payload: dict[str, object] = {}
    for model_code, variants in variants_by_model.items():
        payload[model_code] = [asdict(variant) for variant in variants]
    return payload


def serialize_export_rows(export_rows: dict[str, list[ExportLine]]) -> dict[str, object]:
    return {
        shipment_id: [asdict(line) for line in lines]
        for shipment_id, lines in export_rows.items()
    }


def serialize_stock_snapshot(stock_snapshot: dict[str, list[StockBucket]]) -> dict[str, object]:
    return {
        material_code: [asdict(bucket) for bucket in buckets]
        for material_code, buckets in stock_snapshot.items()
    }


def write_case_workspace(
    case_dir: Path,
    export_rows: dict[str, list[ExportLine]],
    variants_by_model: dict[str, list[BomVariant]],
    stock_snapshot: dict[str, list[StockBucket]],
    scenario_results: dict[str, list[ShipmentScenarioResult]],
) -> None:
    normalized_dir = case_dir / "normalized"
    results_dir = case_dir / "results"
    normalized_dir.mkdir(parents=True, exist_ok=True)
    results_dir.mkdir(parents=True, exist_ok=True)

    (normalized_dir / "export-lines.json").write_text(
        json.dumps(serialize_export_rows(export_rows), indent=2, ensure_ascii=False, default=str)
    )
    (normalized_dir / "bom-variants.json").write_text(
        json.dumps(serialize_variants(variants_by_model), indent=2, ensure_ascii=False, default=str)
    )
    (normalized_dir / "stock-snapshot.json").write_text(
        json.dumps(serialize_stock_snapshot(stock_snapshot), indent=2, ensure_ascii=False, default=str)
    )
    (results_dir / "baseline-scenarios.json").write_text(
        json.dumps(serialize_results(scenario_results), indent=2, ensure_ascii=False, default=str)
    )


def serialize_results(results: dict[str, list[ShipmentScenarioResult]]) -> dict[str, object]:
    return {
        shipment_id: [
            {
                "shipment_id": result.shipment_id,
                "scenario_id": result.scenario_id,
                "valuation_mode": result.valuation_mode,
                "stock_sufficient": result.stock_sufficient,
                "passes_rvc": result.passes_rvc,
                "min_margin": result.min_margin,
                "total_unmet_qty": result.total_unmet_qty,
                "product_results": [
                    {
                        **{k: v for k, v in asdict(product_result).items() if k != "materials"},
                        "materials": [asdict(material) for material in product_result.materials],
                    }
                    for product_result in result.product_results
                ],
            }
            for result in shipment_results
        ]
        for shipment_id, shipment_results in results.items()
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--xk-report", type=Path, default=DEFAULT_XK_REPORT)
    parser.add_argument(
        "--shipments",
        nargs="+",
        default=list(DEFAULT_SHIPMENTS),
        help="Shipment invoice ids such as GIN01426B282",
    )
    parser.add_argument(
        "--valuation-mode",
        choices=("workbook_avg", "weighted"),
        default="workbook_avg",
    )
    parser.add_argument("--top", type=int, default=5)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    shipment_ids = set(args.shipments)
    _, variants_by_model = load_dm_variants(args.workbook)
    export_rows = load_export_rows(args.xk_report, shipment_ids)
    stock_snapshot = load_nk2_stock(args.workbook)

    all_results: dict[str, list[ShipmentScenarioResult]] = {}
    for shipment_id in args.shipments:
        if shipment_id not in export_rows:
            raise SystemExit(f"Shipment not found in XK report: {shipment_id}")
        scenario_results, variant_options = enumerate_shipment_scenarios(
            shipment_id=shipment_id,
            export_lines=export_rows[shipment_id],
            variants_by_model=variants_by_model,
            stock_snapshot=stock_snapshot,
            valuation_mode=args.valuation_mode,
        )
        all_results[shipment_id] = scenario_results
        print_shipment_summary(
            shipment_id=shipment_id,
            export_lines=export_rows[shipment_id],
            variant_options=variant_options,
            scenario_results=scenario_results,
            top_n=args.top,
        )

    if args.json_out:
        args.json_out.write_text(json.dumps(serialize_results(all_results), indent=2, ensure_ascii=False, default=str))
        print(f"\nWrote {args.json_out}")

    if args.case_dir:
        write_case_workspace(
            case_dir=args.case_dir,
            export_rows=export_rows,
            variants_by_model=variants_by_model,
            stock_snapshot=stock_snapshot,
            scenario_results=all_results,
        )
        print(f"Wrote case workspace {args.case_dir}")


if __name__ == "__main__":
    main()
