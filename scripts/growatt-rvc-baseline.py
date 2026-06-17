#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from copy import deepcopy
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from itertools import product
from pathlib import Path

from growatt_bom_sources import (
    DEFAULT_CASE_DIR,
    DM_BOM_SOURCE_ID,
    bom_source_choices,
    relative_to_case,
    resolve_bom_source,
)
from growatt_case_config import resolve_shipment_policy


DEFAULT_SHIPMENTS = ("GIN01426B282",)
TARGET_RVC = 35.0
VN_ORIGINS = {"VIETNAM", "VIỆT NAM", "VN"}


@dataclass
class ExportLine:
    shipment_id: str
    declaration_no: str
    declaration_item_no: int
    export_row_no: int
    export_date: date
    invoice_no: str
    invoice_date: date | None
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
    tracking_key: str
    source: str
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
    admissibility_status: str
    admissible_variant_ids: frozenset[str]


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
    admissibility_status: str


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


@dataclass(frozen=True)
class StartingPointScenario:
    starting_point_id: str
    variant_strategy: str
    sequence_strategy: str
    export_order: list[str]
    result: ShipmentScenarioResult


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_numeric(value: object) -> float:
    text = clean_text(value)
    if not text:
        return 0.0
    return float(text)


def parse_int(value: object) -> int:
    text = clean_text(value)
    if not text:
        return 0
    return int(float(text))


def to_date(value: object) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, date):
        return value
    text = clean_text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def shipment_slug(shipment_id: str) -> str:
    match = re.search(r"([A-Z]\d{3,})$", shipment_id)
    if match:
        return match.group(1).lower()
    return shipment_id.lower().replace("/", "-")


def normalize_origin(origin: str | None) -> str:
    if origin is None:
        return ""
    return str(origin).strip().upper()


def is_vietnam_origin(origin: str | None) -> bool:
    return normalize_origin(origin) in VN_ORIGINS


def load_bom_variants(variant_csv_path: Path) -> tuple[list[BomVariant], dict[str, list[BomVariant]]]:
    rows_by_variant: dict[str, list[dict[str, str]]] = defaultdict(list)
    with variant_csv_path.open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            rows_by_variant[clean_text(row["bom_variant_id"])].append(row)

    variants: list[BomVariant] = []
    variants_by_model: dict[str, list[BomVariant]] = defaultdict(list)
    for variant_id, rows in rows_by_variant.items():
        rows.sort(key=lambda item: parse_int(item["dm_row_no"]))
        bom_code = clean_text(rows[0]["bom_code"])
        lines = [
            BomLine(
                material_code=clean_text(row["material_code"]),
                qty_per_unit=parse_numeric(row["qty_per_unit"]),
                row_no=parse_int(row.get("dm_row_no") or row.get("source_row_no")),
                ordinal_key=clean_text(row["ordinal_key"]),
            )
            for row in rows
        ]
        block_match = re.search(r"__block(\d+)$", variant_id)
        block_index = int(block_match.group(1)) if block_match else 1
        variant = BomVariant(
            variant_id=variant_id,
            model_code=clean_text(rows[0]["product_family_code"]),
            bom_code=bom_code,
            block_index=block_index,
            start_row=lines[0].row_no,
            end_row=lines[-1].row_no,
            line_count=len(lines),
            lines=lines,
        )
        variants.append(variant)
        variants_by_model[variant.model_code].append(variant)

    for items in variants_by_model.values():
        items.sort(key=lambda item: (item.bom_code, item.block_index))
    return variants, variants_by_model


def load_export_rows(normalized_dir: Path, shipment_ids: set[str]) -> dict[str, list[ExportLine]]:
    shipments: dict[str, list[ExportLine]] = defaultdict(list)
    with (normalized_dir / "exports-normalized.csv").open(encoding="utf-8") as handle:
        for row_idx, row in enumerate(csv.DictReader(handle), start=2):
            shipment_id = clean_text(row["shipment_id"])
            if shipment_id not in shipment_ids:
                continue
            model_code = clean_text(row["final_lookup_key"])
            export_date = to_date(row["declaration_date"]) or to_date(row["invoice_date"])
            if not model_code or export_date is None:
                continue
            shipments[shipment_id].append(
                ExportLine(
                    shipment_id=shipment_id,
                    declaration_no=clean_text(row["declaration_no"]),
                    declaration_item_no=parse_int(row["declaration_item_no"]),
                    export_row_no=row_idx,
                    export_date=export_date,
                    invoice_no=clean_text(row["invoice_no"]),
                    invoice_date=to_date(row["invoice_date"]),
                    internal_code=clean_text(row["declared_code"]),
                    model_code=model_code,
                    hs_code=clean_text(row["hs_code"]),
                    quantity=parse_numeric(row["quantity"]),
                    unit_price_usd=parse_numeric(row["unit_price_usd"]),
                    incoterm=clean_text(row["incoterm"]).upper(),
                    name=clean_text(row["name"]),
                )
            )

    for shipment_id in shipments:
        shipments[shipment_id].sort(key=lambda item: (item.declaration_no, item.declaration_item_no, item.export_row_no))
    return shipments


def load_candidate_admissibility(
    case_dir: Path,
    shipment_id: str,
    workspace_dir: Path | None = None,
) -> dict[tuple[str, str], set[str]]:
    slug = shipment_slug(shipment_id)
    shipment_workspace_dir = workspace_dir or (case_dir / slug)
    admissibility_path = shipment_workspace_dir / "normalized" / f"{slug}-variant-admissibility.csv"
    if not admissibility_path.exists():
        raise SystemExit(f"Missing admissibility artifact: {admissibility_path}")

    candidate_variants: dict[tuple[str, str], set[str]] = defaultdict(set)
    with admissibility_path.open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            if clean_text(row["admissibility_status"]) != "admissible_candidate_for_variant":
                continue
            key = (
                clean_text(row["tracking_key"]),
                clean_text(row["admissible_material_code"]),
            )
            candidate_variants[key].add(clean_text(row["variant_id"]))
    return candidate_variants


def load_stock_snapshot(normalized_dir: Path, candidate_admissibility: dict[tuple[str, str], set[str]]) -> dict[str, list[StockBucket]]:
    stock_by_code: dict[str, list[StockBucket]] = defaultdict(list)
    with (normalized_dir / "co-stock-tracking-updated.csv").open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            remaining_qty = parse_numeric(row["remaining_qty"])
            if remaining_qty <= 0:
                continue

            tracking_key = clean_text(row["tracking_key"])
            confirmed_code = clean_text(row["confirmed_lookup_code"])
            candidate_code = clean_text(row["lookup_material_code"])
            mapping_status = clean_text(row["mapping_status"])

            admissibility_status = ""
            admissible_variant_ids: frozenset[str] = frozenset()
            material_code = ""

            if confirmed_code:
                material_code = confirmed_code
                admissibility_status = "confirmed_for_variant"
            elif mapping_status == "candidate_exact_dm_match":
                allowed_variants = candidate_admissibility.get((tracking_key, candidate_code), set())
                if not allowed_variants:
                    continue
                material_code = candidate_code
                admissibility_status = "admissible_candidate_for_variant"
                admissible_variant_ids = frozenset(sorted(allowed_variants))
            else:
                continue

            source_row_hint = clean_text(row["nk2_source_row_no"]) or clean_text(row["bcct_source_row_no"]) or "na"
            bucket = StockBucket(
                bucket_id=f"{tracking_key}-{material_code}-{source_row_hint}",
                tracking_key=tracking_key,
                source=clean_text(row["source"]),
                declaration_no=clean_text(row["declaration_no"]),
                declaration_item_no=parse_int(row["declaration_item_no"]),
                import_date=to_date(row["declaration_date"]),
                material_code=material_code,
                hs_code=clean_text(row["hs_code"]),
                name=clean_text(row["name"]),
                origin=clean_text(row["origin"]),
                unit_price_usd=parse_numeric(row["unit_price_usd"]),
                exchange_rate=parse_numeric(row["exchange_rate"]),
                remaining_qty=remaining_qty,
                admissibility_status=admissibility_status,
                admissible_variant_ids=admissible_variant_ids,
            )
            stock_by_code[material_code].append(bucket)

    for buckets in stock_by_code.values():
        buckets.sort(
            key=lambda item: (
                item.import_date or date.max,
                item.declaration_no,
                item.declaration_item_no,
                item.bucket_id,
            )
        )
    return stock_by_code


def related_variants_for_model(model_code: str, variants_by_model: dict[str, list[BomVariant]]) -> list[BomVariant]:
    variants: list[BomVariant] = []
    for key, items in variants_by_model.items():
        if key == model_code:
            variants.extend(items)
    return variants


def bucket_is_admissible_for_variant(bucket: StockBucket, variant_id: str) -> bool:
    if bucket.admissibility_status == "confirmed_for_variant":
        return True
    return variant_id in bucket.admissible_variant_ids


def is_bucket_eligible_for_export(
    bucket: StockBucket,
    export_line: ExportLine,
    import_lead_days: int,
    max_import_age_days: int,
) -> bool:
    if bucket.import_date is None:
        return False
    latest_allowed_date = export_line.export_date - timedelta(days=import_lead_days)
    if bucket.import_date > latest_allowed_date:
        return False
    if max_import_age_days > 0:
        earliest_allowed_date = export_line.export_date - timedelta(days=max_import_age_days)
        if bucket.import_date < earliest_allowed_date:
            return False
    return True


def scenario_sort_key(result: ShipmentScenarioResult) -> tuple[float, float, float, float]:
    min_margin = result.min_margin if result.min_margin is not None else -9999.0
    return (
        0.0 if result.stock_sufficient else 1.0,
        0.0 if result.passes_rvc else 1.0,
        -min_margin,
        result.total_unmet_qty,
    )


def export_line_sort_key(line: ExportLine, sequence_strategy: str) -> tuple[object, ...]:
    if sequence_strategy == "invoice_then_declaration":
        return (
            line.invoice_date or line.export_date,
            line.invoice_no,
            line.declaration_no,
            line.declaration_item_no,
            line.export_row_no,
        )
    return (
        line.declaration_no,
        line.declaration_item_no,
        line.export_row_no,
    )


def pick_latest_variant(options: list[BomVariant]) -> BomVariant:
    return max(
        options,
        key=lambda item: (
            item.end_row,
            item.start_row,
            item.block_index,
            item.bom_code,
            item.variant_id,
        ),
    )


def compute_material_unit_price(
    allocations: list[tuple[StockBucket, float]],
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
    import_lead_days: int,
    max_import_age_days: int,
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
                if not bucket_is_admissible_for_variant(bucket, variant.variant_id):
                    continue
                if not is_bucket_eligible_for_export(
                    bucket,
                    export_line,
                    import_lead_days=import_lead_days,
                    max_import_age_days=max_import_age_days,
                ):
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

            unit_price_usd = compute_material_unit_price(allocations, valuation_mode)
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
                    admissibility_status=bucket.admissibility_status,
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
    import_lead_days: int,
    max_import_age_days: int,
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
                import_lead_days=import_lead_days,
                max_import_age_days=max_import_age_days,
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
                f"  option {variant.variant_id} | bom_code={variant.bom_code} | rows={variant.start_row}-{variant.end_row} | "
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
    payload: dict[str, object] = {}
    for material_code, buckets in stock_snapshot.items():
        rows: list[dict[str, object]] = []
        for bucket in buckets:
            item = asdict(bucket)
            item["admissible_variant_ids"] = sorted(bucket.admissible_variant_ids)
            rows.append(item)
        payload[material_code] = rows
    return payload


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


def serialize_scenario(result: ShipmentScenarioResult) -> dict[str, object]:
    return {
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


def build_starting_points(
    shipment_id: str,
    export_lines: list[ExportLine],
    variants_by_model: dict[str, list[BomVariant]],
    scenario_results: list[ShipmentScenarioResult],
    stock_snapshot: dict[str, list[StockBucket]],
    valuation_mode: str,
    import_lead_days: int,
    max_import_age_days: int,
) -> dict[str, StartingPointScenario]:
    if not scenario_results:
        raise SystemExit(f"No baseline scenarios available for {shipment_id}")

    heuristic_best = StartingPointScenario(
        starting_point_id="heuristic_best",
        variant_strategy="global_best_after_exhaustive_sort",
        sequence_strategy="declaration_order",
        export_order=[
            f"{line.model_code}:{line.declaration_no}:{line.declaration_item_no}"
            for line in export_lines
        ],
        result=scenario_results[0],
    )

    staff_export_lines = sorted(
        export_lines,
        key=lambda line: export_line_sort_key(line, "invoice_then_declaration"),
    )
    staff_variants = {
        model_code: pick_latest_variant(options)
        for model_code, options in variants_by_model.items()
    }
    staff_result = evaluate_scenario(
        shipment_id=shipment_id,
        export_lines=staff_export_lines,
        scenario_variants=staff_variants,
        stock_snapshot=stock_snapshot,
        valuation_mode=valuation_mode,
        import_lead_days=import_lead_days,
        max_import_age_days=max_import_age_days,
    )
    staff_seed = StartingPointScenario(
        starting_point_id="staff_latest_bom_invoice_order",
        variant_strategy="latest_dm_block_per_model",
        sequence_strategy="invoice_then_declaration",
        export_order=[
            f"{line.model_code}:{line.invoice_no or 'no-invoice'}:{line.declaration_no}:{line.declaration_item_no}"
            for line in staff_export_lines
        ],
        result=staff_result,
    )

    return {
        heuristic_best.starting_point_id: heuristic_best,
        staff_seed.starting_point_id: staff_seed,
    }


def serialize_starting_points(
    payload: dict[str, dict[str, StartingPointScenario]],
) -> dict[str, object]:
    return {
        shipment_id: {
            starting_point_id: {
                "starting_point_id": item.starting_point_id,
                "variant_strategy": item.variant_strategy,
                "sequence_strategy": item.sequence_strategy,
                "export_order": item.export_order,
                "scenario": serialize_scenario(item.result),
            }
            for starting_point_id, item in starting_points.items()
        }
        for shipment_id, starting_points in payload.items()
    }


def write_case_workspace(
    shipment_dir: Path,
    export_rows: dict[str, list[ExportLine]],
    variants_by_model: dict[str, list[BomVariant]],
    stock_snapshot: dict[str, list[StockBucket]],
    scenario_results: dict[str, list[ShipmentScenarioResult]],
    starting_points: dict[str, dict[str, StartingPointScenario]],
    run_config: dict[str, object],
) -> None:
    normalized_dir = shipment_dir / "normalized"
    results_dir = shipment_dir / "results"
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
    (results_dir / "baseline-starting-points.json").write_text(
        json.dumps(serialize_starting_points(starting_points), indent=2, ensure_ascii=False, default=str)
    )
    (results_dir / "run-config.json").write_text(
        json.dumps(run_config, indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--shipments",
        nargs="+",
        default=list(DEFAULT_SHIPMENTS),
        help="Shipment invoice ids such as GIN01426B282",
    )
    parser.add_argument(
        "--bom-source",
        choices=bom_source_choices(),
        default=DM_BOM_SOURCE_ID,
    )
    parser.add_argument("--workspace-dir", type=Path)
    parser.add_argument("--config-path", type=Path)
    parser.add_argument(
        "--valuation-mode",
        choices=("workbook_avg", "weighted"),
    )
    parser.add_argument("--policy-version")
    parser.add_argument("--import-lead-days", type=int)
    parser.add_argument("--max-import-age-days", type=int)
    parser.add_argument("--top", type=int, default=5)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    shipment_ids = set(args.shipments)
    if args.workspace_dir is not None and len(shipment_ids) != 1:
        raise SystemExit("--workspace-dir can only be used with a single shipment")

    shared_normalized_dir = args.case_dir / "shared" / "normalized"
    bom_source = resolve_bom_source(args.case_dir, args.bom_source)
    _, variants_by_model = load_bom_variants(bom_source.variant_csv_path)
    export_rows = load_export_rows(shared_normalized_dir, shipment_ids)

    all_results: dict[str, list[ShipmentScenarioResult]] = {}
    all_starting_points: dict[str, dict[str, StartingPointScenario]] = {}
    for shipment_id in args.shipments:
        if shipment_id not in export_rows:
            raise SystemExit(f"Shipment not found in normalized exports: {shipment_id}")

        policy = resolve_shipment_policy(
            args.case_dir,
            shipment_id,
            config_path_override=args.config_path,
            policy_version_override=args.policy_version,
            import_lead_days_override=args.import_lead_days,
            max_import_age_days_override=args.max_import_age_days,
            valuation_mode_override=args.valuation_mode,
        )
        shipment_workspace_dir = args.workspace_dir or (args.case_dir / shipment_slug(shipment_id))
        candidate_admissibility = load_candidate_admissibility(
            args.case_dir,
            shipment_id,
            workspace_dir=shipment_workspace_dir,
        )
        stock_snapshot = load_stock_snapshot(shared_normalized_dir, candidate_admissibility)
        scenario_results, variant_options = enumerate_shipment_scenarios(
            shipment_id=shipment_id,
            export_lines=export_rows[shipment_id],
            variants_by_model=variants_by_model,
            stock_snapshot=stock_snapshot,
            valuation_mode=policy.valuation_mode,
            import_lead_days=policy.import_lead_days,
            max_import_age_days=policy.max_import_age_days,
        )
        all_results[shipment_id] = scenario_results
        starting_points = build_starting_points(
            shipment_id=shipment_id,
            export_lines=export_rows[shipment_id],
            variants_by_model=variant_options,
            scenario_results=scenario_results,
            stock_snapshot=stock_snapshot,
            valuation_mode=policy.valuation_mode,
            import_lead_days=policy.import_lead_days,
            max_import_age_days=policy.max_import_age_days,
        )
        all_starting_points[shipment_id] = starting_points
        print_shipment_summary(
            shipment_id=shipment_id,
            export_lines=export_rows[shipment_id],
            variant_options=variant_options,
            scenario_results=scenario_results,
            top_n=args.top,
        )

        write_case_workspace(
            shipment_dir=shipment_workspace_dir,
            export_rows={shipment_id: export_rows[shipment_id]},
            variants_by_model=variant_options,
            stock_snapshot=stock_snapshot,
            scenario_results={shipment_id: scenario_results},
            starting_points={shipment_id: starting_points},
            run_config={
                "stage": "baseline",
                **policy.to_dict(),
                "bom_source_id": bom_source.source_id,
                "bom_source_label": bom_source.label,
                "bom_source_variant_csv": relative_to_case(args.case_dir, bom_source.variant_csv_path),
                "bom_reference_comparison": relative_to_case(
                    args.case_dir,
                    bom_source.comparison_summary_path,
                ),
                "workspace_dir": relative_to_case(args.case_dir, shipment_workspace_dir),
                "starting_points": {
                    "heuristic_best": {
                        "variant_strategy": "global_best_after_exhaustive_sort",
                        "sequence_strategy": "declaration_order",
                    },
                    "staff_latest_bom_invoice_order": {
                        "variant_strategy": (
                            "latest_dm_block_per_model"
                            if bom_source.source_id == DM_BOM_SOURCE_ID
                            else "canonical_variant_per_model"
                        ),
                        "sequence_strategy": "invoice_then_declaration",
                    },
                },
            },
        )
        print(
            f"Wrote case workspace {shipment_workspace_dir}"
            f" | bom_source={bom_source.source_id}"
            f" | policy={policy.policy_version}"
            f" | valuation_mode={policy.valuation_mode}"
            f" | import_lead_days={policy.import_lead_days}"
            f" | max_import_age_days={policy.max_import_age_days}"
        )

    if args.json_out:
        args.json_out.write_text(json.dumps(serialize_results(all_results), indent=2, ensure_ascii=False, default=str))
        print(f"\nWrote {args.json_out}")


if __name__ == "__main__":
    main()
