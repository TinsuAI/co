#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path


IMPORT_LEAD_DAYS = 2
MAX_IMPORT_AGE_DAYS = 365
DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENT = "GIN01426B282"


@dataclass(frozen=True)
class ExportLine:
    shipment_id: str
    model_code: str
    declaration_no: str
    declaration_item_no: int
    export_date: date
    quantity: float


@dataclass(frozen=True)
class VariantLine:
    variant_id: str
    bom_code: str
    export_model_code: str
    material_code: str
    qty_per_unit: float


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--shipment", default=DEFAULT_SHIPMENT)
    parser.add_argument("--import-lead-days", type=int, default=IMPORT_LEAD_DAYS)
    parser.add_argument("--max-import-age-days", type=int, default=MAX_IMPORT_AGE_DAYS)
    return parser.parse_args()


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_float(value: object) -> float:
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


def is_import_eligible(import_date: date | None, export_date: date, import_lead_days: int, max_import_age_days: int) -> bool:
    if import_date is None:
        return False
    latest_allowed_date = export_date - timedelta(days=import_lead_days)
    if import_date > latest_allowed_date:
        return False
    if max_import_age_days > 0:
        earliest_allowed_date = export_date - timedelta(days=max_import_age_days)
        if import_date < earliest_allowed_date:
            return False
    return True


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def load_exports(exports_path: Path, shipment_id: str) -> list[ExportLine]:
    rows: list[ExportLine] = []
    with exports_path.open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            if row["shipment_id"] != shipment_id:
                continue
            model_code = clean_text(row["final_lookup_key"])
            export_date = to_date(row["declaration_date"]) or to_date(row["invoice_date"])
            if not model_code or export_date is None:
                continue
            rows.append(
                ExportLine(
                    shipment_id=shipment_id,
                    model_code=model_code,
                    declaration_no=clean_text(row["declaration_no"]),
                    declaration_item_no=parse_int(row["declaration_item_no"]),
                    export_date=export_date,
                    quantity=parse_float(row["quantity"]),
                )
            )
    rows.sort(key=lambda item: (item.declaration_no, item.declaration_item_no, item.model_code))
    return rows


def load_variant_lines(dm_path: Path, export_models: set[str]) -> tuple[dict[str, list[VariantLine]], dict[str, str]]:
    variant_lines: dict[str, list[VariantLine]] = defaultdict(list)
    variant_to_model: dict[str, str] = {}

    with dm_path.open(encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            bom_code = clean_text(row["bom_code"])
            export_model = next(
                (
                    model
                    for model in export_models
                    if bom_code == model or bom_code.startswith(f"{model}-")
                ),
                "",
            )
            if not export_model:
                continue
            variant_id = clean_text(row["bom_variant_id"])
            variant_to_model[variant_id] = export_model
            variant_lines[variant_id].append(
                VariantLine(
                    variant_id=variant_id,
                    bom_code=bom_code,
                    export_model_code=export_model,
                    material_code=clean_text(row["material_code"]),
                    qty_per_unit=parse_float(row["qty_per_unit"]),
                )
            )

    return variant_lines, variant_to_model


def main() -> None:
    args = parse_args()
    shared_dir = args.case_dir / "shared" / "normalized"
    shipment_dir = args.case_dir / shipment_slug(args.shipment)
    normalized_dir = shipment_dir / "normalized"
    normalized_dir.mkdir(parents=True, exist_ok=True)

    exports = load_exports(shared_dir / "exports-normalized.csv", args.shipment)
    if not exports:
        raise SystemExit(f"No normalized export rows found for {args.shipment}")

    export_models = {row.model_code for row in exports}
    export_by_model = {row.model_code: row for row in exports}
    variant_lines, variant_to_model = load_variant_lines(shared_dir / "dm-variants.csv", export_models)
    if not variant_lines:
        raise SystemExit(f"No DM variants found for shipment models in {args.shipment}")

    variant_materials = {
        variant_id: {line.material_code for line in lines}
        for variant_id, lines in variant_lines.items()
    }
    material_variant_ids: dict[str, set[str]] = defaultdict(set)
    for variant_id, material_codes in variant_materials.items():
        for material_code in material_codes:
            material_variant_ids[material_code].add(variant_id)

    admissibility_rows: list[dict[str, object]] = []
    ambiguity_rows: list[dict[str, object]] = []

    with (shared_dir / "co-stock-tracking-updated.csv").open(encoding="utf-8") as handle:
        stock_reader = csv.DictReader(handle)
        for row in stock_reader:
            remaining_qty = parse_float(row["remaining_qty"])
            if remaining_qty <= 0:
                continue

            confirmed_code = clean_text(row["confirmed_lookup_code"])
            candidate_code = clean_text(row["lookup_material_code"])
            mapping_status = clean_text(row["mapping_status"])

            if confirmed_code:
                admissibility_code = confirmed_code
                admissibility_basis = "confirmed_lookup_code"
                admissibility_status = "confirmed_for_variant"
            elif mapping_status == "candidate_exact_dm_match" and candidate_code:
                admissibility_code = candidate_code
                admissibility_basis = "candidate_exact_dm_match"
                admissibility_status = "admissible_candidate_for_variant"
            else:
                continue

            hit_variants = sorted(material_variant_ids.get(admissibility_code, set()))
            if not hit_variants:
                continue

            for variant_id in hit_variants:
                export_line = export_by_model[variant_to_model[variant_id]]
                import_date = to_date(row["declaration_date"])
                eligible_by_date = is_import_eligible(
                    import_date=import_date,
                    export_date=export_line.export_date,
                    import_lead_days=args.import_lead_days,
                    max_import_age_days=args.max_import_age_days,
                )
                admissibility_rows.append(
                    {
                        "shipment_id": args.shipment,
                        "variant_id": variant_id,
                        "bom_code": variant_lines[variant_id][0].bom_code,
                        "export_model_code": export_line.model_code,
                        "export_date": export_line.export_date.isoformat(),
                        "tracking_key": clean_text(row["tracking_key"]),
                        "declaration_no": clean_text(row["declaration_no"]),
                        "declaration_item_no": parse_int(row["declaration_item_no"]),
                        "admissible_material_code": admissibility_code,
                        "admissibility_basis": admissibility_basis,
                        "admissibility_status": admissibility_status,
                        "shipment_variant_scope": (
                            "confirmed"
                            if confirmed_code
                            else "ambiguous_across_variants"
                            if len(hit_variants) > 1
                            else "single_variant_only"
                        ),
                        "variant_hit_count": len(hit_variants),
                        "variant_hit_ids": ";".join(hit_variants),
                        "eligible_by_date": eligible_by_date,
                        "mapping_status": mapping_status,
                        "lookup_material_code": candidate_code,
                        "confirmed_lookup_code": confirmed_code,
                        "remaining_qty": remaining_qty,
                        "import_qty": parse_float(row["import_qty"]),
                        "import_date": clean_text(row["declaration_date"]),
                        "invoice_date": clean_text(row["invoice_date"]),
                        "origin": clean_text(row["origin"]),
                        "unit_price_usd": parse_float(row["unit_price_usd"]),
                        "source": clean_text(row["source"]),
                        "nk2_source_row_no": clean_text(row["nk2_source_row_no"]),
                        "bcct_source_row_no": clean_text(row["bcct_source_row_no"]),
                    }
                )

            if not confirmed_code and len(hit_variants) > 1:
                ambiguity_rows.append(
                    {
                        "shipment_id": args.shipment,
                        "tracking_key": clean_text(row["tracking_key"]),
                        "declaration_no": clean_text(row["declaration_no"]),
                        "declaration_item_no": parse_int(row["declaration_item_no"]),
                        "admissible_material_code": admissibility_code,
                        "variant_hit_count": len(hit_variants),
                        "variant_hit_ids": ";".join(hit_variants),
                        "remaining_qty": remaining_qty,
                        "mapping_status": mapping_status,
                        "origin": clean_text(row["origin"]),
                        "unit_price_usd": parse_float(row["unit_price_usd"]),
                        "import_date": clean_text(row["declaration_date"]),
                        "invoice_date": clean_text(row["invoice_date"]),
                        "ambiguity_reason": "candidate_row_hits_multiple_b282_variants",
                    }
                )

    coverage_rows: list[dict[str, object]] = []
    coverage_index: dict[tuple[str, str], dict[str, object]] = {}
    for variant_id, lines in variant_lines.items():
        export_line = export_by_model[variant_to_model[variant_id]]
        for line in lines:
            key = (variant_id, line.material_code)
            if key not in coverage_index:
                coverage_index[key] = {
                    "shipment_id": args.shipment,
                    "variant_id": variant_id,
                    "bom_code": line.bom_code,
                    "export_model_code": export_line.model_code,
                    "export_date": export_line.export_date.isoformat(),
                    "material_code": line.material_code,
                    "qty_per_unit": 0.0,
                    "export_qty": export_line.quantity,
                    "demand_qty": 0.0,
                    "eligible_confirmed_qty": 0.0,
                    "eligible_candidate_qty": 0.0,
                    "date_blocked_confirmed_qty": 0.0,
                    "date_blocked_candidate_qty": 0.0,
                    "eligible_confirmed_row_count": 0,
                    "eligible_candidate_row_count": 0,
                }
                coverage_rows.append(coverage_index[key])
            coverage_index[key]["qty_per_unit"] += line.qty_per_unit
            coverage_index[key]["demand_qty"] += line.qty_per_unit * export_line.quantity

    for row in admissibility_rows:
        coverage = coverage_index[(clean_text(row["variant_id"]), clean_text(row["admissible_material_code"]))]
        qty = parse_float(row["remaining_qty"])
        if row["eligible_by_date"] in (True, "True", "true"):
            if row["admissibility_status"] == "confirmed_for_variant":
                coverage["eligible_confirmed_qty"] += qty
                coverage["eligible_confirmed_row_count"] += 1
            else:
                coverage["eligible_candidate_qty"] += qty
                coverage["eligible_candidate_row_count"] += 1
        else:
            if row["admissibility_status"] == "confirmed_for_variant":
                coverage["date_blocked_confirmed_qty"] += qty
            else:
                coverage["date_blocked_candidate_qty"] += qty

    for row in coverage_rows:
        eligible_total_qty = row["eligible_confirmed_qty"] + row["eligible_candidate_qty"]
        row["eligible_total_qty"] = eligible_total_qty
        row["shortage_after_confirmed_qty"] = max(0.0, row["demand_qty"] - row["eligible_confirmed_qty"])
        row["shortage_after_total_qty"] = max(0.0, row["demand_qty"] - eligible_total_qty)

    slug = shipment_slug(args.shipment)
    write_csv(
        normalized_dir / f"{slug}-variant-admissibility.csv",
        admissibility_rows,
        [
            "shipment_id",
            "variant_id",
            "bom_code",
            "export_model_code",
            "export_date",
            "tracking_key",
            "declaration_no",
            "declaration_item_no",
            "admissible_material_code",
            "admissibility_basis",
            "admissibility_status",
            "shipment_variant_scope",
            "variant_hit_count",
            "variant_hit_ids",
            "eligible_by_date",
            "mapping_status",
            "lookup_material_code",
            "confirmed_lookup_code",
            "remaining_qty",
            "import_qty",
            "import_date",
            "invoice_date",
            "origin",
            "unit_price_usd",
            "source",
            "nk2_source_row_no",
            "bcct_source_row_no",
        ],
    )
    write_csv(
        normalized_dir / f"{slug}-material-coverage.csv",
        coverage_rows,
        [
            "shipment_id",
            "variant_id",
            "bom_code",
            "export_model_code",
            "export_date",
            "material_code",
            "qty_per_unit",
            "export_qty",
            "demand_qty",
            "eligible_confirmed_qty",
            "eligible_candidate_qty",
            "eligible_total_qty",
            "shortage_after_confirmed_qty",
            "shortage_after_total_qty",
            "date_blocked_confirmed_qty",
            "date_blocked_candidate_qty",
            "eligible_confirmed_row_count",
            "eligible_candidate_row_count",
        ],
    )
    write_csv(
        normalized_dir / f"{slug}-ambiguity-report.csv",
        ambiguity_rows,
        [
            "shipment_id",
            "tracking_key",
            "declaration_no",
            "declaration_item_no",
            "admissible_material_code",
            "variant_hit_count",
            "variant_hit_ids",
            "remaining_qty",
            "mapping_status",
            "origin",
            "unit_price_usd",
            "import_date",
            "invoice_date",
            "ambiguity_reason",
        ],
    )

    print(f"Wrote admissibility artifacts to {normalized_dir}")
    print(f"Variants: {len(variant_lines)}")
    print(f"Admissibility rows: {len(admissibility_rows)}")
    print(f"Ambiguity rows: {len(ambiguity_rows)}")


if __name__ == "__main__":
    main()
