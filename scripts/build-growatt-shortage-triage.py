#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter, defaultdict
from pathlib import Path


DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENT = "GIN01426B282"
EPS = 1e-9


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--shipment", default=DEFAULT_SHIPMENT)
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


def shipment_slug(shipment_id: str) -> str:
    tail = shipment_id[-4:]
    if tail.isalnum():
        return tail.lower()
    return shipment_id.lower().replace("/", "-")


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def load_best_scenario(results_dir: Path, shipment_id: str) -> dict[str, object]:
    baseline_path = results_dir / "baseline-scenarios.json"
    with baseline_path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    scenarios = payload.get(shipment_id, [])
    if not scenarios:
        raise SystemExit(f"No scenarios found for {shipment_id} in {baseline_path}")
    return scenarios[0]


def load_variant_options(normalized_dir: Path) -> dict[str, list[str]]:
    path = normalized_dir / "bom-variants.json"
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    variant_options: dict[str, list[str]] = {}
    for model_code, items in payload.items():
        variant_options[model_code] = [clean_text(item["variant_id"]) for item in items]
    return variant_options


def classify_line(
    unmet_row: dict[str, object],
    chosen_coverage: dict[str, object] | None,
    alt_candidates: list[dict[str, object]],
    ambiguity_count: int,
    ambiguity_qty: float,
) -> tuple[str, str, dict[str, object]]:
    unmet_qty = parse_float(unmet_row["unmet_qty"])
    blocked_by_date_qty = parse_float(unmet_row["blocked_by_date_qty"])
    chosen_demand_qty = parse_float(chosen_coverage["demand_qty"]) if chosen_coverage else 0.0
    chosen_eligible_total_qty = parse_float(chosen_coverage["eligible_total_qty"]) if chosen_coverage else 0.0
    chosen_shortage_after_total_qty = (
        parse_float(chosen_coverage["shortage_after_total_qty"]) if chosen_coverage else unmet_qty
    )
    chosen_date_blocked_total_qty = (
        parse_float(chosen_coverage["date_blocked_confirmed_qty"]) + parse_float(chosen_coverage["date_blocked_candidate_qty"])
        if chosen_coverage
        else blocked_by_date_qty
    )

    best_alt = None
    if alt_candidates:
        best_alt = min(
            alt_candidates,
            key=lambda item: (
                parse_float(item["shortage_after_total_qty"]),
                parse_float(item["demand_qty"]),
                clean_text(item["variant_id"]),
            ),
        )
    variant_improvement_qty = 0.0
    if best_alt is not None:
        variant_improvement_qty = max(
            0.0,
            chosen_shortage_after_total_qty - parse_float(best_alt["shortage_after_total_qty"]),
        )
    has_variant_choice_signal = best_alt is not None and variant_improvement_qty > EPS
    has_date_block_signal = blocked_by_date_qty > EPS or chosen_date_blocked_total_qty > EPS
    has_ambiguity_signal = ambiguity_qty > EPS

    details = {
        "chosen_demand_qty": chosen_demand_qty,
        "chosen_eligible_total_qty": chosen_eligible_total_qty,
        "chosen_shortage_after_total_qty": chosen_shortage_after_total_qty,
        "chosen_date_blocked_total_qty": chosen_date_blocked_total_qty,
        "best_alt_variant_id": clean_text(best_alt["variant_id"]) if best_alt else "",
        "best_alt_material_status": clean_text(best_alt["material_status"]) if best_alt else "",
        "best_alt_demand_qty": parse_float(best_alt["demand_qty"]) if best_alt else 0.0,
        "best_alt_shortage_after_total_qty": parse_float(best_alt["shortage_after_total_qty"]) if best_alt else 0.0,
        "variant_shortage_improvement_qty": variant_improvement_qty,
        "ambiguity_row_count": ambiguity_count,
        "ambiguity_total_qty": ambiguity_qty,
        "has_variant_choice_signal": has_variant_choice_signal,
        "has_date_block_signal": has_date_block_signal,
        "has_ambiguity_signal": has_ambiguity_signal,
    }

    if has_variant_choice_signal:
        alt_variant = clean_text(best_alt["variant_id"])
        if clean_text(best_alt["material_status"]) == "material_not_in_variant":
            reason = f"Alternative variant {alt_variant} removes this material requirement."
        else:
            alt_shortage = parse_float(best_alt["shortage_after_total_qty"])
            reason = (
                f"Alternative variant {alt_variant} lowers standalone shortage "
                f"from {chosen_shortage_after_total_qty:.4f} to {alt_shortage:.4f}."
            )
        return "variant-choice-driven", reason, details

    if blocked_by_date_qty > EPS and blocked_by_date_qty + EPS >= unmet_qty:
        reason = (
            f"Date-blocked quantity {blocked_by_date_qty:.4f} is enough to cover current unmet "
            f"quantity {unmet_qty:.4f} under the active policy window."
        )
        return "date-blocked", reason, details

    if ambiguity_qty > EPS and chosen_eligible_total_qty <= EPS and blocked_by_date_qty <= EPS:
        reason = (
            f"No eligible supply is available in the chosen variant, but ambiguity rows for this material total "
            f"{ambiguity_qty:.4f} and still need resolution."
        )
        return "ambiguity-blocked", reason, details

    reason = (
        f"Remaining unmet quantity {unmet_qty:.4f} is not fully explained by date blocking or a clearly better "
        "existing variant."
    )
    return "true-shortage", reason, details


def main() -> None:
    args = parse_args()
    slug = shipment_slug(args.shipment)
    shipment_dir = args.case_dir / slug
    results_dir = shipment_dir / "results"
    normalized_dir = shipment_dir / "normalized"

    best_scenario = load_best_scenario(results_dir, args.shipment)
    selected_variants = {
        clean_text(product["model_code"]): clean_text(product["variant_id"])
        for product in best_scenario["product_results"]
    }
    variant_options = load_variant_options(normalized_dir)

    unmet_rows = list(csv.DictReader((results_dir / "top-scenario-unmet-materials.csv").open(encoding="utf-8")))
    coverage_rows = list(csv.DictReader((normalized_dir / f"{slug}-material-coverage.csv").open(encoding="utf-8")))
    ambiguity_rows = list(csv.DictReader((normalized_dir / f"{slug}-ambiguity-report.csv").open(encoding="utf-8")))

    coverage_index = {
        (clean_text(row["variant_id"]), clean_text(row["material_code"])): row
        for row in coverage_rows
    }
    coverage_by_model_material: dict[tuple[str, str], dict[str, dict[str, object]]] = defaultdict(dict)
    for row in coverage_rows:
        coverage_by_model_material[(clean_text(row["export_model_code"]), clean_text(row["material_code"]))][
            clean_text(row["variant_id"])
        ] = row

    ambiguity_by_material: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in ambiguity_rows:
        ambiguity_by_material[clean_text(row["admissible_material_code"])].append(row)

    line_rows: list[dict[str, object]] = []
    material_groups: dict[str, list[dict[str, object]]] = defaultdict(list)

    for unmet_row in unmet_rows:
        model_code = clean_text(unmet_row["model_code"])
        chosen_variant_id = clean_text(unmet_row["variant_id"])
        material_code = clean_text(unmet_row["material_code"])
        chosen_coverage = coverage_index.get((chosen_variant_id, material_code))

        alt_candidates: list[dict[str, object]] = []
        for variant_id in variant_options.get(model_code, []):
            if variant_id == chosen_variant_id:
                continue
            alt_row = coverage_by_model_material.get((model_code, material_code), {}).get(variant_id)
            if alt_row is not None:
                alt_candidates.append(
                    {
                        "variant_id": variant_id,
                        "material_status": "same_material",
                        "demand_qty": parse_float(alt_row["demand_qty"]),
                        "shortage_after_total_qty": parse_float(alt_row["shortage_after_total_qty"]),
                    }
                )
            else:
                alt_candidates.append(
                    {
                        "variant_id": variant_id,
                        "material_status": "material_not_in_variant",
                        "demand_qty": 0.0,
                        "shortage_after_total_qty": 0.0,
                    }
                )

        ambiguity_items = ambiguity_by_material.get(material_code, [])
        bucket, reason, details = classify_line(
            unmet_row=unmet_row,
            chosen_coverage=chosen_coverage,
            alt_candidates=alt_candidates,
            ambiguity_count=len(ambiguity_items),
            ambiguity_qty=sum(parse_float(row["remaining_qty"]) for row in ambiguity_items),
        )

        row = {
            "shipment_id": args.shipment,
            "model_code": model_code,
            "bom_code": clean_text(unmet_row["bom_code"]),
            "chosen_variant_id": chosen_variant_id,
            "material_code": material_code,
            "need_qty": parse_float(unmet_row["need_qty"]),
            "allocated_qty": parse_float(unmet_row["allocated_qty"]),
            "unmet_qty": parse_float(unmet_row["unmet_qty"]),
            "blocked_by_date_qty": parse_float(unmet_row["blocked_by_date_qty"]),
            "source_bucket_count": clean_text(unmet_row["source_bucket_count"]),
            "primary_bucket": bucket,
            "bucket_reason": reason,
            **details,
        }
        secondary_signals: list[str] = []
        if row["has_variant_choice_signal"] and bucket != "variant-choice-driven":
            secondary_signals.append("variant-choice")
        if row["has_date_block_signal"] and bucket != "date-blocked":
            secondary_signals.append("date-block")
        if row["has_ambiguity_signal"] and bucket != "ambiguity-blocked":
            secondary_signals.append("ambiguity")
        row["secondary_signals"] = ";".join(secondary_signals)
        line_rows.append(row)
        material_groups[material_code].append(row)

    line_rows.sort(
        key=lambda row: (
            {"true-shortage": 0, "date-blocked": 1, "variant-choice-driven": 2, "ambiguity-blocked": 3}.get(
                clean_text(row["primary_bucket"]), 9
            ),
            -parse_float(row["unmet_qty"]),
            clean_text(row["model_code"]),
            clean_text(row["material_code"]),
        )
    )

    material_rows: list[dict[str, object]] = []
    for material_code, rows in material_groups.items():
        models = sorted({clean_text(row["model_code"]) for row in rows})
        bucket_counter = Counter(clean_text(row["primary_bucket"]) for row in rows)
        qty_by_bucket: dict[str, float] = defaultdict(float)
        for row in rows:
            qty_by_bucket[clean_text(row["primary_bucket"])] += parse_float(row["unmet_qty"])
        dominant_bucket = max(
            qty_by_bucket.items(),
            key=lambda item: (item[1], item[0]),
        )[0]
        material_rows.append(
            {
                "shipment_id": args.shipment,
                "material_code": material_code,
                "affected_models": ";".join(models),
                "line_count": len(rows),
                "total_unmet_qty": sum(parse_float(row["unmet_qty"]) for row in rows),
                "max_blocked_by_date_qty": max(parse_float(row["blocked_by_date_qty"]) for row in rows),
                "dominant_bucket": dominant_bucket,
                "bucket_breakdown": ";".join(f"{bucket}:{count}" for bucket, count in sorted(bucket_counter.items())),
                "best_alt_variants": ";".join(
                    sorted({clean_text(row["best_alt_variant_id"]) for row in rows if clean_text(row["best_alt_variant_id"])})
                ),
                "ambiguity_total_qty": max(parse_float(row["ambiguity_total_qty"]) for row in rows),
            }
        )

    material_rows.sort(key=lambda row: (-parse_float(row["total_unmet_qty"]), clean_text(row["material_code"])))

    bucket_qty = Counter()
    bucket_lines = Counter()
    for row in line_rows:
        bucket = clean_text(row["primary_bucket"])
        bucket_qty[bucket] += parse_float(row["unmet_qty"])
        bucket_lines[bucket] += 1

    write_csv(
        results_dir / "best-scenario-shortage-triage.csv",
        line_rows,
        [
            "shipment_id",
            "model_code",
            "bom_code",
            "chosen_variant_id",
            "material_code",
            "need_qty",
            "allocated_qty",
            "unmet_qty",
            "blocked_by_date_qty",
            "source_bucket_count",
            "primary_bucket",
            "bucket_reason",
            "secondary_signals",
            "chosen_demand_qty",
            "chosen_eligible_total_qty",
            "chosen_shortage_after_total_qty",
            "chosen_date_blocked_total_qty",
            "best_alt_variant_id",
            "best_alt_material_status",
            "best_alt_demand_qty",
            "best_alt_shortage_after_total_qty",
            "variant_shortage_improvement_qty",
            "ambiguity_row_count",
            "ambiguity_total_qty",
            "has_variant_choice_signal",
            "has_date_block_signal",
            "has_ambiguity_signal",
        ],
    )
    write_csv(
        results_dir / "best-scenario-shortage-material-summary.csv",
        material_rows,
        [
            "shipment_id",
            "material_code",
            "affected_models",
            "line_count",
            "total_unmet_qty",
            "max_blocked_by_date_qty",
            "dominant_bucket",
            "bucket_breakdown",
            "best_alt_variants",
            "ambiguity_total_qty",
        ],
    )

    lines: list[str] = []
    lines.append(f"# {args.shipment} Shortage Triage")
    lines.append("")
    lines.append("## Bucket Totals")
    lines.append("")
    lines.append("| Bucket | Line Count | Unmet Qty |")
    lines.append("| --- | ---: | ---: |")
    for bucket in sorted(bucket_qty, key=lambda key: (-bucket_qty[key], key)):
        lines.append(f"| {bucket} | {bucket_lines[bucket]} | {bucket_qty[bucket]:.4f} |")

    lines.append("")
    lines.append("## Top Materials")
    lines.append("")
    lines.append("| Material | Models | Total Unmet Qty | Dominant Bucket | Max Date-Blocked Qty | Alt Variants |")
    lines.append("| --- | --- | ---: | --- | ---: | --- |")
    for row in material_rows[:20]:
        lines.append(
            f"| {row['material_code']} | {row['affected_models']} | {parse_float(row['total_unmet_qty']):.4f} | "
            f"{row['dominant_bucket']} | {parse_float(row['max_blocked_by_date_qty']):.4f} | {row['best_alt_variants']} |"
        )

    lines.append("")
    lines.append("## Top Line-Level Findings")
    lines.append("")
    lines.append("| Model | Material | Unmet Qty | Bucket | Date-Blocked Qty | Best Alt Variant |")
    lines.append("| --- | --- | ---: | --- | ---: | --- |")
    for row in sorted(line_rows, key=lambda item: (-parse_float(item["unmet_qty"]), clean_text(item["model_code"])))[:25]:
        lines.append(
            f"| {row['model_code']} | {row['material_code']} | {parse_float(row['unmet_qty']):.4f} | "
            f"{row['primary_bucket']} | {parse_float(row['blocked_by_date_qty']):.4f} | {row['best_alt_variant_id']} |"
        )

    (results_dir / "best-scenario-shortage-triage.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {results_dir / 'best-scenario-shortage-triage.csv'}")
    print(f"Wrote {results_dir / 'best-scenario-shortage-material-summary.csv'}")
    print(f"Wrote {results_dir / 'best-scenario-shortage-triage.md'}")


if __name__ == "__main__":
    main()
