#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict
from pathlib import Path


DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENT = "GIN01426B282"
DEFAULT_OUTPUT = Path("docs/growatt-b282-true-shortage-report.md")
EPS = 1e-9


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--shipment", default=DEFAULT_SHIPMENT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
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


def load_best_scenario(results_dir: Path, shipment_id: str) -> dict[str, object]:
    with (results_dir / "baseline-scenarios.json").open(encoding="utf-8") as handle:
        payload = json.load(handle)
    scenarios = payload.get(shipment_id, [])
    if not scenarios:
        raise SystemExit(f"No scenarios found for {shipment_id}")
    return scenarios[0]


def fmt(value: float) -> str:
    return f"{value:.4f}"


def main() -> None:
    args = parse_args()
    slug = shipment_slug(args.shipment)
    shipment_dir = args.case_dir / slug
    results_dir = shipment_dir / "results"
    normalized_dir = shipment_dir / "normalized"

    best_scenario = load_best_scenario(results_dir, args.shipment)
    run_config = json.loads((results_dir / "run-config.json").read_text(encoding="utf-8"))
    triage_rows = list(csv.DictReader((results_dir / "best-scenario-shortage-triage.csv").open(encoding="utf-8")))
    material_rows = list(
        csv.DictReader((results_dir / "best-scenario-shortage-material-summary.csv").open(encoding="utf-8"))
    )
    coverage_rows = list(csv.DictReader((normalized_dir / f"{slug}-material-coverage.csv").open(encoding="utf-8")))

    chosen_variants = {
        clean_text(product["model_code"]): clean_text(product["variant_id"])
        for product in best_scenario["product_results"]
    }
    chosen_coverage_by_material: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in coverage_rows:
        model_code = clean_text(row["export_model_code"])
        if chosen_variants.get(model_code) != clean_text(row["variant_id"]):
            continue
        chosen_coverage_by_material[clean_text(row["material_code"])].append(row)

    true_rows = [row for row in triage_rows if clean_text(row["primary_bucket"]) == "true-shortage"]
    true_rows_by_material: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in true_rows:
        true_rows_by_material[clean_text(row["material_code"])].append(row)

    summary_by_material = {
        clean_text(row["material_code"]): row
        for row in material_rows
    }

    report_items: list[dict[str, object]] = []
    for material_code, rows in true_rows_by_material.items():
        chosen_rows = chosen_coverage_by_material.get(material_code, [])
        if not chosen_rows:
            continue

        shipment_total_demand = sum(parse_float(row["demand_qty"]) for row in chosen_rows)
        shared_eligible_qty = max(parse_float(row["eligible_total_qty"]) for row in chosen_rows)
        shipment_level_gap = max(0.0, shipment_total_demand - shared_eligible_qty)
        shared_date_blocked_qty = max(
            parse_float(row["date_blocked_confirmed_qty"]) + parse_float(row["date_blocked_candidate_qty"])
            for row in chosen_rows
        )
        true_unmet_qty = sum(parse_float(row["unmet_qty"]) for row in rows)
        affected_models = sorted({clean_text(row["model_code"]) for row in rows})
        alt_variants = sorted({clean_text(row["best_alt_variant_id"]) for row in rows if clean_text(row["best_alt_variant_id"])})
        max_variant_improvement = max(parse_float(row["variant_shortage_improvement_qty"]) for row in rows)
        has_competition = len(chosen_rows) > 1

        if max_variant_improvement > EPS:
            action = "Review existing variant choices before any replacement BOM work."
        elif has_competition:
            action = "Check shared-stock allocation and stock evidence for this material across all affected models."
        else:
            action = "Check stock evidence first; if unchanged, treat this as a candidate for replacement-driven BOM analysis."

        if shared_date_blocked_qty > EPS:
            explanation = (
                f"Still shows {fmt(shared_date_blocked_qty)} date-blocked quantity under the current lead-time rule, "
                "but remaining gap is not explained away by removing the one-year cap."
            )
        elif has_competition:
            explanation = (
                f"Shipment-level demand {fmt(shipment_total_demand)} exceeds shared eligible supply {fmt(shared_eligible_qty)}. "
                "This is a shared-pool shortage across multiple products."
            )
        else:
            explanation = (
                f"Eligible supply {fmt(shared_eligible_qty)} is below the chosen variant demand {fmt(shipment_total_demand)} "
                "even after removing the one-year cap."
            )

        report_items.append(
            {
                "material_code": material_code,
                "affected_models": affected_models,
                "true_unmet_qty": true_unmet_qty,
                "shipment_total_demand": shipment_total_demand,
                "shared_eligible_qty": shared_eligible_qty,
                "shipment_level_gap": shipment_level_gap,
                "shared_date_blocked_qty": shared_date_blocked_qty,
                "alt_variants": alt_variants,
                "max_variant_improvement": max_variant_improvement,
                "action": action,
                "explanation": explanation,
                "bucket_breakdown": clean_text(summary_by_material.get(material_code, {}).get("bucket_breakdown")),
            }
        )

    report_items.sort(key=lambda item: (-float(item["true_unmet_qty"]), item["material_code"]))

    bucket_totals = defaultdict(float)
    for row in triage_rows:
        bucket_totals[clean_text(row["primary_bucket"])] += parse_float(row["unmet_qty"])

    lines: list[str] = []
    lines.append("# Growatt B282 True Shortage Report")
    lines.append("")
    lines.append(f"- Shipment: `{args.shipment}`")
    lines.append(f"- Policy version: `{run_config['policy_version']}`")
    lines.append(
        f"- Policy: import at least `{run_config['import_lead_days']}` days before export, "
        f"max import age `{run_config['max_import_age_days']}` (`0` means no cap)"
    )
    lines.append(f"- Best scenario unmet qty: `{fmt(best_scenario['total_unmet_qty'])}`")
    lines.append(
        f"- Bucket totals: true-shortage `{fmt(bucket_totals['true-shortage'])}`, "
        f"variant-choice-driven `{fmt(bucket_totals['variant-choice-driven'])}`, "
        f"date-blocked `{fmt(bucket_totals['date-blocked'])}`"
    )
    lines.append("")
    lines.append("## What Counts As True Shortage")
    lines.append("")
    lines.append(
        "These are the gaps that still remain after removing the one-year stock window. "
        "They are not primarily explained by the date policy anymore. Some are single-material shortages; "
        "others are shipment-level shortages where multiple products compete for the same eligible stock pool."
    )
    lines.append("")
    lines.append("## Highest-Impact Materials")
    lines.append("")
    lines.append("| Material | Affected Models | True Unmet Qty | Shipment Demand | Eligible Supply | Shipment Gap | Alt Variants |")
    lines.append("| --- | --- | ---: | ---: | ---: | ---: | --- |")
    for item in report_items[:12]:
        lines.append(
            f"| {item['material_code']} | {';'.join(item['affected_models'])} | {fmt(item['true_unmet_qty'])} | "
            f"{fmt(item['shipment_total_demand'])} | {fmt(item['shared_eligible_qty'])} | {fmt(item['shipment_level_gap'])} | "
            f"{';'.join(item['alt_variants'])} |"
        )

    lines.append("")
    lines.append("## Findings")
    lines.append("")
    for item in report_items[:8]:
        lines.append(f"### `{item['material_code']}`")
        lines.append(f"- Affected models: `{'`, `'.join(item['affected_models'])}`")
        lines.append(f"- True unmet quantity in best scenario: `{fmt(item['true_unmet_qty'])}`")
        lines.append(
            f"- Shipment demand vs eligible supply: `{fmt(item['shipment_total_demand'])}` vs `{fmt(item['shared_eligible_qty'])}` "
            f"(gap `{fmt(item['shipment_level_gap'])}`)"
        )
        if item["bucket_breakdown"]:
            lines.append(f"- Current triage mix: `{item['bucket_breakdown']}`")
        if item["alt_variants"]:
            lines.append(f"- Existing variants worth checking: `{'`, `'.join(item['alt_variants'])}`")
            if item["max_variant_improvement"] > EPS:
                lines.append(
                    f"- Best observed variant-driven improvement on a line: `{fmt(item['max_variant_improvement'])}`"
                )
        else:
            lines.append("- Existing variants worth checking: none observed")
        lines.append(f"- Why it is still a true shortage: {item['explanation']}")
        lines.append(f"- Operational read: {item['action']}")
        lines.append("")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
