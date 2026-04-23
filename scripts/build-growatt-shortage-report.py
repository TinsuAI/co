#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter, defaultdict
from pathlib import Path


DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENT = "GIN01426B282"
DEFAULT_DOC = Path("docs/growatt-b282-shortage-report.md")
DEFAULT_CSV = "best-scenario-shortage-report-material-summary.csv"
TARGET_RVC = 35.0
EPS = 1e-9


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--shipment", default=DEFAULT_SHIPMENT)
    parser.add_argument("--output-doc", type=Path, default=DEFAULT_DOC)
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


def fmt(value: float | None, digits: int = 4) -> str:
    if value is None:
        return "n/a"
    return f"{value:.{digits}f}"


def join_codes(values: list[str]) -> str:
    return ";".join(sorted({value for value in values if value}))


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def load_best_scenario(results_dir: Path, shipment_id: str) -> dict[str, object]:
    payload = json.loads((results_dir / "baseline-scenarios.json").read_text(encoding="utf-8"))
    scenarios = payload.get(shipment_id, [])
    if not scenarios:
        raise SystemExit(f"No baseline scenarios found for {shipment_id}")
    return scenarios[0]


def build_line_material_index(best_scenario: dict[str, object]) -> dict[tuple[str, str], dict[str, float]]:
    line_items: dict[tuple[str, str], dict[str, float]] = {}
    for product in best_scenario["product_results"]:
        model_code = clean_text(product["model_code"])
        for material in product["materials"]:
            material_code = clean_text(material["material_code"])
            key = (model_code, material_code)
            row = line_items.setdefault(
                key,
                {
                    "need_qty": 0.0,
                    "allocated_qty": 0.0,
                    "unmet_qty": 0.0,
                    "blocked_by_date_qty": 0.0,
                    "non_origin_value_usd": 0.0,
                },
            )
            row["need_qty"] += parse_float(material["need_qty"])
            row["allocated_qty"] += parse_float(material["allocated_qty"])
            row["unmet_qty"] += parse_float(material["unmet_qty"])
            row["blocked_by_date_qty"] += parse_float(material["blocked_by_date_qty"])
            row["non_origin_value_usd"] += parse_float(material["non_origin_value_usd"])
    return line_items


def aggregate_triage_rows(
    triage_rows: list[dict[str, str]],
    line_material_index: dict[tuple[str, str], dict[str, float]],
) -> dict[str, list[dict[str, object]]]:
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    for triage_row in triage_rows:
        model_code = clean_text(triage_row["model_code"])
        material_code = clean_text(triage_row["material_code"])
        key = (material_code, model_code)
        line = line_material_index.get((model_code, material_code), {})
        item = grouped.setdefault(
            key,
            {
                "model_code": model_code,
                "material_code": material_code,
                "variant_id": clean_text(triage_row["chosen_variant_id"]),
                "need_qty": line.get("need_qty", 0.0),
                "allocated_qty": line.get("allocated_qty", 0.0),
                "line_unmet_qty": line.get("unmet_qty", 0.0),
                "line_blocked_by_date_qty": line.get("blocked_by_date_qty", 0.0),
                "non_origin_value_usd": line.get("non_origin_value_usd", 0.0),
                "triage_unmet_qty": 0.0,
                "true_shortage_qty": 0.0,
                "date_blocked_unmet_qty": 0.0,
                "variant_choice_unmet_qty": 0.0,
                "ambiguity_qty": 0.0,
                "best_alt_variants": set(),
                "primary_buckets": Counter(),
                "secondary_signals": set(),
            },
        )
        unmet_qty = parse_float(triage_row["unmet_qty"])
        item["triage_unmet_qty"] += unmet_qty
        bucket = clean_text(triage_row["primary_bucket"])
        item["primary_buckets"][bucket] += 1
        if bucket == "true-shortage":
            item["true_shortage_qty"] += unmet_qty
        elif bucket == "date-blocked":
            item["date_blocked_unmet_qty"] += unmet_qty
        elif bucket == "variant-choice-driven":
            item["variant_choice_unmet_qty"] += unmet_qty
        item["ambiguity_qty"] = max(item["ambiguity_qty"], parse_float(triage_row["ambiguity_total_qty"]))
        best_alt_variant_id = clean_text(triage_row["best_alt_variant_id"])
        if best_alt_variant_id:
            item["best_alt_variants"].add(best_alt_variant_id)
        secondary = clean_text(triage_row["secondary_signals"])
        if secondary:
            item["secondary_signals"].update(bit for bit in secondary.split(";") if bit)

    by_material: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in grouped.values():
        row["best_alt_variants"] = sorted(row["best_alt_variants"])
        row["secondary_signals"] = sorted(row["secondary_signals"])
        by_material[clean_text(row["material_code"])].append(row)
    for rows in by_material.values():
        rows.sort(key=lambda row: (-float(row["line_unmet_qty"]), clean_text(row["model_code"])))
    return by_material


def pick_band(
    unlockable_models: list[str],
    has_true_shortage: bool,
    has_date_blocked: bool,
    has_ambiguity_signal: bool,
    has_variant_choice: bool,
) -> str:
    if not unlockable_models:
        return "low-yield"
    if has_true_shortage:
        return "true-shortage"
    if has_date_blocked:
        return "date-blocked"
    if has_ambiguity_signal:
        return "ambiguity-blocked"
    if has_variant_choice:
        return "variant-choice-driven"
    return "review"


def operational_read(
    band: str,
    has_date_blocked: bool,
    has_ambiguity_signal: bool,
    has_variant_choice: bool,
) -> str:
    if band == "low-yield":
        return "Affected models are still below RVC 35 before fixing this material, so this gap is not a near-term unlock."
    if band == "true-shortage":
        if has_variant_choice:
            return "Same-code supply is still short; some existing BOM variants reduce part of the gap but do not clear it."
        if has_date_blocked or has_ambiguity_signal:
            return "Same-code supply remains short even after visible date-blocked or ambiguity evidence is taken into account."
        return "Same-code supply is genuinely short under the active policy window."
    if band == "date-blocked":
        if has_ambiguity_signal:
            return "Visible same-code supply exists, but it is currently blocked by the date rule and still carries ambiguity evidence."
        return "Visible same-code supply exists, but it is blocked by the date rule under the active policy window."
    if band == "ambiguity-blocked":
        return "Visible same-code supply exists, but candidate or ambiguity rows still need evidence resolution."
    if band == "variant-choice-driven":
        return "Current unmet quantity is driven by the chosen BOM variant; an existing variant can remove or reduce this requirement."
    return "This material needs manual review because the current signals do not fit a single clean bucket."


def band_order(band: str) -> int:
    order = {
        "true-shortage": 0,
        "date-blocked": 1,
        "ambiguity-blocked": 2,
        "variant-choice-driven": 3,
        "low-yield": 4,
        "review": 5,
    }
    return order.get(band, 99)


def main() -> None:
    args = parse_args()
    slug = shipment_slug(args.shipment)
    shipment_dir = args.case_dir / slug
    results_dir = shipment_dir / "results"
    normalized_dir = shipment_dir / "normalized"

    best_scenario = load_best_scenario(results_dir, args.shipment)
    run_config = json.loads((results_dir / "run-config.json").read_text(encoding="utf-8"))
    triage_rows = list(csv.DictReader((results_dir / "best-scenario-shortage-triage.csv").open(encoding="utf-8")))
    material_summary_rows = list(
        csv.DictReader((results_dir / "best-scenario-shortage-material-summary.csv").open(encoding="utf-8"))
    )
    coverage_rows = list(csv.DictReader((normalized_dir / f"{slug}-material-coverage.csv").open(encoding="utf-8")))
    ambiguity_rows = list(csv.DictReader((normalized_dir / f"{slug}-ambiguity-report.csv").open(encoding="utf-8")))

    line_material_index = build_line_material_index(best_scenario)
    material_lines = aggregate_triage_rows(triage_rows, line_material_index)
    material_summary = {clean_text(row["material_code"]): row for row in material_summary_rows}
    product_metrics = {
        clean_text(product["model_code"]): {
            "rvc_percent": parse_float(product["rvc_percent"]) if product["rvc_percent"] is not None else None,
            "margin_to_threshold": (
                parse_float(product["rvc_percent"]) - TARGET_RVC
                if product["rvc_percent"] is not None
                else None
            ),
        }
        for product in best_scenario["product_results"]
    }
    chosen_variants = {
        clean_text(product["model_code"]): clean_text(product["variant_id"])
        for product in best_scenario["product_results"]
    }

    chosen_coverage: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in coverage_rows:
        model_code = clean_text(row["export_model_code"])
        if chosen_variants.get(model_code) != clean_text(row["variant_id"]):
            continue
        chosen_coverage[clean_text(row["material_code"])].append(row)

    ambiguity_by_material: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in ambiguity_rows:
        ambiguity_by_material[clean_text(row["admissible_material_code"])].append(row)

    summary_rows: list[dict[str, object]] = []
    for material_code, lines in material_lines.items():
        coverage_items = chosen_coverage.get(material_code, [])
        summary_row = material_summary.get(material_code, {})
        shipment_demand_qty = sum(parse_float(row["demand_qty"]) for row in coverage_items)
        eligible_confirmed_qty = max((parse_float(row["eligible_confirmed_qty"]) for row in coverage_items), default=0.0)
        eligible_candidate_qty = max((parse_float(row["eligible_candidate_qty"]) for row in coverage_items), default=0.0)
        eligible_total_qty = max((parse_float(row["eligible_total_qty"]) for row in coverage_items), default=0.0)
        date_blocked_pool_qty = max(
            (
                parse_float(row["date_blocked_confirmed_qty"]) + parse_float(row["date_blocked_candidate_qty"])
                for row in coverage_items
            ),
            default=0.0,
        )
        confirmed_gap_qty = max(0.0, shipment_demand_qty - eligible_confirmed_qty)
        total_gap_qty = max(0.0, shipment_demand_qty - eligible_total_qty)
        ambiguity_total_qty = sum(parse_float(row["remaining_qty"]) for row in ambiguity_by_material.get(material_code, []))
        ambiguity_variant_hits = join_codes(
            [
                clean_text(row["variant_hit_ids"])
                for row in ambiguity_by_material.get(material_code, [])
                if clean_text(row["variant_hit_ids"])
            ]
        )
        affected_models = sorted({clean_text(row["model_code"]) for row in lines})
        unlockable_models = sorted(
            model_code
            for model_code in affected_models
            if product_metrics.get(model_code, {}).get("margin_to_threshold") is not None
            and float(product_metrics[model_code]["margin_to_threshold"]) >= 0.0
        )
        blocked_models = sorted(set(affected_models) - set(unlockable_models))
        has_true_shortage = any(float(row["true_shortage_qty"]) > EPS for row in lines)
        has_date_blocked = (
            any(float(row["date_blocked_unmet_qty"]) > EPS for row in lines) or date_blocked_pool_qty > EPS
        )
        has_ambiguity_signal = ambiguity_total_qty > EPS or eligible_candidate_qty > EPS or any(
            bool(row["secondary_signals"]) for row in lines
        )
        has_variant_choice = any(float(row["variant_choice_unmet_qty"]) > EPS for row in lines)
        band = pick_band(
            unlockable_models=unlockable_models,
            has_true_shortage=has_true_shortage,
            has_date_blocked=has_date_blocked,
            has_ambiguity_signal=has_ambiguity_signal,
            has_variant_choice=has_variant_choice,
        )

        summary_rows.append(
            {
                "shipment_id": args.shipment,
                "material_code": material_code,
                "band": band,
                "dominant_bucket": clean_text(summary_row.get("dominant_bucket")),
                "bucket_breakdown": clean_text(summary_row.get("bucket_breakdown")),
                "affected_models": ";".join(affected_models),
                "unlockable_models": ";".join(unlockable_models),
                "blocked_models": ";".join(blocked_models),
                "shipment_demand_qty": shipment_demand_qty,
                "eligible_confirmed_qty": eligible_confirmed_qty,
                "eligible_candidate_qty": eligible_candidate_qty,
                "eligible_total_qty": eligible_total_qty,
                "confirmed_gap_qty": confirmed_gap_qty,
                "total_gap_qty": total_gap_qty,
                "total_unmet_qty": parse_float(summary_row.get("total_unmet_qty")),
                "true_shortage_qty": sum(parse_float(row["true_shortage_qty"]) for row in lines),
                "date_blocked_unmet_qty": sum(parse_float(row["date_blocked_unmet_qty"]) for row in lines),
                "variant_choice_unmet_qty": sum(parse_float(row["variant_choice_unmet_qty"]) for row in lines),
                "date_blocked_pool_qty": date_blocked_pool_qty,
                "ambiguity_total_qty": ambiguity_total_qty,
                "ambiguity_variant_hits": ambiguity_variant_hits,
                "best_alt_variants": join_codes(
                    [variant_id for row in lines for variant_id in row["best_alt_variants"]]
                ),
                "non_origin_gap_exposure_usd": sum(parse_float(row["non_origin_value_usd"]) for row in lines),
                "operational_read": operational_read(
                    band=band,
                    has_date_blocked=has_date_blocked,
                    has_ambiguity_signal=has_ambiguity_signal,
                    has_variant_choice=has_variant_choice,
                ),
            }
        )

    summary_rows.sort(
        key=lambda row: (
            band_order(clean_text(row["band"])),
            -float(row["total_unmet_qty"]),
            clean_text(row["material_code"]),
        )
    )

    csv_fields = [
        "shipment_id",
        "material_code",
        "band",
        "dominant_bucket",
        "bucket_breakdown",
        "affected_models",
        "unlockable_models",
        "blocked_models",
        "shipment_demand_qty",
        "eligible_confirmed_qty",
        "eligible_candidate_qty",
        "eligible_total_qty",
        "confirmed_gap_qty",
        "total_gap_qty",
        "total_unmet_qty",
        "true_shortage_qty",
        "date_blocked_unmet_qty",
        "variant_choice_unmet_qty",
        "date_blocked_pool_qty",
        "ambiguity_total_qty",
        "ambiguity_variant_hits",
        "best_alt_variants",
        "non_origin_gap_exposure_usd",
        "operational_read",
    ]
    write_csv(results_dir / DEFAULT_CSV, summary_rows, csv_fields)

    lane_counts = Counter(clean_text(row["band"]) for row in summary_rows)
    lane_unmet = defaultdict(float)
    for row in summary_rows:
        lane_unmet[clean_text(row["band"])] += parse_float(row["total_unmet_qty"])

    lines_out: list[str] = []
    lines_out.append("# Growatt B282 Shortage Report")
    lines_out.append("")
    lines_out.append(f"- Shipment: `{args.shipment}`")
    lines_out.append(f"- Policy version: `{run_config['policy_version']}`")
    lines_out.append(
        f"- Policy: import at least `{run_config['import_lead_days']}` days before export, "
        f"max import age `{run_config['max_import_age_days']}` (`0` means no cap)"
    )
    lines_out.append(
        "- Scope: this report covers all unmet materials in the current best B282 baseline. "
        "It distinguishes real same-code shortage from date-blocked, ambiguity, variant-choice, and low-yield lanes."
    )
    lines_out.append("")
    lines_out.append("## Lane Summary")
    lines_out.append("")
    lines_out.append("| Band | Material Count | Total Unmet |")
    lines_out.append("| --- | ---: | ---: |")
    for band in ["true-shortage", "date-blocked", "ambiguity-blocked", "variant-choice-driven", "low-yield"]:
        if lane_counts[band] <= 0:
            continue
        lines_out.append(f"| {band} | {lane_counts[band]} | {fmt(lane_unmet[band])} |")

    lines_out.append("")
    lines_out.append("## Material Summary")
    lines_out.append("")
    lines_out.append(
        "| Material | Band | Affected Models | Total Unmet | Confirmed Gap | Total Gap | Date-Blocked Pool | Ambiguity | Bucket Mix |"
    )
    lines_out.append("| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |")
    for row in summary_rows:
        lines_out.append(
            f"| {row['material_code']} | {row['band']} | "
            f"{clean_text(row['affected_models']) or 'none'} | "
            f"{fmt(row['total_unmet_qty'])} | {fmt(row['confirmed_gap_qty'])} | {fmt(row['total_gap_qty'])} | "
            f"{fmt(row['date_blocked_pool_qty'])} | {fmt(row['ambiguity_total_qty'])} | "
            f"{clean_text(row['bucket_breakdown']) or clean_text(row['dominant_bucket']) or 'n/a'} |"
        )

    lines_out.append("")
    lines_out.append("## Findings")
    lines_out.append("")
    for row in summary_rows:
        lines_out.append(f"### `{row['material_code']}`")
        lines_out.append(f"- Band: `{row['band']}`")
        lines_out.append(
            f"- Gap view: total unmet `{fmt(row['total_unmet_qty'])}`, true-shortage `{fmt(row['true_shortage_qty'])}`, "
            f"date-blocked unmet `{fmt(row['date_blocked_unmet_qty'])}`, variant-choice unmet `{fmt(row['variant_choice_unmet_qty'])}`"
        )
        lines_out.append(
            f"- Same-code pool: demand `{fmt(row['shipment_demand_qty'])}`, confirmed `{fmt(row['eligible_confirmed_qty'])}`, "
            f"candidate `{fmt(row['eligible_candidate_qty'])}`, total eligible `{fmt(row['eligible_total_qty'])}`, "
            f"date-blocked pool `{fmt(row['date_blocked_pool_qty'])}`, ambiguity `{fmt(row['ambiguity_total_qty'])}`"
        )
        lines_out.append(
            f"- Models: affected `{clean_text(row['affected_models']) or 'none'}`, "
            f"unlockable `{clean_text(row['unlockable_models']) or 'none'}`, "
            f"below-threshold `{clean_text(row['blocked_models']) or 'none'}`"
        )
        if clean_text(row["best_alt_variants"]):
            lines_out.append(f"- Existing alt variants: `{row['best_alt_variants']}`")
        if clean_text(row["ambiguity_variant_hits"]):
            lines_out.append(f"- Ambiguity variant hits: `{row['ambiguity_variant_hits']}`")
        lines_out.append(f"- Triage mix: `{clean_text(row['bucket_breakdown']) or clean_text(row['dominant_bucket']) or 'n/a'}`")
        lines_out.append(f"- Operational read: {row['operational_read']}")
        lines_out.append("")

    args.output_doc.parent.mkdir(parents=True, exist_ok=True)
    args.output_doc.write_text("\n".join(lines_out), encoding="utf-8")
    print(f"Wrote {args.output_doc}")


if __name__ == "__main__":
    main()
