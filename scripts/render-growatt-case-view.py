#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("case_dir", type=Path)
    parser.add_argument("--top", type=int, default=5)
    return parser.parse_args()


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def fmt_num(value: object, digits: int = 2) -> str:
    if value is None:
        return "n/a"
    if isinstance(value, (int, float)):
        return f"{value:.{digits}f}"
    return str(value)


def main() -> None:
    args = parse_args()
    results_dir = args.case_dir / "results"
    baseline_path = results_dir / "baseline-scenarios.json"
    if not baseline_path.exists():
        raise SystemExit(f"Missing {baseline_path}")

    with baseline_path.open(encoding="utf-8") as handle:
        payload = json.load(handle)

    shipment_id = next(iter(payload))
    scenarios = payload[shipment_id]
    top_scenarios = scenarios[: args.top]

    scenario_rows: list[dict[str, object]] = []
    product_rows: list[dict[str, object]] = []

    for rank, scenario in enumerate(scenarios, start=1):
        scenario_rows.append(
            {
                "shipment_id": shipment_id,
                "rank": rank,
                "stock_sufficient": scenario["stock_sufficient"],
                "passes_rvc": scenario["passes_rvc"],
                "min_margin": scenario["min_margin"],
                "total_unmet_qty": scenario["total_unmet_qty"],
                "valuation_mode": scenario["valuation_mode"],
                "scenario_id": scenario["scenario_id"],
            }
        )
        for product in scenario["product_results"]:
            product_rows.append(
                {
                    "shipment_id": shipment_id,
                    "scenario_rank": rank,
                    "model_code": product["model_code"],
                    "bom_code": product["bom_code"],
                    "variant_id": product["variant_id"],
                    "export_date": product["export_date"],
                    "export_qty": product["export_qty"],
                    "rvc_percent": product["rvc_percent"],
                    "passes_rvc": product["passes_rvc"],
                    "stock_sufficient": product["stock_sufficient"],
                    "unmet_material_count": product["unmet_material_count"],
                    "unmet_qty_total": product["unmet_qty_total"],
                    "multi_source_material_count": product["multi_source_material_count"],
                }
            )

    top_unmet_material_rows: list[dict[str, object]] = []
    top_blocked_material_rows: list[dict[str, object]] = []
    if scenarios:
        best = scenarios[0]
        for product in best["product_results"]:
            for material in product["materials"]:
                if material["unmet_qty"] > 0:
                    top_unmet_material_rows.append(
                        {
                            "shipment_id": shipment_id,
                            "model_code": product["model_code"],
                            "bom_code": product["bom_code"],
                            "variant_id": product["variant_id"],
                            "material_code": material["material_code"],
                            "need_qty": material["need_qty"],
                            "allocated_qty": material["allocated_qty"],
                            "unmet_qty": material["unmet_qty"],
                            "blocked_by_date_qty": material["blocked_by_date_qty"],
                            "source_bucket_count": material["source_bucket_count"],
                            "source_declarations": ";".join(material["source_declarations"]),
                        }
                    )
                if material["blocked_by_date_qty"] > 0:
                    top_blocked_material_rows.append(
                        {
                            "shipment_id": shipment_id,
                            "model_code": product["model_code"],
                            "bom_code": product["bom_code"],
                            "variant_id": product["variant_id"],
                            "material_code": material["material_code"],
                            "need_qty": material["need_qty"],
                            "blocked_by_date_qty": material["blocked_by_date_qty"],
                            "unmet_qty": material["unmet_qty"],
                        }
                    )

    top_unmet_material_rows.sort(key=lambda row: (-float(row["unmet_qty"]), str(row["model_code"]), str(row["material_code"])))
    top_blocked_material_rows.sort(
        key=lambda row: (-float(row["blocked_by_date_qty"]), str(row["model_code"]), str(row["material_code"]))
    )

    write_csv(
        results_dir / "scenario-summary.csv",
        scenario_rows,
        [
            "shipment_id",
            "rank",
            "stock_sufficient",
            "passes_rvc",
            "min_margin",
            "total_unmet_qty",
            "valuation_mode",
            "scenario_id",
        ],
    )
    write_csv(
        results_dir / "product-summary.csv",
        product_rows,
        [
            "shipment_id",
            "scenario_rank",
            "model_code",
            "bom_code",
            "variant_id",
            "export_date",
            "export_qty",
            "rvc_percent",
            "passes_rvc",
            "stock_sufficient",
            "unmet_material_count",
            "unmet_qty_total",
            "multi_source_material_count",
        ],
    )
    write_csv(
        results_dir / "top-scenario-unmet-materials.csv",
        top_unmet_material_rows,
        [
            "shipment_id",
            "model_code",
            "bom_code",
            "variant_id",
            "material_code",
            "need_qty",
            "allocated_qty",
            "unmet_qty",
            "blocked_by_date_qty",
            "source_bucket_count",
            "source_declarations",
        ],
    )
    write_csv(
        results_dir / "top-scenario-date-blocked-materials.csv",
        top_blocked_material_rows,
        [
            "shipment_id",
            "model_code",
            "bom_code",
            "variant_id",
            "material_code",
            "need_qty",
            "blocked_by_date_qty",
            "unmet_qty",
        ],
    )

    lines: list[str] = []
    lines.append(f"# Shipment {shipment_id} Baseline View")
    lines.append("")
    lines.append("## Top Scenarios")
    lines.append("")
    lines.append("| Rank | Stock | RVC | Min Margin | Unmet Qty | Scenario |")
    lines.append("| --- | --- | --- | ---: | ---: | --- |")
    for row in top_scenarios:
        rank = scenarios.index(row) + 1
        lines.append(
            f"| {rank} | {'ok' if row['stock_sufficient'] else 'fail'} | "
            f"{'pass' if row['passes_rvc'] else 'fail'} | {fmt_num(row['min_margin'])} | "
            f"{fmt_num(row['total_unmet_qty'], 4)} | {row['scenario_id']} |"
        )

    if scenarios:
        best = scenarios[0]
        lines.append("")
        lines.append("## Best Scenario Product View")
        lines.append("")
        lines.append("| Model | BOM Code | Variant | RVC | Stock | Unmet Materials | Unmet Qty | Multi-Source |")
        lines.append("| --- | --- | --- | ---: | --- | ---: | ---: | ---: |")
        for product in best["product_results"]:
            lines.append(
                f"| {product['model_code']} | {product['bom_code']} | {product['variant_id']} | "
                f"{fmt_num(product['rvc_percent'])}% | {'ok' if product['stock_sufficient'] else 'fail'} | "
                f"{product['unmet_material_count']} | {fmt_num(product['unmet_qty_total'], 4)} | "
                f"{product['multi_source_material_count']} |"
            )

        lines.append("")
        lines.append("## Top Unmet Materials In Best Scenario")
        lines.append("")
        lines.append("| Model | Material | Need Qty | Allocated Qty | Unmet Qty | Date-Blocked Qty | Source Buckets |")
        lines.append("| --- | --- | ---: | ---: | ---: | ---: | ---: |")
        for row in top_unmet_material_rows[:25]:
            lines.append(
                f"| {row['model_code']} | {row['material_code']} | {fmt_num(row['need_qty'], 4)} | "
                f"{fmt_num(row['allocated_qty'], 4)} | {fmt_num(row['unmet_qty'], 4)} | "
                f"{fmt_num(row['blocked_by_date_qty'], 4)} | {row['source_bucket_count']} |"
            )

        lines.append("")
        lines.append("## Top Date-Blocked Materials In Best Scenario")
        lines.append("")
        lines.append("| Model | Material | Need Qty | Date-Blocked Qty | Unmet Qty |")
        lines.append("| --- | --- | ---: | ---: | ---: |")
        for row in top_blocked_material_rows[:25]:
            lines.append(
                f"| {row['model_code']} | {row['material_code']} | {fmt_num(row['need_qty'], 4)} | "
                f"{fmt_num(row['blocked_by_date_qty'], 4)} | {fmt_num(row['unmet_qty'], 4)} |"
            )

    (results_dir / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {results_dir / 'summary.md'}")


if __name__ == "__main__":
    main()
