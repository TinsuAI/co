#!/usr/bin/env python3

from __future__ import annotations

import argparse
import copy
import importlib.util
import itertools
import json
import sys
from pathlib import Path

from growatt_case_config import resolve_shipment_policy


DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DEFAULT_SHIPMENT = "GIN01426B282"
DEFAULT_DOC = Path("docs/growatt-b282-sequential-priority-report.md")
EPS = 1e-9


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--shipment", default=DEFAULT_SHIPMENT)
    parser.add_argument("--output-doc", type=Path, default=DEFAULT_DOC)
    parser.add_argument("--top", type=int, default=20)
    parser.add_argument("--max-scenarios", type=int, default=0)
    return parser.parse_args()


def load_baseline_module() -> object:
    module_path = Path(__file__).with_name("growatt-rvc-baseline.py")
    spec = importlib.util.spec_from_file_location("growatt_rvc_baseline_runtime", module_path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"Unable to load baseline module from {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def shipment_slug(shipment_id: str) -> str:
    tail = shipment_id[-4:]
    if tail.isalnum():
        return tail.lower()
    return shipment_id.lower().replace("/", "-")


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    import csv

    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def fmt(value: float | None, digits: int = 4) -> str:
    if value is None:
        return "n/a"
    return f"{value:.{digits}f}"


def clone_stock_snapshot(stock_snapshot: dict[str, list[object]]) -> dict[str, list[object]]:
    return {
        material_code: [copy.copy(bucket) for bucket in buckets]
        for material_code, buckets in stock_snapshot.items()
    }


def evaluate_order_fast(
    baseline: object,
    *,
    shipment_id: str,
    ordered_export_lines: list[object],
    scenario_variants: dict[str, object],
    stock_snapshot: dict[str, list[object]],
    valuation_mode: str,
    import_lead_days: int,
    max_import_age_days: int,
) -> dict[str, object]:
    stock_state = clone_stock_snapshot(stock_snapshot)
    product_details: list[dict[str, object]] = []

    for sequence_no, export_line in enumerate(ordered_export_lines, start=1):
        variant = scenario_variants[export_line.model_code]
        non_origin_total = 0.0
        unmet_material_count = 0
        unmet_qty_total = 0.0

        for bom_line in variant.lines:
            need_qty = bom_line.qty_per_unit * export_line.quantity
            remaining_need = need_qty
            allocations: list[tuple[object, float]] = []

            for bucket in stock_state.get(bom_line.material_code, []):
                if remaining_need <= EPS:
                    break
                if bucket.remaining_qty <= EPS:
                    continue
                if not baseline.bucket_is_admissible_for_variant(bucket, variant.variant_id):
                    continue
                if not baseline.is_bucket_eligible_for_export(
                    bucket,
                    export_line,
                    import_lead_days=import_lead_days,
                    max_import_age_days=max_import_age_days,
                ):
                    continue
                take_qty = min(bucket.remaining_qty, remaining_need)
                if take_qty <= EPS:
                    continue
                bucket.remaining_qty -= take_qty
                remaining_need -= take_qty
                allocations.append((bucket, take_qty))

            if remaining_need > EPS:
                unmet_material_count += 1
                unmet_qty_total += remaining_need

            unit_price_usd = baseline.compute_material_unit_price(allocations, valuation_mode)
            all_vn = bool(allocations) and all(
                baseline.is_vietnam_origin(bucket.origin)
                for bucket, _ in allocations
            )
            if unit_price_usd is not None and not all_vn:
                non_origin_total += need_qty * unit_price_usd

        fob_value_usd = export_line.quantity * export_line.unit_price_usd
        rvc_percent = None
        if fob_value_usd > 0:
            rvc_percent = ((fob_value_usd - non_origin_total) / fob_value_usd) * 100.0
        stock_sufficient = unmet_material_count == 0
        passes_rvc = stock_sufficient and rvc_percent is not None and rvc_percent >= baseline.TARGET_RVC
        margin_to_threshold = None
        if rvc_percent is not None:
            margin_to_threshold = rvc_percent - baseline.TARGET_RVC

        product_details.append(
            {
                "sequence_no": sequence_no,
                "model_code": export_line.model_code,
                "variant_id": variant.variant_id,
                "bom_code": variant.bom_code,
                "passes_rvc": passes_rvc,
                "stock_sufficient": stock_sufficient,
                "rvc_percent": rvc_percent,
                "margin_to_threshold": margin_to_threshold,
                "unmet_qty_total": unmet_qty_total,
            }
        )

    passed_models = [item["model_code"] for item in product_details if item["passes_rvc"]]
    failed_models = [item["model_code"] for item in product_details if not item["passes_rvc"]]
    pass_margins = [
        float(item["margin_to_threshold"])
        for item in product_details
        if item["passes_rvc"] and item["margin_to_threshold"] is not None
    ]
    total_unmet_qty = sum(float(item["unmet_qty_total"]) for item in product_details)
    return {
        "passed_product_count": len(passed_models),
        "passed_models": passed_models,
        "failed_models": failed_models,
        "total_pass_margin": sum(pass_margins),
        "min_pass_margin": min(pass_margins) if pass_margins else None,
        "total_unmet_qty": total_unmet_qty,
        "total_failed_unmet_qty": sum(
            float(item["unmet_qty_total"]) for item in product_details if not item["passes_rvc"]
        ),
        "product_details": product_details,
        "order_signature": " -> ".join(item["model_code"] for item in product_details),
    }


def load_top_variant_signatures(results_dir: Path, shipment_id: str, max_scenarios: int) -> list[str]:
    baseline_path = results_dir / "baseline-scenarios.json"
    with baseline_path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    scenarios = payload.get(shipment_id, [])
    ordered: list[str] = []
    seen: set[str] = set()
    for scenario in scenarios:
        bits = []
        for product in scenario["product_results"]:
            bits.append((str(product["model_code"]), str(product["variant_id"])))
        signature = " | ".join(f"{model}:{variant}" for model, variant in sorted(bits))
        if signature in seen:
            continue
        seen.add(signature)
        ordered.append(signature)
        if max_scenarios > 0 and len(ordered) >= max_scenarios:
            break
    return ordered


def main() -> None:
    args = parse_args()
    baseline = load_baseline_module()
    policy = resolve_shipment_policy(args.case_dir, args.shipment)

    shared_normalized_dir = args.case_dir / "shared" / "normalized"
    shipment_dir = args.case_dir / shipment_slug(args.shipment)
    results_dir = shipment_dir / "results"
    results_dir.mkdir(parents=True, exist_ok=True)

    _, variants_by_model = baseline.load_dm_variants(shared_normalized_dir)
    export_rows = baseline.load_export_rows(shared_normalized_dir, {args.shipment})
    if args.shipment not in export_rows:
        raise SystemExit(f"Shipment not found in normalized exports: {args.shipment}")
    export_lines = export_rows[args.shipment]

    candidate_admissibility = baseline.load_candidate_admissibility(args.case_dir, args.shipment)
    stock_snapshot = baseline.load_stock_snapshot(shared_normalized_dir, candidate_admissibility)

    variant_options: dict[str, list[object]] = {}
    variant_by_id: dict[str, object] = {}
    for export_line in export_lines:
        options = baseline.related_variants_for_model(export_line.model_code, variants_by_model)
        if not options:
            raise SystemExit(f"No BOM variants found for {export_line.model_code}")
        variant_options[export_line.model_code] = options
        for option in options:
            variant_by_id[option.variant_id] = option

    export_permutations = list(itertools.permutations(export_lines))
    top_variant_signatures = load_top_variant_signatures(results_dir, args.shipment, args.max_scenarios)
    total_evaluations = len(export_permutations) * len(top_variant_signatures)

    best_per_scenario: dict[str, dict[str, object]] = {}
    for scenario_index, variant_signature in enumerate(top_variant_signatures, start=1):
        scenario_variants: dict[str, object] = {}
        for item in variant_signature.split(" | "):
            model_code, variant_id = item.split(":", 1)
            scenario_variants[model_code] = variant_by_id[variant_id]

        best_row: dict[str, object] | None = None
        best_sort_key: tuple[float, float, float, float, str] | None = None
        for permutation in export_permutations:
            fast_result = evaluate_order_fast(
                baseline,
                shipment_id=args.shipment,
                ordered_export_lines=list(permutation),
                scenario_variants=scenario_variants,
                stock_snapshot=stock_snapshot,
                valuation_mode=policy.valuation_mode,
                import_lead_days=policy.import_lead_days,
                max_import_age_days=policy.max_import_age_days,
            )
            sort_key = (
                float(fast_result["passed_product_count"]),
                float(fast_result["total_pass_margin"]),
                -float(fast_result["total_failed_unmet_qty"]),
                -float(fast_result["total_unmet_qty"]),
                str(fast_result["order_signature"]),
            )
            if best_sort_key is None or sort_key > best_sort_key:
                best_sort_key = sort_key
                best_row = {
                    "scenario_index": scenario_index,
                    "variant_signature": variant_signature,
                    **fast_result,
                }

        if best_row is None:
            continue
        best_per_scenario[variant_signature] = best_row

    ranked_rows = sorted(
        best_per_scenario.values(),
        key=lambda row: (
            -int(row["passed_product_count"]),
            -float(row["total_pass_margin"]),
            float(row["total_failed_unmet_qty"]),
            float(row["total_unmet_qty"]),
            str(row["order_signature"]),
        ),
    )
    global_ranked = ranked_rows[: args.top]

    write_csv(
        results_dir / "sequential-priority-best-per-scenario.csv",
        [
            {
                **{k: v for k, v in row.items() if k != "product_details"},
                "passed_models": ";".join(row["passed_models"]),
                "failed_models": ";".join(row["failed_models"]),
                "product_details": json.dumps(row["product_details"], ensure_ascii=False),
            }
            for row in ranked_rows
        ],
        [
            "scenario_index",
            "variant_signature",
            "order_signature",
            "passed_product_count",
            "passed_models",
            "failed_models",
            "total_pass_margin",
            "min_pass_margin",
            "total_unmet_qty",
            "total_failed_unmet_qty",
            "product_details",
        ],
    )
    (results_dir / "sequential-priority-best-orders.json").write_text(
        json.dumps(global_ranked, indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )

    best_product_rows = global_ranked[0]["product_details"] if global_ranked else []
    write_csv(
        results_dir / "sequential-priority-best-product-order.csv",
        best_product_rows,
        [
            "sequence_no",
            "model_code",
            "variant_id",
            "bom_code",
            "passes_rvc",
            "stock_sufficient",
            "rvc_percent",
            "margin_to_threshold",
            "unmet_qty_total",
        ],
    )

    lines: list[str] = []
    lines.append("# Báo Cáo Ưu Tiên Xử Lý Tuần Tự Growatt B282")
    lines.append("")
    lines.append(f"- Lô hàng: `{args.shipment}`")
    lines.append(f"- Phiên bản policy: `{policy.policy_version}`")
    if args.max_scenarios > 0:
        lines.append(
            f"- Phạm vi search dùng cho báo cáo này: top `{len(best_per_scenario)}` BOM scenario theo baseline x "
            f"`{len(export_permutations)}` thứ tự product = `{total_evaluations}` tổ hợp exact"
        )
    else:
        lines.append(
            f"- Phạm vi search dùng cho báo cáo này: toàn bộ `{len(best_per_scenario)}` BOM scenario x "
            f"`{len(export_permutations)}` thứ tự product = `{total_evaluations}` tổ hợp exact"
        )
    lines.append(
        "- Objective: tối đa số product đạt C/O trước, sau đó tối đa tổng margin vượt ngưỡng, rồi tối thiểu unmet còn lại"
    )
    lines.append("")

    if global_ranked:
        best = global_ranked[0]
        lines.append("## Thứ Tự Tốt Nhất")
        lines.append("")
        lines.append(f"- Số product đạt C/O: `{best['passed_product_count']}`")
        lines.append(f"- Thứ tự xử lý: `{best['order_signature']}`")
        lines.append(f"- Tổ hợp BOM variant: `{best['variant_signature']}`")
        lines.append(f"- Product đạt C/O: `{';'.join(best['passed_models']) or 'không có'}`")
        lines.append(f"- Product không đạt C/O: `{';'.join(best['failed_models']) or 'không có'}`")
        lines.append(f"- Tổng margin của các product đạt C/O: `{fmt(best['total_pass_margin'])}`")
        lines.append(f"- Tổng unmet còn lại sau khi chạy tuần tự: `{fmt(best['total_unmet_qty'])}`")
        lines.append("")
        lines.append("## Chi Tiết Theo Thứ Tự Product")
        lines.append("")
        lines.append("| STT | Model | Variant | Đủ Stock | Đạt C/O | RVC | Margin So Với 35 | Unmet Qty |")
        lines.append("| --- | --- | --- | --- | --- | ---: | ---: | ---: |")
        for product in best["product_details"]:
            lines.append(
                f"| {product['sequence_no']} | {product['model_code']} | {product['variant_id']} | "
                f"{'có' if product['stock_sufficient'] else 'không'} | "
                f"{'có' if product['passes_rvc'] else 'không'} | {fmt(product['rvc_percent'], 2)} | "
                f"{fmt(product['margin_to_threshold'], 2)} | {fmt(product['unmet_qty_total'])} |"
            )
        lines.append("")
        lines.append("## Top Tổ Hợp Scenario/Order")
        lines.append("")
        lines.append("| Hạng | Số Product Đạt C/O | Thứ Tự | Tổ Hợp Variant | Tổng Margin Đạt C/O | Unmet Qty |")
        lines.append("| --- | ---: | --- | --- | ---: | ---: |")
        for rank, row in enumerate(global_ranked, start=1):
            lines.append(
                f"| {rank} | {row['passed_product_count']} | {row['order_signature']} | {row['variant_signature']} | "
                f"{fmt(row['total_pass_margin'])} | {fmt(row['total_unmet_qty'])} |"
            )

    args.output_doc.parent.mkdir(parents=True, exist_ok=True)
    args.output_doc.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"Wrote {results_dir / 'sequential-priority-best-per-scenario.csv'}")
    print(f"Wrote {results_dir / 'sequential-priority-best-orders.json'}")
    print(f"Wrote {results_dir / 'sequential-priority-best-product-order.csv'}")
    print(f"Wrote {args.output_doc}")


if __name__ == "__main__":
    main()
