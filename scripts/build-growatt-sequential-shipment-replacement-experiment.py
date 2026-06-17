#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from dataclasses import dataclass
from pathlib import Path

import growatt_replacement_runner as replacement
from growatt_bom_sources import DEFAULT_CASE_DIR, technical_priority_artifact_dir
from growatt_case_config import resolve_shipment_policy


DEFAULT_SHIPMENTS = ("GIN01426C171", "GIN01426B282")
DEFAULT_OUTPUT_DIRNAME = "sequential-c171-first-2026-04-23"


@dataclass(frozen=True)
class LaneConfig:
    lane_id: str
    label: str
    starting_point_by_shipment: dict[str, str]
    workspace_by_shipment: dict[str, Path]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument(
        "--shipments",
        nargs="+",
        default=list(DEFAULT_SHIPMENTS),
    )
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--lanes", nargs="+")
    return parser.parse_args()


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field) for field in fieldnames})


def build_lane_configs(case_dir: Path) -> list[LaneConfig]:
    return [
        LaneConfig(
            lane_id="dm-heuristic-best",
            label="DM variants | heuristic_best",
            starting_point_by_shipment={
                "GIN01426C171": "heuristic_best",
                "GIN01426B282": "heuristic_best",
            },
            workspace_by_shipment={
                "GIN01426C171": case_dir / "c171",
                "GIN01426B282": case_dir / "b282",
            },
        ),
        LaneConfig(
            lane_id="dm-staff-latest",
            label="DM variants | staff_latest_bom_invoice_order",
            starting_point_by_shipment={
                "GIN01426C171": "staff_latest_bom_invoice_order",
                "GIN01426B282": "staff_latest_bom_invoice_order",
            },
            workspace_by_shipment={
                "GIN01426C171": case_dir / "c171",
                "GIN01426B282": case_dir / "b282",
            },
        ),
        LaneConfig(
            lane_id="technical-synthesized",
            label="Technical synthesized BOM | heuristic_best",
            starting_point_by_shipment={
                "GIN01426C171": "heuristic_best",
                "GIN01426B282": "heuristic_best",
            },
            workspace_by_shipment={
                "GIN01426C171": case_dir / "c171-tech-bom-2026-04-23",
                "GIN01426B282": case_dir / "b282-tech-bom-2026-04-23",
            },
        ),
        LaneConfig(
            lane_id="c171-tech-pass3-then-b282-dm-heuristic",
            label="C171 technical pass3 | B282 DM heuristic_best",
            starting_point_by_shipment={
                "GIN01426C171": "heuristic_best",
                "GIN01426B282": "heuristic_best",
            },
            workspace_by_shipment={
                "GIN01426C171": case_dir / "c171-tech-bom-2026-04-23-pass3of3",
                "GIN01426B282": case_dir / "b282",
            },
        ),
        LaneConfig(
            lane_id="c171-tech-pass3-then-b282-dm-staff",
            label="C171 technical pass3 | B282 DM staff_latest_bom_invoice_order",
            starting_point_by_shipment={
                "GIN01426C171": "heuristic_best",
                "GIN01426B282": "staff_latest_bom_invoice_order",
            },
            workspace_by_shipment={
                "GIN01426C171": case_dir / "c171-tech-bom-2026-04-23-pass3of3",
                "GIN01426B282": case_dir / "b282",
            },
        ),
    ]


def load_selected_starting_point(workspace_dir: Path, shipment_id: str, starting_point_id: str) -> dict[str, object]:
    path = workspace_dir / "results" / "baseline-starting-points.json"
    payload = json.loads(path.read_text(encoding="utf-8"))
    shipment_payload = payload[shipment_id]
    if starting_point_id not in shipment_payload:
        raise SystemExit(f"Missing starting point {starting_point_id} in {path}")
    return {starting_point_id: shipment_payload[starting_point_id]}


def build_variant_lookup(
    variants_by_model: dict[str, list[object]],
    selected_payload: dict[str, object],
) -> dict[str, dict[str, object]]:
    variant_by_id = {
        variant.variant_id: variant
        for items in variants_by_model.values()
        for variant in items
    }
    return {
        starting_point_id: {
            item["model_code"]: variant_by_id[item["variant_id"]]
            for item in payload["scenario"]["product_results"]
        }
        for starting_point_id, payload in selected_payload.items()
    }


def apply_residual_qty(
    basis_buckets: list[replacement.ReplacementBasisBucket],
    residual_qty_by_bucket: dict[str, float],
) -> None:
    if not residual_qty_by_bucket:
        return
    for bucket in basis_buckets:
        if bucket.bucket_id in residual_qty_by_bucket:
            bucket.remaining_qty = residual_qty_by_bucket[bucket.bucket_id]


def summarize_seed(seed_summary: dict[str, object]) -> dict[str, object]:
    return {
        "starting_point_id": seed_summary["starting_point_id"],
        "pass_count": seed_summary["pass_count"],
        "final_unmet_qty": seed_summary["final_unmet_qty"],
        "min_margin": seed_summary["min_margin"],
        "changed_products": ";".join(seed_summary["changed_products"]),
        "changed_materials": ";".join(seed_summary["changed_materials"]),
    }


def write_selected_payload(path: Path, shipment_id: str, payload: dict[str, object]) -> None:
    path.write_text(
        json.dumps({shipment_id: payload}, indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )


def render_summary_doc(
    path: Path,
    shipment_order: list[str],
    lane_rows: list[dict[str, object]],
) -> None:
    lines = [
        "# Growatt Sequential Shipment Replacement Experiment",
        "",
        f"- Shipment order: `{' -> '.join(shipment_order)}`",
        "",
        "| lane_id | shipment_id | starting_point_id | pass_count | final_unmet_qty | min_margin |",
        "| --- | --- | --- | ---: | ---: | ---: |",
    ]
    for row in lane_rows:
        min_margin = row["min_margin"]
        min_margin_text = "" if min_margin in (None, "") else f"{float(min_margin):.6f}"
        lines.append(
            "| "
            + " | ".join(
                [
                    str(row["lane_id"]),
                    str(row["shipment_id"]),
                    str(row["starting_point_id"]),
                    str(row["pass_count"]),
                    f"{float(row['final_unmet_qty']):.6f}",
                    min_margin_text,
                ]
            )
            + " |"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_lane(
    baseline: object,
    case_dir: Path,
    lane: LaneConfig,
    shipment_order: list[str],
    output_dir: Path,
) -> list[dict[str, object]]:
    shared_normalized_dir = case_dir / "shared" / "normalized"
    staff_substitutes = replacement.read_staff_substitutes()
    substitute_reference_path = (
        technical_priority_artifact_dir(case_dir)
        / "substitute-reference"
        / "explicit-substitute-links.csv"
    )
    explicit_substitute_index = replacement.load_explicit_substitute_index(substitute_reference_path)
    residual_qty_by_bucket: dict[str, float] = {}
    lane_rows: list[dict[str, object]] = []

    for sequence_index, shipment_id in enumerate(shipment_order, start=1):
        workspace_dir = lane.workspace_by_shipment.get(shipment_id)
        if workspace_dir is None:
            raise SystemExit(f"Lane {lane.lane_id} does not define a workspace for {shipment_id}")
        starting_point_id = lane.starting_point_by_shipment.get(shipment_id)
        if starting_point_id is None:
            raise SystemExit(f"Lane {lane.lane_id} does not define a starting point for {shipment_id}")
        shipment_slug = replacement.shipment_slug(shipment_id)
        shipment_output_dir = output_dir / lane.lane_id / shipment_slug
        shipment_output_dir.mkdir(parents=True, exist_ok=True)

        export_rows = baseline.load_export_rows(shared_normalized_dir, {shipment_id})
        selected_payload = load_selected_starting_point(workspace_dir, shipment_id, starting_point_id)
        starting_point_exports = replacement.resolve_starting_point_export_lines(
            selected_payload,
            shipment_id,
            export_rows,
        )
        variants_by_model = replacement.load_workspace_variants(workspace_dir / "normalized")
        start_variant_by_seed = build_variant_lookup(variants_by_model, selected_payload)

        basis_buckets = replacement.build_replacement_basis(
            shared_normalized_dir,
            workspace_dir,
            shipment_slug,
        )
        apply_residual_qty(basis_buckets, residual_qty_by_bucket)
        stock_by_material, _, _ = replacement.clone_stock_indexes(basis_buckets)
        material_basis_stats = replacement.build_material_basis_stats(basis_buckets)

        write_selected_payload(
            shipment_output_dir / "replacement-starting-points.json",
            shipment_id,
            selected_payload,
        )
        replacement.write_csv(
            shipment_output_dir / "replacement-custom-code-basis.csv",
            replacement.build_basis_csv_rows(basis_buckets),
            [
                "custom_code_basis",
                "erp_material_code",
                "candidate_material_code",
                "bucket_id",
                "tracking_key",
                "candidate_class",
                "mapping_status",
                "mapping_confidence",
                "evidence_source",
                "remaining_qty",
                "unit_price_usd",
                "origin",
                "shipment_variant_scope",
                "variant_scope",
                "variant_hit_count",
                "source",
                "source_row_no",
                "declaration_no",
                "declaration_item_no",
                "import_date",
                "hs_code",
                "name",
            ],
        )

        class VariantLookup(dict):
            def __missing__(self, key: str) -> object:
                raise KeyError(key)

        policy = resolve_shipment_policy(case_dir, shipment_id)
        result = replacement.run_starting_point(
            baseline,
            shipment_id=shipment_id,
            starting_point_id=starting_point_id,
            export_lines=starting_point_exports[starting_point_id],
            variant_by_id=VariantLookup(start_variant_by_seed[starting_point_id]),
            initial_stock_by_material=stock_by_material,
            material_basis_stats=material_basis_stats,
            staff_substitutes=staff_substitutes,
            import_lead_days=policy.import_lead_days,
            max_import_age_days=policy.max_import_age_days,
            output_dir=shipment_output_dir,
            replacement_mode=replacement.REPLACEMENT_MODE_HEURISTIC_CUSTOM_BASIS,
            explicit_substitute_index=explicit_substitute_index,
            collect_candidate_rows=False,
        )

        replacement.write_csv(
            shipment_output_dir / "replacement-seed-summary.csv",
            [summarize_seed(result["seed_summary"])],
            [
                "starting_point_id",
                "pass_count",
                "final_unmet_qty",
                "min_margin",
                "changed_products",
                "changed_materials",
            ],
        )
        replacement.write_csv(
            shipment_output_dir / "replacement-product-status.csv",
            result["product_status_rows"],
            [
                "starting_point_id",
                "sequence_no",
                "model_code",
                "variant_id",
                "bom_code",
                "declaration_no",
                "rvc_before",
                "rvc_after",
                "stock_before",
                "stock_after",
                "passes_before",
                "passes_after",
                "unmet_before",
                "unmet_after",
                "changed_materials",
                "snapshot_path",
            ],
        )
        replacement.write_csv(
            shipment_output_dir / "replacement-final-stock-by-bucket.csv",
            result["stock_exports"]["final_bucket_rows"],
            [
                "material_code",
                "bucket_material_code",
                "candidate_material_code",
                "custom_code_basis",
                "bucket_id",
                "tracking_key",
                "declaration_no",
                "declaration_item_no",
                "import_date",
                "source",
                "source_row_no",
                "hs_code",
                "name",
                "origin",
                "unit_price_usd",
                "exchange_rate",
                "remaining_qty",
                "candidate_class",
                "mapping_status",
                "mapping_confidence",
                "evidence_source",
                "shipment_variant_scope",
                "variant_scope",
                "variant_hit_count",
            ],
        )
        replacement.write_csv(
            shipment_output_dir / "replacement-final-stock-by-material.csv",
            result["stock_exports"]["final_material_rows"],
            ["material_code", "remaining_qty", "bucket_count"],
        )

        final_summary = result["seed_summary"]
        row = {
            "lane_id": lane.lane_id,
            "lane_label": lane.label,
            "sequence_index": sequence_index,
            "shipment_id": shipment_id,
            "workspace_dir": str(workspace_dir),
            "starting_point_id": starting_point_id,
            "pass_count": final_summary["pass_count"],
            "final_unmet_qty": final_summary["final_unmet_qty"],
            "min_margin": final_summary["min_margin"],
            "changed_products": ";".join(final_summary["changed_products"]),
            "changed_material_count": len(final_summary["changed_materials"]),
            "output_dir": str(shipment_output_dir),
        }
        lane_rows.append(row)

        residual_qty_by_bucket.update(
            {
                str(item["bucket_id"]): float(item["remaining_qty"])
                for item in result["stock_exports"]["final_bucket_rows"]
            }
        )

        (shipment_output_dir / "replacement-run-config.json").write_text(
            json.dumps(
                {
                    "stage": "sequential_shipment_replacement_experiment",
                    "lane_id": lane.lane_id,
                    "lane_label": lane.label,
                    "shipment_order": shipment_order,
                    "sequence_index": sequence_index,
                    "shipment_id": shipment_id,
                    "workspace_dir": str(workspace_dir),
                    "output_dir": str(shipment_output_dir),
                    "starting_point_id": starting_point_id,
                    "replacement_mode": replacement.REPLACEMENT_MODE_HEURISTIC_CUSTOM_BASIS,
                    "substitute_reference_path": str(substitute_reference_path),
                    "carryover_bucket_count": len(residual_qty_by_bucket),
                    "policy_version": policy.policy_version,
                    "import_lead_days": policy.import_lead_days,
                    "max_import_age_days": policy.max_import_age_days,
                    "valuation_mode": policy.valuation_mode,
                },
                indent=2,
                ensure_ascii=False,
                default=str,
            )
            + "\n",
            encoding="utf-8",
        )

    return lane_rows


def main() -> None:
    args = parse_args()
    case_dir = args.case_dir
    shipment_order = list(args.shipments)
    output_dir = args.output_dir or (case_dir / DEFAULT_OUTPUT_DIRNAME)
    output_dir.mkdir(parents=True, exist_ok=True)

    baseline = replacement.load_baseline_module()
    selected_lane_ids = set(args.lanes or [])
    lane_rows: list[dict[str, object]] = []
    for lane in build_lane_configs(case_dir):
        if selected_lane_ids and lane.lane_id not in selected_lane_ids:
            continue
        lane_rows.extend(run_lane(baseline, case_dir, lane, shipment_order, output_dir))

    write_csv(
        output_dir / "sequence-summary.csv",
        lane_rows,
        [
            "lane_id",
            "lane_label",
            "sequence_index",
            "shipment_id",
            "workspace_dir",
            "starting_point_id",
            "pass_count",
            "final_unmet_qty",
            "min_margin",
            "changed_products",
            "changed_material_count",
            "output_dir",
        ],
    )
    render_summary_doc(output_dir / "summary.md", shipment_order, lane_rows)
    (output_dir / "manifest.json").write_text(
        json.dumps(
            {
                "shipment_order": shipment_order,
                "lanes": lane_rows,
            },
            indent=2,
            ensure_ascii=False,
            default=str,
        )
        + "\n",
        encoding="utf-8",
    )
    print(f"Wrote sequential shipment experiment to {output_dir}")


if __name__ == "__main__":
    main()
