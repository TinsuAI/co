#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict
from pathlib import Path

from growatt_bom_sources import (
    DEFAULT_CASE_DIR,
    TECHNICAL_PRIORITY_BOM_SOURCE_ID,
    TECHNICAL_PRIORITY_VARIANT_SUFFIX,
    technical_priority_artifact_dir,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--artifact-dir", type=Path)
    parser.add_argument(
        "--source-id",
        default=TECHNICAL_PRIORITY_BOM_SOURCE_ID,
    )
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


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def load_json(path: Path) -> object:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def build_variant_rows(
    per_code_dir: Path,
    *,
    source_id: str = TECHNICAL_PRIORITY_BOM_SOURCE_ID,
    variant_suffix: str = TECHNICAL_PRIORITY_VARIANT_SUFFIX,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    rows_out: list[dict[str, object]] = []
    code_summaries: list[dict[str, object]] = []

    for csv_path in sorted(per_code_dir.glob("*.csv")):
        with csv_path.open(encoding="utf-8") as handle:
            source_rows = list(csv.DictReader(handle))
        if not source_rows:
            continue

        root_codes = {clean_text(row["rootProductCode"]) for row in source_rows}
        if len(root_codes) != 1:
            raise ValueError(f"Expected one rootProductCode in {csv_path}, found {sorted(root_codes)}")
        root_code = next(iter(root_codes))
        variant_id = f"{root_code}__{variant_suffix}"
        material_counts: dict[str, int] = defaultdict(int)

        for row_idx, row in enumerate(source_rows, start=1):
            material_code = clean_text(row["leafComponentCode"])
            material_counts[material_code] += 1
            rows_out.append(
                {
                    "bom_source_id": source_id,
                    "bom_source_kind": "technical_flatten",
                    "bom_code": root_code,
                    "product_family_code": root_code,
                    "bom_variant_id": variant_id,
                    "dm_row_no": row_idx,
                    "ordinal_key": f"T{row_idx}{root_code}",
                    "material_code": material_code,
                    "qty_per_unit": parse_float(row["quantity"]),
                    "unit": clean_text(row["unit"]),
                    "leaf_component_description": clean_text(row["leafComponentDescription"]),
                    "brand": clean_text(row["brand"]),
                    "sample_path": clean_text(row["samplePath"]),
                    "path_count": clean_text(row["pathCount"]),
                    "source_file_count": clean_text(row["sourceFileCount"]),
                    "source_files": clean_text(row["sourceFiles"]),
                }
            )

        code_summaries.append(
            {
                "code": root_code,
                "variant_id": variant_id,
                "row_count": len(source_rows),
                "unique_material_count": len(material_counts),
                "duplicate_leaf_material_count": sum(
                    1 for count in material_counts.values() if count > 1
                ),
                "input_csv": csv_path.name,
            }
        )

    return rows_out, code_summaries


def build_manifest(
    artifact_dir: Path,
    variant_csv_path: Path,
    code_summaries: list[dict[str, object]],
) -> dict[str, object]:
    flattened_manifest_path = artifact_dir / "flattened-crossworkbook" / "manifest.json"
    comparison_summary_path = artifact_dir / "flattened-crossworkbook-vs-dm" / "summary.json"
    flattened_manifest = load_json(flattened_manifest_path) if flattened_manifest_path.exists() else []
    comparison_summary = load_json(comparison_summary_path) if comparison_summary_path.exists() else []

    flattened_by_code = {clean_text(item["code"]): item for item in flattened_manifest}
    comparison_by_code = {clean_text(item["code"]): item for item in comparison_summary}

    codes: list[dict[str, object]] = []
    for item in code_summaries:
        code = clean_text(item["code"])
        combined = dict(item)
        combined["flattened_summary"] = flattened_by_code.get(code, {})
        combined["dm_comparison_summary"] = comparison_by_code.get(code, {})
        codes.append(combined)

    return {
        "bom_source_id": TECHNICAL_PRIORITY_BOM_SOURCE_ID,
        "bom_source_kind": "technical_flatten",
        "variant_csv": variant_csv_path.name,
        "flattened_input_dir": "flattened-crossworkbook/per-code",
        "reference_comparison_dir": "flattened-crossworkbook-vs-dm",
        "notes": [
            "Use this CSV as the working canonical BOM input for the 9 priority Growatt product codes.",
            "Treat ST06.0010800 as a valid imported black-box BTP input, not a missing parse bug.",
            "Do not use PV02.0229100 DM quantities as a clean oracle because DM contains a duplicated block for that code.",
            "Remaining gaps against DM are primarily substitute or alternative semantics, not basic workbook coverage.",
        ],
        "codes": codes,
    }


def main() -> None:
    args = parse_args()
    artifact_dir = args.artifact_dir or technical_priority_artifact_dir(args.case_dir)
    per_code_dir = artifact_dir / "flattened-crossworkbook" / "per-code"
    if not per_code_dir.exists():
        raise SystemExit(f"Missing technical BOM per-code directory: {per_code_dir}")

    rows_out, code_summaries = build_variant_rows(per_code_dir, source_id=args.source_id)
    if not rows_out:
        raise SystemExit(f"No technical BOM rows found in {per_code_dir}")

    variant_csv_path = artifact_dir / "technical-bom-variants.csv"
    manifest_path = artifact_dir / "technical-bom-variants-manifest.json"

    write_csv(
        variant_csv_path,
        rows_out,
        [
            "bom_source_id",
            "bom_source_kind",
            "bom_code",
            "product_family_code",
            "bom_variant_id",
            "dm_row_no",
            "ordinal_key",
            "material_code",
            "qty_per_unit",
            "unit",
            "leaf_component_description",
            "brand",
            "sample_path",
            "path_count",
            "source_file_count",
            "source_files",
        ],
    )
    manifest = build_manifest(artifact_dir, variant_csv_path, code_summaries)
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    print(f"Wrote {variant_csv_path}")
    print(f"Wrote {manifest_path}")
    print(f"Codes: {len(code_summaries)}")
    print(f"Rows: {len(rows_out)}")


if __name__ == "__main__":
    main()
