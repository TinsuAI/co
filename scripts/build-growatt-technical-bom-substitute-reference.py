#!/usr/bin/env python3

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from pathlib import Path

from growatt_bom_sources import DEFAULT_CASE_DIR, technical_priority_artifact_dir


SPLIT_CODES_PATTERN = re.compile(r"[;,，；\r\n]+")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case-dir", type=Path, default=DEFAULT_CASE_DIR)
    parser.add_argument("--artifact-dir", type=Path)
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


def split_material_codes(raw_text: str) -> tuple[str, ...]:
    seen: set[str] = set()
    items: list[str] = []
    for part in SPLIT_CODES_PATTERN.split(clean_text(raw_text)):
        token = clean_text(part)
        if not token or token in seen:
            continue
        seen.add(token)
        items.append(token)
    return tuple(items)


def collect_signal_fields(row: dict[str, object]) -> list[str]:
    signal_fields: list[str] = []
    if clean_text(row.get("替代项目组")):
        signal_fields.append("substitute_group")
    if clean_text(row.get("可替代物料")):
        signal_fields.append("explicit_substitute_materials")
    if clean_text(row.get("优先级")) not in {"", "0"}:
        signal_fields.append("priority")
    if clean_text(row.get("策略")):
        signal_fields.append("strategy")
    if clean_text(row.get("使用概率")) not in {"", "0"}:
        signal_fields.append("usage_probability")
    if clean_text(row.get("替代组合标识")):
        signal_fields.append("combo_marker")
    if clean_text(row.get("临时替代预留行号")):
        signal_fields.append("temp_row_reference")
    return signal_fields


def classify_signal(
    *,
    substitute_group: str,
    substitute_material_codes: tuple[str, ...],
    combo_marker: str,
    temp_row_reference: str,
    priority: str,
    strategy: str,
    usage_probability: str,
) -> str:
    has_explicit_materials = bool(substitute_material_codes)
    has_group = bool(substitute_group)
    has_combo = bool(combo_marker)
    has_temp_ref = bool(temp_row_reference)
    has_policy = priority not in {"", "0"} or bool(strategy) or usage_probability not in {"", "0"}

    if has_group and has_explicit_materials:
        return "group_with_explicit_materials"
    if has_explicit_materials:
        return "explicit_materials_only"
    if has_group and has_combo:
        return "group_and_combo_semantics"
    if has_group:
        return "group_semantics_only"
    if has_combo:
        return "combo_marker_only"
    if has_temp_ref:
        return "temp_row_reference_only"
    if has_policy:
        return "policy_only"
    return "other_substitute_signal"


def build_reference_group_key(
    *,
    root_product_code: str,
    source_file_code: str,
    substitute_group: str,
    combo_marker: str,
    temp_row_reference: str,
    component_code: str,
    source_row_index: int,
) -> str:
    if substitute_group:
        return f"{root_product_code}|{source_file_code}|group|{substitute_group}"
    if combo_marker:
        return f"{root_product_code}|{source_file_code}|combo|{combo_marker}"
    if temp_row_reference:
        return f"{root_product_code}|{source_file_code}|temp|{temp_row_reference}"
    return f"{root_product_code}|{source_file_code}|row|{source_row_index:04d}|{component_code}"


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([{field: row.get(field) for field in fieldnames} for row in rows])


def collect_code_reference_rows(
    root_product_code: str,
    payload: dict[str, object],
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    row_signals: list[dict[str, object]] = []
    explicit_links: list[dict[str, object]] = []
    total_source_rows = 0

    canonical_source_chain = payload.get("canonicalSourceChain", [])
    for source_chain_index, chain in enumerate(canonical_source_chain, start=1):
        source_parent_code = clean_text(chain.get("parentCode"))
        source_file_code = clean_text(chain.get("normalizedFileCode"))
        source_kind = clean_text(chain.get("kind"))
        source_file_name = Path(clean_text(chain.get("filePath"))).name if clean_text(chain.get("filePath")) else ""

        for source_row_index, row in enumerate(chain.get("rows", []), start=1):
            total_source_rows += 1
            signal_fields = collect_signal_fields(row)
            if not signal_fields:
                continue

            substitute_group = clean_text(row.get("替代项目组"))
            substitute_materials_raw = clean_text(row.get("可替代物料"))
            substitute_material_codes = split_material_codes(substitute_materials_raw)
            substitute_priority = clean_text(row.get("优先级"))
            substitute_strategy = clean_text(row.get("策略"))
            usage_probability = clean_text(row.get("使用概率"))
            same_level_group_marker = clean_text(row.get("同层同组替代标识"))
            combo_marker = clean_text(row.get("替代组合标识"))
            temp_row_reference = clean_text(row.get("临时替代预留行号"))
            component_code = clean_text(row.get("组件物料"))
            signal_class = classify_signal(
                substitute_group=substitute_group,
                substitute_material_codes=substitute_material_codes,
                combo_marker=combo_marker,
                temp_row_reference=temp_row_reference,
                priority=substitute_priority,
                strategy=substitute_strategy,
                usage_probability=usage_probability,
            )
            reference_group_key = build_reference_group_key(
                root_product_code=root_product_code,
                source_file_code=source_file_code,
                substitute_group=substitute_group,
                combo_marker=combo_marker,
                temp_row_reference=temp_row_reference,
                component_code=component_code,
                source_row_index=source_row_index,
            )

            row_signal = {
                "root_product_code": root_product_code,
                "source_chain_index": source_chain_index,
                "source_row_index": source_row_index,
                "source_parent_code": source_parent_code,
                "source_file_code": source_file_code,
                "source_file_name": source_file_name,
                "source_kind": source_kind,
                "component_code": component_code,
                "component_description": clean_text(row.get("组件物料描述")),
                "component_brand": clean_text(row.get("组件物料品牌")),
                "unit": clean_text(row.get("单位")),
                "standard_qty": parse_float(row.get("标准用量")),
                "position_no": clean_text(row.get("位置号")),
                "substitute_group": substitute_group,
                "substitute_materials_raw": substitute_materials_raw,
                "substitute_material_count": len(substitute_material_codes),
                "substitute_material_codes": ";".join(substitute_material_codes),
                "substitute_priority": substitute_priority,
                "substitute_strategy": substitute_strategy,
                "usage_probability": usage_probability,
                "same_level_group_marker": same_level_group_marker,
                "substitute_combo_marker": combo_marker,
                "temp_substitute_row_no": temp_row_reference,
                "signal_fields": ";".join(signal_fields),
                "signal_class": signal_class,
                "reference_group_key": reference_group_key,
            }
            row_signals.append(row_signal)

            for substitute_material_code in substitute_material_codes:
                explicit_links.append(
                    {
                        "root_product_code": root_product_code,
                        "source_parent_code": source_parent_code,
                        "source_file_code": source_file_code,
                        "source_file_name": source_file_name,
                        "source_kind": source_kind,
                        "component_code": component_code,
                        "component_description": clean_text(row.get("组件物料描述")),
                        "substitute_group": substitute_group,
                        "substitute_priority": substitute_priority,
                        "substitute_strategy": substitute_strategy,
                        "usage_probability": usage_probability,
                        "same_level_group_marker": same_level_group_marker,
                        "substitute_combo_marker": combo_marker,
                        "temp_substitute_row_no": temp_row_reference,
                        "substitute_material_code": substitute_material_code,
                        "reference_group_key": reference_group_key,
                    }
                )

    summary = {
        "root_product_code": root_product_code,
        "source_chain_count": len(canonical_source_chain),
        "total_source_row_count": total_source_rows,
        "substitute_signal_row_count": len(row_signals),
        "rows_with_explicit_substitute_materials": sum(
            1 for row in row_signals if row["substitute_material_count"] > 0
        ),
        "explicit_substitute_link_count": len(explicit_links),
        "substitute_group_count": len({row["substitute_group"] for row in row_signals if row["substitute_group"]}),
        "reference_group_count": len({row["reference_group_key"] for row in row_signals}),
        "combo_marker_count": sum(1 for row in row_signals if row["substitute_combo_marker"]),
        "temp_row_reference_count": sum(1 for row in row_signals if row["temp_substitute_row_no"]),
    }
    return row_signals, explicit_links, summary


def build_reference_groups(row_signals: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped_rows: dict[str, dict[str, object]] = {}
    group_sets: dict[str, dict[str, set[str]]] = defaultdict(
        lambda: {
            "member_component_codes": set(),
            "explicit_substitute_material_codes": set(),
            "priority_values": set(),
            "strategy_values": set(),
            "usage_probability_values": set(),
            "same_level_group_markers": set(),
            "substitute_combo_markers": set(),
            "temp_substitute_row_refs": set(),
            "signal_classes": set(),
            "source_positions": set(),
        }
    )

    for row in row_signals:
        key = str(row["reference_group_key"])
        if key not in grouped_rows:
            grouped_rows[key] = {
                "reference_group_key": key,
                "root_product_code": row["root_product_code"],
                "source_parent_code": row["source_parent_code"],
                "source_file_code": row["source_file_code"],
                "source_file_name": row["source_file_name"],
                "source_kind": row["source_kind"],
                "substitute_group": row["substitute_group"],
                "member_row_count": 0,
            }
        grouped_rows[key]["member_row_count"] += 1

        group_sets[key]["member_component_codes"].add(str(row["component_code"]))
        group_sets[key]["signal_classes"].add(str(row["signal_class"]))
        if row["position_no"]:
            group_sets[key]["source_positions"].add(str(row["position_no"]))
        if row["substitute_priority"]:
            group_sets[key]["priority_values"].add(str(row["substitute_priority"]))
        if row["substitute_strategy"]:
            group_sets[key]["strategy_values"].add(str(row["substitute_strategy"]))
        if row["usage_probability"]:
            group_sets[key]["usage_probability_values"].add(str(row["usage_probability"]))
        if row["same_level_group_marker"]:
            group_sets[key]["same_level_group_markers"].add(str(row["same_level_group_marker"]))
        if row["substitute_combo_marker"]:
            group_sets[key]["substitute_combo_markers"].add(str(row["substitute_combo_marker"]))
        if row["temp_substitute_row_no"]:
            group_sets[key]["temp_substitute_row_refs"].add(str(row["temp_substitute_row_no"]))
        for code in split_material_codes(str(row["substitute_materials_raw"])):
            group_sets[key]["explicit_substitute_material_codes"].add(code)

    out_rows: list[dict[str, object]] = []
    for key, row in sorted(grouped_rows.items()):
        sets = group_sets[key]
        row["member_component_count"] = len(sets["member_component_codes"])
        row["member_component_codes"] = ";".join(sorted(sets["member_component_codes"]))
        row["explicit_substitute_material_count"] = len(sets["explicit_substitute_material_codes"])
        row["explicit_substitute_material_codes"] = ";".join(sorted(sets["explicit_substitute_material_codes"]))
        row["priority_values"] = ";".join(sorted(sets["priority_values"]))
        row["strategy_values"] = ";".join(sorted(sets["strategy_values"]))
        row["usage_probability_values"] = ";".join(sorted(sets["usage_probability_values"]))
        row["same_level_group_markers"] = ";".join(sorted(sets["same_level_group_markers"]))
        row["substitute_combo_markers"] = ";".join(sorted(sets["substitute_combo_markers"]))
        row["temp_substitute_row_refs"] = ";".join(sorted(sets["temp_substitute_row_refs"]))
        row["signal_classes"] = ";".join(sorted(sets["signal_classes"]))
        row["source_positions"] = ";".join(sorted(sets["source_positions"]))
        out_rows.append(row)
    return out_rows


def build_manifest(
    *,
    output_dir: Path,
    summary_rows: list[dict[str, object]],
    row_signals: list[dict[str, object]],
    explicit_links: list[dict[str, object]],
    reference_groups: list[dict[str, object]],
) -> dict[str, object]:
    return {
        "artifact_kind": "growatt-technical-bom-substitute-reference",
        "root_product_count": len(summary_rows),
        "substitute_signal_row_count": len(row_signals),
        "explicit_substitute_link_count": len(explicit_links),
        "reference_group_count": len(reference_groups),
        "output_dir": output_dir.name,
        "files": {
            "summary_csv": "summary.csv",
            "signal_rows_csv": "signal-rows.csv",
            "explicit_links_csv": "explicit-substitute-links.csv",
            "reference_groups_csv": "reference-groups.csv",
        },
        "notes": [
            "This is a reference layer extracted from the technical BOM source rows, not a substitute-resolved final BOM.",
            "Rows are included when they carry substitute-group, explicit substitute material, policy, combo, or temp-row-reference signals.",
            "Same-level markers are preserved as attributes but do not by themselves trigger row inclusion.",
        ],
        "per_product_summary": summary_rows,
    }


def main() -> None:
    args = parse_args()
    artifact_dir = args.artifact_dir or technical_priority_artifact_dir(args.case_dir)
    per_code_dir = artifact_dir / "flattened-crossworkbook" / "per-code"
    if not per_code_dir.exists():
        raise SystemExit(f"Missing technical BOM per-code directory: {per_code_dir}")

    output_dir = artifact_dir / "substitute-reference"
    output_dir.mkdir(parents=True, exist_ok=True)

    summary_rows: list[dict[str, object]] = []
    row_signals: list[dict[str, object]] = []
    explicit_links: list[dict[str, object]] = []

    for json_path in sorted(per_code_dir.glob("*.json")):
        payload = json.loads(json_path.read_text(encoding="utf-8"))
        code = clean_text(payload.get("code")) or json_path.stem
        code_row_signals, code_explicit_links, code_summary = collect_code_reference_rows(code, payload)
        row_signals.extend(code_row_signals)
        explicit_links.extend(code_explicit_links)
        summary_rows.append(code_summary)

    reference_groups = build_reference_groups(row_signals)

    write_csv(
        output_dir / "summary.csv",
        summary_rows,
        [
            "root_product_code",
            "source_chain_count",
            "total_source_row_count",
            "substitute_signal_row_count",
            "rows_with_explicit_substitute_materials",
            "explicit_substitute_link_count",
            "substitute_group_count",
            "reference_group_count",
            "combo_marker_count",
            "temp_row_reference_count",
        ],
    )
    write_csv(
        output_dir / "signal-rows.csv",
        row_signals,
        [
            "root_product_code",
            "source_chain_index",
            "source_row_index",
            "source_parent_code",
            "source_file_code",
            "source_file_name",
            "source_kind",
            "component_code",
            "component_description",
            "component_brand",
            "unit",
            "standard_qty",
            "position_no",
            "substitute_group",
            "substitute_materials_raw",
            "substitute_material_count",
            "substitute_material_codes",
            "substitute_priority",
            "substitute_strategy",
            "usage_probability",
            "same_level_group_marker",
            "substitute_combo_marker",
            "temp_substitute_row_no",
            "signal_fields",
            "signal_class",
            "reference_group_key",
        ],
    )
    write_csv(
        output_dir / "explicit-substitute-links.csv",
        explicit_links,
        [
            "root_product_code",
            "source_parent_code",
            "source_file_code",
            "source_file_name",
            "source_kind",
            "component_code",
            "component_description",
            "substitute_group",
            "substitute_priority",
            "substitute_strategy",
            "usage_probability",
            "same_level_group_marker",
            "substitute_combo_marker",
            "temp_substitute_row_no",
            "substitute_material_code",
            "reference_group_key",
        ],
    )
    write_csv(
        output_dir / "reference-groups.csv",
        reference_groups,
        [
            "reference_group_key",
            "root_product_code",
            "source_parent_code",
            "source_file_code",
            "source_file_name",
            "source_kind",
            "substitute_group",
            "member_row_count",
            "member_component_count",
            "member_component_codes",
            "explicit_substitute_material_count",
            "explicit_substitute_material_codes",
            "priority_values",
            "strategy_values",
            "usage_probability_values",
            "same_level_group_markers",
            "substitute_combo_markers",
            "temp_substitute_row_refs",
            "signal_classes",
            "source_positions",
        ],
    )

    manifest = build_manifest(
        output_dir=output_dir,
        summary_rows=summary_rows,
        row_signals=row_signals,
        explicit_links=explicit_links,
        reference_groups=reference_groups,
    )
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    print(f"Wrote {output_dir / 'summary.csv'}")
    print(f"Wrote {output_dir / 'signal-rows.csv'}")
    print(f"Wrote {output_dir / 'explicit-substitute-links.csv'}")
    print(f"Wrote {output_dir / 'reference-groups.csv'}")
    print(f"Wrote {output_dir / 'manifest.json'}")
    print(f"Signal rows: {len(row_signals)}")
    print(f"Explicit links: {len(explicit_links)}")
    print(f"Reference groups: {len(reference_groups)}")


if __name__ == "__main__":
    main()
