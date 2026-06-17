from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


DEFAULT_CASE_DIR = Path("data/cases/growatt-rvc-20260421")
DM_BOM_SOURCE_ID = "dm"
TECHNICAL_PRIORITY_BOM_SOURCE_ID = "technical-priority-2026-04-23"
TECHNICAL_PRIORITY_VARIANT_SUFFIX = "technical_flatten_20260423"


@dataclass(frozen=True)
class GrowattBomSource:
    source_id: str
    label: str
    variant_csv_path: Path
    comparison_summary_path: Path | None = None
    manifest_path: Path | None = None


def bom_source_choices() -> tuple[str, ...]:
    return (DM_BOM_SOURCE_ID, TECHNICAL_PRIORITY_BOM_SOURCE_ID)


def technical_priority_artifact_dir(case_dir: Path) -> Path:
    return (
        case_dir
        / "shared"
        / "normalized"
        / "growatt-priority-technical-bom-2026-04-23"
    )


def resolve_bom_source(case_dir: Path, source_id: str) -> GrowattBomSource:
    shared_normalized_dir = case_dir / "shared" / "normalized"
    if source_id == DM_BOM_SOURCE_ID:
        variant_csv_path = shared_normalized_dir / "dm-variants.csv"
        comparison_summary_path = None
        manifest_path = shared_normalized_dir / "manifest.json"
        label = "DM variants"
    elif source_id == TECHNICAL_PRIORITY_BOM_SOURCE_ID:
        artifact_dir = technical_priority_artifact_dir(case_dir)
        variant_csv_path = artifact_dir / "technical-bom-variants.csv"
        comparison_summary_path = artifact_dir / "flattened-crossworkbook-vs-dm" / "summary.json"
        manifest_path = artifact_dir / "technical-bom-variants-manifest.json"
        label = "Priority technical BOM 2026-04-23"
    else:
        raise ValueError(f"Unsupported Growatt BOM source: {source_id}")

    if not variant_csv_path.exists():
        raise FileNotFoundError(
            f"Missing BOM source artifact for {source_id}: {variant_csv_path}"
        )
    if comparison_summary_path is not None and not comparison_summary_path.exists():
        comparison_summary_path = None
    if manifest_path is not None and not manifest_path.exists():
        manifest_path = None

    return GrowattBomSource(
        source_id=source_id,
        label=label,
        variant_csv_path=variant_csv_path,
        comparison_summary_path=comparison_summary_path,
        manifest_path=manifest_path,
    )


def relative_to_case(case_dir: Path, path: Path | None) -> str:
    if path is None:
        return ""
    try:
        return str(path.relative_to(case_dir))
    except ValueError:
        return str(path)
