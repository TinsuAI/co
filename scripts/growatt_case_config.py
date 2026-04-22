#!/usr/bin/env python3

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path


DEFAULT_CONFIG_PATH = Path("config/growatt-rvc-20260421.json")
LEGACY_IMPORT_LEAD_DAYS = 2
LEGACY_MAX_IMPORT_AGE_DAYS = 365
LEGACY_VALUATION_MODE = "workbook_avg"
VALID_VALUATION_MODES = {"workbook_avg", "weighted"}


@dataclass(frozen=True)
class ShipmentPolicy:
    shipment_id: str
    policy_version: str
    import_lead_days: int
    max_import_age_days: int
    valuation_mode: str
    config_path: str | None
    policy_source: str

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def _load_case_config(case_dir: Path, config_path_override: Path | None) -> tuple[Path, dict[str, object] | None]:
    candidate_paths: list[Path] = []
    if config_path_override is not None:
        candidate_paths.append(config_path_override)
    candidate_paths.append(DEFAULT_CONFIG_PATH)
    candidate_paths.append(case_dir / "case-config.json")

    seen: set[Path] = set()
    for config_path in candidate_paths:
        resolved = config_path
        if resolved in seen:
            continue
        seen.add(resolved)
        if resolved.exists():
            return resolved, json.loads(resolved.read_text(encoding="utf-8"))

    return candidate_paths[0], None


def _coerce_int(value: object, field_name: str) -> int:
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise SystemExit(f"Invalid {field_name}: {value!r}") from exc


def _coerce_valuation_mode(value: object) -> str:
    text = str(value).strip()
    if text not in VALID_VALUATION_MODES:
        allowed = ", ".join(sorted(VALID_VALUATION_MODES))
        raise SystemExit(f"Invalid valuation_mode {text!r}; expected one of: {allowed}")
    return text


def resolve_shipment_policy(
    case_dir: Path,
    shipment_id: str,
    *,
    config_path_override: Path | None = None,
    policy_version_override: str | None = None,
    import_lead_days_override: int | None = None,
    max_import_age_days_override: int | None = None,
    valuation_mode_override: str | None = None,
) -> ShipmentPolicy:
    config_path, payload = _load_case_config(case_dir, config_path_override)
    if payload is None:
        return ShipmentPolicy(
            shipment_id=shipment_id,
            policy_version="legacy-default",
            import_lead_days=(
                LEGACY_IMPORT_LEAD_DAYS
                if import_lead_days_override is None
                else int(import_lead_days_override)
            ),
            max_import_age_days=(
                LEGACY_MAX_IMPORT_AGE_DAYS
                if max_import_age_days_override is None
                else int(max_import_age_days_override)
            ),
            valuation_mode=(
                LEGACY_VALUATION_MODE
                if valuation_mode_override is None
                else _coerce_valuation_mode(valuation_mode_override)
            ),
            config_path=None,
            policy_source="legacy-defaults",
        )

    policy_versions = payload.get("policy_versions", {})
    shipments = payload.get("shipments", {})
    shipment_payload = shipments.get(shipment_id, {})
    configured_policy_version = (
        policy_version_override
        or shipment_payload.get("policy_version")
        or payload.get("active_policy_version")
    )
    if not configured_policy_version:
        raise SystemExit(
            f"Missing policy version for shipment {shipment_id} in {config_path}"
        )

    if configured_policy_version not in policy_versions:
        raise SystemExit(
            f"Unknown policy version {configured_policy_version!r} for shipment {shipment_id} in {config_path}"
        )

    policy_payload = policy_versions[configured_policy_version]
    merged = dict(policy_payload)
    for key in ("import_lead_days", "max_import_age_days", "valuation_mode"):
        if key in shipment_payload:
            merged[key] = shipment_payload[key]

    import_lead_days = (
        _coerce_int(merged.get("import_lead_days", LEGACY_IMPORT_LEAD_DAYS), "import_lead_days")
        if import_lead_days_override is None
        else int(import_lead_days_override)
    )
    max_import_age_days = (
        _coerce_int(merged.get("max_import_age_days", LEGACY_MAX_IMPORT_AGE_DAYS), "max_import_age_days")
        if max_import_age_days_override is None
        else int(max_import_age_days_override)
    )
    valuation_mode = (
        _coerce_valuation_mode(merged.get("valuation_mode", LEGACY_VALUATION_MODE))
        if valuation_mode_override is None
        else _coerce_valuation_mode(valuation_mode_override)
    )

    return ShipmentPolicy(
        shipment_id=shipment_id,
        policy_version=configured_policy_version,
        import_lead_days=import_lead_days,
        max_import_age_days=max_import_age_days,
        valuation_mode=valuation_mode,
        config_path=str(config_path),
        policy_source="config-file",
    )
