# Project Status

## Current State
- Active work remains on branch `case/growatt-rvc-20260421`, but the `B282` case pipeline is now materially more trustworthy than the previous workbook-driven baseline.
- Shared normalized case data under `data/cases/growatt-rvc-20260421/shared/normalized/` now carries USD-normalized valuation fields with explicit audit basis, plus the prior candidate-vs-confirmed code-state split.
- `B282` now has shipment-specific admissibility artifacts in `data/cases/growatt-rvc-20260421/b282/normalized/`, and the baseline runner now consumes shared normalized CSVs plus admissibility instead of raw `NK2`.
- Import-age policy is now explicit and configurable:
  - latest import date allowed: `export_date - 2 days`
  - oldest import date allowed by default: `export_date - 365 days`
- Current `B282` output is still a failure, but it is now based on normalized inputs, variant-scoped admissibility, and explicit stock-age limits:
  - best scenario `total_unmet_qty = 10807.002139`
  - best scenario `min_margin = -7.132388635440243`
- `BCCT` import evidence for this Growatt case is only `2026`; older `2023-2024` dates in the merged stock ledger come from `NK2` carry-over stock, not the current BCCT import report.

## Recent Changes
- Added [scripts/build-growatt-shipment-admissibility.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-shipment-admissibility.py) to derive variant-scoped `B282` admissibility, coverage, and ambiguity artifacts from shared normalized data.
- Rewrote [scripts/growatt-rvc-baseline.py](/home/vp/workspace/client/barry-CO/scripts/growatt-rvc-baseline.py) to:
  - read normalized exports / stock instead of raw workbook sheets
  - gate candidate rows through shipment-specific admissibility artifacts
  - apply configurable import lead-time and max-age rules
- Extended [scripts/build-growatt-case-data.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-case-data.py) so import-side valuation is normalized to USD with an explicit `price_normalization_basis`.
- Added [scripts/download-customs-fx-rates.py](/home/vp/workspace/client/barry-CO/scripts/download-customs-fx-rates.py) to refresh customs FX data directly from `customs.gov.vn`.
- Rebuilt shared normalized data, rebuilt `B282` admissibility and baseline outputs, and added [docs/growatt-valuation-and-stock-insights.md](/home/vp/workspace/client/barry-CO/docs/growatt-valuation-and-stock-insights.md) to preserve the valuation and stock-aging findings.

## Next Steps
- Review the `B282` failure outputs by material group and separate true stock shortages from date-window exclusions.
- Decide whether the default `365`-day inventory window is the final business rule or only the current operational default.
- If `B282` logic is accepted, repeat the same pipeline for `C171`.

## Notes for Next AI Session
- User wants the work grounded in operational data, not blind trust in the macro workbook. Workbook columns and formulas are often manually overwritten and should be treated cautiously.
- User explicitly wanted the stock-age rule configurable, with the default set to “within the last year, and at least 2 days before export”.
- Preserve the valuation distinction:
  - customs-workbook USD conversion is preferred
  - source exchange rate is fallback only
  - raw unit price is last fallback
- Current customs FX workbook lives at `data/reference/DS_ty_gia_ngoai_te.xlsx` and is local-only; refresh it with `python3 scripts/download-customs-fx-rates.py`.
- `B282` currently has `756` ambiguity rows, `27963` variant-admissibility rows, and `3884` material-coverage rows.
- There is still a large amount of old `NK2` stock on the ledger, but under the default one-year window the current best `B282` scenario uses none of the `2023-2024` rows.
