# Project Status

## Current State
- Active work is still on Growatt RVC case handling, with `B282` as the current decision case and `C171` intentionally deferred.
- Growatt shipment policy is now tracked in [config/growatt-rvc-20260421.json](/home/vp/workspace/client/barry-CO/config/growatt-rvc-20260421.json) instead of ignored `data/`, and the active policy is `2026-04-22-no-max-age`:
  - import must be at least `2` days before export
  - `max_import_age_days = 0` means no one-year cap
  - valuation mode remains `workbook_avg`
- `B282` has been rebuilt under that policy through admissibility, baseline, rendering, shortage triage, true-shortage reporting, and full exact sequential search.
- Current `B282` result still fails under all existing BOM scenarios:
  - best baseline scenario `total_unmet_qty = 9174.0021`
  - bucket totals: `true-shortage = 7970.0757`, `variant-choice-driven = 926.0380`, `date-blocked = 277.8885`
  - full exact sequential search over `18` BOM scenarios x `720` product orders found `0` product-level C/O passes
- The current best baseline BOM scenario is:
  - `PV01.0117300 -> PV01.0117300__block1`
  - `PV01.0117400 -> PV01.0117400__block1`
  - `PV02.0228801 -> PV02.0228801__block1`
  - `PV02.0228901 -> PV02.0228901-NEW__block1`
  - `PV02.0229000 -> PV02.0229000-NEW__block1`
  - `PV02.0229100 -> PV02.0229100-NEW__block1`
- The best current sequential order is:
  - `PV02.0229100 -> PV02.0229000 -> PV02.0228901 -> PV02.0228801 -> PV01.0117400 -> PV01.0117300`
  - but it still yields `0` products that both have enough stock and meet `RVC >= 35`

## Recent Changes
- Added tracked policy/config support for Growatt case runs:
  - [config/growatt-rvc-20260421.json](/home/vp/workspace/client/barry-CO/config/growatt-rvc-20260421.json)
  - [scripts/growatt_case_config.py](/home/vp/workspace/client/barry-CO/scripts/growatt_case_config.py)
  - updated [scripts/build-growatt-shipment-admissibility.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-shipment-admissibility.py), [scripts/growatt-rvc-baseline.py](/home/vp/workspace/client/barry-CO/scripts/growatt-rvc-baseline.py), and [scripts/render-growatt-case-view.py](/home/vp/workspace/client/barry-CO/scripts/render-growatt-case-view.py)
- Added shortage and true-shortage reporting:
  - [scripts/build-growatt-shortage-triage.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-shortage-triage.py)
  - [scripts/build-growatt-true-shortage-report.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-true-shortage-report.py)
  - [docs/growatt-b282-true-shortage-report.md](/home/vp/workspace/client/barry-CO/docs/growatt-b282-true-shortage-report.md)
- Added and fixed full-search sequential reporting:
  - [scripts/build-growatt-sequential-priority-report.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-sequential-priority-report.py)
  - [docs/growatt-b282-sequential-priority-report.md](/home/vp/workspace/client/barry-CO/docs/growatt-b282-sequential-priority-report.md)
  - fixed prior issues where the report only searched the top `5` scenarios and where stock sufficiency was mislabeled as product pass
  - report is now in Vietnamese and separates `Đủ Stock` from `Đạt C/O`
- Documented generic replacement rules and sprint framing in [docs/growatt-replacement-rules-and-sprints.md](/home/vp/workspace/client/barry-CO/docs/growatt-replacement-rules-and-sprints.md), then linked it from [docs/README.md](/home/vp/workspace/client/barry-CO/docs/README.md)

## Next Steps
- Build a replaceability-aware report for the main true-shortage materials in the locked `B282` baseline scenario, instead of ranking only by unmet quantity.
- Use that report to decide which material shortages are suitable for Sprint 1 same-customs-code replacement work.
- Keep `C171` untouched until the replacement strategy for `B282` is clearer.

## Blockers
- HQ/ERP mapping quality is improved but not fully closed for replacement decisions. There are still large candidate and ambiguity populations in the normalized stock/admissibility outputs, so replacement decisions on ambiguous materials remain risky without targeted review.

## Notes for Next AI Session
- User wants optimization aimed at maximizing the number of product-level passes, not only minimizing total unmet quantity.
- User confirmed that the one-year stock window is not a real business restriction for this case; materials from `2023+` can still be used if they satisfy the lead-time rule.
- Do not treat workbook parity as the main truth source. The current normalized/admissibility pipeline is the operational baseline; workbook comparison is only supporting context.
- The current sequential result is a feasibility proof, not an operational rescue plan: no order works with the existing BOM variants.
- Replacement rules captured this session are intended as a general project pattern, not a Growatt-only workaround.
