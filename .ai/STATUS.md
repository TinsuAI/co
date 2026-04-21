# Project Status

## Current State
- Active work is on branch `case/growatt-rvc-20260421` for the Growatt RVC case, focused on unfinished shipments `GIN01426B282` and `GIN01426C171`.
- A case plan now exists in [docs/growatt-rvc-case-execution-plan.md](/home/vp/workspace/client/barry-CO/docs/growatt-rvc-case-execution-plan.md), and the repo has case scripts for baseline evaluation, normalized data building, and result rendering.
- Shared normalized case data is now built under `data/cases/growatt-rvc-20260421/shared/normalized/` from workbook `DM`/`NK2` plus BCCT import/export reports.
- Import/customs-line identity has been corrected to `declaration_no + declaration_item_no`; the merged stock ledger no longer duplicates the same customs line when BCCT and `NK2` use different code representations.
- Code extraction is now stateful:
  - `lookup_material_code` / `internal_code_for_dm` are candidate codes only
  - `confirmed_lookup_code` is filled only when confirmation is defensible
  - `final_lookup_key` is filled only for confirmed rows
- Current import normalization status:
  - `3045` BCCT import rows total
  - `904` confirmed rows
  - `792` confirmed via same customs line from `NK2`
  - `1738` still `candidate_exact_dm_match`
  - `228` still `candidate_not_in_dm`
  - `175` still `no_lookup_code`
- Current merged stock ledger status:
  - `33065` rows in `co-stock-tracking-updated.csv`
  - `0` duplicate customs-line keys
- The old baseline runner is still not trustworthy for decision-making yet because it still reads raw workbook `NK2` directly instead of the normalized shared stock ledger and still lacks the new admissibility layer.

## Recent Changes
- Added Growatt case plan doc and linked it from `docs/README.md`.
- Added [scripts/growatt-rvc-baseline.py](/home/vp/workspace/client/barry-CO/scripts/growatt-rvc-baseline.py) to enumerate `DM` BOM blocks as `bom_variant_id` scenarios and run shipment-level allocation/RVC baselines.
- Added [scripts/render-growatt-case-view.py](/home/vp/workspace/client/barry-CO/scripts/render-growatt-case-view.py) to produce summary markdown/CSV views for case results.
- Added [scripts/build-growatt-case-data.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-case-data.py) to build shared normalized data from workbook + BCCT reports.
- Refined normalization so candidate-vs-confirmed code states are explicit and import rows can inherit confirmation from `NK2` only when the same customs line has one unambiguous confirmed code.
- Critic reviews established that the next logic layer must be `variant-scoped admissibility`, not shipment-wide confirmation.

## Next Steps
- Build a `B282`-specific admissibility layer without mutating shared normalized data.
- Produce artifacts at least equivalent to:
  - `b282-variant-admissibility.csv`
  - `b282-material-coverage.csv`
  - `b282-ambiguity-report.csv`
- Make admissibility state variant-scoped (`admissible_for_variant` / ambiguous), not shipment-scoped confirmation.
- Repoint the baseline runner to consume shared normalized data instead of raw `NK2`.
- Ensure allocation uses only `confirmed_*` or explicit variant-admissible rows, never raw `candidate_*`.
- Only after `B282` scenario logic is defensible should work continue to `C171`.

## Blockers
- The remaining `1966` unresolved import rows cannot be auto-confirmed further by same-customs-line matching to `NK2`; they need variant-scoped admissibility logic or manual review.
- Baseline/RVC outputs already generated for `B282` are provisional/stale because they were computed before the normalized stock/reconciliation corrections were integrated into the runner.

## Notes for Next AI Session
- User wants the work grounded in operational data, not blind trust in the macro workbook. Workbook columns and formulas are often manually overwritten and should be treated cautiously.
- `DM` should be modeled as layered identity:
  - `export_product_code`
  - `bom_code`
  - `bom_variant_id`
  - optional `product_family_code`
- Growatt names often contain both customs-facing code and internal ERP code. Keep raw `declared_code`, extracted candidate code, and confirmed code separate.
- User explicitly wants sequence, not parallel shipment processing: `B282` first, then `C171`.
- Import eligibility rule is currently `import_date <= export_date - 2 days`.
- Critic review result to preserve: do not “confirm for shipment”; preserve ambiguity as data and adjudicate against explicit `bom_variant_id` scope.
