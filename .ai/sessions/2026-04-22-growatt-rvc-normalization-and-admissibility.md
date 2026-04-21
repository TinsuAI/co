# What Was Done
- Created a Growatt case execution plan in [docs/growatt-rvc-case-execution-plan.md](/home/vp/workspace/client/barry-CO/docs/growatt-rvc-case-execution-plan.md) and indexed it from [docs/README.md](/home/vp/workspace/client/barry-CO/docs/README.md).
- Added [scripts/growatt-rvc-baseline.py](/home/vp/workspace/client/barry-CO/scripts/growatt-rvc-baseline.py) to:
  - parse `DM` into contiguous BOM blocks
  - assign stable `bom_variant_id` identities
  - enumerate shipment scenarios
  - enforce the `import <= export - 2 days` rule
  - compute baseline allocation and product-level RVC
- Added [scripts/render-growatt-case-view.py](/home/vp/workspace/client/barry-CO/scripts/render-growatt-case-view.py) to render summary markdown and CSV views for case results.
- Added [scripts/build-growatt-case-data.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-case-data.py) to normalize workbook and BCCT data into:
  - `imports-normalized.csv`
  - `exports-normalized.csv`
  - `co-stock-tracking-updated.csv`
  - `dm-variants.csv`
  - `growatt-case-data.xlsx`
  - `manifest.json`
- Refactored normalization so extracted import/export codes are not treated as automatically final:
  - candidate code remains in `lookup_material_code` / `internal_code_for_dm`
  - `confirmed_lookup_code` is separate
  - `final_lookup_key` is only filled when the row is confirmed
  - `mapping_status` and `lookup_confidence` are now explicit
- Corrected customs-line identity from `declaration + item + code` to `declaration + item`.
- Added reconcile logic so BCCT import rows can inherit a confirmed code from `NK2` only when the same customs line has one unambiguous confirmed code.
- Rebuilt the shared normalized dataset after the reconciliation fix.
- Ran critic/explorer reviews on the remaining import ambiguity and on the proposed next-step logic.

# Decisions Made
- Treat workbook `DM` blocks as the real BOM identity layer; do not collapse by export product code alone.
- Treat extracted import codes as candidates, not confirmed BOM codes.
- Use `declaration_no + declaration_item_no` as customs-line identity; code representation differences between BCCT and `NK2` are attributes, not identity.
- Allow BCCT import rows to inherit confirmation from `NK2` only under a strict same-customs-line, single-confirmed-code guard.
- Do not implement shipment-wide `confirmed_for_b282`. The next admissibility logic must be variant-scoped and preserve ambiguity.
- Keep shared normalized data immutable; shipment-specific interpretation should live in separate derived artifacts.

# What Didn't Work
- Early merged stock logic used `declaration + item + declared_code` as the key. This created duplicate rows for the same customs line whenever BCCT and `NK2` used different code representations such as `TEM` vs `940.0012600`.
- The first normalized pass treated `lookup_material_code` / `final_lookup_key` too optimistically. That overstated how many import rows were truly confirmed.
- The baseline runner currently still reads raw `NK2` instead of the normalized shared stock ledger, so current `B282` scenario outputs should not be treated as final evidence.
- The idea of “confirming imports for `B282`” at shipment scope was rejected by critic review because it bakes scenario decisions into the evidence too early.

# Open Items
- Build `B282` variant-scoped admissibility artifacts without changing shared normalized data.
- Quantify, for each `bom_variant_id`, which import rows are:
  - already confirmed
  - admissible but ambiguous
  - unusable
- Rewire `growatt-rvc-baseline.py` to consume the normalized shared dataset and the new admissibility layer instead of raw workbook `NK2`.
- Re-run `B282` baseline only after the admissibility layer exists.
- Defer `C171` until `B282` has a defensible scenario pipeline.
- Keep the remaining unresolved import groups visible:
  - `1738 candidate_exact_dm_match`
  - `228 candidate_not_in_dm`
  - `175 no_lookup_code`
