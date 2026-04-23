# What Was Done
- Re-investigated the extracted Growatt VBA to determine how the workbook handles repeated product codes and multiple BOM-like blocks in sheet `DM`.
- Confirmed that `DM.sumifDM` only calculates totals and shipment/run key `K6`; it does not choose a BOM version.
- Confirmed that `HideCopy.update` collapses `DM` product codes into unique `Xuat` product entries and sums quantities by product code.
- Confirmed that `Run1` / `Run3` use `CountIf` plus `VLookup(stt + product_code)` against `DM`, which is positional row-order lookup rather than explicit BOM-block selection.
- Confirmed that `RunUpgrade` rebuilds the same pattern with dictionaries, which means duplicate helper keys can be overwritten by later rows instead of treated as separate BOM variants.
- Checked the nine active Growatt target product codes against normalized exports and `dm-variants.csv`, and confirmed that all are present in current artifacts.
- Added [docs/growatt-dm-workbook-behavior.md](/home/vp/workspace/client/barry-CO/docs/growatt-dm-workbook-behavior.md) and linked it from [docs/README.md](/home/vp/workspace/client/barry-CO/docs/README.md).
- Recorded the current multi-worktree setup in project state:
  - `main` at `c41db7e`
  - `feature/bom-builder` at `d50ab92`
  - current Growatt case branch remains separate

# Decisions Made
- Treat the current workbook as operational evidence of behavior, not as a reliable domain model for BOM variants.
- Treat repeated exact-code blocks in `DM` as unresolved and unsafe for automatic interpretation.
- Keep the project-facing doc focused on behavior and design implications, not on asserting that the workbook is “correct”.
- Preserve the split between Growatt case work and BOM Builder work through separate worktrees instead of trying to reuse one dirty branch for both.

# What Didn't Work
- The workbook does not expose any explicit concept of contiguous BOM block, version selection, or exact-code duplicate block identity.
- Different macro families can resolve the same repeated-code `DM` data differently:
  - old macros can take the first matching rows
  - dictionary-based macros can overwrite toward the last matching rows
- This means workbook behavior cannot be relied on to prove which BOM version was intended when exact product codes repeat without suffix changes.

# Open Items
- Get original technical BOMs from staff/factory for the active `B282` and `C171` products.
- Decide whether future Growatt replacement runs should use:
  - rebuilt technical BOMs
  - or a curated merge of the two current workbook sources
- Continue BOM Builder implementation in `/home/vp/workspace/client/barry-CO-bom-builder` without waiting for the Growatt case to close.
- If future workbook investigation continues, compare whether any user workflow outside VBA manually cleans `DM` before runs; that may explain why the legacy workbook tolerated repeated exact product codes operationally.
