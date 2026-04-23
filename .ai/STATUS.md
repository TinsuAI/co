# Project Status

## Current State
- Active Growatt work is still in replacement execution for `B282`; `C171` is identified and its export product list is pinned, but it is still deferred as an optimization case.
- Current `B282` replacement results remain the latest operational baseline from `data/cases/growatt-rvc-20260421/b282/results/replacement-seed-summary.csv`:
  - `heuristic_best`: `4` product passes, `final_unmet_qty = 1318.1440`
  - `staff_latest_bom_invoice_order`: `4` product passes, `final_unmet_qty = 1356.1440`
  - `replaceability_cost_down_best`: `3` product passes
  - `replaceability_unmet_best`: `2` product passes
- Source-of-truth status is still split:
  - root workbook is newer for `XK` and `NK2`
  - `lo-da-lam` workbook still contains BOM edits absent from the root workbook
- The workbook `DM` logic has now been re-checked. It does not model explicit BOM variants; repeated exact-code BOM blocks are handled implicitly by row-order lookup or dictionary overwrite, depending on the macro path.
- The nine Growatt target product codes for the two active shipments are all present in current normalized Growatt artifacts, and exact `dm-variants.csv` coverage exists for all of them.
- Separate implementation worktrees now exist:
  - `main` at `/home/vp/workspace/client/barry-CO-main` on commit `c41db7e`
  - `feature/bom-builder` at `/home/vp/workspace/client/barry-CO-bom-builder` on commit `d50ab92`

## Recent Changes
- Added project-facing documentation for current workbook behavior in [docs/growatt-dm-workbook-behavior.md](/home/vp/workspace/client/barry-CO/docs/growatt-dm-workbook-behavior.md).
- Indexed that note from [docs/README.md](/home/vp/workspace/client/barry-CO/docs/README.md).
- Re-investigated extracted VBA around `DM`, `Run1`, `Run3`, `RunUpgrade`, `HideCopy`, and `LocmaSapxepLaydata` to confirm how repeated BOM-like blocks are actually resolved.
- Verified that the target codes
  - `PV01.0117300`, `PV01.0117400`, `PV02.0228801`, `PV02.0228901`, `PV02.0229000`, `PV02.0229100`
  - `PV00.0048400`, `PV00.0048500`, `PV01.0117600`
  are all present in the Growatt normalized exports / DM-derived artifacts.

## Next Steps
- Request factory/staff BOMs for all export products in `GIN01426B282` and `GIN01426C171`, plus a consolidated DM replacement sheet.
- Decide the replacement input BOM source before any further optimization:
  - rebuild from technical factory BOMs
  - or reconcile root vs `lo-da-lam` workbook edits into one curated BOM source
- Keep using root workbook `XK`/`NK2` as the stock/export baseline unless better evidence appears.
- Continue Growatt work on the current case branch and BOM Builder work on the separate `feature/bom-builder` worktree.

## Blockers
- DM/BOM source quality is still unresolved. Duplicate material lines inside BOM blocks are widespread, and the two available workbooks disagree on BOM content.
- The current workbook logic does not safely distinguish multiple exact-code BOM blocks for the same product.
- Ambiguity/no-lookup replacement candidates can be reviewed, but they still carry evidence risk unless staff confirms BOM and substitute rules.

## Notes for Next AI Session
- User preference: stock tracking is the critical artifact; DM can be rebuilt from a cleaner technical BOM source if needed.
- For RVC workbook export, the temporary rule is to treat all materials as `Không có xuất xứ` unless documentary proof is explicitly available.
- Replacement logic currently stops replacing a product as soon as it first reaches `đủ stock + RVC >= 35`, to preserve substitute stock for later products.
- The most important workbook finding from this session: old `VLookup`-based macros tend to resolve repeated `DM` keys toward the first match, while newer dictionary-based macros can overwrite toward the last match. The workbook therefore does not provide a stable BOM-variant identity model.
