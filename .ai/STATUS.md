# Project Status

## Current State
- Active Growatt work has moved beyond baseline feasibility into replacement execution for `B282`; `C171` is still not optimized yet, but its export product list is already pinned for BOM collection.
- The repo now contains a replacement runner and audit tooling committed at `a87d47f` (`Add Growatt replacement runner and audit tooling`).
- Current `B282` replacement results from `data/cases/growatt-rvc-20260421/b282/results/replacement-seed-summary.csv`:
  - `heuristic_best`: `4` product passes, `final_unmet_qty = 1318.1440`
  - `staff_latest_bom_invoice_order`: `4` product passes, `final_unmet_qty = 1356.1440`
  - `replaceability_cost_down_best`: `3` product passes
  - `replaceability_unmet_best`: `2` product passes
- The currently passing `B282` products under both main seeds are:
  - `PV01.0117300`
  - `PV01.0117400`
  - `PV02.0228901`
  - `PV02.0229000`
- Per-seed replacement outputs already exist under `data/cases/growatt-rvc-20260421/b282/results/replacement-excel/`, including:
  - `replacement-rvc-standard.xlsx`
  - `replacement-bom-before-after.xlsx`
  - `replacement-stock-tracking.xlsx`
- Source-of-truth status is still split:
  - root workbook `data/extracted/Growatt-20260421/Growatt/tru lui CO final  SXXK - 2025 commercial-MAC - Huyền đúng.xlsm` is more up to date for `XK` and `NK2`
  - `unpacked/lo-da-lam/.../tru lui CO final  SXXK - 2025 commercial-MAC - Huyền đúng.xlsm` contains BOM edits not present in the root workbook

## Recent Changes
- Added the replacement runner and CLI entrypoint:
  - `scripts/growatt_replacement_runner.py`
  - `scripts/build-growatt-replacement-runner.py`
- Added tests for replacement basis resolution, ambiguity handling, stop-after-pass behavior, and workbook generation in `tests/test_growatt_replacement_runner.py`.
- Added shortage reporting in `scripts/build-growatt-shortage-report.py` and `docs/growatt-b282-shortage-report.md`.
- Added DM audit tooling in `scripts/build-growatt-dm-audit.py` for:
  - products with multiple BOM blocks
  - duplicate material codes inside a BOM block
- Hardened workbook normalization in `scripts/build-growatt-case-data.py`:
  - preserve text-form codes and integral identifiers
  - reject numeric DM/XK code cells where text is required
  - force text formatting for identifier/code columns in generated xlsx outputs
- Extended `scripts/growatt-rvc-baseline.py` with export invoice metadata and starting-point helpers used by replacement planning.
- Confirmed the export product codes for the two lots now in scope:
  - `GIN01426B282`: `PV01.0117300`, `PV01.0117400`, `PV02.0228801`, `PV02.0228901`, `PV02.0229000`, `PV02.0229100`
  - `GIN01426C171`: `PV00.0048400`, `PV00.0048500`, `PV01.0117600`

## Next Steps
- Request factory/staff BOMs for all export products in `GIN01426B282` and `GIN01426C171`, plus a consolidated DM replacement sheet.
- Decide the replacement input BOM source before any further optimization:
  - rebuild from technical factory BOMs
  - or reconcile root vs `lo-da-lam` workbook edits into one curated BOM source
- Keep using root workbook `XK`/`NK2` as the stock/export baseline unless newer evidence appears.
- If work continues on replacement tooling, separate exporter-only workflows from optimizer reruns and add explicit progress/debug logging.

## Blockers
- DM/BOM source quality is still unresolved. Duplicate material lines inside BOM blocks are widespread, and the two available workbooks disagree on BOM content.
- Ambiguity/no-lookup replacement candidates can be reviewed, but they still carry evidence risk unless staff confirms BOM and substitute rules.

## Notes for Next AI Session
- User preference: stock tracking is the critical artifact; DM can be rebuilt from a cleaner technical BOM source if needed.
- For RVC workbook export, the temporary rule is to treat all materials as `Không có xuất xứ` unless documentary proof is explicitly available.
- Replacement logic now stops replacing a product as soon as it first reaches `đủ stock + RVC >= 35`, to preserve substitute stock for later products.
- The user wants a clean BOM request to staff for the two active lots, not more speculative DM interpretation from the current mixed workbook sources.
