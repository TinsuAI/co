# What Was Done
- Added a full Growatt replacement lane on top of the existing baseline flow:
  - `scripts/growatt_replacement_runner.py`
  - `scripts/build-growatt-replacement-runner.py`
- Implemented replacement basis grouping, candidate generation, sequential stock-local replacement execution, per-product BOM snapshots, and per-seed workbook exports for `B282`.
- Added `tests/test_growatt_replacement_runner.py` covering:
  - customs-basis resolution for `NK2_ONLY`
  - review lanes for clean/date-blocked/ambiguity candidates
  - stop-after-pass behavior
  - workbook generation and stock tracking export
- Added shortage material reporting in `scripts/build-growatt-shortage-report.py` and `docs/growatt-b282-shortage-report.md`.
- Added DM audit tooling in `scripts/build-growatt-dm-audit.py` to compare both Growatt DM workbooks and export Excel summaries for:
  - product codes with multiple BOM blocks
  - duplicate material codes within a single BOM block
- Hardened workbook parsing in `scripts/build-growatt-case-data.py` so BOM/XK/NK identifiers and material codes stay as text and integral identifiers do not silently coerce through numeric cells.
- Extended `scripts/growatt-rvc-baseline.py` so export rows carry invoice metadata and starting-point selection helpers used by the replacement runner.
- Verified the code with:
  - `python3 -m unittest tests.test_growatt_replacement_runner`
  - `python3 -m py_compile scripts/build-growatt-dm-audit.py scripts/build-growatt-replacement-runner.py scripts/build-growatt-shortage-report.py scripts/growatt_replacement_runner.py scripts/growatt-rvc-baseline.py scripts/build-growatt-case-data.py`
- Committed the code as `a87d47f Add Growatt replacement runner and audit tooling`.

# Decisions Made
- Replacement optimization is no longer pure unmet-qty rescue. It now supports same-family replacement search, but a product stops consuming substitute stock as soon as it first satisfies `stock sufficient + RVC >= 35`.
- For workbook output, `Có xuất xứ` must not be inferred from `origin = VN`; the temporary safe rule is to export all materials as `Không có xuất xứ` until documentary proof exists.
- Root workbook remains the operational source for `XK` and `NK2` because it is newer and fuller there.
- The `unpacked/lo-da-lam` workbook is still useful because it contains BOM edits absent from the root workbook.
- The user considers stock tracking the most critical artifact. DM is replaceable if a cleaner technical BOM source is obtained from the factory.
- The two lots currently driving BOM collection are:
  - `GIN01426B282`: `PV01.0117300`, `PV01.0117400`, `PV02.0228801`, `PV02.0228901`, `PV02.0229000`, `PV02.0229100`
  - `GIN01426C171`: `PV00.0048400`, `PV00.0048500`, `PV01.0117600`

# What Didn't Work
- Earlier replacement logic was too permissive:
  - it treated `declared_code` as if it were always the customs-family basis
  - it allowed optimization to continue after a product had already passed
  - it mixed analytical partial-improvement moves with declaration-grade feasible moves
- Export-only artifact generation still triggered full reruns because exporter and optimizer are not cleanly separated in the current code structure.
- DM interpretation from the mixed workbook sources remains unreliable. Duplicate material lines are widespread inside BOM blocks, and BOM corrections in `lo-da-lam` do not align cleanly with the newer `XK`/`NK2` data in the root workbook.

# Open Items
- Get BOMs for all export products in `GIN01426B282` and `GIN01426C171` from staff/factory, plus a replacement DM sheet.
- Decide whether future optimization should run on:
  - a rebuilt technical BOM source from the factory
  - or a curated merged BOM derived from the two current workbooks
- If replacement tooling continues, split exporter-only commands from optimization reruns and add progress/debug logging so long runs are observable.
