# What Was Done
- Added [`scripts/build-growatt-shipment-admissibility.py`](/home/vp/workspace/client/barry-CO/scripts/build-growatt-shipment-admissibility.py) to derive shipment-specific `B282` admissibility data from shared normalized inputs without mutating the shared ledger.
- Reworked [`scripts/growatt-rvc-baseline.py`](/home/vp/workspace/client/barry-CO/scripts/growatt-rvc-baseline.py) so it now reads:
  - `shared/normalized/imports-normalized.csv`
  - `shared/normalized/exports-normalized.csv`
  - `shared/normalized/co-stock-tracking-updated.csv`
  - shipment-specific admissibility artifacts
- Added configurable stock-age controls to admissibility and baseline logic:
  - `--import-lead-days` default `2`
  - `--max-import-age-days` default `365`
- Extended [`scripts/build-growatt-case-data.py`](/home/vp/workspace/client/barry-CO/scripts/build-growatt-case-data.py) to normalize import-side values to USD with explicit provenance fields and a customs-rate-first conversion path.
- Added [`scripts/download-customs-fx-rates.py`](/home/vp/workspace/client/barry-CO/scripts/download-customs-fx-rates.py) to download the customs FX workbook directly from public `customs.gov.vn` JSON endpoints.
- Rebuilt shared normalized data, rebuilt `B282` admissibility outputs, reran the `B282` baseline, rerendered the case view, and wrote [`docs/growatt-valuation-and-stock-insights.md`](/home/vp/workspace/client/barry-CO/docs/growatt-valuation-and-stock-insights.md).

# Decisions Made
- Do not interpret the source `exchange_rate` field as a guaranteed USD conversion basis. Use the customs USD table by declaration date whenever possible.
- Keep valuation auditability in the normalized data by storing `price_normalization_basis`, `source_unit_price`, and both source and applied exchange rates.
- Keep stock-age policy explicit and configurable. The current default is:
  - imported at least `2` days before export
  - imported within the prior `365` days
- Treat old `NK2` carry-over stock as possible ledger evidence, not automatically valid evidence for a `2026` CO run.
- Keep shipment interpretation in derived artifacts and keep shared normalized evidence immutable.

# What Didn't Work
- The earlier shortcut `unit_price_usd = tax_unit_price / exchange_rate` was not defensible as a general rule because rows with `source_exchange_rate = 1` still had VND-denominated `tax_unit_price`.
- Assuming the customs rate page required browser/CAPTCHA automation was unnecessary; the page data is available through simple JSON endpoints.
- The earlier `B282` baseline that read raw `NK2` and accepted older stock without a max-age rule was too optimistic and should not be reused as evidence.

# Open Items
- Review the current `B282` shortages by material family and decide which failures are true stock gaps versus policy-driven date exclusions.
- Confirm whether the one-year stock-age window is the intended business rule or just a practical default.
- Only after `B282` is accepted should the same pipeline be run for `C171`.
- If a future run depends on updated customs rates, rerun `python3 scripts/download-customs-fx-rates.py` before rebuilding shared normalized data.
