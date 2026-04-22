# Growatt Valuation And Stock Insights

This note captures the highest-value findings from the current Growatt RVC case work so later case runs do not repeat the same valuation and stock-assumption mistakes.

## 1. BCCT import history is only 2026

- [`imports-normalized.csv`](</home/vp/workspace/client/barry-CO/data/cases/growatt-rvc-20260421/shared/normalized/imports-normalized.csv>) contains `3045` BCCT import rows and every row is dated in `2026`.
- Any `2023` or `2024` inventory showing up in the merged CO stock ledger comes from workbook `NK2` carry-over stock, not from the BCCT import report for this case.
- This matters because “old stock exists in the ledger” and “old stock is evidenced by the current BCCT import report” are not the same claim.

## 2. Source `exchange_rate` cannot be treated as “USD rate” by default

- The raw workbook / report fields use mixed price representations.
- Some rows have `source_exchange_rate = 1` while `tax_unit_price` is clearly in VND, so the old rule `unit_price_usd = tax_unit_price / source_exchange_rate` was not defensible as a general normalization rule.
- The safe interpretation is:
  - `tax_unit_price` is often the customs-tax basis in VND.
  - `raw unit price` may already be USD on some rows.
  - `source_exchange_rate` may be useful fallback data, but it is not reliable enough to define USD conversion semantics by itself.

## 3. Import valuation now normalizes to USD with explicit audit basis

- [`scripts/build-growatt-case-data.py`](../scripts/build-growatt-case-data.py) now converts import-side values to USD using the customs USD rate table by effective date whenever coverage exists.
- The normalization order is:
  1. Use `tax_unit_price / customs_usd_rate` when the customs table covers the declaration date.
  2. Fallback to `tax_unit_price / source_exchange_rate` only when the customs table does not cover that date and the source rate is meaningful.
  3. Fallback to the raw unit price field when no credible tax-based conversion is available.
- The normalized data now keeps audit fields instead of hiding the derivation:
  - `source_unit_price`
  - `source_exchange_rate`
  - `exchange_rate` used for USD normalization
  - `price_normalization_basis`
- Current shared manifest shows:
  - `3045` import rows normalized via `tax_vnd_over_customs_usd_rate`
  - merged stock rows split across `25352` customs-rate rows, `7707` source-rate fallback rows, and `6` raw-price rows

## 4. Customs USD rates can be refreshed automatically

- [`scripts/download-customs-fx-rates.py`](../scripts/download-customs-fx-rates.py) downloads the customs FX tables directly from `customs.gov.vn` and writes `data/reference/DS_ty_gia_ngoai_te.xlsx`.
- The script does not need browser automation or CAPTCHA solving; it calls the same public JSON endpoints the page uses.
- This removes the manual dependency on copying a workbook from Downloads every time the customs table needs refresh.

## 5. Old `NK2` stock exists in material volume, but should be treated cautiously

- The merged stock ledger still shows positive `NK2` carry-over from prior years:
  - `2023`: `678` rows, `353,894.54` remaining quantity
  - `2024`: `74` rows, `372,069.22` remaining quantity
  - `2025`: `2795` rows, `65,213,087.72` remaining quantity
  - `2026`: `4131` rows, `56,395,673.73` remaining quantity
- So the question is not “does old stock exist?”; it does.
- The real question is whether that old stock should still be admissible for a `2026` CO case.

## 6. Stock-age admissibility is now explicit and configurable

- Shipment admissibility and baseline allocation now use two time gates:
  - `import_date <= export_date - 2 days`
  - `import_date >= export_date - 365 days` by default
- Both are configurable in:
  - [`scripts/build-growatt-shipment-admissibility.py`](../scripts/build-growatt-shipment-admissibility.py)
  - [`scripts/growatt-rvc-baseline.py`](../scripts/growatt-rvc-baseline.py)
- The important change is not just configurability. The pipeline now makes the stock-age policy explicit instead of silently consuming very old ledger rows.

## 7. Old stock was technically relevant to `B282`, but is no longer used under the default one-year rule

- Before adding the one-year window, `B282` could draw a small amount of `2023-2024` stock from the ledger.
- After the default `365`-day window was applied, the best current `B282` scenario uses `0` source allocations from `2023-2024`.
- That means the current `B282` result is no longer being propped up by very old carry-over material.

## 8. The corrected `B282` pipeline still fails, and the shortage pattern is now more trustworthy

- `B282` now runs from:
  - shared normalized CSVs
  - variant-scoped admissibility artifacts
  - explicit stock-age rules
  - USD-normalized valuation inputs
- Current best scenario still fails with:
  - `total_unmet_qty = 10807.002139`
  - `min_margin = -7.132388635440243`
- Shortage evidence is now easier to trust because it is no longer mixing raw `NK2` logic, shipment-wide candidate confirmation, and ambiguous currency assumptions.
- The largest remaining gaps are concentrated in a few material groups, with several lines also showing large date-blocked quantities, including `009.0002801`, `005.0005000`, `940.0180200`, `940.0637900`, `030.0122600`, and `006.0088400`.

## 9. Practical implication for future case work

- Do not treat workbook prices as already normalized to USD without an audit basis.
- Do not treat old `NK2` stock as automatically usable just because it remains on the ledger.
- Keep shipment interpretation in derived artifacts; keep shared normalized evidence immutable.
- When a case fails after the stricter pipeline, the failure is more meaningful and should drive the next review.
