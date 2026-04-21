# Growatt RVC Case Execution Plan

## Purpose
This document captures the agreed execution plan for the Growatt case covering the unfinished shipments:

- `GIN01426B282`
- `GIN01426C171`

The immediate goal is not to redesign the whole project.
The immediate goal is to produce a defensible baseline for the current case, then decide whether existing BOM variants are enough or whether new replacement-driven variants are required.

## Case-Specific Constraints
- Growatt uses both internal product/material codes and customs HS codes.
- The workbook is operationally useful, but it is not a clean source of truth in every column.
- `DM` should be treated as the source of available BOM variants.
- `NK2` is the usable source of current `CO stock` snapshot at import-row granularity.
- Materials used for one exported product must come from import rows registered at least `2` days before the export date of that product.
- Completed RVC files are reference examples, not authoritative source data.
- Existing workbook output supports one material drawing from multiple import declarations, but it currently renders those sources inside one output row.
- Future logic should remain configurable enough to support:
  - one material code using multiple import buckets
  - one original material being replaced by one or many alternative material codes

## Working Rule Assumption
The completed Growatt references show a product-level rule pattern of `RVC 35% + CTSH`.

For this case, the immediate operational bottleneck raised by staff is the `RVC 35%` threshold.
The baseline implementation therefore focuses on:

- shipment-scoped allocation
- product-level `RVC` calculation
- traceability to import rows

`CTSH` remains a case gate and reporting note, but it is not yet auto-evaluated by the first baseline script.

## Data Sources
### Shipment / export data
- External BCCT export report:
  - `data/extracted/Growatt-20260421/Growatt/BaoCaoHangChiTietXK GRW T3-T4.2026 moi.xls`

### BOM variants
- Workbook `DM` sheet:
  - `data/extracted/Growatt-20260421/Growatt/tru lui CO final  SXXK - 2025 commercial-MAC - Huyền đúng.xlsm`

### Current CO stock snapshot
- Workbook `NK2` sheet from the same workbook.

### Case-specific normalized working data
- Extracted and normalized outputs for this task should be stored separately from raw source files.
- The intended working area is a case-specific directory under `data/` so the task can run on stable normalized data instead of repeatedly parsing raw files in-place.

## Shipment Priority
The unfinished shipments should not be worked in parallel.

They should be processed sequentially, with the earlier export shipment first:

1. `GIN01426B282`
   - export declaration `308416454160`
   - invoice date `2026-04-09`
2. `GIN01426C171`
   - export declaration `308438331410`
   - invoice date `2026-04-16`

The practical reason is to keep stock allocation, scenario review, and replacement logic focused on one case at a time.

## Identity Model
The workbook cannot safely be modeled as `one product code -> one BOM`.

For this case, product/BOM identity should be split into separate layers:

- `export_product_code`
  - the product code actually exported in the shipment, such as `PV02.0229100`
- `bom_code`
  - the code written in `DM`, preserving suffixes exactly as observed, such as `PV02.0229100` or `PV02.0229100-NEW`
- `bom_variant_id`
  - the unique identity of one contiguous `DM` block, such as `PV02.0229100__block1`
- optional `product_family_code`
  - a helper grouping key used only for variant discovery, not as a primary identity

For this case:
- every contiguous `DM` block is treated as one BOM variant
- every BOM variant gets a stable `bom_variant_id`
- suffix BOMs are kept distinct at the `bom_code` layer
- exact duplicate product codes in separate `DM` blocks remain separate `bom_variant_id` values unless later evidence proves they are the same BOM

This avoids two failure modes:

- collapsing suffix BOMs into the unsuffixed export code
- collapsing repeated exact-code blocks such as duplicated `PV02.0229100` blocks into one identity

At the current evidence level, repeated exact-code blocks should be treated as distinct BOM blocks, not auto-merged.

## Baseline Evaluation Flow
1. Start with one shipment only, not both:
   - first `GIN01426B282`
   - then `GIN01426C171`
2. Resolve the shipment to its export declaration rows, product internal codes, quantities, unit export prices, and finished-good HS.
3. Freeze the current `NK2` stock state as the baseline allocation snapshot for that shipment run.
4. For each `export_product_code`, enumerate related BOM options from `DM`:
   - exact `bom_code` matches
   - suffix `bom_code` variants whose code starts with the exported code plus `-`
   - keep every matching contiguous `DM` block as a separate `bom_variant_id`
5. Build shipment scenarios as combinations of those per-product `bom_variant_id` choices.
6. For each scenario, run one joint stock allocation across the whole shipment:
   - allocation unit is the import row in `NK2`
   - only import rows with `import_date <= export_date - 2 days` are eligible
   - one material code may draw from multiple import rows
   - products in the same shipment compete for the same remaining stock pool
7. Compute product-level `RVC` using:
   - `FOB = export quantity x export unit price`
   - non-originating input value derived from allocated material rows
8. Rank scenarios by:
   - stock sufficiency
   - product pass/fail against `35%`
   - minimum margin to threshold
   - traceability quality

## Valuation Convention
The first implementation keeps a configurable valuation mode.

Supported modes:
- `workbook_avg`
  - compatibility mode that mimics the current workbook more closely
  - if one product-material line uses multiple import rows, the line unit price is the simple average of the contributing row unit prices
- `weighted`
  - uses quantity-weighted average of the contributing row unit prices
  - this is closer to the real allocated material value when source rows contribute different quantities

Default mode for the case runner is still `workbook_avg` so the baseline stays closer to current staff output behavior, but it should be treated as a workbook-compatibility baseline rather than the final valuation truth.

## Origin Convention
The baseline runner uses a conservative operational rule:

- if all allocated rows for a product-material line are marked Vietnam origin, treat the line as originating
- otherwise treat the line as non-originating

This is still weaker than a full proof-aware origin engine, but it is explicit and reproducible.

## Decision Gates
### Gate A: existing variant selection
If a shipment scenario built only from existing `DM` variants:
- has sufficient stock
- clears the `35%` threshold for every product in scope
- remains traceable to import rows

then stop at scenario selection.

### Gate B: replacement search
Only if all existing-variant scenarios fail should the workflow continue to replacement analysis:
- identify the cost-driving materials
- search alternative materials under configurable rules
- build new BOM variants
- update the case BOM data with the newly accepted BOM variant
- rerun the same shipment-level allocation and `RVC` evaluation

The same gate applies sequentially:
- finish the decision for `B282` first
- only then move to `C171`

## Post-Decision Operational Updates
After a shipment scenario is accepted for actual use:
- update the `CO stock` tracking to reflect the allocated import-row consumption
- persist the accepted `bom_variant_id` choice for the shipment
- if the accepted solution required replacement-driven BOM changes, create the new BOM variant and update the working BOM dataset before the next shipment run

## First Implementation Output
The first case runner should produce:
- shipment product inventory
- variant options per exported product
- normalized extracted data for the case workspace
- ranked baseline scenarios
- per-product `RVC` result for each scenario
- stock sufficiency and unmet-demand flags
- enough allocation detail to support later `CO stock` updates
- enough variant detail to support later BOM updates

## Implementation Boundary
This plan is intentionally narrow.
It is designed to solve the current Growatt case without forcing a final architecture for the future product.
