# Growatt DM Workbook Behavior

This note captures what the current VBA workbook actually does when `DM` contains repeated product codes or multiple BOM-like blocks for the same product.

It is a behavior note, not a statement that the workbook model is correct for the real business domain.

## Key Finding

The workbook does not model `DM` as explicit BOM variants.

Instead, the VBA treats `DM` as a flat row list keyed by:
- product code
- row order within that product
- helper keys such as `stt + product_code`

There is no stable concept of:
- contiguous BOM block
- BOM version selection
- exact-code duplicate block identity

## What The VBA Actually Does

### `DM.sumifDM`
- Fills derived totals in `DM`
- Builds shipment/run key `K6`
- Pushes `K6` into `NK2` and `Xuat`

This step does not choose a BOM version.

### `HideCopy.update`
- Reads product codes from `DM`
- Adds them into `Xuat` as unique product codes
- Aggregates quantities with `SumIf`

This means multiple `DM` blocks with the same product code are collapsed into one `Xuat` product entry.

### `Run1` and `Run3`
- Count BOM rows for a product with `CountIf(DM!A:A, product_code)`
- Read BOM lines with `VLookup(stt & product_code, DM!B:I, ...)`

This is positional lookup, not block-aware lookup.

If the same exact `stt + product_code` key appears in more than one `DM` block, classic `VLookup` behavior means the first matching row wins.

### `RunUpgrade`
- Rebuilds the same lookup model with dictionaries
- Uses dictionary keys from `DM` row helper columns rather than a distinct BOM variant id

If duplicate keys exist, later rows overwrite earlier rows.

This means newer-style macros can resolve the same repeated-code data differently from older `VLookup` macros.

## Practical Consequence

If the real workbook convention is:
- one active BOM per product code
- or manual cleanup before every run
- or suffix-based code separation for every version

then the workbook can still produce operational output.

But if the real data contains multiple valid BOM versions under the same exact product code, the workbook model is under-specified:
- old macros tend to read the first matching rows
- newer dictionary-based macros can overwrite toward the last matching rows
- `Xuat` still sees one product code, not multiple BOM variants

So repeated exact-code blocks in `DM` are not safely distinguishable in the current workbook logic.

## Current Target Product Coverage

The active Growatt target codes are present in the current normalized artifacts:

- `GIN01426B282`
  - `PV01.0117300`
  - `PV01.0117400`
  - `PV02.0228801`
  - `PV02.0228901`
  - `PV02.0229000`
  - `PV02.0229100`
- `GIN01426C171`
  - `PV00.0048400`
  - `PV00.0048500`
  - `PV01.0117600`

In `dm-variants.csv`, all nine codes have exact `bom_code` coverage.
Eight currently show one exact block.
`PV02.0229100` currently shows two exact blocks:
- `PV02.0229100__block1`
- `PV02.0229100__block2`

Some `B282` products also have suffix variants such as `-NEW`, which reinforces that the future system must preserve BOM identity more explicitly than the workbook does.

## Design Implication

The future system should treat BOM identity explicitly:
- `export_product_code`
- `bom_code`
- `bom_variant_id`

It should not depend on row order or implicit `VLookup` precedence to decide which BOM is active.
