# Growatt B282 True Shortage Report

- Shipment: `GIN01426B282`
- Policy version: `2026-04-22-no-max-age`
- Policy: import at least `2` days before export, max import age `0` (`0` means no cap)
- Best scenario unmet qty: `9174.0021`
- Bucket totals: true-shortage `7970.0757`, variant-choice-driven `926.0380`, date-blocked `277.8885`

## What Counts As True Shortage

These are the gaps that still remain after removing the one-year stock window. They are not primarily explained by the date policy anymore. Some are single-material shortages; others are shipment-level shortages where multiple products compete for the same eligible stock pool.

## Highest-Impact Materials

| Material | Affected Models | True Unmet Qty | Shipment Demand | Eligible Supply | Shipment Gap | Alt Variants |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| 001.0031200 | PV01.0117300;PV01.0117400 | 5393.0360 | 10011.9460 | 4618.9100 | 5393.0360 |  |
| 001.0036000 | PV02.0229100 | 885.9240 | 1041.1140 | 155.1900 | 885.9240 | PV02.0229100__block1 |
| 018.0688900 | PV02.0228801;PV02.0229000 | 521.0000 | 1124.0000 | 84.0000 | 1040.0000 | PV02.0229000__block1 |
| 006.0089100 | PV02.0229100 | 448.1157 | 1728.2857 | 1280.1700 | 448.1157 | PV02.0229100__block1 |
| 018.0649700 | PV02.0229000;PV02.0229100 | 179.0000 | 202.0000 | 23.0000 | 179.0000 | PV02.0229000__block1;PV02.0229100__block1 |
| 007.0047000 | PV01.0117300 | 113.0000 | 130.0000 | 17.0000 | 113.0000 |  |
| 018.0683100 | PV02.0228801 | 102.0000 | 144.0000 | 42.0000 | 102.0000 |  |
| 018.0699100 | PV02.0228901 | 85.0000 | 231.0000 | 146.0000 | 85.0000 | PV02.0228901-NEW__block2 |
| 018.0649600 | PV02.0229100 | 81.0000 | 202.0000 | 121.0000 | 81.0000 | PV02.0229100__block1 |
| 018.0699200 | PV02.0229100 | 54.0000 | 664.0000 | 610.0000 | 54.0000 | PV02.0229100__block1 |
| 030.0145500 | PV02.0229100 | 50.0000 | 634.0000 | 584.0000 | 50.0000 | PV02.0229100__block1 |
| 030.0001100 | PV02.0229100 | 32.0000 | 404.0000 | 372.0000 | 32.0000 | PV02.0229100__block1 |

## Findings

### `001.0031200`
- Affected models: `PV01.0117300`, `PV01.0117400`
- True unmet quantity in best scenario: `5393.0360`
- Shipment demand vs eligible supply: `10011.9460` vs `4618.9100` (gap `5393.0360`)
- Current triage mix: `true-shortage:2`
- Existing variants worth checking: none observed
- Why it is still a true shortage: Shipment-level demand 10011.9460 exceeds shared eligible supply 4618.9100. This is a shared-pool shortage across multiple products.
- Operational read: Check shared-stock allocation and stock evidence for this material across all affected models.

### `001.0036000`
- Affected models: `PV02.0229100`
- True unmet quantity in best scenario: `885.9240`
- Shipment demand vs eligible supply: `1041.1140` vs `155.1900` (gap `885.9240`)
- Current triage mix: `true-shortage:1`
- Existing variants worth checking: `PV02.0229100__block1`
- Why it is still a true shortage: Eligible supply 155.1900 is below the chosen variant demand 1041.1140 even after removing the one-year cap.
- Operational read: Check stock evidence first; if unchanged, treat this as a candidate for replacement-driven BOM analysis.

### `018.0688900`
- Affected models: `PV02.0228801`, `PV02.0229000`
- True unmet quantity in best scenario: `521.0000`
- Shipment demand vs eligible supply: `1124.0000` vs `84.0000` (gap `1040.0000`)
- Current triage mix: `true-shortage:5;variant-choice-driven:3`
- Existing variants worth checking: `PV02.0229000__block1`
- Why it is still a true shortage: Shipment-level demand 1124.0000 exceeds shared eligible supply 84.0000. This is a shared-pool shortage across multiple products.
- Operational read: Check shared-stock allocation and stock evidence for this material across all affected models.

### `006.0089100`
- Affected models: `PV02.0229100`
- True unmet quantity in best scenario: `448.1157`
- Shipment demand vs eligible supply: `1728.2857` vs `1280.1700` (gap `448.1157`)
- Current triage mix: `true-shortage:1`
- Existing variants worth checking: `PV02.0229100__block1`
- Why it is still a true shortage: Shipment-level demand 1728.2857 exceeds shared eligible supply 1280.1700. This is a shared-pool shortage across multiple products.
- Operational read: Check shared-stock allocation and stock evidence for this material across all affected models.

### `018.0649700`
- Affected models: `PV02.0229000`, `PV02.0229100`
- True unmet quantity in best scenario: `179.0000`
- Shipment demand vs eligible supply: `202.0000` vs `23.0000` (gap `179.0000`)
- Current triage mix: `true-shortage:2`
- Existing variants worth checking: `PV02.0229000__block1`, `PV02.0229100__block1`
- Why it is still a true shortage: Shipment-level demand 202.0000 exceeds shared eligible supply 23.0000. This is a shared-pool shortage across multiple products.
- Operational read: Check shared-stock allocation and stock evidence for this material across all affected models.

### `007.0047000`
- Affected models: `PV01.0117300`
- True unmet quantity in best scenario: `113.0000`
- Shipment demand vs eligible supply: `130.0000` vs `17.0000` (gap `113.0000`)
- Current triage mix: `true-shortage:1`
- Existing variants worth checking: none observed
- Why it is still a true shortage: Eligible supply 17.0000 is below the chosen variant demand 130.0000 even after removing the one-year cap.
- Operational read: Check stock evidence first; if unchanged, treat this as a candidate for replacement-driven BOM analysis.

### `018.0683100`
- Affected models: `PV02.0228801`
- True unmet quantity in best scenario: `102.0000`
- Shipment demand vs eligible supply: `144.0000` vs `42.0000` (gap `102.0000`)
- Current triage mix: `true-shortage:1`
- Existing variants worth checking: none observed
- Why it is still a true shortage: Eligible supply 42.0000 is below the chosen variant demand 144.0000 even after removing the one-year cap.
- Operational read: Check stock evidence first; if unchanged, treat this as a candidate for replacement-driven BOM analysis.

### `018.0699100`
- Affected models: `PV02.0228901`
- True unmet quantity in best scenario: `85.0000`
- Shipment demand vs eligible supply: `231.0000` vs `146.0000` (gap `85.0000`)
- Current triage mix: `true-shortage:1`
- Existing variants worth checking: `PV02.0228901-NEW__block2`
- Why it is still a true shortage: Shipment-level demand 231.0000 exceeds shared eligible supply 146.0000. This is a shared-pool shortage across multiple products.
- Operational read: Check shared-stock allocation and stock evidence for this material across all affected models.
