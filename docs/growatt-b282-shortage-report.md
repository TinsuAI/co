# Growatt B282 Shortage Report

- Shipment: `GIN01426B282`
- Policy version: `2026-04-22-no-max-age`
- Policy: import at least `2` days before export, max import age `0` (`0` means no cap)
- Scope: this report covers all unmet materials in the current best B282 baseline. It distinguishes real same-code shortage from date-blocked, ambiguity, variant-choice, and low-yield lanes.

## Lane Summary

| Band | Material Count | Total Unmet |
| --- | ---: | ---: |
| true-shortage | 11 | 2983.0397 |
| date-blocked | 1 | 624.9265 |
| variant-choice-driven | 1 | 60.0000 |
| low-yield | 2 | 5506.0360 |

## Material Summary

| Material | Band | Affected Models | Total Unmet | Confirmed Gap | Total Gap | Date-Blocked Pool | Ambiguity | Bucket Mix |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| 018.0688900 | true-shortage | PV02.0228801;PV02.0229000;PV02.0229100 | 1040.0000 | 1040.0000 | 1040.0000 | 0.0000 | 0.0000 | true-shortage:5;variant-choice-driven:3 |
| 001.0036000 | true-shortage | PV02.0229100 | 885.9240 | 885.9240 | 885.9240 | 0.0000 | 0.0000 | true-shortage:1 |
| 006.0089100 | true-shortage | PV02.0229100 | 448.1157 | 448.1157 | 448.1157 | 0.0000 | 0.0000 | true-shortage:1 |
| 018.0649700 | true-shortage | PV02.0229000;PV02.0229100 | 179.0000 | 179.0000 | 179.0000 | 0.0000 | 0.0000 | true-shortage:2 |
| 018.0683100 | true-shortage | PV02.0228801 | 102.0000 | 102.0000 | 102.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 018.0699100 | true-shortage | PV02.0228901 | 85.0000 | 85.0000 | 85.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 018.0649600 | true-shortage | PV02.0229100 | 81.0000 | 81.0000 | 81.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 018.0699200 | true-shortage | PV02.0229100 | 54.0000 | 54.0000 | 54.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 030.0145500 | true-shortage | PV02.0229100 | 50.0000 | 50.0000 | 50.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 030.0001100 | true-shortage | PV02.0229100 | 32.0000 | 32.0000 | 32.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 007.0030600 | true-shortage | PV02.0228801 | 26.0000 | 26.0000 | 26.0000 | 0.0000 | 0.0000 | true-shortage:1 |
| 005.0005000 | date-blocked | PV02.0228801;PV02.0228901;PV02.0229000;PV02.0229100 | 624.9265 | 624.9265 | 624.9265 | 9000.0000 | 9000.0000 | date-blocked:3;variant-choice-driven:1 |
| 018.0649800 | variant-choice-driven | PV02.0228901 | 60.0000 | 60.0000 | 60.0000 | 0.0000 | 0.0000 | variant-choice-driven:1 |
| 001.0031200 | low-yield | PV01.0117300;PV01.0117400 | 5393.0360 | 5393.0360 | 5393.0360 | 0.0000 | 0.0000 | true-shortage:2 |
| 007.0047000 | low-yield | PV01.0117300 | 113.0000 | 113.0000 | 113.0000 | 0.0000 | 0.0000 | true-shortage:1 |

## Findings

### `018.0688900`
- Band: `true-shortage`
- Gap view: total unmet `1040.0000`, true-shortage `521.0000`, date-blocked unmet `0.0000`, variant-choice unmet `519.0000`
- Same-code pool: demand `1124.0000`, confirmed `84.0000`, candidate `0.0000`, total eligible `84.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0228801;PV02.0229000;PV02.0229100`, unlockable `PV02.0228801;PV02.0229000;PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229000__block1;PV02.0229100__block1`
- Triage mix: `true-shortage:5;variant-choice-driven:3`
- Operational read: Same-code supply is still short; some existing BOM variants reduce part of the gap but do not clear it.

### `001.0036000`
- Band: `true-shortage`
- Gap view: total unmet `885.9240`, true-shortage `885.9240`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `1041.1140`, confirmed `155.1900`, candidate `0.0000`, total eligible `155.1900`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `006.0089100`
- Band: `true-shortage`
- Gap view: total unmet `448.1157`, true-shortage `448.1157`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `1728.2857`, confirmed `1280.1700`, candidate `0.0000`, total eligible `1280.1700`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `018.0649700`
- Band: `true-shortage`
- Gap view: total unmet `179.0000`, true-shortage `179.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `202.0000`, confirmed `23.0000`, candidate `0.0000`, total eligible `23.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229000;PV02.0229100`, unlockable `PV02.0229000;PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229000__block1;PV02.0229100__block1`
- Triage mix: `true-shortage:2`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `018.0683100`
- Band: `true-shortage`
- Gap view: total unmet `102.0000`, true-shortage `102.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `144.0000`, confirmed `42.0000`, candidate `0.0000`, total eligible `42.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0228801`, unlockable `PV02.0228801`, below-threshold `none`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `018.0699100`
- Band: `true-shortage`
- Gap view: total unmet `85.0000`, true-shortage `85.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `231.0000`, confirmed `146.0000`, candidate `0.0000`, total eligible `146.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0228901`, unlockable `PV02.0228901`, below-threshold `none`
- Existing alt variants: `PV02.0228901-NEW__block2`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `018.0649600`
- Band: `true-shortage`
- Gap view: total unmet `81.0000`, true-shortage `81.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `202.0000`, confirmed `121.0000`, candidate `0.0000`, total eligible `121.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `018.0699200`
- Band: `true-shortage`
- Gap view: total unmet `54.0000`, true-shortage `54.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `664.0000`, confirmed `610.0000`, candidate `0.0000`, total eligible `610.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `030.0145500`
- Band: `true-shortage`
- Gap view: total unmet `50.0000`, true-shortage `50.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `634.0000`, confirmed `584.0000`, candidate `0.0000`, total eligible `584.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `030.0001100`
- Band: `true-shortage`
- Gap view: total unmet `32.0000`, true-shortage `32.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `404.0000`, confirmed `372.0000`, candidate `0.0000`, total eligible `372.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0229100`, unlockable `PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0229100__block1`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `007.0030600`
- Band: `true-shortage`
- Gap view: total unmet `26.0000`, true-shortage `26.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `288.0000`, confirmed `262.0000`, candidate `0.0000`, total eligible `262.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0228801`, unlockable `PV02.0228801`, below-threshold `none`
- Triage mix: `true-shortage:1`
- Operational read: Same-code supply is genuinely short under the active policy window.

### `005.0005000`
- Band: `date-blocked`
- Gap view: total unmet `624.9265`, true-shortage `0.0000`, date-blocked unmet `277.8885`, variant-choice unmet `347.0380`
- Same-code pool: demand `868.6465`, confirmed `243.7200`, candidate `0.0000`, total eligible `243.7200`, date-blocked pool `9000.0000`, ambiguity `9000.0000`
- Models: affected `PV02.0228801;PV02.0228901;PV02.0229000;PV02.0229100`, unlockable `PV02.0228801;PV02.0228901;PV02.0229000;PV02.0229100`, below-threshold `none`
- Existing alt variants: `PV02.0228901-NEW__block2;PV02.0229000__block1;PV02.0229100__block1`
- Ambiguity variant hits: `PV02.0228801__block1;PV02.0228901-NEW__block1;PV02.0228901-NEW__block2;PV02.0228901__block1;PV02.0229000-NEW__block1;PV02.0229100-NEW__block1`
- Triage mix: `date-blocked:3;variant-choice-driven:1`
- Operational read: Visible same-code supply exists, but it is currently blocked by the date rule and still carries ambiguity evidence.

### `018.0649800`
- Band: `variant-choice-driven`
- Gap view: total unmet `60.0000`, true-shortage `0.0000`, date-blocked unmet `0.0000`, variant-choice unmet `60.0000`
- Same-code pool: demand `87.0000`, confirmed `27.0000`, candidate `0.0000`, total eligible `27.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV02.0228901`, unlockable `PV02.0228901`, below-threshold `none`
- Existing alt variants: `PV02.0228901-NEW__block2`
- Triage mix: `variant-choice-driven:1`
- Operational read: Current unmet quantity is driven by the chosen BOM variant; an existing variant can remove or reduce this requirement.

### `001.0031200`
- Band: `low-yield`
- Gap view: total unmet `5393.0360`, true-shortage `5393.0360`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `10011.9460`, confirmed `4618.9100`, candidate `0.0000`, total eligible `4618.9100`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV01.0117300;PV01.0117400`, unlockable `none`, below-threshold `PV01.0117300;PV01.0117400`
- Triage mix: `true-shortage:2`
- Operational read: Affected models are still below RVC 35 before fixing this material, so this gap is not a near-term unlock.

### `007.0047000`
- Band: `low-yield`
- Gap view: total unmet `113.0000`, true-shortage `113.0000`, date-blocked unmet `0.0000`, variant-choice unmet `0.0000`
- Same-code pool: demand `130.0000`, confirmed `17.0000`, candidate `0.0000`, total eligible `17.0000`, date-blocked pool `0.0000`, ambiguity `0.0000`
- Models: affected `PV01.0117300`, unlockable `none`, below-threshold `PV01.0117300`
- Triage mix: `true-shortage:1`
- Operational read: Affected models are still below RVC 35 before fixing this material, so this gap is not a near-term unlock.
