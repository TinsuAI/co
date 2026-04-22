# Session Summary: Growatt Policy, Sequential Search, And Replacement Planning

## What Was Done
- Moved the hand-authored Growatt case policy out of ignored `data/` and into tracked config at [config/growatt-rvc-20260421.json](/home/vp/workspace/client/barry-CO/config/growatt-rvc-20260421.json).
- Added [scripts/growatt_case_config.py](/home/vp/workspace/client/barry-CO/scripts/growatt_case_config.py) and updated the Growatt admissibility, baseline, and render scripts so case policy is resolved from tracked config and stamped into local run artifacts.
- Rebuilt `B282` under the updated staff-confirmed policy `2026-04-22-no-max-age`, which keeps `import_lead_days = 2` and removes the one-year stock cap.
- Added shortage-triage tooling in [scripts/build-growatt-shortage-triage.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-shortage-triage.py) to classify unmet lines by `true-shortage`, `date-blocked`, `ambiguity-blocked`, or `variant-choice-driven`.
- Added [scripts/build-growatt-true-shortage-report.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-true-shortage-report.py) and generated [docs/growatt-b282-true-shortage-report.md](/home/vp/workspace/client/barry-CO/docs/growatt-b282-true-shortage-report.md) to isolate the highest-impact shortages that remain after removing the one-year policy cap.
- Added [scripts/build-growatt-sequential-priority-report.py](/home/vp/workspace/client/barry-CO/scripts/build-growatt-sequential-priority-report.py) and generated [docs/growatt-b282-sequential-priority-report.md](/home/vp/workspace/client/barry-CO/docs/growatt-b282-sequential-priority-report.md) for exact product-order search.
- Reviewed the sequential code/report and fixed two quality issues:
  - the search now covers all BOM scenarios by default instead of only the top `5`
  - the report now distinguishes `Đủ Stock` from `Đạt C/O` instead of showing stock sufficiency as if it were product pass/fail
- Documented generalized replacement rules and sprint framing in [docs/growatt-replacement-rules-and-sprints.md](/home/vp/workspace/client/barry-CO/docs/growatt-replacement-rules-and-sprints.md), then linked it from [docs/README.md](/home/vp/workspace/client/barry-CO/docs/README.md).

## Decisions Made
- The Growatt shipment policy is now versioned and tracked in repo config, not stored only in ignored local data.
- Staff-confirmed policy for the current case is:
  - import date must be at least `2` days before export
  - no maximum stock age cap
- `B282` remains the only active decision case. `C171` stays deferred until `B282` replacement direction is clearer.
- The working optimization goal is to maximize the number of products that can qualify at product level, not merely minimize shipment-level unmet quantity.
- For replacement work, the team should lock one current best baseline BOM scenario and compare all future BOM changes against that fixed starting point.
- Replacement planning notes were rewritten as a general project pattern rather than a Growatt-only workaround.

## What Didn't Work
- Relaxing the policy by removing the one-year stock cap did not make `B282` pass. It reduced the best unmet quantity from `10807.0021` to `9174.0021`, but true shortage still dominates.
- Sequential allocation was tested as an exact search across the full current search space (`18` BOM scenarios x `720` product orders = `12960` combinations). No combination produced even one product that both had enough stock and met `RVC >= 35`.
- A shipment/order that minimizes unmet quantity is not automatically the best baseline for replacement planning, because some shortages may be much harder to replace than others. That gap was identified but not solved in this session.

## Open Items
- Build a replaceability-aware report for the main true-shortage materials in the locked `B282` baseline scenario.
- Use that report to identify which shortages are realistic Sprint 1 same-customs-code replacement targets.
- Revisit HQ/ERP mapping confidence for the materials that look like replacement bottlenecks before treating every shortage as replacement-ready.
- Leave `C171` alone until the `B282` replacement path is clearer.
