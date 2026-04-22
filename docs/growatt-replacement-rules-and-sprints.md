# Replacement Rules And Sprint Plan

This note captures the current material-replacement rules as a general project problem, not only a Growatt-specific workaround.
The goal is to define reusable replacement logic for CO case work where BOM materials, customs evidence, and stock-allocation constraints interact.

## General Rules
### 1. Exact customs-code replacement is the first-class path
- A replacement material may use a different internal ERP code if it keeps the same customs code.
- In practice, this means ERP-code equality is not required for replacement eligibility.
- Customs-code equality is the primary hard constraint for the first replacement implementation.

### 2. Near customs-code replacement is a later phase
- If no suitable replacement is available under the same customs code, the later fallback is:
  - customs code is close or similar
  - goods description is also similar or near-equivalent
- This should not be implemented as an on-demand full scan for every search.
- It requires a prebuilt search/index layer so candidate retrieval is cheap enough to run repeatedly.

### 3. Unit-of-measure alignment is mandatory
- The BCCT unit of measure and the BOM unit of measure must be checked before a replacement candidate is accepted.
- If they differ, the candidate must be converted into the BCCT unit-of-measure basis before quantity and stock checks are evaluated.
- Replacement ranking or allocation must not compare raw quantities across mismatched units.

### 4. Value-impact priority matters
- For one specific product, replacement work should prioritize the materials that create the largest value difference first.
- The point is not only to remove unmet quantity.
- The point is to improve the chance of product-level qualification with the fewest and most meaningful substitutions.

### 5. Replacement must stay auditable
- The system must preserve:
  - original BOM material
  - chosen replacement material
  - customs-code basis
  - description-similarity basis when used
  - unit conversion basis when needed
  - stock evidence used for the replacement run
- A replacement proposal is not only an optimization output.
- It must also be explainable to staff and reviewable later.

## Suggested Sprint Layout
### Sprint 1. Exact-code replacement baseline
- Scope:
  - only allow replacements with the same customs code
  - add explicit unit-of-measure conversion into the BCCT unit basis
  - rank candidate substitutions by value impact for each product
- Goal:
  - produce a defensible first replacement pass without fuzzy matching
- Expected output:
  - replacement candidate list per product-material line
  - converted comparable quantities
  - value-impact-ranked suggestions
  - explicit audit fields for why a candidate is considered valid

### Sprint 2. Offline similarity index for near-code replacement
- Scope:
  - build a reusable material index for customs-code proximity and description similarity
  - use the index to retrieve near-code candidates when same-code replacement is exhausted
- Goal:
  - support broader replacement search without expensive on-demand scans
- Expected output:
  - indexed candidate corpus
  - similarity search rules
  - auditable candidate retrieval step
  - reusable search foundation for multiple companies and product lines

### Sprint 3. Operational ranking and review layer
- Scope:
  - combine stock sufficiency, value impact, unit conversion confidence, and replacement plausibility into one review view
  - support user review of which substitutions are worth trying first
- Goal:
  - make replacement runs operationally usable instead of producing a raw candidate dump
  - support case-level optimization such as “maximize number of products that can qualify before new material changes are required”

### Sprint 4. Cross-product optimization
- Scope:
  - evaluate replacement choices at shipment or case level, not only one product at a time
  - account for shared stock pools, competing products, and priority objectives
- Goal:
  - choose replacement plans that maximize business outcome across the whole case instead of locally improving one product while hurting another
- Expected output:
  - optimization mode for case-level replacement planning
  - clear objective settings such as:
    - maximize number of qualifying products
    - minimize number of BOM changes
    - minimize residual shortage

## Current Working Assumption
- For the current Growatt replacement work, Sprint 1 is the active scope.
- Same-customs-code replacement is in scope now.
- Near-customs-code search with description similarity is deferred until the offline index exists.
