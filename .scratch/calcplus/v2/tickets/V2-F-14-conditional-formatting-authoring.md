# V2-F-14 — Conditional-formatting authoring

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 3–4 pw (provisional)
- **Prerequisites:** V2-V-03 (for the unsupported rule types)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §3; known
  deviation D-5

## Scope

Add the Rule Manager and the New/Edit Rule dialogs for cell-value, text, date,
top/bottom, duplicate/unique, data-bar, colour-scale and icon-set rules, plus
stop-if-true, priority and apply-to-range. Formula-based, above-average,
blank/error and formula-threshold rules become authorable once V2-V-03 renders
them. v1 renders supported rules but cannot author any.

## Acceptance

- [ ] Each supported rule type can be created, edited, reordered and deleted.
- [ ] Rules paint immediately and match Excel on the fixture corpus.
- [ ] Format changes and rule edits are one undo step each.
- [ ] Rule XML round-trips and opens correctly in Excel.
- [ ] The known deviation D-5 is narrowed or removed for authored types.

## Notes

The core rule types can start before V2-V-03; the ticket cannot finish until
the remaining rule types render, because authoring a rule that never paints
would be a false promise.
