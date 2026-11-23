# V2-V-03 — Remaining conditional-format rule rendering

- **Stream:** Visual
- **Status:** proposed
- **Effort:** 3–5 pw (provisional)
- **Prerequisites:** none
- **Blocks:** V2-F-14 (for these rule types)
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §3; known
  deviation D-5

## Scope

Render the rule types v1 preserves but does not paint: formula-based rules,
above-average and below-average rules, blank and error rules, and
formula-threshold rules. Keep the v1 support for cell-value, text, date,
top/bottom, duplicate/unique, data-bar, colour-scale and icon-set rules.

## Acceptance

- [ ] Each rule type paints on the fixture corpus as Excel does.
- [ ] Rule priority and stop-if-true still resolve correctly.
- [ ] Structural edits transform the rule ranges as v1 does.
- [ ] Rule XML stays untouched when the rule is not edited.
- [ ] The D-5 deviation is narrowed to authoring only.

## Notes

Evaluation runs in the existing conditional-format engine. The new part is the
predicate evaluation and the paint for these rule types. V2-F-14 can then
author them.
