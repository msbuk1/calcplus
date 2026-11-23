# V2-F-06 — Entry AutoComplete

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 0.5–1 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §1

## Scope

Suggest existing text values from the same column while the user types in a
cell or the formula bar. Enter or Tab accepts the suggestion, Delete removes
it, and Esc keeps the typed text. Do not suggest for formulas, numbers or
dates, and do not interfere with IME composition.

## Acceptance

- [ ] Typing a prefix offers matching column values in Excel's order.
- [ ] Enter, Tab and Delete behave as Excel does; Esc cancels the suggestion.
- [ ] Formulas, numbers and dates are not completed.
- [ ] IME input is unaffected.
- [ ] The feature can be turned off in options.

## Notes

The list comes from the column's existing text values, case-insensitively,
with the first match highlighted.
