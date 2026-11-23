# V2-F-05 — Fill series

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §1

## Scope

Add the Fill Series dialog and drag options: linear, growth and date series,
step and stop values, and custom lists. The v1 fill handle keeps copy and
relative-reference adjustment; this ticket adds the series behaviours.

## Acceptance

- [ ] Linear, growth and date series match Excel's values and formatting.
- [ ] Custom lists fill in their declared order and wrap as Excel does.
- [ ] Drag options (copy, fill series, fill formats, fill without formatting)
      behave as Excel does.
- [ ] Each fill is one undo step; formulas adjust relatively.
- [ ] The result round-trips and opens in Excel without repair.

## Notes

Series detection during a drag should reuse the same planner as the dialog so
the two paths cannot drift.
