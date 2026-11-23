# V2-S-05 — Iterative and manual calculation

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [formula semantics](../../spec/formula-semantics.md) §6, §8

## Scope

Add the workbook calculation settings: iterative calculation (enable, maximum
iterations, maximum change) and manual calculation mode with Calculate Now,
Calculate Sheet and F9. Settings round-trip with the workbook.

## Acceptance

- [ ] With iteration on, a circular reference converges to Excel's values
      within the declared limits.
- [ ] With iteration off, the v1 `#CIRC!` warning and error remain.
- [ ] Manual mode defers recalculation until Calculate Now, Calculate Sheet or
      F9; the UI shows the pending state.
- [ ] Calculation settings round-trip and are honoured on open.
- [ ] Excel 365 opens the saved file without a repair prompt.

## Notes

The v1 deviation (show `#CIRC!` instead of iterating) is replaced only when
iteration is enabled. Update the known deviations and pins with this ticket.
