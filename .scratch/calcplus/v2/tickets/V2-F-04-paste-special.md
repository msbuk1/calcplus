# V2-F-04 — Paste Special

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §1

## Scope

Add the Paste Special dialog and its modes: all, formulas, values, formats,
comments and notes, validation, column widths, transpose, arithmetic
(add/subtract/multiply/divide), skip blanks and paste link where v1 supports
it. The v1 clipboard and undo path stays the base.

## Acceptance

- [ ] Each mode pastes what Excel pastes for values, formulas and formats.
- [ ] Transpose and the arithmetic modes match Excel on a test matrix.
- [ ] Skip blanks and column widths behave as Excel does.
- [ ] A paste is one undo step; paste from the OS clipboard still works.
- [ ] Pasting into a protected sheet respects V2-S-06.

## Notes

Multi-range paste depends on V2-F-07; the single-range dialog ships without
it.
