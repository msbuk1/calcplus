# V2-F-09 — Sheet management and tab colours

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 3–4 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §5; ADR 0008

## Scope

Add sheet reorder by drag and by the Move or Copy dialog, move and copy to
another position or workbook position, hide and unhide sheets, and tab colours.
Insert, delete and rename already work in v1; sheet delete stays terminal as
ADR 0008 requires.

## Acceptance

- [ ] Dragging a tab reorders sheets; the Move or Copy dialog does the same
      with an index.
- [ ] Copy duplicates the sheet and its formulas with Excel's reference rules.
- [ ] Hide and unhide work; at least one sheet stays visible.
- [ ] Tab colours render and round-trip.
- [ ] Each operation is undoable except the terminal sheet delete.
- [ ] The workbook opens in Excel without repair.

## Notes

Imported order, hidden state and tab colours already round-trip; this ticket
adds authoring. Coordinate with V2-S-01 if a table is copied.
