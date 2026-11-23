# V2-F-07 — Discontiguous selection

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none (soft: V2-F-04)
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §1

## Scope

Support Ctrl+click and Ctrl+drag to add, remove and extend multiple ranges.
Paint every range with the active cell distinct. Copy, cut, clear, format and
drag operations apply to all ranges in Excel's order.

## Acceptance

- [ ] Ctrl+click and Ctrl+drag build and edit a multi-range selection.
- [ ] Every range paints correctly, including across freeze panes.
- [ ] Copy and paste preserve Excel's range order; one undo restores.
- [ ] Clear and formatting apply to every range.
- [ ] Keyboard selection and the status bar reflect the whole selection.

## Notes

Selection state currently assumes one rectangle. This ticket widens it without
changing single-range behaviour. Multi-range paste in V2-F-04 uses it.
