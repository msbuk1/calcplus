# V2-F-08 — Hide/unhide rows and columns

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1 pw (provisional)
- **Prerequisites:** none
- **Blocks:** V2-F-10, V2-F-19
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §1, §8

## Scope

Author hidden rows and columns: Hide and Unhide from the context menu and
ribbon, unhide a selection, and unhide all. Imported hidden state already
renders and round-trips; this ticket adds the commands.

## Acceptance

- [ ] Hidden rows and columns collapse out of the geometry.
- [ ] Unhide restores the selected rows or columns; Unhide All restores all.
- [ ] The hidden state round-trips and opens correctly in Excel.
- [ ] Hide and unhide are one undo step each.
- [ ] Filter-hidden rows stay distinct from author-hidden rows.

## Notes

Grouping and outline (V2-F-10) and custom views (V2-F-19) depend on this
state.
