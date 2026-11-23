# V2-F-20 — Page-break preview

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1 pw (provisional)
- **Prerequisites:** V2-F-21
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §6; known
  deviation D-3

## Scope

Add Page Break Preview: show automatic and manual page breaks on the sheet,
drag breaks to new positions, insert and remove a break, and reset all breaks.
The print preview and PDF export honour the result.

## Acceptance

- [ ] Page Break Preview shows the current breaks and page numbers.
- [ ] Dragging a break moves it; Insert and Remove work from the menu.
- [ ] Reset All Breaks restores automatic pagination.
- [ ] The print preview and PDF export honour manual breaks.
- [ ] Each change is one undo step; breaks round-trip.

## Notes

Manual breaks are part of V2-F-21; this ticket is the visible editing surface.
The D-3 deviation for manual breaks is removed when both land.
