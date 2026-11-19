# 31 — Selection active cell and copy-marquee interaction parity

Status: resolved
Type: oracle parity
Blocked by: the context-menu submenu work (shared files: `App.tsx`, `Menu.tsx`)

## Question

Two interaction gaps against the Excel 365 oracle, reported by the maintainer
with a capture (2026-11-19):

1. **The active cell of a dragged selection sits at the drag end.** The
   selection draws its last cell as an independently highlighted box. Excel
   keeps the active cell at the **first** cell of the drag (the anchor) and
   renders it as the one *unfilled* cell inside the selected range, with no
   separate outline: the range keeps its border and the fill handle stays at
   the range corner. The wrong end misbehaves everything anchored to "the
   active cell": formula reference insertion, the formula-bar binding, the
   style readouts, the fill source and Ctrl+D/R semantics. Related to ticket
   29 (selection-aware command targeting).

2. **The copy marquee is inert.** Excel's behaviour after Ctrl+C / Ctrl+X: a
   dashed marquee marks the copied range and **becomes the paste destination**;
   the arrow keys move the marquee (not the selection), Enter pastes at the
   marquee position (one transaction; the pasted range becomes selected), Esc
   cancels the copy mode and removes the marquee without pasting. Any other
   navigation dismisses the copy mode. The current app draws a dashed box that
   the arrow keys leave behind.

## Ruled behaviours (2026-11-19, maintainer)

**A. Active cell**
- The active cell is where the selection began: click or drag starts at a cell
  that remains the active cell for the whole gesture; Shift+arrow extends the
  range and the active cell does not move; Ctrl+A selects the sheet with the
  active cell unchanged; collapsing the selection (plain arrow) moves the
  active cell normally.
- Rendering: the active cell paints **unfilled** (the grid background) inside
  the selected range; no independent outline; the range keeps its border and
  the fill handle; the row/column headers shade for the full extent.
- Everything that reads "the active cell" reads this one: the name box, the
  formula bar, in-cell editing, point-and-click reference insertion, style
  readouts, and the command contexts (ticket 29 `activeCell`).

**B. Copy marquee**
- Ctrl+C / Ctrl+X enters copy mode: a dashed marquee marks the source range.
- Arrow keys (and only arrows) move the marquee by one row/column per press;
  the selection does not move. The marquee wraps at the sheet edge? NO — it
  clamps (oracle parity).
- Enter pastes the in-memory snapshot at the marquee position in one
  transaction (values, formulas with adjusted references, formats, merges per
  clipboard.md §6), selects the pasted range, and ends copy mode.
- Esc ends copy mode without pasting; any other key or mouse action that
  changes the selection dismisses copy mode first (then performs its action).
- Cut mode is identical but clears the source on the paste transaction.

Output: the selection model change (`activeCell` semantics in
`src/grid/selection.ts` and its consumers), the renderer change (active-cell
paint), the marquee state (view state in `CanvasGrid`/`App`), the keyboard
path, and tests at all three levels (selection unit tests, command
integration, e2e gesture tests).

## Answer

Pending implementation; the behaviours above are the ruling.
