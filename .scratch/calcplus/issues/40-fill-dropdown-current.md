# 40 — Ribbon Fill drop-down does not show the current fill

Status: resolved
Type: defect (visual parity against the Excel 365 oracle)
Blocked by: none

## Question

Reported by the maintainer with a capture (2026-11-20): open the ribbon's
**Fill colour** drop-down while a cell with a fill is selected. The
**Current** swatch at the bottom shows the slashed "No Fill" swatch instead
of the cell's fill colour. The capture shows a red-filled cell selected
while Current reads "No Fill". The same happens on **merged cells**.

In the Excel 365 oracle the colour drop-down's Current swatch previews the
active cell's actual fill (or "No Fill" when it has none).

## Ruled behaviour

- **Current** previews the active cell's fill colour (ticket 31's active
  cell — the gesture origin), or the No Fill swatch when the cell has none.
- A merged cell counts as one cell: the preview shows the merged cell's
  fill (its anchor's style). A selection whose active cell is inside a merge
  reads the merge's fill.
- After Apply (or any fill change), the Current swatch reflects the new
  state without reopening the drop-down.
- The New swatch keeps showing the colour about to be applied, unchanged.

## Acceptance

1. Single cell with a fill: Current shows that fill.
2. No fill on the active cell: Current shows the No Fill swatch.
3. Merged cell with a fill: Current shows that fill.
4. After applying a fill, Current updates.
5. The ribbon's other style readouts stay consistent (same active-cell
   source); no command or model changes.

## Resolution (2026-11-20)

Root cause was the command map: the Fill Colour placement carried no
`controlId`, so the ribbon always passed an empty value to the colour picker
and empty reads as No Fill. The font colour control was fine because its
placement carried the id. The map now declares `fill-color`, and one helper
(`src/chrome/ribbon-readouts.ts`) resolves all five ribbon readouts through
the merge anchor (a merged cell is one cell; a covered cell's hidden style is
ignored). The Current and New swatches gained test ids.

Reading of "updates live" accepted: Apply-mode closes the drop-down (existing
behaviour), and reopening shows the new fill without a reload.

Follow-up recorded: a projection-level `cellAt` that folds merges in would
replace the chrome-side merge resolution; left to the view layer's owner.
