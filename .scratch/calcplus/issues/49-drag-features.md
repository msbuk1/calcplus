# 49 — Drag features: move, Shift-drag insert and file drop

Status: open
Type: feature (three v1 IN features; split from ticket 36 item 4)
Blocked by: none

## Question

Ticket 36 item 4 recorded three version-one features confirmed missing after
the chrome and data-tools sweep, all mouse drags. The grid's drag handling is
the fill handle, the Format Painter and the range autoscroll; none of these
three exists. All three are promised IN in the
[feature matrix](../calcplus/spec/feature-matrix.md) and all three are still
`implemented: false` in `src/commands/command-inventory.json`.

| Feature | Feature-matrix row | Command id |
|---|---|---|
| Drag-and-drop move | §1 Grid and editing — "Dropping moves cells and rewrites references" | `edit.dragMove` |
| Shift-drag insert | §1 Grid and editing — "Shift-drag inserts and shifts cells" | `edit.dragInsert` |
| Drop a file to open | §7 File I/O and printing — "Dropping `.xlsx`, `.xlsm` or `.csv` opens the file; unsupported types show an error." | `file.openDropped` |

## Ruled behaviour

- **Drag-and-drop move (`edit.dragMove`).** Press inside the selection and
  drag: a destination outline follows the pointer and the drop moves the
  range there. Dropping moves cells and rewrites references through the
  structural-edit range transform (transform or block, never a silently
  stale anchor). One undo transaction. Esc cancels the drag. A drop back
  on the source does nothing.
- **Shift-drag insert (`edit.dragInsert`).** Shift-drag on the selection
  edge inserts and shifts cells down or right, with the same reference
  rewriting and one undo transaction.
- **File drop (`file.openDropped`).** Dropping a supported file on the app
  opens it through the same path as File — Open, including the unsaved-
  changes and external-change prompts. `.xlsx`, `.xlsm` and `.csv` are
  supported; an unsupported type shows an error and opens nothing.

## Sources

- [Feature matrix](../calcplus/spec/feature-matrix.md) §1 and §7 rows above.
- [Ticket 36](36-followup-chrome-data-tools.md) item 4, which this ticket
  takes over.
- [Command system](../calcplus/spec/command-system.md) for the undo and
  transaction rules; the M5 structural-edit range transform for rewriting.
- `src/grid/CanvasGrid.tsx`: the pointer-drag pattern and the range
  autoscroll from the ticket 36 sweep are the place to start.

## Acceptance

1. Drag-move moves the range, rewrites references and is one undo
   transaction; Esc cancels; a drop on the source changes nothing.
2. Shift-drag inserts and shifts cells down or right with the same
   reference rewriting and one undo transaction.
3. Dropping `.xlsx`, `.xlsm` or `.csv` opens the file; an unsupported type
   shows an error and opens nothing.
4. `edit.dragMove`, `edit.dragInsert` and `file.openDropped` read
   `implemented: true`; `npm run inventory` is green and
   `docs/command-inventory.md` is regenerated.
5. Unit and integration tests cover the move and insert commands; e2e
   covers the gestures and the drop; existing suites stay green.

## Placement

Split from ticket 36 item 4 on 2026-11-22, per that ticket's placement
rule ("item 4 is feature work and should be split out before it starts").
This is v1 scope: the M8 release gate requires every feature-matrix IN
item to be present and meet its acceptance criterion.
