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

## Progress note — 2026-11-22

All three features are implemented and tested. The work is in the working
tree, uncommitted.

**What was built**

- `edit.dragMove` and `edit.dragInsert` (`src/commands/move.ts`): each drop is
  one `structure.apply` transaction carrying the moved cells, the rewritten
  formulas and the moved merges, comments and hyperlinks, plus one `name.set`
  per rewritten defined name. Undo restores everything in one step.
- The move's reference rules live in the range-transform module
  (`src/model/range-transform.ts`, `planMoveTransform` and friends), next to
  the structural-edit transform it reuses: the same formula scanner, the same
  name and hyperlink handling, the same record shapes and the same block
  messages. Shift-drag insert composes a banded Insert Cells & Shift edit into
  the move (`MoveRangeEdit.prior`), so references to shifted cells follow the
  shift and references to carried cells follow the carry.
- `file.openDropped` (`src/commands/chrome.ts`, the App drop handler): a drop
  runs the File → Open path (`loadPickedFile`), prompts and all. `.csv` opens
  as a one-sheet workbook through the same `openWorkbook` path
  (`src/ooxml/csv-open.ts`); `.xlsx` and `.xlsm` open as before. Any other
  extension shows the refusal message and opens nothing.
- The gestures live in `src/grid/CanvasGrid.tsx`: a press inside a multi-cell
  selection and a drag shows a destination outline that follows the pointer
  (with edge auto-scroll) and drops on release; Esc or a pointer cancel
  abandons the gesture; a press and release without movement keeps the click
  behaviour that existed before. Shift-drag from the selection's outer ring
  shows the gap and the shift direction instead.

**Reference rules (recorded because the contracts do not spell them out)**

- Every reference keeps pointing at the same content. A reference whose whole
  area lies inside the carried rectangle is rewritten to the cells' new
  address; a reference that only partly covers the rectangle keeps its
  authored text, which is what Excel keeps there too (the "sometimes NOT
  update" behaviour people report when dragging columns).
- A moved formula keeps its text when its targets did not move (Microsoft's
  "when you move a formula the references stay the same") and follows its
  targets when they moved with it. Both rules are one rule: the text changes
  exactly when the referenced cells moved.
- Transform or block: a merge the move would split is blocked with Excel's
  "Cannot change part of a merged cell."; a shared or array formula that would
  move or change is blocked; a formula or name with a construct the transform
  cannot rewrite safely blocks when it can reference the moved sheet. A move
  that would strand a chart series, pivot source or pivot output blocks with
  an explanation instead of leaving it stale (v1 does not rewrite Tier B parts
  for moves; insert and delete still do). Comments, threaded comments and
  legacy note anchors follow their cells.
- Overwritten destination cells are replaced; references to them keep their
  text, as any cell edit leaves them. A merge wholly inside the overwritten
  rectangle is dropped (the cells are gone); one that only partly overlaps it
  blocks.
- CSV is a value channel (file-compat §11): fields come in as numbers or text,
  never as formulas. The sheet takes the file's stem.

**Checks run (all on this PC)**

- `npm.cmd run test`: 98 files, 1506 passed (14 new in
  `tests/ticket49-drag-commands.test.ts`, covering both commands: moves,
  reference rewriting, name rewrites, the blocks, comment and hyperlink moves,
  one-step undo and the no-op drops).
- `npm.cmd run typecheck`, `npm.cmd run check`, `npm.cmd run build`: green.
- Playwright: `tests/e2e/drag-move.spec.ts` (9 tests: both gestures, the
  outline, Esc, the source drop, one-step undo, the three drop kinds and the
  unsaved-changes prompt) plus grid, grid-interaction, selection-autoscroll,
  fill-handle, feature-gaps, freeze-scroll, format-dialog, copy-marquee,
  hyperlink-comments and smoke — all green.
- `npm.cmd run inventory` is green and `docs/command-inventory.md` is
  regenerated; the three entries read "Implemented".
- ESLint and `prettier --check` pass on every changed file. The repo-wide
  `npm run format:check` fails on 180 files on this PC, including files this
  work never touched: the working tree has CRLF endings and prettier's default
  expects LF. That predates this work and is left alone.
