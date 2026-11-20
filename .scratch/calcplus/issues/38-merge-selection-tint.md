# 38 — After Merge Cells the first cell keeps a selection tint

Status: resolved
Type: defect (selection paint for merged ranges)
Blocked by: none (same lane as ticket 37)

## Question

Reported by the maintainer with a capture (2026-11-20): run **Merge Cells**
over a range, then look at the result. The merged range is selected — it
draws the selection border — but the **first original cell also keeps a
selection tint**, so the merge reads as "the first cell is still selected as
well as the merge range". Two selection indications appear where there must
be one.

The merge authoring commands are new (`src/model/merge.ts` and
`src/commands/merge.ts`, landed in the chrome/data-tools sweep `4051f8b`), so
the bug is either in the selection the merge command returns or in how
`paintSelection` in `src/grid/canvas-renderer.ts` handles a merged range
(ticket 31's merge-aware hole in `fillRectExcept`).

## Ruled behaviour

- A merged cell is **one cell**. A selection containing it renders like any
  other selection of its size: one range border, the selection tint across
  the range, and the ticket-31 rule that the active cell paints unfilled
  inside a multi-cell range.
- No pre-merge residue: after the merge, no paint may show a sub-cell of the
  merged range as separately selected. The pre-merge cell geometry must not
  appear in the selection paint at all.
- The selection state after Merge Cells is the merged range (its top-left is
  the active cell), and the name box shows the merged range.

## Acceptance

1. Merge a range: exactly one selection indication (the merged rect).
2. A multi-cell selection that contains a merged cell tints once and punches
   one hole (the active cell), never a pre-merge cell rect.
3. The ticket-31 pixel assertions extend to merged ranges (the existing
   unfilled-active-cell pixel test shows the shape to follow).
4. One undo of the merge restores the old cells and their selection exactly.

## Resolution (2026-11-20)

Two-part root cause, both fixed. The selection state: a press inside a merged
cell collapsed to the merge anchor, so the paint tinted the anchor sub-cell
while the border covered the merge — exactly the capture. The merge command
also returned no `selectionAfter`. Now the merge command returns the merged
range (top-left active, name box shows the range) and a press inside a merge
selects the whole merged range. The paint: a collapsed selection inside a
merge now expands to the merged rectangle before painting, so no pre-merge
sub-cell geometry can appear. Ticket-31 behaviour is unchanged elsewhere.

Follow-ups: keyboard landings (Enter/Tab/Home) on a merge still leave an
anchored collapsed selection in state (the paint shows it as one cell) —
they should select the whole merge like a press does. Undo restores the same
selected range but as the command's selection shape (`extent: range`) rather
than the literal pre-merge drag shape; restoring the literal object would
need view selection in transactions.

Tests: `tests/m5-merge-commands.test.ts`, `tests/e2e/grid.spec.ts` (merged
pixel assertions), and the Format Cells dialog label assertions updated for
the merged-range target.
