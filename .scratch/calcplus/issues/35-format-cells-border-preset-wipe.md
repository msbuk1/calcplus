# 35 — Format Cells border presets with no edge change wipe the cell style

Status: resolved
Type: defect (data loss in an accepted feature)
Blocked by: none

## Question

In the Format Cells dialog, choosing Inside on a single cell, or Outline on
interior cells, reports no edge change but writes a default style anyway. That
wipe destroys the cell's font, fill, border, alignment and number format.

Where it lives: the null-patch path in
[`src/commands/format.ts`](../../../src/commands/format.ts), the `buildChange`
and `styleForCellAfter` functions.

Current-state evidence (2026-11-19):

1. The border preset builds a per-cell patch. Inside on a single cell, and
   Outline on an interior cell, produce an empty border patch, so the patch
   callback returns `null` ("no edge for this cell").
2. `buildChange` turns that `null` into style id 0 with `?? 0`, and
   `styleForCellAfter` passes the 0 through. The change then rewrites the cell
   with the default style.
3. The default style has no font, fill, border, alignment or number format, so
   the rewrite is a wipe. The command still reports a change, so the dialog
   reads as successful.

Why it matters: a user can lose all formatting on a cell by opening a dialog
and picking a preset that does nothing visible. This is data loss in a feature
milestone 4 marked accepted.

## Ruled direction

When a preset produces no edge change, write nothing. The cell keeps its
existing style untouched. The dialog still reports the outcome honestly.

## Acceptance shape

1. Inside on a single cell writes no style change and the cell keeps font,
   fill, border, alignment and number format.
2. Outline on interior cells does the same.
3. A preset that does change edges still writes only those edges and leaves the
   rest of the style intact.
4. A regression test covers the null-patch path and proves no wipe.
5. The dialog's status text matches what was written.

## Placement

Small bounded fix in `src/commands/format.ts` plus one test. Not milestone 7
work.

## Answer

Pending.

## Resolution (2026-11-20)

The classic wipe (a null patch writing style 0) was already fixed by the
chrome sweep (`4051f8b`). The remaining leak was subtler: a patch that
resolved to the cell's **current appearance** still went through
`findOrAppendStyle`, which can pick a different record with the same look —
rewriting the cell's style id (losing its own record) and opening a spurious
transaction.

Fixed once at the shared choke point (`applyPatchForCell` in
`src/commands/format.ts`): the patched style is compared with the cell's
current record by `styleSignature`, and a match leaves the cell untouched.
This closes the no-op write for every patch command — borders, font, fill,
alignment, number format and toggles — not just the presets.

Accepted limitations (pre-existing, recorded): an applied style write still
creates a record without the source `xf` attributes outside the model (needs
the styles-writer work in `style-table.ts`), and `formatBorderCommand`'s
`preset` type should widen to include `'custom'` in a cleanup.

Tests: `tests/m4c-format-border-noop.test.ts` (red first, green after).
