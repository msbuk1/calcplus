# 43 — Freeze Panes does not flip to Unfreeze Panes

Status: resolved
Type: feature parity (discoverability against the Excel 365 oracle)
Blocked by: none (folded into the ticket 42 chrome lane)

## Question

Maintainer question (2026-11-20): "should freeze panes have an unfreeze
option?" The behaviour already exists — pressing Freeze Panes again at the
same split point unfreezes, and pressing at a new cell moves the split
(`src/commands/view.ts` documents the repeat-press rule; the unit and e2e
freeze suites cover it). What is missing is the oracle's discoverability:
in Excel the entry changes its label to **Unfreeze Panes** when a freeze is
active and back to Freeze Panes when it is not.

## Ruled behaviour (2026-11-20, maintainer)

- The Freeze Panes entry shows **"Unfreeze Panes"** whenever the active
  sheet has a frozen split, and "Freeze Panes" otherwise (label and
  accessible name; the icon may keep the freeze glyph as the oracle does).
- The behaviour underneath is unchanged: activating it from the frozen
  state unfreezes (one undoable transaction, as today).
- The enablement stays "a sheet is active".

## Acceptance

1. Freeze a sheet: the ribbon/menu entry reads "Unfreeze Panes".
2. Unfreeze: it reads "Freeze Panes" again.
3. The existing freeze behaviour and suites stay green.

## Resolution (2026-11-20)

A new map field `activeLabel` on ribbon placements carries the checked
label (the generator validates it: non-empty, different, only on commands
with a checked state). The ribbon model picks the label and accessible name
from the checked state. The direction is armed by a fresh-workbook test.
Note: `m2-grid.xlsx` ships a frozen pane, so its entry correctly reads
Unfreeze at open — the old spec pinned the wrong expectation there and now
asserts the real toggle flow instead.
