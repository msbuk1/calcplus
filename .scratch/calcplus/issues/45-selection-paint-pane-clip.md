# 45 — Active cell box and fill handle paint outside the pane clip

Status: resolved
Type: defect (selection paint across freeze panes; family of ticket 41)
Blocked by: none (the ticket 41 lane owns it)

## Question

Found by the ticket 41 lane (2026-11-20): the active-cell box and the fill
handle are painted with **no pane clip**. Repro: freeze two rows, select A1,
wheel down — the blue selection box is drawn inside the frozen strip, over
the frozen rows' cells. The same would apply to the fill handle and any
selection paint that scrolls under a frozen strip.

## Ruled behaviour

- Every selection overlay clips to the pane it belongs to, exactly like the
  cells and the row/column headers (ticket 41's rule): a selection scrolled
  under a frozen strip never paints inside it.
- Ticket-31 semantics are unchanged: the active cell paints unfilled inside
  a multi-cell range, the fill handle sits at the range corner — both
  clipped to the visible pane.

## Acceptance

1. Freeze rows, select a cell, scroll it under the strip: no selection ink
   inside the frozen band (pixel or geometry assertion, the freeze-scroll
   spec's signature pattern).
2. The fill handle follows the same clip.
3. A selection inside the visible area paints exactly as today (ticket-31
   pixel tests stay green).

## Resolution (2026-11-22)

Fixed in commit `d8939fc`: the active-cell box, range tint and border, the
tint's active-cell hole and the fill handle (ink and hit target) now paint
inside their pane's clip, exactly like the cells and headers. Ticket-31
semantics and merged selections render unchanged.

Verified 2026-11-22 on the maintainer's PC: the e2e freeze suite
(`tests/e2e/freeze-scroll.spec.ts`) passes **8/8**, including the ticket-45
repro ("a selection scrolled under the frozen strip leaves no ink in it",
acceptance 1) and the fill-handle outline clip (acceptance 2). The
ticket-31 selection tests are green in the same run (acceptance 3), and the
full unit suite (97 files, 1493 tests) passes.

The status line stayed open after the fix landed; this record closes it.
The parent family is ticket 41.
