# 45 — Active cell box and fill handle paint outside the pane clip

Status: open
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
