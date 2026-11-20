# 41 — Freeze panes: scrolling up paints into the frozen row strip

Status: open
Type: defect (freeze-pane rendering)
Blocked by: none

## Question

Reported by the maintainer with a capture (2026-11-20): with frozen rows,
scrolling **up** past the boundary paints the incoming row's number and
content into the frozen strip — the capture shows two row numbers drawn on
top of each other in the gutter and row text bleeding under the frozen
header. Frozen **columns** do not show the problem: horizontal scrolling
clips correctly. So the vertical frozen band's clip is wrong in one
direction only.

## Ruled behaviour

- The frozen row strip is a strict clip band: nothing that scrolls beneath
  it is ever painted inside it — no row numbers, no cell text, no fills.
- Symmetry with frozen columns: both axes clip identically in both scroll
  directions.
- The row numbers in the frozen strip belong to the frozen rows only.

## Acceptance

1. Scroll up past the frozen boundary: the frozen strip shows only its own
   rows (pixel or geometry assertion — the cf-paint freeze-divider tests
   show the shape to follow).
2. Scrolling down, left and right stay correct (no regressions).
3. The existing freeze suites (`tests/m5-freeze-panes.test.ts` and the
   e2e freeze specs) stay green.
