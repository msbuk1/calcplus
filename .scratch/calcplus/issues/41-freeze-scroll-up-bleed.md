# 41 — Freeze panes: scrolling up paints into the frozen row strip

Status: resolved
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

## Resolution (2026-11-22)

Fixed in commit `a5ab706`: the row-number gutter painted the scrolling
numbers with no clip, so a row half-scrolled under the divider drew its
number over the frozen strip's. The scrolling column headers already
clipped at the divider; the row headers now mirror them, so both axes clip
identically in both directions.

Verified 2026-11-22 on the maintainer's PC: the e2e freeze suite
(`tests/e2e/freeze-scroll.spec.ts`) passes **8/8**, covering every vertical
and horizontal scroll offset for both strips (acceptance 1 and 2) and the
selection, marquee, Format Painter and fill-outline clips that belong to
the same family. `tests/m5-freeze-panes.test.ts` and the full unit suite
(97 files, 1493 tests) are green (acceptance 3).

The status line stayed open after the fix landed; this record closes it.
The companion selection-paint defect is ticket 45.
