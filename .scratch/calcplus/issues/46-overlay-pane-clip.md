# 46 — Copy marquee and highlight overlays lack the pane clip

Status: resolved
Type: defect (overlay bleed across freeze panes; family of tickets 41 and 45)
Blocked by: none (the ticket 45 lane owns it)

## Question

Found by the ticket 45 lane (2026-11-20): two overlays still bleed into
frozen strips because they are positioned from unclipped cell rectangles.

1. **The copy marquee** (the dashed paste-destination overlay): its DOM
   rectangle comes straight from the cell rects, so a marquee destination
   scrolled under a frozen strip draws over the strip. It is DOM chrome, not
   canvas ink, so the pixel checks do not see it; the fix needs per-pane
   clipping (a clip-path or per-pane positioning) in `positionMarquee`.
2. **Pointing and Format Painter highlights** (`paintRangeHighlight`): the
   same no-clip geometry can paint a few pixels into the strip.

## Ruled behaviour

- Every overlay clips to its pane exactly like the cells, headers and
  selection paint: the copy marquee and the pointing/painter highlights
  never draw inside a frozen strip.
- Ticket-31 marquee semantics are unchanged (arrows move it, clamps at the
  sheet edges, Enter pastes, Esc cancels).

## Acceptance

1. Move the marquee under a frozen strip: no marquee ink in the strip; the
   debug hit/geometry state stays consistent.
2. Pointing and painter highlights follow the same clip.
3. The copy-marquee and painter suites stay green.

## Resolution (2026-11-20)

The copy marquee keeps one DOM element with the full paste-destination box
(ticket-31 geometry, movement, clamps and the tests' box checks intact) and
clips only its painted dashes to the pane rectangle — applied in both the
render and the per-frame scroll tracking. The pointing and Format Painter
highlights run inside the same paintInPane clip as the selection paint.
Both are proved by tests that fail without their fix. The pre-existing
freeze-filter pixel-read flake is stabilised with a two-frame paint settle
(test-only; reproduced independently of this change).
