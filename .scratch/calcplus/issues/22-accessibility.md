# 22 — Accessibility target for the grid

Status: open
Type: grilling
Blocked by: none

## Question

What accessibility bar must v1 meet?

Decide:

- WCAG target (for example 2.2 AA) and which criteria apply to a canvas grid.
- Screen-reader semantics: ARIA `role="grid"`, row/column counts, active cell, selection, headers, and announcements for edits and structural changes.
- Keyboard-only operation: every command reachable, focus order, visible focus, no keyboard traps.
- High-contrast and forced-colours behaviour; zoom and reflow.
- How the canvas renderer and the DOM accessibility mirror stay in sync (ticket 05).
- What is explicitly out of scope for v1, and what must be tested with real screen readers.

Output: the accessibility contract for tickets 10, 11 and 13.
