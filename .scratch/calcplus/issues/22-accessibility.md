# 22 — Accessibility target for the grid

Status: resolved
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

## Answer

The contract is in [accessibility.md](../spec/accessibility.md).

**Target:** WCAG 2.2 AA for the app interface and the DOM accessibility mirror. The canvas is `aria-hidden`; the mirror is the conformant representation. User-authored cell and conditional formatting is a content caveat: the value is always exposed, but its contrast and colour-only meaning are the author's.

**Screen reader:** `role="grid"` with `aria-rowcount`/`aria-colcount`, absolute `aria-rowindex`/`aria-colindex`, `columnheader`/`rowheader`, merged spans, roving tabindex with a persistent active cell, `aria-selected` for the selection, and a polite live region for edits, structural changes, sort, filter, save and errors. Only the visible range plus the active cell is mirrored.

**Keyboard:** every command reachable through the ribbon, menus, context menu or a shortcut. F6 cycles zones (the grid uses Tab for cells, so F6 is the documented exit), Esc returns to the grid, dialogs trap focus intentionally, and every drag has a keyboard alternative.

**Focus and contrast:** 2 px focus ring at 3:1, never obscured; forced-colours mode repaints the canvas with system colours; 200% browser zoom and 10–400% app zoom; chrome reflows below 900 px. Chrome controls meet 24 × 24; grid cells and resize gutters use the essential two-dimensional exceptions and have keyboard alternatives.

**Testing:** axe-core and Playwright keyboard/live-region/zoom tests in CI; manual NVDA/Chrome, NVDA/Firefox, JAWS/Chrome and VoiceOver/Safari before the grid milestone is accepted. If the mirror fails, the DOM-cell renderer fallback is built behind `RenderFrame` and the milestone is blocked without it.

**Out of scope:** mobile screen readers, braille, sonification, AAA, dark mode, tagged PDF, charts/images, split panes, localized announcements.
