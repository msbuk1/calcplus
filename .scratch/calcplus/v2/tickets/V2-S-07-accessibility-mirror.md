# V2-S-07 — Accessibility mirror and WCAG claim

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 4–7 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [accessibility](../../spec/accessibility.md)

## Scope

Build the DOM accessibility mirror the v1 contract describes: `role="grid"`
with absolute counts and indices, headers, merged spans, roving tabindex, a
persistent active cell, `aria-selected`, a polite live region, keyboard access
to every command, F6 zone exit, a 2 px 3:1 focus ring, forced-colours palette,
200% zoom and 24 × 24 chrome targets. The canvas stays `aria-hidden`.

## Acceptance

- [ ] A screen reader reads the grid, headers, selection and edits.
- [ ] Every command is reachable by keyboard, with visible focus.
- [ ] axe-core and Playwright accessibility runs pass locally.
- [ ] The four-combination manual assistive-technology matrix passes.
- [ ] The WCAG 2.2 AA claim is recorded in the docs and release notes.

## Notes

v1 makes no accessibility claim (limitation L-4). This ticket replaces L-4
with the claim and its evidence. It does not change the visual design.
