# V2-F-12 — Hyperlink authoring

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1 pw (provisional)
- **Prerequisites:** V2-S-06 (for protected sheets)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §8; known
  deviation D-9

## Scope

Insert, edit and remove hyperlinks: internal cell and range targets, external
URLs, tooltips and display text. v1 renders and follows links but cannot author
them. Structural edits keep the v1 rewrite behaviour (D-9).

## Acceptance

- [ ] The Insert and Edit Hyperlink dialogs create and change both link kinds.
- [ ] Removing a link keeps the display text.
- [ ] Internal links follow their targets; a removed target reads `#REF!` per
      D-9.
- [ ] Links respect sheet protection and are one undo step.
- [ ] The hyperlink XML round-trips and opens correctly in Excel.

## Notes

The Excel-oracle capture for D-9 stays the reference for the rewrite rule.
