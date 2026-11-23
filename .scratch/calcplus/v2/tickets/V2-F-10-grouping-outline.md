# V2-F-10 — Grouping and outline

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** V2-F-08
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §5

## Scope

Group and ungroup rows and columns, show outline levels, expand and collapse
groups, and edit the outline settings (summary rows/columns, automatic
outline). Imported outline parts already round-trip; this ticket adds the
commands and the outline bar.

## Acceptance

- [ ] Group and ungroup work from the ribbon and the context menu.
- [ ] Collapse hides the grouped rows or columns; expand restores them.
- [ ] Outline levels and settings round-trip and open correctly in Excel.
- [ ] Each operation is one undo step.
- [ ] Collapsed state stays distinct from filter-hidden rows.

## Notes

Collapse uses the hide/unhide state from V2-F-08, which is why that ticket
comes first.
