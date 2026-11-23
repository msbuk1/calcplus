# V2-F-18 — Split view

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** none
- **Blocks:** V2-F-19
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §6

## Scope

Split the sheet into two or four independently scrolling panes at the active
cell. Selection and the formula bar stay in sync across panes; freeze panes
and split view interact as they do in Excel.

## Acceptance

- [ ] Split at the active cell creates two or four panes.
- [ ] Each pane scrolls independently; the selection is visible in all.
- [ ] Editing in one pane updates the others.
- [ ] Removing the split restores one pane without changing the workbook.
- [ ] Freeze panes and split view do not produce conflicting dividers.

## Notes

Split view is window state, not document state. Custom views (V2-F-19) can
store it if Excel does.
