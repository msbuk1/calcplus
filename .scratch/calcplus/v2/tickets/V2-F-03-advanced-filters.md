# V2-F-03 — Advanced, colour and Top 10 filters

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §2

## Scope

Extend AutoFilter beyond the v1 value checklist and text/number/date
conditions: filter by cell or font colour, Top 10 (top/bottom, items or
percent), and the custom advanced filter dialog with AND/OR criteria. Clear
filters and keep the filter state in the file.

## Acceptance

- [ ] Colour, Top 10 and advanced conditions hide the right rows.
- [ ] Multiple conditions combine with Excel's AND/OR rules.
- [ ] Clearing restores every row and is one undo step.
- [ ] Filter state round-trips and survives a save/open in Excel.
- [ ] Large ranges filter within the performance budget.

## Notes

The v1 AutoFilter engine and geometry already exist; this ticket adds the
condition types and UI. Filtering by colour needs the style table to expose
cell and font colour.
