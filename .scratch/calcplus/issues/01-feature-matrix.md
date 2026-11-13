# 01 — Define the v1 feature matrix and acceptance criteria

Status: open
Type: grilling
Blocked by: none

## Question

What exactly is in v1?

Produce the authoritative feature matrix for calcplus v1, grouped by area:

- Grid and editing
- Formatting
- Formulas and names
- Workbook and sheet structure
- File I/O and printing
- View and navigation

For each feature: include/exclude, and if included, a testable acceptance criterion. Classify everything excluded as **v2** (next effort) or **never** (beyond the product). Resolve boundary cases explicitly, including at minimum:

- Basic sort/filter, tables, and data validation — the scope cut defers "data tools" to v2; does that include basic sort/filter?
- Conditional formatting
- Freeze panes, zoom, split view
- Number formats and custom format strings
- Print and PDF export
- Comments/notes
- Find & replace, Go To
- Hyperlinks
- Cell and sheet protection

This ticket defines the scope contract every other ticket hangs on. Record the boundary, not the implementation.
