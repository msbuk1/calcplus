# 01 — Define the v1 feature matrix and acceptance criteria

Status: resolved
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

## Answer

The v1 scope contract is fixed in [feature-matrix.md](../spec/feature-matrix.md).

- **IN:** core grid editing (no Paste Special or fill series), range sort + AutoFilter on plain ranges, full cell and number formatting, render-only conditional formatting/comments/hyperlinks, the core view set, Find & Replace + Go To, multi-sheet insert/delete/rename, `.xlsx`/CSV round-trip, `.xlsm` macros preserved, basic print/PDF.
- **Read-only fidelity:** defined names, theme colours and named styles, protection (preserved, not enforced), hidden rows/columns/sheets, imported validation/tables/pivots.
- **V2 and NEVER** are classified per area in the matrix.

Open decisions graduated: [18 — Number-format engine](18-number-format-engine.md), [19 — Sort and AutoFilter semantics](19-sort-autofilter.md), [20 — Conditional-formatting rendering](20-conditional-formatting.md), [21 — Print and PDF pipeline](21-print-pdf.md), [22 — Accessibility target for the grid](22-accessibility.md).
