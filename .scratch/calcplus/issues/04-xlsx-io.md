# 04 — .xlsx import/export approach

Status: claimed
Type: research
Blocked by: none

## Question

How should calcplus read and write `.xlsx` with high fidelity in the browser?

Evaluate at least SheetJS, ExcelJS, and direct OOXML handling (JSZip + XML), plus any credible alternative. Compare coverage and round-trip fidelity for:

- Cell values, formulas, and cached results
- Styles: fonts, fills, borders, alignment, number formats
- Merged cells, row heights, column widths, hidden rows/columns
- Defined names, sheet metadata, multiple sheets
- Comments, hyperlinks, data validation, conditional formatting
- Charts, images, and shapes (preserve vs drop)
- External links, pivot caches, macros, encryption
- Large-file performance, streaming, memory
- Licensing and browser constraints

Recommend an approach and enumerate what it cannot preserve.
