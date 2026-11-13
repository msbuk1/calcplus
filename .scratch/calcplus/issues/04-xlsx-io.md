# 04 — .xlsx import/export approach

Status: resolved
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

## Answer

**Recommendation: build a preservation-first OOXML layer in-house** (`fflate`/JSZip + `saxes`), not SheetJS CE or ExcelJS as the storage engine. Model the V1 features (values, formulas + cached results, styles, merges, dimensions, names, sheets, comments, hyperlinks, validation, conditional formatting) and keep charts, images, shapes, pivots, external links, macros and unknown parts as opaque parts copied verbatim. No-op save returns the original bytes; only dirty parts are rewritten; style and shared-string tables are append-only.

- SheetJS CE is a data extractor: no cell fonts/fills/borders, validation, conditional formatting or freeze panes, and it dropped charts and external links in testing. SheetJS Pro covers these but is closed-source and quote-only.
- ExcelJS is excellent for data/styles/images/validation/conditional formatting, but drops charts, external links, pivots, VBA and custom properties; its streaming writer is Node-only.
- xlsx-populate preserves every package part but intentionally drops cached formula results; useful as a reference/fork, not as-is.
- Trade-off: the I/O layer plus the range-transform module for structural edits is the largest V1 engineering item. Charts and pivots will be preserved but not rendered until v2; anything unmodelled must be preserved, never silently dropped. If build capacity fails, SheetJS Pro (Pro Edit + styling/image/chart builds + pivot/encryption add-ons) is the only complete commercial fallback.

Full research: [04-xlsx-io.md](../research/04-xlsx-io.md)
