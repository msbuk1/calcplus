# 21 — Print and PDF pipeline

Status: resolved
Type: research
Blocked by: none

## Question

How should calcplus produce Excel-like printed pages and PDFs in a browser?

Evaluate:

- Browser print via a dedicated print stylesheet/DOM, versus rendering to canvas and emitting PDF with a library (`pdf-lib`, `jsPDF`, others). A headless server is not an option: no backend.
- Honouring print area, orientation, fit-to-width, margins, scaling, and pagination across pages.
- Fidelity of fonts, borders, fills, merged cells, row heights and images (images are v2, but layout should not break).
- Headers/footers and print titles as a v2 extension.
- Browser differences (Chrome/Firefox/Safari) and PDF output consistency.
- Memory and performance on large print areas.

Recommend an architecture and list the fidelity limits of the chosen path.

## Answer

**Recommendation:** own pagination in a pure `PrintPlanner` that emits a `PrintPlan` (page boxes in points, page→row/column bands, uniform scale, vector display list), then render that one plan three ways: inline SVG per page for preview and `window.print()`, PDFKit for a vector PDF download, and a 300 DPI canvas raster only for pages whose fonts/scripts cannot be embedded.

- **SVG page boxes, not a CSS cell grid or canvas, for print.** Vector text with the browser's own fonts, one element per page, exact geometry, and no dependence on CSS fragmentation. Pre-sized boxes with `@page { size: W H; margin: 0 }` produced exact page counts; a `break-inside: avoid` table row still split in Chromium 155, so breaks must be computed. SVG shapes also print with Chrome's "Background graphics" off, where CSS backgrounds are suppressed.
- **PDFKit for export, not pdf-lib or jsPDF.** MIT, actively released (0.20.2, 2026-08-30), exact page boxes, subsetted embedded fonts, positioned text, images, tagged-PDF path. pdf-lib's last release was 2021-11-06. A vector PDF of 20 A4 pages measured ~195 KB (browser print-to-PDF) versus ~16 MB for 300 DPI raster, and its text is selectable.
- **Raster is the fallback, not the default.** Pages needing an unembeddable font or uncertified script (Arabic/Hebrew/Indic/Thai) render at 300 DPI from the same display list, so content is never wrong, at ~800 KB/page.
- **Trade-offs:** paper printing is advisory because the browser dialog can override paper, margins, scale, headers/footers and backgrounds (Excel has the same printer caveat); PDF substitutes fonts (no Aptos metric clone); Safari <18.2 does not honour `@page`; Excel's exact fit-to-page edge cases need golden workbooks.
- **V2 is incremental:** headers/footers, print titles, manual breaks and page order are additions to the plan, not a new pipeline. Do not use `@page` margin boxes for headers/footers (Chrome-only).

Full research: [21-print-pdf.md](../research/21-print-pdf.md)
