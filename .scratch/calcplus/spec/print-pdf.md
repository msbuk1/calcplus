# calcplus v1 — print and PDF

Resolved by [21 — Print and PDF pipeline](../issues/21-print-pdf.md) on 2026-11-13.

## 1. Promise

v1 prints and exports the **print area**, with **orientation**, **fit-to-width** and **margins**, and the PDF matches the print preview. Headers/footers, print titles, manual page breaks, multi-page scaling and page order are v2.

## 2. Pipeline

One pure planner, three renderers:

```
Workbook model + print settings
            │
            ▼
      PrintPlanner  ──►  PrintPlan
                          ├─ page boxes in points
                          ├─ page → row/column bands
                          ├─ uniform scale
                          └─ vector display list
            │
    ┌───────┼───────────────────────┐
    ▼       ▼                       ▼
  SVG    PDFKit                 300 DPI canvas
preview  vector PDF            font/script fallback
window.print()
```

- **`PrintPlanner`** is pure and testable: it takes the model, the print area and the page setup, and returns a `PrintPlan`.
- **`PrintPlan`** is the single source of truth for preview, paper and PDF. All three renderers consume it.
- **SVG** renders one element per page for preview and `window.print()`. Vector text uses the browser's fonts, geometry is exact, and it does not depend on CSS fragmentation.
- **PDFKit** writes the vector PDF download. It gives exact page boxes, subsetted embedded fonts, positioned text and images.
- **300 DPI canvas** is the fallback for pages whose fonts or scripts cannot be embedded. It renders from the same display list, so content is never wrong.

## 3. Why this pipeline

- CSS cell grids split rows unpredictably; a `break-inside: avoid` row still split in Chromium. Breaks must be computed.
- SVG shapes print even when Chrome's "Background graphics" is off, where CSS backgrounds are suppressed.
- PDFKit is MIT, actively maintained, and produced a ~195 KB vector PDF for 20 A4 pages versus ~16 MB for 300 DPI raster. Its text is selectable.
- `pdf-lib`'s last release was 2021 and `jsPDF` is weaker for positioned text and font subsetting.

## 4. Page setup

- **Print area:** the authored range; if none, the used range.
- **Orientation:** portrait or landscape.
- **Fit-to-width:** scale all columns to one page width; the row count follows.
- **Margins:** top, bottom, left, right in points.
- **Scaling:** uniform; no independent X/Y scale in v1.
- **Page count:** computed from the plan, not by the browser.
- **Gridlines and headings:** follow the sheet’s **print** settings (`printOptions.gridLines` / `gridLinesSet` / `headings`); the view-gridlines state never drives print output (research/21).
- **Multiple print areas:** each defined print area prints as its own page set, matching Excel (research/21).
- **Font subsets are renamed at export** (for example `CALP+Selawik`) so no modified font ships under a reserved name (licensing.md); PDFKit’s default subsetting is kept.
- **Merged cells:** clipped per page; content appears on the page containing the anchor.
- **Row heights, column widths, borders, fills, fonts and number formats:** rendered from the same cell display records as the grid.
- **Conditional formatting:** rendered from the same CF layer.
- **Dates:** 1900 and 1904 serials format correctly.

## 5. Fidelity limits

- Paper printing is advisory: the browser dialog can override paper size, margins, scale, headers/footers and background graphics. Excel has the same printer caveat.
- PDF substitutes fonts. There is no Aptos metric clone in v1; the closest bundled metric-compatible font is used.
- Safari before 18.2 does not honour `@page`; the preview still works, but paper geometry may differ.
- Excel's exact fit-to-page edge cases need golden workbooks; the plan follows the documented behaviour and is pinned by fixtures.
- Images and charts are v2; the layout reserves their anchors but does not render them.
- Tagged PDF is v2.

## 6. Performance

- The plan is computed once per print/export.
- A 20-page A4 report is expected to stay under 2 seconds to preview and under 3 seconds to PDF on the reference machine.
- Large print areas are paginated in row bands; the renderer never builds a page-sized canvas for the whole sheet.
- The 300 DPI fallback is limited to pages that need it.

## 7. Acceptance criteria

1. Print preview and PDF use the same `PrintPlan` and match each other.
2. Print area, orientation, fit-to-width and margins are honoured.
3. Page count and page boxes match the plan.
4. PDF text is selectable and fonts are embedded or subsetted.
5. Conditional formatting, merges, borders, fills and number formats render as on screen.
6. 1900 and 1904 dates render correctly.
7. The 300 DPI fallback produces readable output for unembeddable fonts/scripts.
8. `window.print()` produces the same pages as the preview in Chromium.
9. The About/help text explains the browser-dialog caveat.
10. A 20-page report meets the performance target.

## 8. V2 extensions

Headers and footers, print titles, manual page breaks, page order, multi-page scaling, images and charts, tagged PDF. These are additions to the plan, not a new pipeline.
