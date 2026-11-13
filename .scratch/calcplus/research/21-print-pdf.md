# Print and PDF pipeline — research

- **Ticket:** 21 — Print and PDF pipeline
- **Date:** 2026-11-13
- **Status:** complete
- **Scope:** how calcplus produces Excel-like printed pages and PDFs in a browser, with no backend. Charts and images are v2; the layout must leave room for them.

## Recommendation

**Build one pure print model, then render it twice: inline SVG for preview and paper, and vector PDF via PDFKit for the download button. Use a per-page canvas raster as the PDF fallback for fonts and scripts we cannot embed.**

Concretely:

1. **`PrintPlanner` (pure TypeScript)** owns page setup and pagination. It reads the sheet projection, `pageSetup`/`pageMargins`/`printOptions`/`pageSetUpPr`, the `_xlnm.Print_Area` and `_xlnm.Print_Titles` names, and emits a `PrintPlan`: page boxes in points, the row/column band each page covers, a uniform scale, and a display list of vector primitives (rect, line, text, image placeholder). Nothing in it touches the DOM.
2. **Preview and printing** render the plan as **one inline SVG per page** inside a hidden print root. SVG is the print medium, not a CSS cell grid: text stays vector and uses the browser's own fonts and shaping, fills and borders are shapes, and one page is one element. `@page` gets an explicit two-value `size` and `margin: 0`; calcplus bakes the Excel margins into the page geometry. This avoids CSS fragmentation entirely, which is the part browsers control and where spreadsheets break.
3. **PDF export** renders the same plan with **PDFKit** (MIT, actively released) to exact PDF page boxes. Rectangles, lines and text are vector; fonts are bundled metric-compatible faces (Carlito for Calibri, Liberation for Arial/Times) and subsetted. A page that needs a font or script we cannot embed is rasterised at 300 DPI from the same display list, so content never becomes wrong.
4. **Do not** rely on the browser's print dialog for the product's PDF. It stays as a secondary path ("Print → Save as PDF"), because it is the highest-fidelity vector output when the user wants it, but it is multi-step, browser-specific and not automatable.

The main trade-off is that print fidelity is ultimately subject to the browser's print dialog (paper size, margins, scale, headers/footers, background graphics), exactly as Excel's own output is subject to the printer driver.[15] PDF export is the deterministic channel, and the preview is generated from the same plan, so "PDF matches the print preview" holds by construction for geometry, content and colour.

## 1. What Excel's print model requires

### 1.1 The settings that matter in v1

Excel stores page setup in the worksheet XML, not in the UI:

- `pageSetup` carries `orientation`, `paperSize`, `scale` (10–400), `fitToWidth`, `fitToHeight`, `pageOrder`, `blackAndWhite`, `draft`, `cellComments`, `errors`, `firstPageNumber`, DPI and custom `paperWidth`/`paperHeight`.[19] `scale` is **overridden** when fit-to-width or fit-to-height is in use.[19]
- `pageSetUpPr@fitToPage` is the switch that enables the Fit to Page option; without it the `scale` value applies.[20]
- `pageMargins` is in **inches**, with header and footer margins separate from the body margins. The schema's own example is the Excel default: left/right 0.7, top/bottom 0.75, header/footer 0.3.[21]
- `printOptions` carries `gridLines` + `gridLinesSet` (both must be true to print gridlines), `headings` (row/column headings), and `horizontalCentered`/`verticalCentered`.[22]
- The print area is a **defined name**, `_xlnm.Print_Area`, scoped to the sheet with `localSheetId`; print titles are `_xlnm.Print_Titles`.[18] Both are reserved SpreadsheetML names, not user names.[18]
- A worksheet may have **multiple print areas**, and each one prints as a separate page.[14]
- Print titles repeat chosen rows and/or columns at the top/left of every printed page.[16]
- Headers and footers have odd/even/first variants and are stored as `headerFooter` with `oddHeader`, `oddFooter`, `evenHeader`, `evenFooter`, `firstHeader`, `firstFooter`, plus `differentFirst`, `differentOddEven`, `alignWithMargins` and `scaleWithDoc`.[23] Excel's UI exposes field tokens such as `&[Page]`, `&[Pages]` and `&[Date]`;[17] the stored strings use the raw `&`-code form (`&L`/`&C`/`&R`, `&P`, `&N`, `&D`), which the v2 parser must handle.
- Page order is `downThenOver` or `overThenDown`.[24]

V1 authoring is limited to print area, orientation, fit-to-width and margins (feature matrix §7). Everything else must be **parsed, honoured where it affects output, and preserved on save**. That is a smaller job than it sounds: the planner needs the values, and ticket 04's preservation-first layer already keeps untouched XML.

### 1.2 Fit-to-width semantics

Excel's Fit to Page is a shrink-to-fit, not a stretch-to-fill: Microsoft states that "Excel does not stretch the data to fill the pages", and that Excel **ignores manual page breaks** when Fit to is used.[15] The UI exposes Width and Height as page counts, with "Automatic" for height.[15] The Open XML attributes are the same two counts.[19]

For v1 the planner needs: `fitToWidth = 1`, height automatic. The general algorithm (both counts, 0 = automatic) is barely more work and is needed to honour imported files, so implement it once and expose only the v1 control.

### 1.3 Pagination rules the planner must own

- **Columns and rows are the atomic units.** Excel breaks between columns and between rows; it does not split a row's content across pages in normal use. A column wider than the printable width can still overflow a page, and must then be split (or clipped) rather than silently dropped. Maximums are 255 characters per column and 409 points per row,[36] so an oversized column is reachable and an oversized row is possible only on small paper.
- **Hidden rows and columns do not print.** Their sizes must not contribute to pagination. This follows from Excel's print behaviour and needs a golden test rather than a doc quote.
- **Page order** defaults to down-then-over; `overThenDown` must be honoured when present.[24]
- **Multiple print areas** each start a new page sequence.[14]
- **Print titles** repeat bands and therefore reduce the usable body area on every page after the first (and on the first page too, in Excel's implementation — the titles appear on page 1 as well).
- **Scaling is uniform** in both axes; Excel does not distort the sheet.

This is a small, deterministic, unit-testable algorithm. It is also the reason not to delegate pagination to CSS.

## 2. Options compared

Ratings are relative to calcplus. "Effort" means implementation and test effort inside this product.

| Option | Print area / pagination | Orientation, margins, scale | Fills and borders | Text | Cross-browser | Large areas | PDF text | Effort | Verdict |
|---|---|---|---|---|---|---|---|---|---|
| **A. Print stylesheet on the live grid DOM** | Browser fragmentation decides breaks; no page count or break positions are queryable | `@page size`/`margin` where supported | CSS backgrounds suppressed unless `print-color-adjust: exact`; user setting can win | Native, vector | Safari <18.2 does not honour `@page`; fragmentation differs | Whole sheet DOM must exist | n/a (Save as PDF) | Low | **Rejected** |
| **B. Pre-paginated DOM page boxes** | Ours, exact | `@page size` + baked margins | CSS backgrounds, suppressed unless `exact` | Native, vector | Good | One element per cell | n/a | Medium | Viable fallback |
| **C. Pre-paginated inline SVG pages** | Ours, exact | `@page size` + baked margins | SVG shapes print with backgrounds off (Chromium measured) | Native, vector, browser shaping | Good; Safari <18.2 `@page` gap | ~2.2 nodes/cell, measured | n/a | Medium | **Recommended for preview + print** |
| **D. Canvas → raster PDF** | Ours, exact | Exact in PDF | Exact | Raster 150–600 DPI | Best (no print CSS) | 8.7 MP per A4 page at 300 DPI | None; not selectable | Low–medium | Fallback only |
| **E. Vector PDF from the plan (PDFKit)** | Ours, exact | Exact in PDF | Vector | Embedded fonts, selectable | Best | Small vector files | Yes | Medium | **Recommended for export** |
| **F. C + E (one plan, two renderers)** | Ours, exact | Exact in PDF; best-effort in print | Both | Both | Good | Good | Yes, with raster fallback | Medium | **Recommended overall** |
| **G. Browser print-to-PDF only** | Browser fragmentation unless pre-paginated | Print dialog decides | As print | Native, vector | Good | Good | Yes | None | Keep as secondary path |

### 2.1 Why not print the live grid DOM (A)

Three independent failures:

1. **No control over breaks.** CSS fragmentation is a user-agent algorithm. `break-inside: avoid` is a request ("Avoid a break inside the principal box"), and the UA chooses the actual break points.[8] In a Chromium 155 test, a table row 120 mm tall with `break-inside: avoid` was split at the page boundary when it did not fit, and the continuation was not rendered on the next page (measured; §3.2). A spreadsheet cannot ship with page breaks that depend on which engine rendered the table.
2. **No page count.** `window.print()` returns `undefined` and exposes no page count or break positions.[5] A print preview, a "page X of Y" footer, and a test suite all need those numbers before printing.
3. **The grid is a canvas.** The v1 grid is a viewport-sized canvas with DOM overlays (ticket 05). There is no sheet-sized DOM to print; printing the visible viewport is obviously wrong.

### 2.2 Why SVG page boxes (C) rather than DOM page boxes (B)

Both are pre-paginated, so both fix the break problem. SVG wins on two measured points:

- **Fills survive the print dialog's "Background graphics" setting.** Chromium's print preview exposes "Background graphics" and "Headers and footers" as user options.[12] In a headless Chromium 155 test with backgrounds disabled, a CSS `background-color` was suppressed, a `print-color-adjust: exact` element still printed, and **SVG `<rect fill>` printed in both cases** (measured; §3.3). SVG shapes are document content, not CSS backgrounds, so spreadsheet fills and borders are not hostage to an ink-saving checkbox.
- **Node count.** A DOM page needs a positioned element per cell plus text; an SVG page needs a rect and a text element per cell. For 20 A4 pages of 1,200 cells the SVG tree was 52,860 nodes, built in 31 ms, inserted in 186 ms, and printed to PDF in 860 ms (measured; §3.6). That is acceptable; a DOM equivalent is larger and triggers layout for every positioned box.

SVG also gives exact geometry, one element per page, and a preview that is literally the print output.

**Caveats.** Use an explicit two-value `size` (`11in 8.5in`), not the `landscape` keyword: the syntax is defined by the `size` descriptor,[2] and MDN's compatibility data lists `landscape`/`portrait` as unsupported in Safari even where `size` is supported.[3] Do not rely on `dominant-baseline`; position text by its alphabetic baseline. Avoid `vh`/`vw` inside `@page`, which MDN explicitly warns against.[1]

### 2.3 Why a library PDF rather than only browser print-to-PDF (G)

`window.print()` + "Save as PDF" is genuinely attractive: it is vector, it embeds the font subset, it produced a 10.5 KB two-page PDF with selectable text in our test, and it needs no dependency. But:

- It is a multi-step user flow whose UI differs per browser (Chrome's destination list,[13] Firefox's print preview, Safari's PDF menu).
- The output name, page numbering and header/footer defaults are browser-controlled.
- It cannot be invoked headlessly for tests or automation.
- The browser can change paper size, margins, scale and background printing in the same dialog.[12][13]

Keep it as an explicit secondary action. It is also the escape hatch when a user needs a vector PDF of a page our PDF path would rasterise.

### 2.4 Why not raster PDF as the primary (D)

It is the simplest exact path: draw each page with the existing canvas painter and embed a PNG. It is also the worst output for a spreadsheet:

- A4 at 300 DPI is 2,481 × 3,507 px = 8.7 MP. Our test page (1,200 cells) encoded to a 798 KB PNG in 106 ms; 600 DPI was 34.8 MP, 2.13 MB and 402 ms (measured; §3.6). A 20-page export is ~16 MB and a few seconds.
- Text is not selectable, searchable or accessible, and softens when zoomed.
- For comparison, a vector PDF of 20 similar pages was **195 KB** and 60 pages was **586 KB** (measured; §3.6).

Raster is the right *fallback* for a page whose fonts or scripts cannot be embedded, not the default.

### 2.5 Why not print a generated PDF

One could generate the PDF first and hand it to the browser's PDF viewer to print. That guarantees paper and PDF are literally the same document, but it adds a second application and a second print dialog, requires a blob URL or download, and puts the final scaling and margins under the viewer's control rather than the page's. It also throws away the vector-text advantage on pages that had to be rasterised. The direct SVG path is simpler and keeps paper output at the browser's native text quality.

## 3. Evidence by axis

All "measured" numbers come from throwaway Playwright tests against Chromium 155 (Skia/PDF m155) on Linux, run for this ticket. The harness is not committed. No Safari or Firefox print run was possible in this environment; those claims rest on compatibility data and vendor docs and are flagged in §3.5 and §7.

### 3.1 What the browser can and cannot control

**Supported, with version caveats.** `@page` exists in Chrome 2+, Firefox 19+ and Safari 18.2+.[3] The `size` descriptor exists in Chrome 15+, Firefox 95+ and Safari 18.2+.[3] Safari 18.2 added the `@page` margin descriptors and changed `size` to parse as a descriptor rather than a global property.[11] `page-orientation` exists in Chrome 85+ and Firefox 122+, but not Safari.[3] Page-margin boxes (`@top-left` and friends) exist only in Chrome 131+.[3] Named pages (`page: <name>`) exist in Chrome 85+, Firefox 110+ and Safari 1+.[37]

**The page model.** CSS Paged Media defines the page box, the page area and the margin boxes,[7] and CSS Fragmentation defines how content flows between pages.[8] Those are the concepts the planner reproduces in TypeScript, because the browser's implementation of them is not scriptable.

**Not supported anywhere.** The rest of the page-box properties (backgrounds, borders, padding, font properties on the page box) are named in the spec but "have not been supported by any user agent yet".[1] So a page border or a page background cannot be expressed through `@page`; it must be drawn inside the page.

**The print dialog wins.** Chromium's print preview exposes Paper size, Margins (Default/None/Custom/Minimum), Scale, Pages per sheet, Background graphics and Headers and footers.[12] Chrome's own help tells users to "change your preferred print settings" in that dialog.[13] Excel documents the same class of override: if the printer's paper size differs, "Excel adjusts the printout to fit the paper size in the printer and the printed worksheet might not match your print preview".[15] Any web print path inherits this; it is not a calcplus defect, but it must be in the PRD.

**Practical consequence.** The app can set the *initial* page setup, not guarantee it. The PDF export is the only place where calcplus controls the output exactly. The preview should be labelled as the plan, and the print UI should tell the user which dialog settings reproduce it (margins None/Default as appropriate, background graphics on or off).

### 3.2 Pagination: why CSS fragmentation is not enough

The spec defines `break-inside: avoid` as "Avoid a break inside the principal box", while the UA chooses among allowed break points and must guarantee progress through the content.[8] That is the correct behaviour for prose; it is the wrong contract for a spreadsheet, where a page break is a data-driven decision.

Measured (Chromium 155, A4, 15 mm CSS margins):

- Three `<section>`s with `height: calc(297mm - 30mm)`, `break-after: page`, and a 1 px border produced **exactly three A4 pages**, with the bottom marker on each page. Pre-sized page boxes are reliable, including with CSS margins.
- A 10-row table with `thead { display: table-header-group }`, `tr { break-inside: avoid }` and one 120 mm row produced **two pages**. The tall row was split at the page boundary on page 1, and page 2 began with the repeated header and the following rows — the remainder of the tall row was not rendered. The break landed inside the row despite `avoid`.

The first result is the design we want; the second is why we must compute the break ourselves. It also means the print DOM must be pre-paginated, not a flow of cells.

### 3.3 Rendering: DOM vs SVG vs canvas

Measured (Chromium 155, 11 × 8.5 in pages, `printBackground: false`):

| Renderer | Page background | Cell fill | Text | PDF size (2 pages) |
|---|---|---|---|---|
| DOM, CSS `background-color` | suppressed | suppressed | vector, font subset | 10.5 KB |
| DOM, `print-color-adjust: exact` on the filled element | suppressed | **printed** | vector | 10.5 KB |
| Inline SVG, `<rect fill>` | printed | **printed** | vector, font subset | 10.8 KB |
| Canvas (raster) | printed | printed | raster | ~800 KB/page at 300 DPI |

The `print-color-adjust` result matches the spec's model: `exact` says the colours are significant and should not be tweaked "except at the user's request", but the user preference must be respected more strongly than the hint.[9] It is therefore a hint, not a guarantee. SVG shapes sidestep the mechanism.

Vector text was confirmed with `pdftotext` (text extracted) and `pdffonts` (a subsetted CID TrueType font embedded), for both the DOM and SVG cases.

### 3.4 PDF generation: library landscape

| Library | Licence | Latest npm release | Size (min / gzip) | Text model | Fit for this job |
|---|---|---|---|---|---|
| **pdf-lib** | MIT[25] | **1.17.1, 2021-11-06**[26] | 525 KB / 207 KB | Standard fonts are WinAnsi (218 Latin chars); custom fonts via `@pdf-lib/fontkit`, subsetting optional and not universal; draws strings, with no shaping/bidi API documented[25] | Good API, **stale**: five years without a release, and its fontkit fork dates from 2020[26] |
| **jsPDF** | MIT[27] | 4.2.1, 2026-03-17[28] | 420 KB / 133 KB | Standard fonts ASCII; custom TTF via VFS; `html()` is html2canvas → raster[27] | Active and small, but the text/font model is the weakest; its HTML path is exactly the raster path we are avoiding |
| **PDFKit** | MIT[29] | 0.20.2, 2026-08-30[31] | 1.42 MB / 352 KB | TTF/OTF/WOFF/WOFF2/TTC, subsetting, OpenType features, `lineBreak:false`, `baseline`, rotation, `widthOfString`; tagged PDF[29][30] | **Best fit**: vector primitives, real font embedding, absolute positioning, active |
| **pdfmake** | MIT | 0.3.14, 2026-11-08[32] | 835 KB / 294 KB | Declarative document model on top of PDFKit | Wrong abstraction: a second layout engine to keep in sync |
| **@react-pdf/renderer** | MIT | 4.9.0, 2026-08-27[33] | 320 KB package (plus pdfkit and deps) | React components, Yoga layout, PDFKit underneath | Wrong abstraction and a React dependency for a renderer |

**Choice: PDFKit.** It is the strongest fit among the maintained low-level writers for this exact vocabulary: exact page boxes, vector paths, embedded subsetted fonts, positioned text with a canvas-compatible `baseline`, rotation, and images. It is also the only one of the three with a documented tagged-PDF path, which matters for the v2 accessibility work.

Bundle sizes in the table are the published minified builds measured from unpkg (`pdf-lib.min.js`, `jspdf.umd.min.js`, `pdfkit.standalone.js`, `pdfmake.min.js`), gzipped locally; they exclude lazily loaded fonts.

**pdf-lib is rejected on maintenance, not API.** Its `drawText`/`drawRectangle`/`embedFont` API is the cleanest of the three, and if the project already had it in production the calculus would differ. A new five-year dependency with no security releases is not a good foundation for a document-export feature.

**Pattern fills are a known PDFKit gap.** Its feature list still says "Coming soon! Patterns fills".[29] Excel supports pattern fills in v1, so the PDF renderer must draw the supported patterns as explicit vector primitives (lines, dots, grids) or fall back to raster for those cells/pages. SVG handles them natively with `<pattern>`. This is a real, bounded piece of work, not a blocker.

### 3.5 Cross-browser matrix

| Capability | Chrome / Edge | Firefox | Safari | Consequence |
|---|---|---|---|---|
| `window.print()` | Yes | Yes | Yes | Universal trigger |
| `beforeprint` / `afterprint` | 63+ | 6+ | 13+[6] | Build/tear down the print root |
| `@page` at-rule | 2+ | 19+ | **18.2+**[3] | Pre-18.2 Safari does not honour page setup |
| `@page size` | 15+ | 95+ | 18.2+[3] | Use it for paper + orientation |
| `size: landscape` keyword | Yes | Yes | **No** (per BCD)[3] | Always emit explicit `W H` |
| `page-orientation` | 85+ | 122+ | No[3] | Not needed if `size` is explicit |
| `@page` margin boxes | 131+ | No | No[3] | Do not use for headers/footers; draw them in the page |
| Named pages (`page:`) | 85+ | 110+ | 1+[37] | Available if mixed orientation is ever needed |
| SVG in print | Yes | Yes | Yes (long-standing) | Print medium |
| `print-color-adjust: exact` | Yes (measured) | Presumed | Presumed | Hint only; avoid relying on it |

**Safari is the weak spot.** Before 18.2 it does not honour `@page` page setup (the `size` descriptor only began parsing as a descriptor in 18.2[11]), so a print job gets the dialog's paper size and orientation, not the workbook's. Even on 18.2+, BCD says the orientation keywords are unsupported, so explicit dimensions are mandatory. This is the strongest argument for making PDF export, not paper printing, the fidelity channel. Safari print behaviour was not empirically tested here (no Safari/WebKit browser available); treat §7 as a release gate.

**Firefox** supports the pieces we need on paper (`@page size`, print events) but has no margin boxes and no empirical test here. Its print preview also lets the user override margins and headers/footers, like Chrome's.

### 3.6 Memory and performance on large print areas

**SVG page tree (measured, Chromium 155, A4, 20 columns × 60 rows per page, 1,200 cells):**

| Pages | SVG nodes | HTML size | Build | Insert | Print to PDF | PDF size |
|---|---|---|---|---|---|---|
| 20 | 52,860 | 4.8 MB | 31 ms | 186 ms | 860 ms | 195 KB |
| 60 | 158,580 | 14.5 MB | 66 ms | 501 ms | 2.42 s | 586 KB |

This is comfortably inside a desktop browser's budget for a one-shot print. The print root is built on demand and removed on `afterprint`, so it does not affect steady-state grid performance. For very large jobs (hundreds of pages), build the SVG strings in chunks and insert them in one pass; the measured insert cost is roughly linear and dominated by DOM construction.

The cost is linear in pages: about 0.24 MB of markup and 2,600 nodes per A4 page at 1,200 cells. A 500-page job is ~120 MB of markup and ~1.3 M nodes, which is not a sensible thing to put in one document. Excel itself is not designed for that either. The plan should cap the *preview* thumbnails, warn above a threshold (say 200 pages), and let the PDF path (which never materialises a DOM) handle the large case.

**Raster fallback cost (measured, A4, 1,200 cells):**

| DPI | Pixels | MP | PNG encode | PNG size |
|---|---|---|---|---|
| 150 | 1,241 × 1,754 | 2.18 | 47 ms | 328 KB |
| 200 | 1,654 × 2,338 | 3.87 | 60 ms | 580 KB |
| 300 | 2,481 × 3,507 | 8.70 | 106 ms | 798 KB |
| 600 | 4,962 × 7,014 | 34.8 | 402 ms | 2.13 MB |

PNG beats JPEG for this content (line art and text): at 300 DPI JPEG q0.92 was 1.29 MB versus 798 KB for PNG, and JPEG adds ringing around glyphs. Use PNG.

**Canvas limits.** MDN warns that exceeding a canvas's maximum dimensions makes it unusable, that most desktop browsers exceed 10,000 × 10,000 px, and that iOS caps at 4,096 × 4,096.[10] A4 at 300 DPI is 2,481 × 3,507, so it is safe everywhere; A0 at 300 DPI (9,933 × 14,043, 139 MP) already exceeds the iOS limit, and at 600 DPI it is 558 MP. The raster fallback must clamp DPI by pixel area, not just by a constant.

**Do not parallelise raster pages.** One A4 page at 300 DPI is a 34.8 MB RGBA backing store; four in flight is 139 MB before PNG encoding. Render sequentially, release the canvas, and stream each PNG into the PDF document.

### 3.7 Fonts and text fidelity

- **Print** uses the browser's font stack, so a sheet rendered with an installed Calibri or Aptos prints with that font, shaped by the browser. This is the highest text fidelity available without shipping fonts.
- **PDF** can only embed font files we possess. Bundle metric-compatible open fonts: **Carlito** (metric-compatible with Calibri, OFL)[34] and **Liberation Sans/Serif** (metric-compatible with Arial/Times New Roman, OFL).[35] Map the workbook's family/weight/style to a bundled file, subset on embed.[29][30]
- **Aptos has no known metric-compatible open clone.** Microsoft 365's newer default font is therefore a genuine gap: the PDF will substitute, and glyph shapes and advances will differ slightly. Because the planner positions each line at a fixed x/baseline, the text will not reflow, but right/centre alignment must be recomputed with the PDF font's width (`widthOfString`)[30] so the anchored edge stays put. Document this as a fidelity limit; it is the same class of substitution Excel itself performs when a font is missing.
- **Complex scripts.** None of the three PDF libraries is a full text engine: they draw a string with a font's glyphs, and PDFKit applies OpenType features via fontkit,[29][30] but its published dependency set contains no bidi-reordering library.[31] Arabic, Hebrew, Indic and Thai must be treated as "verify, then fall back". The safe v1 rule: if a page contains a run in a script/font we have not certified, rasterise that page.
- **Rich text** (mixed bold/italic/colour inside a cell) becomes multiple text runs; both SVG and PDFKit support that natively (SVG `tspan`, PDFKit `continued`/separate draws).[30]

### 3.8 Cell-level fidelity mapping

| Excel feature | In the plan | SVG renderer | PDF renderer |
|---|---|---|---|
| Font family, size, bold, italic, underline, strike, colour | `text` run with a resolved font key and the grid's metrics | `<text>`/`<tspan>`; browser shapes with the real font | embedded subsetted font; `widthOfString` for anchors |
| Fill: solid, pattern | `rect` + fill token | `<rect fill>` or `<pattern>` | `rect().fill()`; patterns drawn as primitives |
| Borders: edge, weight, dash, diagonal | one `line` per edge plus diagonals | `<line>` with `stroke-width`/`stroke-dasharray` | `moveTo`/`lineTo`/`stroke` |
| Merged cells | one rect for the region, anchor text, clip to the region | rect + clipped text | rect + clipped text |
| Row height / column width | points from the sheet model; rows are atomic | exact geometry | exact geometry |
| Wrap, shrink-to-fit, rotation | line runs at baseline positions; smaller font for shrink; rotate around the anchor | `<text>` per line, `transform="rotate(...)"` | positioned text, `rotation` |
| Alignment and indent | anchor `start`/`middle`/`end` plus padding | `text-anchor` | computed x |
| Number formats | already-rendered strings from the formatter | text | text |
| Conditional formatting | resolved fill/font/icon primitives | shapes and text | shapes and text |
| Text overflow into empty neighbours | clip rect stops at the first non-empty cell or the print/page edge | `clip-path` | clip rect |

Four details that decide whether this looks like Excel:

- **Merged cells** keep only the anchor value and draw one outline; internal gridlines and borders inside the merge must be suppressed. Ticket 05 already owns the merge map; the planner must consume it rather than re-derive it.
- **Borders on a page boundary** must be drawn on the page that owns the edge, exactly once. Test both sides of a break to catch a doubled or missing hairline.
- **Row height is the only vertical unit.** SVG line boxes and PDFKit's `lineGap` must not add height; the plan fixes baselines and the renderers obey them.
- **Text overflow is an explicit clip rectangle**, not `overflow: hidden` per cell. Excel lets text spill into adjacent empty cells and clips it at the first non-empty neighbour or the page edge; per-cell clipping would truncate almost every label.

Diagonal borders and pattern fills are the two places where the vector vocabulary needs extra work; both are bounded and both are covered by the fidelity corpus.

## 4. Recommended architecture

```
             Workbook model (ticket 07) ──► SheetView projection (ticket 05)
                                                        │
  Page setup (xlsx + v1 UI) ──► PrintPlanner ───────────┘
                                   │
                                   ▼
                         PrintPlan { pageBox, pages[],
                                     scale, decorations }
                                   │
             ┌─────────────────────┼──────────────────────┐
             ▼                     ▼                      ▼
       SvgRenderer          CanvasRenderer          PdfRenderer
     preview + print       raster fallback          (PDFKit)
             │                     │                      │
             ▼                     ▼                      ▼
      hidden print root      PNG per page          PDF Blob → download
```

### 4.1 Modules and seams

| Module | Responsibility | Test surface |
|---|---|---|
| `print/settings.ts` | Normalise `pageSetup`, `pageMargins`, `printOptions`, `pageSetUpPr`, print-area/titles names into a `PrintSettings` | Pure: XML fixtures → settings |
| `print/planner.ts` | Resolve ranges, visible rows/cols, scale search, page bands, display list | **Pure and golden-tested**: page count, break positions, page→range map, primitive list |
| `print/svg.ts` | `PrintPlan` → SVG string per page | Snapshot + geometry assertions |
| `print/print.ts` | Print root lifecycle, `@page` injection, `beforeprint`/`afterprint`, `document.fonts.ready` | Integration: page count, no blank pages |
| `print/pdf.ts` | `PrintPlan` → PDFKit document → `Blob` | PDF page boxes, text extraction, visual diff |
| `print/fonts.ts` | Family mapping, lazy font loading, coverage/shaping certification | Font corpus tests |
| `print/raster.ts` | `PrintPlan` → canvas at chosen DPI (reuses the grid painter) | Pixel diff vs SVG at same size |

`PrintPlanner` is the deep module: a narrow input (settings + sheet projection) and a narrow output (a plan), with all the Excel semantics inside. It is also the only part that needs Excel golden files.

### 4.2 Planner algorithm

1. **Resolve the print range.** Use `_xlnm.Print_Area` if present (comma-separated areas, each a `Sheet!$A$1:$D$20` reference); otherwise the used range. Each area is an independent page sequence.[14] Intersect with the sheet bounds; drop hidden rows/columns from the flow.
2. **Resolve the page box.** `paperSize` (or `paperWidth`/`paperHeight`) and `orientation` → width/height in points (1 pt = 1/72 in; A4 = 595.28 × 841.89, Letter = 612 × 792). Subtract `pageMargins` (inches × 72) to get the body box; reserve header/footer bands in v2.
3. **Resolve the scale.** If `pageSetUpPr@fitToPage`, search for the largest `s ≤ 1` (Excel's Fit to does not enlarge[15]) such that packing the visible columns at `s` yields at most `fitToWidth` pages wide and the rows yield at most `fitToHeight` pages tall (0 = automatic in Excel's UI mapping; verify with a golden file). Otherwise `s = scale / 100`, clamped to 10–400.[19] Page counts are monotone in `s`, so a 20-iteration binary search is enough; log the chosen percentage for the UI.
4. **Pack pages.** Walk columns left-to-right accumulating widths until the body width is exceeded, then start a new page column band; a single column wider than the body box is split across pages (flag this case). Repeat for rows; a row taller than the body box is clipped (Excel's maximum row height makes this a small-paper edge case[36]). Emit bands in `downThenOver` or `overThenDown` order.[24]
5. **Emit the display list.** For each page band, clip the sheet's render records to the band, translate to page coordinates, and emit primitives:
   - `rect` — fill, optional stroke (borders, merged-cell outlines, gridlines when enabled)
   - `line` — border edges with width, colour, dash
   - `text` — x, baseline y, anchor (`start`/`middle`/`end`), font key, size, weight/style, colour, underline/strike, rotation, clip rect
   - `image` — placeholder in v1, drawn in v2 (charts/images)
6. **Decorate.** v1: nothing. v2: header/footer bands from `headerFooter`, print titles, page numbers.

The display list is deliberately the same vocabulary the grid's canvas painter consumes, so the raster fallback is a backend swap rather than a second renderer.

### 4.3 Print and preview renderer (SVG)

- Build a `#print-root` that is `display: none` on screen and `display: block` in `@media print`, using the print-stylesheet pattern MDN's printing guide describes.[4] Inject a per-job `<style>` with `@page { size: <W>pt <H>pt; margin: 0 }`.
- Each page: `<section class="print-page">` with the exact page box, `overflow: hidden`, `break-after: page` (and the legacy `page-break-after`), containing an inline `<svg viewBox="0 0 W H" width="100%" height="100%" preserveAspectRatio="none">` sized to the page box. Last page has no forced break.
- Margins are baked in: content is placed inside `left/top` offsets from the page plan, so the sheet's internal layout is fixed. The dialog can still add its own margin or scale the page box as a whole; that is the documented limit, not a layout reflow.
- Await `document.fonts.ready` before building, so SVG text measures and renders with the final font.
- Build lazily on `beforeprint` (and on "Print" click before calling `window.print()`); remove on `afterprint`. `print()` blocks while the dialog is open per MDN,[5] but the events are the portable lifecycle across browsers.[6]
- Preview uses the same SVG strings in a modal with page thumbnails and zoom; no second layout path.
- For very large jobs, insert the page markup in one `innerHTML` assignment (measured: 186 ms for 20 pages, 501 ms for 60).

If Safari's SVG print proves unreliable in the device test, option B is a drop-in replacement behind the same plan: the same pre-sized page boxes, with DOM cells and `print-color-adjust: exact` instead of SVG. Only `print/svg.ts` changes; the planner, preview geometry and PDF path are unaffected.

### 4.4 PDF renderer (PDFKit)

- `PDFDocument({ size: [w, h], margin: 0, autoFirstPage: false })`; add one page per plan page.
- Draw fills with `rect(...).fill(color)`, borders with `moveTo/lineTo/lineWidth/stroke`, text with `font(...)`, `fontSize(...)`, `fillColor(...)`, `text(str, x, y, { lineBreak: false, baseline: 'alphabetic', rotation, width })`.[30] Use `widthOfString` to resolve `start`/`middle`/`end` anchors so alignment survives font substitution.
- Embed fonts once per document, subsetted (`doc.font(bytes)`); reuse the registered font across pages.[29][30]
- Pattern fills: draw the supported Excel patterns as vector primitives (PDFKit has no pattern-fill API yet[29]).
- Raster fallback: if a page contains an uncertified font/script, render that page with `print/raster.ts` at the chosen DPI and `doc.image(png, ...)` across the page box. Keep vector for the rest of the document.
- Produce a `Blob` with PDFKit's experimental `toBlob`/`toBytes` browser helpers[29] and trigger a download. Lazy-import PDFKit and the fonts on first export; the 352 KB gzip cost is not on the startup path.
- Optional: run the export in a worker (PDFKit's browser build produces bytes, so it can be moved off the main thread) to keep the UI responsive on 50+ page exports.

### 4.5 What "PDF matches the print preview" means

- **Geometry, pagination, content and colour** are identical because all three backends consume the same `PrintPlan`.
- **Text** is vector in both, but the PDF may use a substitute font; shapes may differ slightly while positions and alignment anchors hold.
- **Fallback pages** are raster in the PDF; the preview is vector. They match at 100% but not under magnification.

The acceptance test should compare preview SVG → raster against PDF → raster at the same DPI, with a perceptual threshold, plus exact page-box and page-count assertions.

## 5. Fidelity limits of the chosen path

1. **Print is advisory.** The print dialog can change paper size, orientation, margins, scale, headers/footers and background graphics.[12][13] Excel has the same caveat with printer paper size.[15] Only the PDF export is exact.
2. **Safari before 18.2 does not honour `@page` page setup.** Orientation, paper size and margins come from the dialog. Even on 18.2+, the `landscape`/`portrait` keywords are unsupported per BCD,[3] so explicit dimensions are mandatory. This is untested on device here.
3. **PDF font substitution.** Calibri/Arial/Times have metric-compatible open substitutes; **Aptos has no known substitute**. Non-bundled fonts substitute, and glyph shapes/advances differ. Line breaks are preserved because the planner positions each line; anchored edges are recomputed with the PDF font.
4. **Complex scripts may be raster.** Arabic/Hebrew/Indic/Thai are not certified for vector PDF in v1; those pages fall back to 300 DPI raster, losing text selection/search and increasing file size to ~800 KB/page.
5. **Pattern fills in PDF** must be hand-drawn or rasterised; PDFKit has no pattern-fill support.[29]
6. **Pagination depends on font metrics.** If the user's machine lacks the workbook font and the fallback has different advances, calcplus can paginate differently from Excel. Excel has the same failure mode when a font is missing; bundling metric clones for the common families keeps it rare.
7. **Excel's exact fit-to-page search and edge cases are not yet verified**: `fitToWidth`/`fitToHeight` = 0 semantics, oversized columns, hidden rows/columns, and the interaction with manual breaks.[15][19] These need golden workbooks before the planner is frozen.
8. **No tagged/accessible PDF in v1.** PDFKit can produce tagged PDF,[29] but absolute-positioned cell drawing is not a logical structure. The accessibility target is ticket 22.
9. **Images and charts are v2.** The display list has an `image` primitive and the planner reserves its anchor/size; v1 renders a placeholder and must not break pagination.
10. **Raster fallback quality is a judgement.** 300 DPI is a print-grade resolution, but it is not vector; very small text and hairline borders will be softer than the browser's own PDF.
11. **Exact page-box heights can round differently across engines.** A page element exactly as tall as the page area can spill into a blank trailing page if an engine rounds up. Keep `overflow: hidden` and a sub-pixel safety margin, and test for a blank last page in every engine.

## 6. V2 extension path

The reason to own pagination is that every v2 print feature is a small addition to the plan:

- **Headers/footers.** Parse `headerFooter`[23] and Excel's `&`-code grammar: `&L`/`&C`/`&R` sections, `&P`, `&N`, `&D`, `&T`, `&F`, `&A`, font/size codes; the UI's `&[Page]`-style tokens[17] are the same fields in a friendlier syntax. Lay them out in the reserved margin band and emit text primitives. Do **not** use `@page` margin boxes: Chrome-only.[3] Both renderers already draw text.
- **Print titles.** Parse `_xlnm.Print_Titles`[18] into repeat row/column bands; subtract them from the body box on every page and emit them first. The planner already emits band-based pages, so this is a loop change.
- **Manual page breaks.** Parse `rowBreaks`/`colBreaks`, force band boundaries at those indices, and keep Excel's rule that Fit to ignores manual breaks.[15]
- **Page order and multi-page scaling.** Both are already in the planner (`pageOrder`, `fitToWidth`/`fitToHeight`); v2 only exposes the UI.
- **Page-break preview / page-layout view.** Draw the plan's band boundaries on the grid; no new layout engine.
- **Images, charts, shapes.** Add image primitives; SVG `<image>` and PDFKit `doc.image` both handle PNG/JPEG. Pagination is unaffected because the anchor and size are already reserved.
- **Mixed orientation per sheet set.** Named pages (`page:`) are supported in all three engines,[37] but calcplus prints one sheet at a time in v1, so this can wait.

## 7. Evidence gaps and tests to run before freezing

1. **Safari print run** on 18.2+ and, if supported, an older Safari: `@page size` with explicit dimensions, SVG text and fills, default dialog margins, background graphics off. This is the largest untested risk; if SVG print fails there, switch the print renderer to option B (DOM page boxes) behind the same plan.
2. **Firefox print run**: same page, `@page size`, SVG, dialog margin overrides, and whether an exact-height page box produces a blank trailing page.
3. **Chrome dialog behaviour** for `@page { margin: 0 }` with Margins = Default / None / Custom. Our headless test used the PDF API, which is not the dialog.
4. **PDFKit shaping/bidi** for Arabic, Hebrew, Devanagari, Thai and CJK; decide the exact fallback trigger and certify the fonts.
5. **Font corpus**: Excel files using Calibri, Aptos, Arial, Times New Roman, Cambria, Segoe UI, and a CJK font. Measure width drift and page-count drift versus Excel.
6. **Excel golden workbooks** for pagination: fit-to-width 1/2/automatic, multiple print areas, hidden rows/columns, oversized column, manual breaks, print titles, landscape/portrait, A4/Letter/Legal.
7. **Print quality of the 300 DPI raster fallback** on paper and in a PDF viewer at 200%/400% zoom.
8. **Large-export performance** on a low-end laptop: 100+ pages, PDFKit in a worker, memory high-water mark.

## 8. Consequences for the spec package

- The technical design should introduce **`PrintPlan` as a deep module** with a pure interface, owned by the print workstream, consumed by three renderers. Page count, breaks and ranges are unit-testable without a browser.
- The PRD should say: **print and PDF share one pagination model; PDF is vector-first with a raster fallback; paper printing is best-effort because the browser dialog is authoritative.**
- Acceptance criteria should be: given a workbook with print area, orientation, fit-to-width and margins, preview page count/breaks match Excel within one page; the PDF's page boxes and page count match the preview; PDF text is selectable except on declared fallback pages.
- Adopt **PDFKit** (lazy, MIT) and **not pdf-lib** (last release 2021). Add the font bundle (Carlito/Liberation) and its licence notices to the build.
- Add a **print fidelity corpus** to ticket 13's conformance harness: golden `.xlsx` page setup plus expected page counts/breaks, and visual diffs of preview vs PDF.
- The print DOM must be **built on demand and removed**, never part of the app shell; the `@page` rule must be injected per job, not baked into the app stylesheet.
- The design should record the **Safari `@page` gap** as a known limitation and make "Export PDF" the recommended fidelity path in the UI.
- If the grid decision changes to a third-party SDK (ticket 17), re-check whether its print/PDF path can consume `PrintPlan`; do not let the SDK own pagination.

## Sources

All sources accessed 2026-11-13. Numbers marked "measured" come from throwaway Playwright 1.65 tests against Chromium 155 (Skia/PDF m155) on Linux, run for this ticket; the harness is not part of the repository.

1. MDN — *`@page` CSS at-rule* (page descriptors; unsupported page properties; no `vh`/`vw` in `@page`). https://developer.mozilla.org/en-US/docs/Web/CSS/@page
2. MDN — *`size` CSS at-rule descriptor* (syntax, keyword sizes, orientation). https://developer.mozilla.org/en-US/docs/Web/CSS/Reference/At-rules/@page/size
3. MDN browser-compat-data — `css/at-rules/page.json` (`@page`, `size`, orientation keywords, `page-orientation`, margin boxes). https://github.com/mdn/browser-compat-data/blob/main/css/at-rules/page.json
4. MDN — *Printing* (print stylesheet, `@media print`, `beforeprint`/`afterprint`, hidden-iframe printing). https://developer.mozilla.org/en-US/docs/Web/CSS/Guides/Media_queries/Printing
5. MDN — *Window: `print()` method* (blocks while the dialog is open; no return value). https://developer.mozilla.org/en-US/docs/Web/API/Window/print
6. MDN browser-compat-data — `api/Window.json` (`beforeprint`/`afterprint` support). https://github.com/mdn/browser-compat-data/blob/main/api/Window.json
7. W3C — *CSS Paged Media Module Level 3* (page box, page area, page-margin boxes, page size). https://drafts.csswg.org/css-page-3/
8. W3C — *CSS Fragmentation Module Level 3* (`break-inside: avoid` semantics; UA-chosen break points; progress guarantee). https://drafts.csswg.org/css-break-3/
9. W3C — *CSS Color Adjustment Module Level 1* (`print-color-adjust: economy | exact`; user preference wins). https://drafts.csswg.org/css-color-adjust-1/
10. MDN — *`<canvas>` HTML graphics canvas element* (maximum canvas size; iOS 4,096 × 4,096; canvas content not exposed to assistive technology). https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Elements/canvas
11. WebKit — *WebKit Features in Safari 18.2* (`@page` margin descriptors; `size` parses as a descriptor; JIS paper sizes). https://webkit.org/blog/16301/webkit-features-in-safari-18-2/
12. Chromium — `chrome/app/printing_strings.grdp` (print preview options: paper size, margins, background graphics, headers and footers, scale, pages per sheet). https://chromium.googlesource.com/chromium/src/+/HEAD/chrome/app/printing_strings.grdp
13. Google Chrome Help — *Print from Chrome* (destination and print settings in the dialog). https://support.google.com/chrome/answer/1069693
14. Microsoft Support — *Set or clear a print area on a worksheet* (multiple print areas; each prints as a separate page; saved with the workbook). https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet
15. Microsoft Support — *Scale a worksheet* (Fit to width/height; Excel does not stretch; ignores manual breaks under Fit to; printer paper size can change the output). https://support.microsoft.com/en-us/excel/scale-a-worksheet
16. Microsoft Support — *Repeat specific rows or columns on every printed page* (print titles). https://support.microsoft.com/en-us/excel/repeat-specific-rows-or-columns-on-every-printed-page
17. Microsoft Support — *Headers and footers in a worksheet* (header/footer elements and codes such as `&[Page]`, `&[Pages]`). https://support.microsoft.com/en-us/excel/headers-and-footers-in-a-worksheet
18. ISO/IEC 29500-1 §18.2.6 — `definedName` (built-in names `_xlnm.Print_Area` and `_xlnm.Print_Titles`; reserved `_xlnm.` prefix), reproduced at c-rex.net. https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_definedName_topic_ID0E2JB4.html
19. Microsoft Learn — *PageSetup class* (`orientation`, `paperSize`, `scale` 10–400, `fitToWidth`, `fitToHeight`, `pageOrder`, `paperWidth`/`paperHeight`; fit overrides scale). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagesetup
20. Microsoft Learn — *PageSetupProperties class* (`pageSetUpPr@fitToPage`, `autoPageBreaks`). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagesetupproperties
21. Microsoft Learn — *PageMargins class* (margins in inches; default example 0.7/0.75/0.3). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagemargins
22. Microsoft Learn — *PrintOptions class* (`gridLines` + `gridLinesSet`, `headings`, `horizontalCentered`, `verticalCentered`). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.printoptions
23. Microsoft Learn — *HeaderFooter class* (odd/even/first headers and footers; `differentFirst`, `differentOddEven`, `alignWithMargins`, `scaleWithDoc`). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.headerfooter
24. Microsoft Learn — *PageOrderValues enum* (`downThenOver`, `overThenDown`). https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pageordervalues
25. pdf-lib — README (MIT; standard fonts use WinAnsi; custom fonts via `@pdf-lib/fontkit`; subsetting caveats; drawing API). https://github.com/Hopding/pdf-lib
26. npm registry — `pdf-lib` metadata (latest 1.17.1, published 2021-11-06; MIT; 525 KB minified). https://registry.npmjs.org/pdf-lib
27. jsPDF — README (MIT; standard fonts are ASCII-only; custom TTF via VFS; `html()` depends on html2canvas/dompurify). https://github.com/parallax/jsPDF
28. npm registry — `jspdf` metadata (latest 4.2.1, published 2026-03-17; MIT; 420 KB minified). https://registry.npmjs.org/jspdf
29. PDFKit — README (MIT; vector graphics; TTF/OTF/WOFF/WOFF2/TTC; font subsetting; images; tagged PDF; "Patterns fills" still coming soon; browser `toBlob`/`toBytes`). https://github.com/foliojs/pdfkit
30. PDFKit — *Text in PDFKit* (`lineBreak: false`, `baseline`, `rotation`, `widthOfString`, OpenType `features`, rich text via `continued`). https://github.com/foliojs/pdfkit/blob/master/docs/text.md
31. npm registry — `pdfkit` metadata (latest 0.20.2, published 2026-08-30; MIT; 1.42 MB standalone / 352 KB gzip). https://registry.npmjs.org/pdfkit
32. npm registry — `pdfmake` metadata (latest 0.3.14, published 2026-11-08; MIT; depends on pdfkit). https://registry.npmjs.org/pdfmake
33. npm registry — `@react-pdf/renderer` metadata (latest 4.9.0, published 2026-08-27; MIT; depends on pdfkit). https://registry.npmjs.org/@react-pdf/renderer
34. Carlito — metric-compatible with Calibri, SIL Open Font License. https://github.com/googlefonts/carlito
35. Liberation Fonts — metric-compatible with Arial/Times New Roman/Courier New, SIL Open Font License. https://github.com/liberationfonts/liberation-fonts
36. Microsoft Support — *Excel specifications and limits* (1,048,576 rows × 16,384 columns; column width 255 characters; row height 409 points). https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
37. MDN browser-compat-data — `css/properties/page.json` (`page` property: Chrome 85, Firefox 110, Safari 1). https://github.com/mdn/browser-compat-data/blob/main/css/properties/page.json
