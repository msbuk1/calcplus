# 04 — .xlsx import/export approach

Status: research complete
Date: 2026-11-13
Ticket: [04-xlsx-io.md](../issues/04-xlsx-io.md)
Method: primary vendor docs, source repositories, license files, ECMA-376 / Microsoft open specifications, and a local round-trip experiment against a fixture exercising styles, layout, comments, hyperlinks, validation, conditional formatting, charts, hidden sheets, and defined names.

## TL;DR

**Recommendation: build a thin, preservation-first OOXML package layer in-house** (`fflate` or JSZip for ZIP, `saxes` for streaming XML), rather than adopting a general workbook library as calcplus's storage engine.

The decisive requirement is not "read/write cells" — every candidate can do that. It is **round-trip preservation of everything calcplus does not model**. A library that deserializes into its own rich object model can only write back what that model understands. Our experiment confirms this: ExcelJS silently dropped the chart, external-link and VBA parts; SheetJS CE additionally dropped cell fonts/fills/borders, data validation, conditional formatting, freeze panes and the chart, and even emitted a non-schema column attribute. Neither is acceptable for a product whose core promise is high-fidelity `.xlsx` round-trip.

The recommended design has three rules:

1. **No-op save returns the original bytes.** Opening and saving without edits is byte-identical by construction.
2. **Untouched parts are copied, never reserialized.** Charts, images, pivot caches, external links, macros and unknown extensions survive verbatim.
3. **Only dirty parts are rewritten, and modeled tables are append-only.** Existing style indices, shared-string indices and relationship IDs are never renumbered.

This is more work than `npm install exceljs`, but it is the only approach that can honestly support "no silent loss". It also becomes the moat: the I/O layer is the hardest part of an Excel clone to get right.

If build capacity is not available, **SheetJS Pro (Pro Edit plus the styling/image/chart builds and the pivot/encryption add-ons) is the only commercial fallback that covers all required categories**; ExcelJS is a good data/styles library but cannot preserve charts, external links, pivots or macros, so it is not a sufficient V1 answer on its own.

## 1. What "high fidelity" must mean

Three different verbs are routinely conflated in library marketing. calcplus should define them explicitly:

| Term | Meaning | Acceptance test |
|---|---|---|
| **Read** | Parse a part into the calcplus model so it can be displayed/edited. | Value/style/range visible in the UI. |
| **Write** | Serialize the model into a valid OOXML part. | Excel opens the output without a repair prompt. |
| **Preserve** | Keep a part or construct that calcplus does not model, without changing its meaning. | Unmodified part is byte-identical; IDs/ranges still resolve; Excel shows the same object. |

Two further distinctions matter:

- **Byte preservation vs semantic preservation.** Copying chart XML keeps the chart, but if the user inserts a row above the chart's source range, the chart now points at the wrong cells. Preserving bytes is necessary but not sufficient for objects whose meaning depends on cell ranges.
- **Preserved vs displayed.** V1 can preserve a chart without rendering it. That is a legitimate scope cut for v2, but it must be an explicit product statement, not an accident of the library.

Proposed fidelity contract for V1:

| Tier | Contents | V1 guarantee |
|---|---|---|
| **A — modeled and editable** | Cell values/types, formulas + cached results, number formats, fonts/fills/borders/alignment, merged cells, row/column dimensions and hidden state, freeze panes, defined names, sheet list/visibility, comments (notes), hyperlinks, data validation, conditional formatting | Read, write, and round-trip with semantic equality. |
| **B — preserved, not editable** | Charts/chartsheets, images, shapes, drawings, OLE objects, form controls, pivot tables/caches, external links, VBA projects, tables (until modeled), slicers/timelines, sparklines, custom XML | Parts copied verbatim; survives open→save unchanged. Not rendered/edited in V1 unless separately scoped. Structural edits that invalidate ranges produce a warning. |
| **C — rejected** | Encrypted/IRM workbooks, `.xls`, `.xlsb`, corrupt/non-ZIP containers, files over the size/memory cap | Clear, actionable error. No partial import. |

**No silent loss:** any export that would drop or invalidate a Tier B construct must produce a user-visible export report (and, where possible, refuse the destructive operation).

## 2. What is actually inside an `.xlsx`

An `.xlsx` file is an OPC (Open Packaging Conventions) ZIP package containing XML parts plus relationship parts that connect them. The minimum workbook is a workbook part, a worksheet part, and the relationships between them `[S1][S2]`. A real workbook adds styles, shared strings, drawings, charts, comments, pivot caches, external links and more `[S2]`.

| Part | Carries | V1 treatment |
|---|---|---|
| `xl/workbook.xml` | Sheet list, sheet IDs and relationship IDs, defined names, workbook properties, date system, calculation settings | Model + write |
| `xl/_rels/workbook.xml.rels` | Maps workbook relationship IDs to sheet/styles/sharedStrings/theme parts | Model + write |
| `xl/worksheets/sheetN.xml` | `<sheetData>` cells, rows, columns, merges, hyperlinks, data validations, conditional formatting, views/panes, autoFilter, page setup | Model + surgical write |
| `xl/styles.xml` | `numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `cellStyles`, `dxfs`, `tableStyles` | Model + append-only write |
| `xl/sharedStrings.xml` | Deduplicated text table; cells reference indices | Model + append-only write |
| `xl/theme/theme1.xml` | Theme colors and fonts referenced by styles | Preserve; copy verbatim |
| `xl/comments*.xml` + VML drawing | Legacy notes, authors, anchors | Model + write |
| `xl/threadedComments/*` + `persons/*` | Modern threaded comments, replies, mentions | Preserve (v1); edit later |
| `xl/drawings/*.xml` + `xl/media/*` | Images, shapes, chart anchors | Preserve |
| `xl/charts/*.xml` + `xl/chartsheets/*` | Chart definitions and chart sheets | Preserve (v2 render) |
| `xl/pivotTables/*`, `xl/pivotCache/*` | Pivot definitions and cached source records | Preserve; never refresh in v1 |
| `xl/externalLinks/*` | External workbook references | Preserve; never resolve/refresh |
| `xl/vbaProject.bin` | VBA project | Preserve in `.xlsm`; never execute/edit |
| `docProps/*`, `customXml/*`, `[Content_Types].xml`, `_rels/.rels` | Metadata, custom XML, part content types, package roots | Preserve; update content types only when adding/removing parts |

Two format facts drive the design:

- A formula cell stores the formula in `<f>` and the **last calculated value in `<v>`**; the latter is explicitly the cached value, not a recomputation `[S3]`. Any writer that drops `<v>` makes the workbook show blanks or stale values until Excel recalculates.
- **Excel does not strictly follow ECMA-376.** SheetJS's own format documentation says so `[S7]`. Therefore schema validation is not enough: the test corpus must be files actually produced by Excel (plus LibreOffice/Google Sheets exports), not only files produced by our writer.

The normative references are ECMA-376 (published as ISO/IEC 29500), with Part 1 covering the markup language, Part 2 the Open Packaging Conventions, and Part 4 transitional migration features `[S1]`.

## 3. Options evaluated

| Option | Version evaluated | License | What it is |
|---|---|---|---|
| **SheetJS Community Edition (CE)** | 0.20.3 (CDN distribution) | Apache-2.0, attribution required `[S10]` | Broad-format data extraction/writing library with a plain-object cell model. |
| **SheetJS Pro** | current vendor tiers (not independently tested) | Commercial, quote-based `[S11]` | Paid builds adding styling, validation/conditional formatting, images, charts, fidelity editing, pivot and encryption components. |
| **ExcelJS** | 4.4.0 (latest release, 2023-10-19) | MIT `[S16]` | Rich workbook object model with strong styles, images, validation and conditional formatting. |
| **xlsx-populate** | 1.21.0 | MIT `[S21]` | Thin XML-manipulation library that deliberately preserves unknown parts. |
| **Direct OOXML** | ECMA-376 / MS-OFFCRYPTO | Own code + MIT/ISC deps | Parse and write the package parts directly. |
| **xlsx-js-style** | 0.18.5-based fork | Apache-2.0 `[S28]` | SheetJS 0.18.5 plus basic style writing. Rejected: stale base, no validation/conditional formatting/charts. |
| **Luckyexcel** | current | MIT `[S26]` | One-way `.xlsx` → Luckysheet JSON import. Rejected: no export, no round-trip. |
| **Univer** | current | Apache-2.0 `[S27]` | Full spreadsheet SDK/UI. Rejected as an I/O dependency: adopting it means adopting its model and UI, which conflicts with building calcplus's own Excel-clone model. |

SheetJS CE is distributed primarily from `cdn.sheetjs.com`, not the public npm registry `[S29]`; that is a supply-chain/currency consideration for a long-lived product.

## 4. Comparison

Notation: **R / W / P** = Read / Write / Preserve. ✅ full, ◐ partial, ❌ none. `*` = vendor claim, not independently verified (SheetJS Pro is closed-source and quote-only).

### 4.1 Data, styles and layout

| Capability | SheetJS CE | SheetJS Pro | ExcelJS 4.4 | xlsx-populate 1.21 | Direct OOXML |
|---|---|---|---|---|---|
| Cell values and types (number, text, bool, error, date) | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Formulas | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Cached formula results (`<v>`) | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ❌ | ✅ / ✅ / ✅ |
| Number formats | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Fonts, fills, borders, alignment | ◐ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Merged cells | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Row heights, column widths, hidden rows/columns | ✅ / ✅ / ✅¹ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Freeze panes / sheet views | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Rich text runs | ◐ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ◐ / ◐ / ✅ |

¹ SheetJS CE preserved the values in our test but emitted a `level` attribute on `<col>` (in addition to `outlineLevel`) that openpyxl 3.1.5 rejects as an unexpected keyword (`TypeError: ColumnDimension.__init__() got an unexpected keyword argument 'level'`). LibreOffice opened it, so this is an interoperability smell rather than a confirmed Excel failure `[E1]`.

**What the evidence says.**

- SheetJS CE's own docs state that the CE parsers/writers "focus on raw data", `cellStyles` adds row/column properties and some style metadata, and "SheetJS Pro offers cell / text styling, conditional formatting and additional styling options" `[S8][S9]`. In our test the parsed `s` field held only a fill; fonts, borders and alignment were absent, and the writer flattened all of them `[E1][S14]`.
- ExcelJS documents a full style model (fonts, fills, borders, alignment, number formats, rich text) `[S15]`, and our round-trip preserved the fixture's font family/size/bold/italic/color, solid fill, alignment, wrap, four border styles and number formats. Only normalizations such as `\$#,##0.00` → `$#,##0.00` were observed `[E1]`.
- xlsx-populate describes a focus on "keeping existing workbook features and styles intact" `[S20]`; the fixture round-trip preserved styles and layout exactly `[E1]`.
- Direct OOXML wins by construction, but the style table is the hardest part: `cellXfs` entries are referenced by integer index from each cell's `s` attribute, so existing indices must never be renumbered. New styles should be appended.
- SheetJS CE reads rich-text runs into `r`/`h` fields but ignores them on write: a cell whose `r` field was set to bold+plain runs was written as a single plain `<v>` string in our check `[E1]`. Rich text therefore belongs to Pro for round-trip purposes.

### 4.2 Workbook structure, annotations and embedded objects

| Capability | SheetJS CE | SheetJS Pro | ExcelJS 4.4 | xlsx-populate 1.21 | Direct OOXML |
|---|---|---|---|---|---|
| Defined names (workbook and sheet scope) | ✅ / ✅ / ✅ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Multiple sheets, hidden/veryHidden, order | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Sheet metadata (tab color, code name, views, print setup) | ◐ / ◐ / ◐ | ✅ / ✅ / ✅* | ✅ / ✅ / ◐ | ◐ / ◐ / ✅ | ✅ / ✅ / ✅ |
| Comments / notes | ✅ / ✅ / ◐² | ✅ / ✅ / ✅* | ✅ / ◐ / ◐³ | ◐ / ❌ / ✅ | ✅ / ✅ / ✅ |
| Hyperlinks (external and internal) | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Data validation | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ | ✅ / ✅ / ✅ |
| Conditional formatting | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ❌ / ❌ / ✅ | ✅ / ✅ / ✅ |
| Images | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ✅ / ✅ / ✅ | ◐ / ❌ / ✅ | ✅ / ✅ / ✅ |
| Charts / chart sheets | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ❌ / ❌ / ❌ | ◐ / ❌ / ✅ | ◐ / ❌ / ✅ |
| Shapes, drawings, form controls | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ❌ / ❌ / ❌ | ◐ / ❌ / ✅ | ◐ / ❌ / ✅ |
| External links | ❌ / ❌ / ❌ | ✅ / ✅ / ✅* | ❌ / ❌ / ❌ | ◐ / ❌ / ✅ | ◐ / ❌ / ✅ |
| Pivot tables / caches | ❌ / ❌ / ❌ | ✅ / ✅ / ✅*⁴ | ❌ / ❌ / ❌⁵ | ◐ / ❌ / ✅ | ◐ / ❌ / ✅ |
| VBA / macros | ✅ / ✅ / ✅⁶ | ✅ / ✅ / ✅*⁷ | ❌ / ❌ / ❌ | ◐ / ❌ / ✅ | ◐ / ❌ / ✅ |
| Encryption / password-protected files | ❌ / ❌ / ❌⁸ | ✅ / ✅ / ✅*⁹ | ❌ / ❌ / ❌ | ✅ / ✅ / ✅ | ❌ / ❌ / ❌¹⁰ |

² SheetJS CE supports plain-text notes and threaded comments; note styling and rich text are Pro features `[S13]`. Our test preserved author and text `[E1]`.
³ ExcelJS supports old-style notes (`cell.note`) but has no threaded-comment model; our test preserved the text but reset the author to `"Author"` `[E1][S15]`.
⁴ SheetJS Pro lists "PivotTable and PivotChart generation and processing" as an add-on `[S11]`.
⁵ ExcelJS's only pivot work is PR #2551, merged 2023-10-31, *after* the 4.4.0 release on 2023-10-19; it is unreleased and described as "with limitations" `[S17][S18]`. The 4.4.0 source tree contains no pivot-table model.
⁶ SheetJS CE stores the raw VBA blob in `vbaraw` when `bookVBA: true` and writes it back to `.xlsm`/`.xlsb`/`.xls`; it does not parse or edit VBA `[S12]`. Our test confirmed the blob survives `read` → `write` as `.xlsm` `[E2]`.
⁷ SheetJS Pro sells VBA read/write and UserForms as a separate component `[S11]`.
⁸ SheetJS CE supports only XOR encryption for legacy `.xls`; XLSX/XLSM AES-CBC requires Pro `[S8]`.
⁹ SheetJS Pro lists password-protection and file encryption/decryption as an add-on `[S11]`.
¹⁰ Encrypted OOXML is not a ZIP at all: it is an OLE/CFB structured-storage container (`D0 CF 11 E0 A1 B1 1A E1`) holding the encryption streams `[S6]`, with ECMA-376 encryption described by MS-OFFCRYPTO `[S4][S5]`. A direct implementation needs a CFB reader plus Agile Encryption (AES/SHA) before it can even see the package.

**Key evidence.**

- **SheetJS CE drops validation and conditional formatting entirely.** Our fixture had a list validation on `C1:C5` and a `cellIs` rule on `E1:E5`; neither appeared in the CE output, and `read` exposed no `!dataValidations` or `!conditionalFormatting` field at all `[E1]`. This is consistent with the docs assigning those features to Pro `[S8][S11]`.
- **SheetJS CE drops charts and external links.** The fixture's `xl/charts/chart1.xml` and `xl/drawings/drawing1.xml` were absent from the output; a synthetic external-link part was also dropped `[E1][E2]`.
- **ExcelJS drops charts, external links, VBA and custom properties.** The chart drawing part and chart XML were absent after load→write; the external-link parts were absent; the VBA part was absent; `docProps/custom.xml` was absent `[E1][E2]`. This is an acknowledged open request: "[F] Accept and preserve files with charts" is open issue #2607 `[S19]`.
- **ExcelJS preserves images.** An image added via `addImage`, round-tripped through load→write, kept `xl/media/image1.png` and its drawing/relationship parts `[E3]`. Note that images are not supported in ExcelJS streaming mode `[S15]`.
- **xlsx-populate preserves all package parts.** In our fixture it retained the chart, drawing, custom properties, comments/VML, conditional formatting, validation and even synthetic pivot/external-link/VBA parts `[E1][E2]`. Its README explains why: it manipulates XML trees rather than deserializing into a lossy rich model `[S20]`.
- **xlsx-populate destroys cached formula results.** It intentionally does not write `<v>` for formula cells; the README warns that it "will not write the values to the output" `[S20]`. Our fixture's `A3` (`=SUM(A2:B2)`, cached `30.5`) came back with `<f>` but no `<v>` `[E1]`. This alone disqualifies it as the round-trip engine without a patch.

### 4.3 Non-functional

| Option | Read/write profile | Streaming in a browser | License | Browser constraints |
|---|---|---|---|---|
| SheetJS CE | Fastest and lowest memory in our test; dense worksheets; whole-buffer write | ❌ No streaming API; `writeFile` triggers a download | Apache-2.0 + attribution `[S10]` | Pure JS; works from `ArrayBuffer`; no Node APIs required |
| SheetJS Pro | Vendor claims Node/browser performance enhancements | Not documented publicly | Commercial, quote-based `[S11]` | Same as CE plus paid builds/components |
| ExcelJS 4.4 | Heaviest memory; full object graph; ~2.5× SheetJS RSS in our test | ❌ Browser bundle ships no `WorkbookWriter`; streaming uses Node streams/`archiver` `[S15][E4]` | MIT `[S16]` | Browser bundle needs `core-js` polyfills or a bare build; streams are Node-only |
| xlsx-populate 1.21 | Full XML DOM; ~2.5× SheetJS RSS; no streaming | ❌ No streaming API | MIT `[S21]` | Browserify bundle; encryption build is larger |
| Direct OOXML | Design-dependent; can stream import/export and keep a compact cell store | ✅ `fflate` supports streaming ZIP in the browser `[S22]` | Own code; MIT (fflate, JSZip) / ISC (saxes) `[S23][S24][S25]` | Needs a Web Worker for large files; File/Blob APIs; no Node dependencies |

**Measured (local, Node 24, x64 Linux, 12 cores):** a 50,000-row × 8-column workbook (400,000 cells, 2.1 MB input):

| Library | Read | Write | Peak RSS |
|---|---|---|---|
| SheetJS CE 0.20.3 (`dense: true`) | 1.42 s | 0.54 s | 324 MB |
| ExcelJS 4.4.0 | 1.24 s | 2.06 s | 800 MB |
| xlsx-populate 1.21.0 | 2.24 s | 1.98 s | 799 MB |

These are single-run, Node-only numbers and should be treated as ratios, not browser budgets. The important result is that **a modest 400k-cell file costs 0.3–0.8 GB of RSS** in the model-based libraries. In a browser tab, memory — not CPU — is the binding constraint. Ticket 11 owns the real performance budget; the I/O design should keep the option of lazy per-sheet parsing and streaming export.

## 5. Local round-trip experiment

### 5.1 Method

1. Built `fixture.xlsx` with openpyxl 3.1.5: two visible sheets plus one hidden sheet; text/number/formula/date/percent cells; a bold red 14 pt font, yellow solid fill, four different border styles, center/wrap alignment, `$#,##0.00` and `mm/dd/yyyy` formats; merged `B2:C2`; row 1 height 30; column A width 24; hidden column D; hidden row 10; defined name `TaxRate`; note on `A1`; hyperlink on `A1`; list validation `C1:C5`; `cellIs > 3` conditional formatting `E1:E5`; freeze pane `A2`; autoFilter `A1:E5`; a bar chart anchored at `G10`.
2. Re-saved through LibreOffice 24.2.7.2 to force a real cached formula value (`A3 = 30.5`) and real shared strings.
3. Round-tripped with SheetJS CE 0.20.3 (`cellStyles`, `cellNF`, `bookFiles`, `bookVBA`), ExcelJS 4.4.0 (`xlsx.load` → `writeBuffer`), and xlsx-populate 1.21.0 (`fromFileAsync` → `outputAsync`).
4. Compared ZIP part inventories, then re-opened outputs with openpyxl 3.1.5 (data and style inspection) and LibreOffice (CSV conversion) to confirm readability.
5. Separately injected synthetic external-link, pivot and VBA parts into copies of the fixture to test preservation of unfamiliar parts.

### 5.2 Results

| Check | Source | SheetJS CE | ExcelJS | xlsx-populate |
|---|---|---|---|---|
| Chart part survives | ✅ | ❌ | ❌ | ✅ |
| Cell font/fill/border/alignment | ✅ | ❌ (all flattened) | ✅ | ✅ |
| Number formats | ✅ | ✅ | ✅ | ✅ |
| Formula + cached `<v>` | ✅ `30.5` | ✅ `30.5` | ✅ `30.5` | ❌ `<f>` only |
| Merged cells / row heights / widths / hidden | ✅ | ✅ | ✅ | ✅ |
| Freeze panes | ✅ | ❌ | ✅ | ✅ |
| Defined name `TaxRate` | ✅ | ✅ | ✅ | ✅ |
| Hidden sheet | ✅ | ✅ | ✅ | ✅ |
| Comment text / author | ✅ | ✅ / ✅ | ✅ / ❌ (`Author`) | ✅ / ✅ |
| Hyperlink | ✅ | ✅ | ✅ | ✅ |
| Data validation | ✅ | ❌ | ✅ | ✅ |
| Conditional formatting | ✅ | ❌ | ✅ | ✅ |
| External-link parts | ✅ (injected) | ❌ | ❌ | ✅ |
| Pivot parts | ✅ (injected) | ❌ | ❌ | ✅ |
| VBA part | ✅ (injected) | ✅ (`bookVBA`, `.xlsm`) | ❌ | ✅ |
| `docProps/custom.xml` | ✅ | ❌ | ❌ | ✅ |
| Output opens in LibreOffice | ✅ | ✅ | ✅ | ✅ |
| Output opens in openpyxl | ✅ | ❌ (invalid `level` attribute on `<col>`) | ✅ | ✅ |

The SheetJS output also wrote literal strings as `t="str"` (the "formula string" cell type) rather than shared/inline strings. Excel and LibreOffice tolerated it, but it is another example of an exporter optimised for data extraction rather than schema fidelity.

## 6. Recommendation

### 6.1 Architecture

Build an in-house module — working name `ooxml` — with these boundaries:

```
ooxml/
  package/        OPC ZIP, content types, relationships, part inventory
  parts/          workbook, worksheet, styles, sharedStrings, comments
  model/          mapping to/from calcplus workbook model
  preserve/       opaque PartStore for unmodeled parts
  write/          dirty-part serializer, append-only tables, export report
```

**Dependencies:** `fflate` for ZIP (streaming, browser-native, MIT) `[S22][S23]`; `saxes` for SAX parsing (ISC, worker-safe) `[S25]`. JSZip is acceptable for prototypes and tests (dual MIT/GPLv3 — choose MIT) `[S24]`, but it materialises the whole archive and is not a streaming reader. Use `saxes` rather than `DOMParser` so parsing can live in a Web Worker; `DOMParser` is a `Window` API, not part of `WorkerGlobalScope` `[S30]`.

### 6.2 Load algorithm

1. Read the file as an `ArrayBuffer`. Check the first 8 bytes: CFB magic `D0 CF 11 E0 A1 B1 1A E1` means encrypted/legacy OLE → reject with the Tier C message `[S6]`.
2. Parse the ZIP central directory; build a `PartStore` of `{name, compression, compressed range, uncompressed size, CRC}`. Tier B parts are not decompressed unless they must be inspected.
3. Parse `[Content_Types].xml` and `_rels/.rels`; locate the workbook part via the officeDocument relationship.
4. Parse `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`: sheet names, sheet IDs, relationship IDs, states, defined names, `date1904`, calculation properties.
5. Parse `xl/styles.xml` into typed tables **while retaining raw XML and original indices**. Parse `xl/sharedStrings.xml` into an index-preserving array.
6. Parse worksheets lazily: the active sheet first, others on demand. Maintain a compact cell store (dense row arrays or sparse maps) with `styleIndex`, `valueType`, raw value, formula text, formula attributes, cached value and a dirty flag.
7. Record every unparsed part in the `PreserveSet`. Do not normalize, re-indent or reserialize it.

### 6.3 Save algorithm

1. If the workbook has no dirty state and no model-level edits, **return the original `ArrayBuffer` unchanged**.
2. For each part in the package:
   - Not dirty → copy its raw compressed entry where the ZIP writer permits, otherwise re-deflate the identical uncompressed bytes. (The container is rebuilt, so file-level bytes differ; part content does not.)
   - Dirty → serialize from the model.
3. `workbook.xml`: update sheet names/order/states and defined names. Never renumber `sheetId` values unnecessarily.
4. `styles.xml`: **append-only**. New `numFmt`, `font`, `fill`, `border`, `cellXf` and `dxf` entries go at the end; existing indices are untouched. Update `count` attributes.
5. `sharedStrings.xml`: **append-only**. Keep existing indices; unused entries are permitted. Rebuild the `count`/`uniqueCount` attributes from actual usage.
6. Worksheet parts: emit the original root element with its namespace declarations; preserve unknown children (`sheetPr`, `extLst`, `legacyDrawing`, etc.) and their order; rewrite only dirty rows/cells. For untouched cells, reuse the original `s` index and formula XML. For a dirty formula cell, write the engine's cached result only if the engine actually evaluated it; otherwise retain the original `<v>` and mark it for recalculation. Update `dimension`, `mergeCells`, `hyperlinks`, `dataValidations`, `conditionalFormatting`, `autoFilter`, `sheetViews/pane`.
7. Comments: update the comment parts and VML anchors only for edited/added notes; preserve existing authors and shapes.
8. Delete `xl/calcChain.xml` (and its content-type/relationship) whenever formulas change, and set `fullCalcOnLoad="1"` in `calcPr`. A stale calc chain is worse than a first-open recalculation.
9. Update `[Content_Types].xml` and relationship parts only for added/removed parts. Keep existing relationship IDs stable so chart/drawing/hyperlink references remain valid.
10. Write the ZIP with `fflate`; stream large worksheet XML rather than building a giant string.
11. Produce an **export report**: list Tier B parts present, and any structural edits that may have invalidated their ranges.

### 6.4 Structural edits are the hard part

Range-bearing constructs are spread across many parts: merged ranges, validations, conditional formats, defined names, tables, drawing anchors, chart series references, pivot cache sources and external-link formulas. Inserting or deleting rows/columns must go through a single **range-transform module** that rewrites all of them consistently.

For V1, decide explicitly:

- **If insert/delete rows and columns are in scope:** implement transforms for merged ranges, validations, conditional formats, names, tables and drawing anchors. Chart series references and pivot cache sources should be transformed where feasible; otherwise mark them stale in the export report.
- **If they are not in scope:** disable those commands on sheets containing Tier B objects (or all sheets), and say why in the UI. Do not silently let anchors drift.

This is a cross-ticket dependency with the workbook-model (07) and command/undo (09) tickets.

### 6.5 Policy decisions to record in the spec

| Topic | Recommendation |
|---|---|
| No-op save | Return the original bytes exactly. |
| Unknown parts | Byte-preserve by default; never drop silently. |
| Charts/images/shapes | Preserve, do not render in V1. v2 adds rendering/editing. |
| Pivot tables | Preserve definition and cache; never refresh; export report notes staleness. |
| External links | Preserve; never resolve or refresh. |
| Macros | Preserve `vbaProject.bin` in `.xlsm`, never execute or edit; warn that signatures are invalidated by any save. |
| Encryption/IRM | Reject with a clear message in V1; a CFB+Agile module is a later feature. |
| Digital signatures | Any edit invalidates them; warn on save. |
| File size | Hard cap on compressed and uncompressed size, entry count and compression ratio to resist zip bombs; reject above cap. |
| Export warnings | One report per save; distinguish "dropped", "preserved but stale", "preserved but not displayed". |

### 6.6 Acceptance tests for the I/O layer

1. **Identity:** open→save with no edits returns a byte-identical file.
2. **Repair-free:** every output opens in Excel 365 without a repair prompt (manual/CI job), and in LibreOffice for automated checks.
3. **Part inventory:** for an edited file, every Tier B part is present with identical uncompressed bytes and resolvable relationships.
4. **Semantic equality:** styles, number formats, merges, dimensions, names, sheets, comments, hyperlinks, validations and conditional formats compare equal to the source via an independent reader (openpyxl or a small verifier).
5. **Formula fidelity:** formula text, shared/array attributes and cached values survive an edit to a different cell in the same sheet.
6. **Structural edits:** after inserting/deleting rows, all range-bearing constructs point at the intended ranges; charts/pivots are either transformed or reported stale.
7. **Hostile inputs:** encrypted, truncated, zip-bomb, XXE-style and million-row files fail safely within the size/time budget.
8. **Corpus:** Excel 365, Excel 2019, LibreOffice, Google Sheets, openpyxl, SheetJS and xlsx-populate outputs; files with charts, images, pivots, external links, VBA, threaded comments and hidden sheets.

## 7. What the recommended approach cannot preserve

This is the honest boundary of the recommendation. "Preserved" means the bytes survive; it does **not** mean calcplus understands, renders or keeps the object semantically correct after arbitrary edits.

### 7.1 Cannot open at all

- **Password-encrypted and IRM-protected workbooks.** They are CFB/OLE containers, not ZIP packages `[S4][S5][S6]`. V1 should detect and reject them. Supporting them requires a CFB reader plus ECMA-376 Agile Encryption (AES/SHA-1/SHA-512) in the browser.
- **`.xls` (BIFF8) and `.xlsb` (BIFF12).** Out of scope per the effort map; do not pretend otherwise.
- **Corrupt, truncated, non-ZIP or zip-bomb files.** Reject above explicit caps.
- **ISO Strict OOXML variants** may use different namespaces; this was not tested. Treat as an evidence gap and either map or reject explicitly.

### 7.2 Preserved as bytes, but not understood, rendered or edited

- **Charts and chart sheets.** XML survives; calcplus does not draw them in V1. Chart series references are not updated if source ranges move. This is the biggest visible gap for a user opening a chart-heavy workbook.
- **Images, shapes, drawings, OLE objects and form controls.** Parts survive; nothing is displayed or movable. Drawing anchors are not adjusted on row/column insert/delete unless the anchor transform is built.
- **Pivot tables and pivot caches.** Definitions and cached records survive; calcplus shows only the pivot's output cells, cannot refresh, cannot change fields, and cannot keep the cache in sync with edited source data.
- **External links.** Parts survive; referenced workbooks are never opened or refreshed; formulas keep their cached values.
- **VBA projects.** The blob survives in `.xlsm`; macros are never executed or edited. Any save invalidates digital signatures. Preserving macros in a browser product is also a security/product decision that must be made explicitly.
- **Digital signatures.** Any modification invalidates them.
- **Slicers, timelines, sparklines, custom XML, third-party extensions.** Bytes survive; behavior does not.
- **Tables** (if not modeled in V1). Table definitions survive, but inserted/deleted rows can leave `tablePart` refs stale.

### 7.3 Preserved, but semantically limited

- **Formulas outside calcplus's engine.** The cached `<v>` is preserved, but calcplus cannot recompute the formula. Excel will recalculate on open; other viewers may show the cached value.
- **Shared, array and dynamic-array formula metadata.** Only preserved if the writer re-emits the `<f>` attributes (`t`, `ref`, `si`, `aca`, `ca`) rather than flattening every formula to a plain string. This must be a model requirement, not an afterthought.
- **Rich text.** Untouched runs survive; editing a rich-text cell flattens it to plain text unless run-level formatting is modeled.
- **Styles.** Existing style tables survive, but new cells can only use style properties calcplus models. Gradient fills, named cell styles and theme-color transforms that are not modeled cannot be authored (they can still be preserved on untouched cells).
- **Conditional formatting.** Rules survive; calcplus may not evaluate or display them in V1, so the grid can look different from Excel.
- **Data validation.** Rules survive; input messages, error alerts and dropdown behavior may not be enforced by calcplus's editor.
- **Threaded comments.** The parts survive, but replies, mentions and resolved state are not editable in V1.
- **Sheet protection.** Settings survive; if calcplus does not enforce them, it can create edits Excel would have blocked.
- **Print settings, headers/footers, page breaks.** Preserved; not editable in V1.
- **Locale-specific number formats.** Format strings are preserved; rendering depends on the formatter's coverage of the format language.
- **Document properties.** Core/app properties may be updated; custom properties are preserved but not editable.

### 7.4 Things no approach preserves

- **The user's intent** behind an unmodeled object. If a user deletes the rows a chart points at, the chart's XML can be updated mechanically, but only a human knows whether the chart should shrink, move or disappear.
- **Excel's exact recalculation result** for functions calcplus does not implement. Cached values are a snapshot, not a guarantee.
- **A valid signature** after any modification.

## 8. Fallbacks if in-house I/O is not funded

Ranked by fitness for this product:

1. **SheetJS Pro, Pro Edit plus the styling/image/chart builds and pivot/encryption add-ons.** The only commercial option whose published feature list covers styling, validation, conditional formatting, images, charts, fidelity editing, pivots and encryption `[S11]`. Risks: quote-only pricing, closed source, per-build licensing, unknown bundle size, and "perfect fidelity" is a vendor claim we could not test. This is the pragmatic fallback if the schedule cannot absorb the I/O work.
2. **ExcelJS + a post-process preservation shim.** ExcelJS is excellent for data, styles, images, validation and conditional formatting `[S15]`, but its writer regenerates the package from its own model, so charts, external links, pivots and VBA are gone. A shim would have to merge original parts and patch relationships/drawing anchors — essentially the direct-OOXML work, with two writers fighting. Not recommended, but viable for a "data and styles only" interim release that explicitly disclaims chart/pivot fidelity.
3. **xlsx-populate fork.** It already preserves unknown parts and supports style/DV/encryption `[S20][S21]`, so a fork could patch the one fatal behavior (cached formula values) and add a bulk API. Risks: low maintenance activity, full-DOM memory profile, no streaming, and a cell-by-cell API that does not fit a spreadsheet model. Useful as a **reference implementation** even if not adopted.
4. **SheetJS CE.** Suitable only as a data import/export helper, not as the fidelity engine: no cell styles, validation, conditional formatting, charts or freeze panes; emits a non-schema column attribute `[S8][S9][E1]`.

## 9. Evidence gaps and open questions

- **SheetJS Pro was not tested** (closed-source, quote-only). All Pro ratings are vendor claims from the product page and docs `[S11]`.
- **ExcelJS pivot support is in flux.** PR #2551 is merged to master but unreleased as of 4.4.0; a future release may add limited creation, not necessarily round-trip preservation `[S17][S18]`.
- **Excel behavior under the SheetJS `level` attribute** was not tested in Excel; LibreOffice tolerated it, openpyxl rejected it. Treat as an interoperability risk.
- **Encryption performance in the browser** was not measured. A future CFB/Agile module needs its own spike.
- **Strict OOXML** support was not tested.
- **Chart/pivot range transformation** after structural edits needs a dedicated design spike; it is the largest hidden cost in the preservation-first approach.
- **Browser memory ceilings** vary by browser/device; the Node benchmark should be repeated in Chromium/Firefox/Safari with the real UI loaded.
- **Real-world corpus:** the fixture is synthetic. Before committing to the contract, run a corpus of Excel-generated files, including files with charts, pivots, macros and threaded comments.

## 10. Sources

| ID | Source |
|---|---|
| S1 | ECMA-376, *Office Open XML file formats*, 5th edition — https://ecma-international.org/publications-and-standards/standards/ecma-376/ |
| S2 | Microsoft Learn, *Structure of a SpreadsheetML document* — https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document |
| S3 | Microsoft Learn, *Working with formulas* (CellValue `<v/>` stores the cached formula value) — https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas |
| S4 | Microsoft Learn, *[MS-OFFCRYPTO]: Office Document Cryptography Structure* — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/3c34d72a-1a61-4b52-a893-196f9157f083 |
| S5 | Microsoft Learn, *[MS-OFFCRYPTO]: ECMA-376 Document Encryption* — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/4f2e28a1-624a-4bed-b3ec-816f0df484ff |
| S6 | Microsoft Learn, *[MS-CFB]: Compound File Header* (magic `D0 CF 11 E0 A1 B1 1A E1`) — https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea |
| S7 | SheetJS, *File Formats* — https://docs.sheetjs.com/docs/miscellany/formats |
| S8 | SheetJS, *Reading Files* (parse options; encryption; Pro styling note) — https://docs.sheetjs.com/docs/api/parse-options |
| S9 | SheetJS, *Writing Files* (write options; Pro styling/images/graphs/PivotTables note) — https://docs.sheetjs.com/docs/api/write-options |
| S10 | SheetJS, *License* (Apache-2.0, attribution) — https://docs.sheetjs.com/docs/miscellany/license |
| S11 | SheetJS, *Professional Edition* (Pro Basic/Image/Chart/Edit, pivot and encryption add-ons) — https://sheetjs.com/pro |
| S12 | SheetJS, *VBA and Macros* (`bookVBA` / `vbaraw`) — https://docs.sheetjs.com/docs/csf/features/vba |
| S13 | SheetJS, *Cell Comments and Notes* — https://docs.sheetjs.com/docs/csf/features/comments |
| S14 | SheetJS, *Cell Objects* — https://docs.sheetjs.com/docs/csf/cell |
| S15 | ExcelJS, *README* (features, styles, validation, conditional formatting, images, streaming, browser bundles) — https://github.com/exceljs/exceljs/blob/master/README.md |
| S16 | ExcelJS, *LICENSE* (MIT) — https://github.com/exceljs/exceljs/blob/master/LICENSE |
| S17 | ExcelJS, *v4.4.0 release* (2023-10-19) — https://github.com/exceljs/exceljs/releases/tag/v4.4.0 |
| S18 | ExcelJS PR #2551, *Add pivot table with limitations* (merged 2023-10-31, unreleased) — https://github.com/exceljs/exceljs/pull/2551 |
| S19 | ExcelJS issue #2607, *[F] Accept and preserve files with charts* (open) — https://github.com/exceljs/exceljs/issues/2607 |
| S20 | xlsx-populate, *README* (preservation approach, formula cached-value caveat, styles, validation, encryption) — https://github.com/dtjohnson/xlsx-populate/blob/master/README.md |
| S21 | xlsx-populate, *LICENSE* (MIT) — https://github.com/dtjohnson/xlsx-populate/blob/master/LICENSE |
| S22 | fflate, *README* (streaming ZIP in the browser; performance) — https://github.com/101arrowz/fflate/blob/master/README.md |
| S23 | fflate, *LICENSE* (MIT) — https://github.com/101arrowz/fflate/blob/master/LICENSE |
| S24 | JSZip, *LICENSE* (MIT or GPLv3) — https://github.com/Stuk/jszip/blob/master/LICENSE.markdown |
| S25 | saxes, *package metadata* (ISC) — https://github.com/lddubeau/saxes/blob/master/package.json |
| S26 | Luckyexcel (MIT, one-way xlsx→Luckysheet import) — https://github.com/dream-num/Luckyexcel |
| S27 | Univer (Apache-2.0, full spreadsheet SDK) — https://github.com/dream-num/univer |
| S28 | xlsx-js-style (Apache-2.0, SheetJS 0.18.5-based style fork) — https://github.com/gitbrent/xlsx-js-style |
| S29 | SheetJS, *Installation* (`cdn.sheetjs.com` is the primary distribution site) — https://docs.sheetjs.com/docs/getting-started/installation/ |
| S30 | MDN, *WorkerGlobalScope* (the worker global interface; `DOMParser` is a `Window` API) — https://developer.mozilla.org/en-US/docs/Web/API/WorkerGlobalScope |

### Local experiments

| ID | Description |
|---|---|
| E1 | Round-trip of `fixture.xlsx` (openpyxl 3.1.5 + LibreOffice 24.2.7.2 recalc) through SheetJS CE 0.20.3, ExcelJS 4.4.0 and xlsx-populate 1.21.0; ZIP part inventories, openpyxl style/comment/validation/CF inspection, LibreOffice CSV open. |
| E2 | Synthetic parts injected into copies of the fixture: external-link (`xl/externalLinks/*`), pivot (`xl/pivotTables/*`, `xl/pivotCache/*`), VBA (`xl/vbaProject.bin` + content type/relationship + `.xlsm`). Round-tripped through the same three libraries; parts checked in output. |
| E3 | Image round-trip: ExcelJS `addImage` → load → `writeBuffer`; `xl/media/image1.png` and drawing rels present in output. |
| E4 | Browser-bundle inspection: ExcelJS 4.4.0 `dist/exceljs.min.js` contains no `WorkbookWriter`; streaming implementation lives in `lib/stream` and depends on Node streams/`archiver`. |
| E5 | SheetJS rich-text write check: a cell with an `r` rich-text value was written as `t="str"` with a plain `<v>`; runs were not emitted. |

**Benchmark environment:** Node 24.15.0, x64 Linux, 12 logical cores, 31 GB RAM. Input: 50,000 rows × 8 columns (400,000 cells), 2.1 MB `.xlsx`. Single run per library; RSS is `time -v` maximum resident set size. Node-only, not a browser measurement.

## Appendix A — Reproduction sketch

The fixture can be rebuilt with openpyxl (styles, merges, dimensions, hidden rows/cols, defined name, note, hyperlink, list validation, `cellIs` conditional format, freeze pane, autoFilter, bar chart, hidden sheet) and then recalculated once with LibreOffice to embed a cached formula value. The round-trip harness is three small Node scripts:

```js
// SheetJS CE
const wb = XLSX.read(buf, { type: 'buffer', cellStyles: true, cellNF: true,
  bookFiles: true, bookVBA: true });
fs.writeFileSync(out, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx',
  cellStyles: true, bookVBA: true, compression: true }));

// ExcelJS
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(buf);
fs.writeFileSync(out, await wb.xlsx.writeBuffer());

// xlsx-populate
const wb = await XlsxPopulate.fromFileAsync(src);
fs.writeFileSync(out, await wb.outputAsync());
```

Inspect part inventories with `JSZip.loadAsync(buf).then(z => Object.keys(z.files))`, and inspect semantics with openpyxl (`load_workbook`, `data_only=True` for cached values). For preservation tests, inject synthetic parts into a copy of the fixture with Python's `zipfile` and re-run the three round-trips.
