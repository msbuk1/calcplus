# 27 — Chart, drawing-anchor and pivot range transforms after structural edits

Status: research complete
Date: 2026-11-18
Ticket: milestone 6, slice 6D (structural range transforms)
Method: primary schema and vendor documentation (ECMA-376 via the c-rex rendering of the 1st edition, Microsoft open specifications, Microsoft support articles), the openpyxl source and documentation as an explicit XML-shape mapper, EPPlus source for pivot paths, and the existing calcplus contracts. No local fixture experiments were run; every claim that would need an Excel fixture check is marked UNVERIFIED.

## TL;DR

A single range-transform module can cover every range-bearing construct the file-compatibility contract lists, because each construct stores its range in exactly one predictable place. The module needs three small parsers, not a general formula engine: an A1 range parser for worksheet attributes, a chart-`c:f` classifier (plain range versus name or expression), and an anchor-marker shifter for drawing parts. The recommended decision set:

1. Rewrite every `c:f` whose text parses as a plain workbook, sheet-qualified, or unqualified A1 range. Rewrite rule is the same interval arithmetic as worksheet ranges.
2. Block the edit with an explanation when a `c:f` does not parse as a plain range (defined name, function call such as `OFFSET`, structured reference, or external-workbook reference).
3. Move `xdr:from` and `xdr:to` markers of `twoCellAnchor` and the `xdr:from` marker of `oneCellAnchor` with the grid. Never attempt to move `absoluteAnchor` (positioned in English Metric Units, independent of cells); an edit that must move it is blocked.
4. Never rewrite `pivotCacheDefinition` source references and never refresh caches. Block any insert or delete that intersects the pivot source range or the pivot output location, and mark the pivot "preserved but not refreshed" in the export report otherwise.
5. Leave chart `numCache` and `strCache` bytes untouched when rewriting `c:f`, and report the series as transformed-but-stale so Excel rebuilds the cache on open. Confirm with an Excel fixture that a rewritten `c:f` with a stale cache opens repair-free (UNVERIFIED).

## 1. Chart series reference shapes

### 1.1 Where the range text lives

Every chart data binding in `xl/charts/chart*.xml` stores its source as a child element containing exactly one `c:f` element whose text is the range, plus an optional cache of the last values used [S1][S2]. The concrete nesting depends on the data role:

| Data role | Parent chain | Reference child | Cache child |
|---|---|---|---|
| Category labels | `c:cat` | `c:numRef`, `c:strRef`, or `c:multiLvlStrRef` | `c:numCache` / `c:strCache` / `c:multiLvlStrCache` |
| Values (bar, line, area, pie, radar, surface) | `c:val` | `c:numRef` (or `c:numLit`) | `c:numCache` |
| Scatter X | `c:xVal` | `c:numRef`, `c:strRef`, `c:multiLvlStrRef` (AxDataSource) | matching cache |
| Scatter Y | `c:yVal` | `c:numRef` (or `c:numLit`) | `c:numCache` |
| Bubble size | `c:bubbleSize` | `c:numRef` (or `c:numLit`) | `c:numCache` |
| Series name | `c:tx` | `c:strRef` (or literal `c:v`) | `c:strCache` |
| Error-bar plus/minus | `c:plus` / `c:minus` | `c:numRef` (or `c:numLit`) | `c:numCache` |

Schema facts behind this table: `CT_NumDataSource` is a choice of exactly one `numRef` or `numLit` [S3]; `CT_SerTx` (series text) is a choice of exactly one `strRef` or literal `v` text [S4]; `CT_NumRef` is the sequence `f`, optional `numCache`, optional `extLst` [S5]; `c:numRef` appears under `bubbleSize`, `cat`, `minus`, `plus`, `val`, `xVal`, and `yVal` [S5]. openpyxl models the same split as `NumDataSource(numRef|numLit)`, `AxDataSource(numRef|numLit|strRef|strLit|multiLvlStrRef)`, `NumRef(f, numCache)`, `StrRef(f, strCache)`, and `SeriesLabel(strRef, v)` [S6][S7].

Per chart type, the `c:ser` children that can carry ranges are:

- Bar, line, area, radar, surface, pie, doughnut, stock: `c:tx`, `c:cat`, `c:val`. (`c:val` parents include the bar, line, pie, surface, radar, and area series elements [S3].)
- Scatter: `c:tx`, `c:xVal`, `c:yVal` (plus `c:smooth` flag, not a range). The scatter chart contains a `c:ser` of type `CT_ScatterSer` [S8]; real-world scatter XML shows `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$7</c:f>` beside `<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$7</c:f>` [S9].
- Bubble: `c:tx`, `c:xVal`, `c:yVal`, `c:bubbleSize`.
- Pie has no `c:xVal`/`c:yVal`; its slices come from `c:cat` plus `c:val`.
- Combo and secondary-axis charts are multiple chart-type elements sharing one `c:plotArea`; each `c:ser` is handled by its own chart type's rule above. No special casing is needed beyond iterating all `c:ser` elements in the part.

Literal points need no transform: `c:numLit` and `c:strLit` (with `c:pt` children and `c:ptCount`) and series-name literal `c:v` contain inline values, not cell references [S4][S6].

### 1.2 What the range text looks like

The `c:f` reference form is "a book, sheet, and cell reference or a book, optional sheet, and defined name reference" and it does not include the leading equals sign [S2]. Observed and documented forms:

- `Sheet1!$A$1:$C$1` — unquoted sheet plus absolute range (the worked example in the schema reference) [S2].
- `Blatt1!$A$1:$A$12` — same shape, non-English sheet name (openpyxl test fixture) [S9].
- `'5.Tiers.Hourly'!$F$19:$F$40` — single-quoted sheet name (required when the name contains dots, spaces, or other specials), seen in real scatter-chart XML [S9].
- `Sheet1!Sales` or bare `Sales` — sheet-qualified or workbook-scoped defined name (the Microsoft dynamic-chart article directs users to type exactly `=Sheet1!Sales` into the Series values box, which Excel stores as the `c:f` text without the `=`) [S10].
- `OFFSET($A$2,0,0,COUNTA($A:$A)-1)` wrapped as a name body — the dynamic-range pattern; the `c:f` then holds the name, while the `OFFSET` expression lives in the defined-name record, not in the chart [S10].

Dollar signs are the norm in chart `c:f` text but are not guaranteed; the transform must apply regardless of `$`, because in stored references `$` affects copy and fill behavior, not insert and delete shifting (Excel adjusts references to shifted cells to reflect their new locations [S11]).

### 1.3 Stale caches

`c:numCache` and `c:strCache` hold the last values used: `formatCode`, `ptCount`, and per-point `c:pt` values [S5][S6]. After the module rewrites `c:f`, the cache no longer matches the range. The recommendation is to leave the cache bytes untouched and report the series as transformed-but-stale, letting Excel rebuild the cache on open. Rationale: deleting or zeroing the cache risks tripping `ptCount` consistency checks in strict readers, while a stale cache is the same condition Excel itself produces between edits and recalculation. UNVERIFIED: confirm with an Excel 365 fixture that a rewritten `c:f` beside an untouched stale cache opens without a repair prompt. If Excel repairs, the fallback is to delete the cache element and keep only `c:f` (the schema makes the cache optional: `minOccurs="0"` [S5]), then re-check repair-free opening.

## 2. Defined-name-backed and other non-range chart sources

A chart is name-backed when any `c:f` in its `c:ser` elements names a defined name instead of cells. The Microsoft dynamic-chart technique (`=Sheet1!Sales`, `=Sheet1!Date`) is exactly this case [S10]. Other non-range `c:f` contents possible in the wild include function expressions, structured table references, and external-workbook references of the form `[Book]Sheet!$A$1`.

### Detection rule the implementer can code

For each `c:f` text `t`:

1. Strip one optional sheet qualifier: a leading `'...'` (with `''` escape) or unquoted sheet token followed by `!`. A leading `[book]` bracket qualifier forces the external case immediately.
2. Split the remainder on the union operator (`,` or `;` depending on locale is NOT relevant here; chart `c:f` uses `,` for multi-area references in some producers — treat any remainder containing `,` as multi-area and attempt to classify each area; if any area is non-range, the whole `c:f` is non-range). UNVERIFIED: multi-area chart `c:f` text is rare; confirm against real Excel output before implementing multi-area rewriting, otherwise treat any `c:f` containing `,` as non-range and block.
3. If every area parses as an A1 cell or range (`$?COL$?ROW(:$?COL$?ROW)?`, including whole-column `A:A` and whole-row `3:3` forms), it is a plain range: rewrite.
4. Otherwise it is a non-range source: defined name (bare `Sales`, qualified `Sheet1!Sales` where `Sales` is not a cell address), function call (contains `(`), structured reference, or external reference. Block the edit with an explanation naming the series and the `c:f` text.

The ambiguous case is `Sheet1!Sales` versus `Sheet1!$A$1`: the rule resolves it deterministically because classification happens on the part after the `!`, and `Sales` never parses as an A1 address while `$A$1` always does. A pathological defined name that looks like a cell address (for example a name literally called `A1`) cannot exist, because Excel forbids defined names that collide with cell addresses; the parser therefore never misfires. Whole-column and whole-row references (`$A:$A`, `COUNTA($A:$A)` appears inside name bodies, not in `c:f`) classify as ranges but are only shifted by edits on the perpendicular axis (see section 6).

Error-bar `c:plus`/`c:minus` and `c:bubbleSize` `c:f` values use the same classifier. Trendlines carry no `c:f` range of their own. The chart title (`c:title//c:tx//c:rich`) and data labels are literal text, not ranges, and are never rewritten.

## 3. Drawing anchor shapes

Worksheet drawings (`xl/drawings/drawing*.xml`, root `xdr:wsDr`) position every chart, image, shape, and control with exactly one of three anchor children [S12][S13][S14]:

```xml
<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>
            <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>
          <xdr:row>11</xdr:row><xdr:rowOff>114300</xdr:rowOff></xdr:to>
  <!-- graphicFrame / sp / pic / cxnSp / grpSp / contentPart, then clientData -->
</xdr:twoCellAnchor>
```

### 3.1 The three anchors

| Anchor | Position children | Size | Moves with the grid? | Transform decision |
|---|---|---|---|---|
| `xdr:twoCellAnchor` | `xdr:from` + `xdr:to` (`xdr:col`, `xdr:row` zero-based cell indices; `xdr:colOff`, `xdr:rowOff` in English Metric Units offsets within the cell) | implied by from/to | yes | Transform: shift both markers by the same interval rules as ranges |
| `xdr:oneCellAnchor` | `xdr:from` only | `xdr:ext` (`cx`, `cy` in English Metric Units) | yes, position moves; size fixed | Transform: shift the `from` marker; never touch `ext` |
| `xdr:absoluteAnchor` | `xdr:pos` (`x`, `y` in English Metric Units from the sheet origin) | `xdr:ext` | no, fixed position independent of cells | Cannot transform: block the edit |

The marker model is confirmed by openpyxl's `AnchorMarker(col, colOff, row, rowOff)`, `TwoCellAnchor(_from, to)`, `OneCellAnchor(_from, ext)`, and `AbsoluteAnchor(pos, ext)` where `pos` is an `XDRPoint2D` and `ext` an `XDRPositiveSize2D` [S14]. Primer prose states the semantics directly: absolute anchoring keeps the object in the same position and does not move or size with cells, while one-cell and two-cell anchors move with cells [S12]. Each anchor then carries one content child (`sp`, `grpSp`, `graphicFrame` for charts, `cxnSp`, `pic` for images, or `contentPart`) plus `clientData` (`fLocksWithSheet`, `fPrintsWithSheet`) [S14]; content children never contain cell ranges and are copied verbatim.

### 3.2 The `editAs` attribute

`xdr:twoCellAnchor` carries `editAs` with values `twoCell`, `oneCell`, or `absolute`, controlling how Excel resizes the object when the user edits the sheet [S14]. `editAs` does not change the transform: both markers still live in cell coordinates and both shift. UNVERIFIED: whether Excel-generated files ever emit `editAs="absolute"` on a two-cell anchor whose markers should then be treated as frozen. The safe rule is to transform markers regardless of `editAs` and note the attribute value in the export report; if fixture testing shows Excel freezing `editAs="absolute"` markers, escalate that case to blocked.

### 3.3 What happens to markers on insert and delete

Markers are zero-based cell coordinates, so they obey the same interval arithmetic as range endpoints (section 6), with offsets (`colOff`, `rowOff`) preserved unchanged. Cases:

- Insertion point at or before a marker: the marker shifts by the inserted count.
- Insertion strictly inside a two-cell span (between `from` and `to`): the `from` marker stays, the `to` marker shifts, so the object grows. This matches Excel's move-or-size-with-cells behavior.
- Deletion entirely before a marker: the marker shifts back by the deleted count.
- Deletion covering a marker: clamp the marker to the first surviving row or column at the deletion point and report the anchor as transformed-but-moved. UNVERIFIED: Excel's exact clamping for a chart whose start row is deleted needs a fixture check; the fallback if Excel instead deletes or errors the object is to block that specific edit.
- Deletion covering the entire two-cell span: the anchor collapses to zero height at the deletion point. UNVERIFIED: confirm Excel keeps versus drops the object; if Excel drops it, the module must report the drawing as removed-or-stale, never silently keep a zero-size anchor without a report entry.

`absoluteAnchor` (`xdr:pos` in English Metric Units) has no cell coordinates to shift, so any row or column insert or delete on its sheet is the contract's third blocked case [S15]. In practice absolute anchors are rare in Excel output (Excel normally writes two-cell anchors); the block should therefore fire rarely, but the check must still run on every structural edit.

VML legacy drawing anchors (`xl/drawings/vmlDrawing*.vml`, legacy comments) use a different `x:Anchor` comma-separated coordinate form and are out of scope for transformation in v1: legacy notes are read-only Tier A and threaded comments are preserved Tier B [S15]. If a structural edit would move a legacy comment anchor, report it stale rather than rewriting VML. UNVERIFIED: exact Excel VML anchor update behavior; hence report-only.

## 4. Pivot cache and pivot table source references

### 4.1 Where the source lives

The pivot source range lives in the cache definition part, not in the worksheet:

```xml
<x:cacheSource type="worksheet">
  <x:worksheetSource ref="A1:X2" sheet="dataWorksheet"/>
</x:cacheSource>
```

The `worksheetSource` element carries `ref` (the source range), `sheet` (the source sheet name), and optionally `name` (a defined name) plus `id` [S16]. Office requires exactly one of `ref` and `name`: the standard permits both attributes, but Office requires one and only one of them; `id` in Office names a relationship to an external workbook rather than a sheet part [S17]. EPPlus reads exactly these two paths (`cacheSource/worksheetSource/@sheet` and `@ref`) [S18]. A `name`-form source is a defined-name indirection and is therefore doubly untransformable: the module neither resolves names into ranges in v1 nor rewrites them.

The pivot's on-sheet footprint is separate: `xl/pivotTables/pivotTable*.xml` root `pivotTableDefinition` carries a `location` child with `ref` (for example `ref="B6:G13"`, with `firstHeaderRow`, `firstDataRow`, `firstDataCol`, `rowPageCount`, `colPageCount`) [S19].

### 4.2 The no-rewrite rule and what "report as stale" means

The contract is explicit: definitions and caches are preserved, never refreshed, and a structural edit that would change the pivot's source range is blocked with an explanation [S15]. Concretely:

- Before applying any row or column insert or delete on sheet S, the module checks the pivot source (`worksheetSource/@ref` on S, or `@name` resolving to S) and the pivot output `location/@ref` on S. Any edit whose affected interval intersects either range, or sits before either range (which would shift it), is blocked with a message naming the pivot table and the source range.
- An edit on sheet S that touches neither the source nor the output location may proceed; the export report still lists the pivot as "preserved but not refreshed" because its cached records may no longer match source data the user edited by hand [S15].
- The user-visible report entry must name the part (`xl/pivotTables/pivotTableN.xml`), the source range, and the reason ("pivot caches are never refreshed in v1; the table shows stored values"). This satisfies the "no silent stale anchor" exit criterion by making staleness explicit rather than silent.

Open question for the project (section 8): whether an edit far from both ranges that only shifts the output `location` (insert above the pivot output but below the source data) should rewrite `location/@ref` or also block. The conservative recommendation is to block, because moving the output without refreshing the cache produces a table Excel may flag as inconsistent. UNVERIFIED against Excel behavior.

## 5. Worksheet range-bearing constructs

The contract's rewrite list is confirmed complete with two additions (sort state and ignored-error ranges) and one generalization (unknown extension content). All elements below live in `xl/worksheets/sheetN.xml` unless noted.

| Construct | Element and attribute | Transform-or-block |
|---|---|---|
| Used dimension | `dimension/@ref` | Rewrite: recompute from surviving content after the edit (union of cells, not interval-shifted) |
| Merged ranges | `mergeCell/@ref` (inside `mergeCells`, with `count`) | Rewrite each range; maintain `count`; a merge fully inside a deleted span is dropped and reported, never kept as a zero-height merge |
| Data validation | `dataValidation/@sqref` (space-separated list of ranges) | Rewrite each range in the list; also rewrite `formula1`/`formula2` text as formulas |
| Conditional formatting | `conditionalFormatting/@sqref` | Rewrite each range in the list; also rewrite `cfRule` `formula` children as formulas (UNVERIFIED that Excel shifts CF rule formulas identically to cell formulas; verify by fixture) |
| Defined names | `definedName` text in `xl/workbook.xml` (scope: workbook or `localSheetId`) | Rewrite range text as a formula; system names `_xlnm._FilterDatabase` and `_xlnm.Print_Area` are owned by their features per contract [S15] |
| Comments | legacy note anchor (VML row/column), threaded comment position | V1: report-only, do not rewrite VML (section 3.3) |
| Hyperlinks | `hyperlink/@ref` (anchor cell), `@location` (internal target range) | Rewrite both; external targets untouched |
| Drawing anchors | `xdr:from`/`xdr:to` markers | Rewrite per section 3; `absoluteAnchor` blocks |
| Chart series | `c:f` in `xl/charts/chart*.xml` | Rewrite plain ranges; block non-range sources (section 2) |
| Pivot source | `worksheetSource/@ref` / `@name` | Never rewrite; block intersecting edits (section 4) |
| Pivot output location | `pivotTableDefinition/location/@ref` | Block intersecting or shifting edits (section 4) |
| Tables | `xl/tables/tableN.xml` `table/@ref`, `autoFilter/@ref` | Rewrite the `ref` but report structured references as may-be-stale per contract [S15]; table rewrite without model support is UNVERIFIED against Excel repair behavior and needs a fixture |
| AutoFilter | `autoFilter/@ref` | Rewrite |
| Sort state | `sortState/@ref` (plus `sortCondition/@ref`) | Rewrite (addition to the contract list; same `ref` shape as AutoFilter) |
| Ignored errors | `ignoredErrors/ignoredError/@sqref` | Rewrite (addition to the contract list) |
| Sheet-view selection | `sheetView/selection/@sqref`, `@activeCell` | Rewrite; cosmetic, never blocks |
| Page breaks | `rowBreaks/brk`, `colBreaks/brk` (`@min`/`@max` spans) | Shift break positions with the grid; a break inside a deleted span is dropped |
| Protected ranges | `protectedRanges/protectedRange/@sqref` | Rewrite |
| Print titles/area | `_xlnm.Print_Titles` / `_xlnm.Print_Area` defined names | Owned by the print feature; rewritten as names |
| External-link formulas | `xl/externalLinks/*` cached references | Never rewrite; an edit requiring it is blocked [S15] |
| Unknown `extLst` range-bearing content | any | Block the edit rather than risk a silent stale reference; every blocked item is reported |

The `sqref` attributes hold space-separated range lists (the conditional-formatting schema type is `ST_Sqref`, "sequence of references") [S20]; each item in the list is transformed independently, and an item that is fully deleted is removed from the list (dropping the whole element only if the list becomes empty). The Primer shows the canonical shapes (`mergeCell ref="D13:H14"`, `conditionalFormatting sqref="H6:H11"`) [S21].

## 6. The transform arithmetic

One interval rule covers rows and columns symmetrically; the description below uses rows, with `r` the edit position (zero- or one-based consistently inside the module) and `n` the count. Excel's own documented behavior is that "cell references automatically adjust to match the location of the shifted cells" on insert [S22], and that "Excel keeps formulas up to date by adjusting references to the shifted cells to reflect their new locations; however, a formula that refers to a deleted cell displays the `#REF!` error value" [S11].

### 6.1 Insert n rows before row r

For each range `[a..b]` on the edited sheet (endpoints inclusive; single cells are `a == b`):

- `b < r`: range entirely above the insertion. Unchanged.
- `a >= r`: range entirely at or below the insertion. Shift: `[a+n..b+n]`.
- `a < r <= b`: range spans the insertion point. Grow: `[a..b+n]`.

Whole-column ranges (`A:C`) are unaffected by row inserts; whole-row ranges (`3:5`) shift and grow as above. Ranges on other sheets are untouched unless they explicitly qualify the edited sheet (chart `c:f` with a sheet prefix naming another sheet is skipped). Cross-sheet references in the edited sheet's own formulas that point at other sheets are never rewritten by a same-sheet structural edit.

### 6.2 Delete rows r..r+n-1

- `b < r` or `a > r+n-1`: entirely outside. Shift back by `n` if below, else unchanged.
- `a >= r` and `b <= r+n-1`: entirely inside the deleted span. The range is gone: drop the construct entry (merge item, validation list item, CF list item, anchor span collapses per section 3.3) and report it as removed. For a cell formula reference Excel writes `#REF!` [S11]; for a chart `c:f` whose range is fully deleted, Excel shows a `#REF!` series. Recommendation: keep the `c:ser` with its `c:f` rewritten to a `#REF!` marker form only if fixture testing shows Excel tolerates it; otherwise block the delete. UNVERIFIED: needs an Excel fixture. Never silently keep the old range text.
- Partial overlap (`a < r <= b <= r+n-1` or `r <= a <= r+n-1 < b`): shrink to the surviving remainder (`[a..r-1]` or `[r..b-n]` after re-indexing). UNVERIFIED in full generality: Excel's merge/validation/CF shrink behavior on partial deletes must be confirmed by fixture; shrinking plus a report entry is the conservative choice because it never points at unintended cells.

### 6.3 Chart-specific edge cases

- A series that shrank to nothing (single-cell range deleted): same as fully-deleted range above; block or `#REF!`-mark, never keep stale text.
- `c:f` text on a different sheet than the edited sheet: skip.
- Multi-area `c:f` (contains `,`): treat as non-range and block until fixture evidence exists (section 2).
- `c:f` with an external `[Book]` qualifier: block (external-link rule [S15]).
- After any rewrite, `ptCount` and point caches are left stale-bytes and reported (section 1.3).

### 6.4 Anchors, dimensions, and counts

- Anchor markers shift by the section 6.1/6.2 rules applied to their zero-based `col`/`row` as points (a point `p >= r` shifts; `p < r` stays; a point inside a deleted span clamps to the deletion point with a report entry).
- `dimension/@ref` is recomputed from the surviving grid, not shifted.
- `mergeCells/@count`, `ptCount`, and any element counts are maintainers of truth after edits, not inputs to the transform.

### 6.5 Performance and ordering

The transform runs once per structural edit as a single pass over the affected sheet's range-bearing constructs plus all chart parts referencing that sheet, all drawing parts of that sheet, `workbook.xml` names, table parts, and pivot parts. openpyxl explicitly disclaims this work ("does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted" — client code must implement it) [S23], which confirms there is no off-the-shelf routine to adopt; xlsx-populate's tree-manipulation approach (manipulate the XML trees rather than a lossy rich model [S24]) is the right implementation pattern for the patch writer.

## 7. What the module must never do

1. Never leave a stale range silently: every dropped, shrunk, clamped, or blocked item produces an export-report entry in the contract's five categories (unchanged, transformed, preserved-but-stale, removed, warning) [S15].
2. Never rewrite a pivot source, an external-link reference, or a non-range chart source.
3. Never renumber style, shared-string, or relationship IDs as a side effect of the transform; tables are append-only [S24 equivalent in contract §2/§5].
4. Never apply a same-sheet edit's arithmetic to another sheet's ranges, except through an explicit sheet-qualified reference naming the edited sheet.

## 8. Decisions and open questions

Resolved by the project on 2026-11-18. The recommendations below were adopted as written:

- **Chart caches stay byte-untouched and are reported stale** (§ recommendation 2). Deleting or zeroing `numCache`/`strCache` is the schema-legal fallback if Excel repairs on open, to be settled by the fixture matrix.
- **`twoCellAnchor` and `oneCellAnchor` shift with the grid; `absoluteAnchor` is untransformable and blocks** (recommendation 3).
- **Pivot sources are never rewritten.** An edit that intersects or shifts a pivot source range or its output `location` is blocked; elsewhere the pivot is reported "preserved but not refreshed" (recommendation 4).
- **A pivot output `location` shift caused by an edit that touches neither source nor output blocks rather than rewrites** (question 4, option taken: block). Moving the output beside an unrefreshed cache would produce a table Excel may flag as inconsistent.
- **Unknown `extLst` range-bearing content blocks the edit** (question 5, option taken: block), consistent with the no-silent-loss rule and the conservative edit-availability promise. This adds a fourth blocked case to the three named in the file-compatibility contract, which has been updated.
- **A chart series range that would be entirely deleted blocks the delete** rather than writing a `#REF!` marker, matching the contract's no-silent-loss rule and avoiding a marker Excel may treat as a repair condition.
- **The Excel fixture matrix (§ section 8 items 1-2, "Local experiments") is the first task of slice 6D**, since the edge-case behavior below cannot be sourced from documentation alone.

Still open, to be settled by that fixture matrix or by later work:

1. Stale chart caches beside rewritten `c:f`: repair-free in Excel 365? Needs a fixture: insert a row above a charted range, rewrite `c:f`, keep the old cache, open in Excel. Fallback (delete the cache element) is schema-legal [S5].
2. Exact Excel behavior for markers and ranges fully or partially inside a deleted span (anchors, merges, validations, CF lists, chart series): needs a small Excel fixture matrix before slice 6D starts. The note's shrink/clamp/drop recommendations are conservative but unverified. The decision above fixes the fully-deleted chart series case; partial spans remain to be observed.
3. Multi-area chart `c:f` (comma-separated): how often real Excel emits it, and whether to support or block.
4. CF rule `formula` children and validation `formula1`/`formula2`: confirm Excel shifts them identically to cell formulas.
5. ISO Strict OOXML variants of the chart and drawing namespaces were not examined; the contract already rejects Strict files [S15], so the module only needs Transitional handling, but the detector must run before the transformer.

## 9. Sources

| ID | Source |
|---|---|
| S1 | ECMA-376 1st edition, Part 4, DrawingML Charts reference, via c-rex rendering — `f` (Formula), `numRef`, `val`, `tx` (Series Text), `scatterChart` element pages |
| S2 | c-rex ECMA-376 `f` (Formula) page: "a book, sheet, and cell reference or a book, optional sheet, and defined name reference", no leading `=`; example `Sheet1!$A$1:$C$1` |
| S3 | c-rex ECMA-376 `val` (Values) page: `CT_NumDataSource` choice of `numRef`/`numLit`; parent `ser` list |
| S4 | c-rex ECMA-376 `tx` (Series Text) page: `CT_SerTx` choice of `strRef`/`v`; parents are all eight chart-type `ser` elements |
| S5 | c-rex ECMA-376 `numRef` page: `CT_NumRef` sequence (`f`, optional `numCache`, optional `extLst`); parents `bubbleSize`, `cat`, `minus`, `plus`, `val`, `xVal`, `yVal` |
| S6 | openpyxl `chart.data_source` API docs: `AxDataSource`, `NumDataSource`, `NumRef(f, numCache)`, `StrRef(f, strCache)`, `MultiLevelStrRef` |
| S7 | openpyxl `chart.series` API docs: `SeriesLabel(strRef, v)` |
| S8 | c-rex ECMA-376 `scatterChart` page: `CT_ScatterChart` with `CT_ScatterSer` series |
| S9 | Real-Excel chart XML excerpts: scatter `xVal`/`yVal` with `Sheet1!$A$2:$A$7` (StackOverflow custom-label thread) and quoted `'5.Tiers.Hourly'!$F$19:$F$40`; openpyxl test asserting `Blatt1!$A$1:$A$12` |
| S10 | Microsoft Learn, "Use defined names to automatically update a chart range": `=Sheet1!Sales` series-values pattern and `OFFSET`-based dynamic names |
| S11 | Microsoft Support, "Insert or delete cells, rows, and columns": Excel adjusts references to shifted cells; references to deleted cells become `#REF!` |
| S12 | officeopenxml.com DrawingML primer and Apache OpenOffice cwiki DrawingML export notes: absolute anchors do not move or size with cells; one-cell and two-cell anchors move with cells |
| S13 | Microsoft Learn [MS-XLSX] Drawing: `oneCellAnchor`, `twoCellAnchor`, `absoluteAnchor` (§20.5.2.24/.33/.1 of ISO/IEC 29500-1:2016) plus slicer `AlternateContent` extensions |
| S14 | openpyxl `drawing.spreadsheet_drawing` source: `AnchorMarker(col, colOff, row, rowOff)`, `TwoCellAnchor(editAs, _from, to)`, `OneCellAnchor(_from, ext)`, `AbsoluteAnchor(pos, ext)`; `SpreadsheetDrawing` anchor lists |
| S15 | calcplus file-compatibility contract (`.scratch/calcplus/spec/file-compat.md` §§4–7, criteria 7/10/15–17): Tier B preservation, the three blocked cases, no-silent-loss reporting, external-link rule |
| S16 | StackOverflow pivot-table cache-source thread: `<x:cacheSource type="worksheet"><x:worksheetSource ref="A1:X2" sheet="dataWorksheet"/>` |
| S17 | Microsoft Learn [MS-OI29500] §18.10.1.95 `worksheetSource`: Office requires exactly one of `ref`/`name`; `id` names an external-workbook relationship; `@name`/`@sheet` length caps |
| S18 | EPPlus `ExcelPivotCacheDefinition` source: `_sourceWorksheetPath` and `_sourceAddressPath` constants |
| S19 | c-rex ECMA-376 `location` (PivotTable Location) page: example `ref="B6:G13"` with `firstHeaderRow/firstDataRow/firstDataCol/rowPageCount/colPageCount`; parent `pivotTableDefinition` |
| S20 | c-rex ECMA-376 `conditionalFormatting` page: `CT_ConditionalFormatting` with `sqref` of type `ST_Sqref` |
| S21 | ECMA-376 Primer (c-rex Supporting Features): `mergeCell ref="D13:H14"`, `conditionalFormatting sqref="H6:H11"` examples |
| S22 | Microsoft Support, "Insert or delete rows and columns in Excel": cell references adjust automatically on insert |
| S23 | openpyxl "Inserting and deleting rows and columns" docs: openpyxl does not manage formulae, tables, or charts on insert/delete; client code must implement it |
| S24 | xlsx-populate README (via prior research 04): XML-tree manipulation preserving unknown parts; the implementation pattern for the patch writer |
| S25 | Prior calcplus research 04 (`.scratch/calcplus/research/04-xlsx-io.md`): preservation-first architecture, byte-versus-semantic preservation, the deferred-spike statement |

### Local experiments

None were run for this spike. The UNVERIFIED items in sections 1.3, 2, 3.2, 3.3, 4.2, 5, and 6.2 each name the Excel fixture that would resolve them; running that matrix is the recommended first task of slice 6D.
