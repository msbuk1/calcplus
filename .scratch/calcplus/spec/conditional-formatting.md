# calcplus v1 — conditional-formatting rendering

Resolved by [20 — Conditional-formatting rendering](../issues/20-conditional-formatting.md) on 2026-11-13. It applies the render-only scope in the [feature matrix](feature-matrix.md) and the round-trip rules in [file compatibility](file-compat.md). CF is an overlay: it never changes stored values or styles.

## 1. Scope

**Rendered in v1**

- Cell-value comparisons.
- Text contains / does not contain / begins with / ends with.
- Date comparisons and time-period presets.
- Top/bottom N and N%.
- Duplicate and unique values.
- Two- and three-colour scales.
- Data bars.
- Icon sets.

**Preserved but not rendered**

- Formula-based rules (`expression`).
- Above-average rules.
- Blank/error rules (`containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`).
- Rules whose thresholds use a formula.
- Any unknown rule type or operator.

**Not in v1**

- Authoring, the rule manager, stop-if-true UI, and formula-based rule authoring. All are v2.
- Custom icon images, data-bar gradients beyond Excel's model, and conditional-formatting tables.

## 2. Rule model

```
ConditionalFormatRule {
  id
  sheetId
  ranges: Range[]          // from sqref; may be discontiguous
  priority: number         // 1 is highest
  stopIfTrue: boolean
  kind: cellIs | text | timePeriod | top10 | duplicate | unique
      | colorScale | dataBar | iconSet | unsupported
  dxfId?: number
  raw: unknown             // original XML and attributes
}
```

- Rules are sheet-scoped. `sqref` may contain several ranges.
- **Priority:** rules evaluate in ascending numeric order. Missing or duplicate priorities fall back to document order.
- **`dxf`:** a differential format holding only the properties the rule sets — fill, font, border and number format. Existing `dxfId` values are never renumbered; new records append.
- CF never mutates the cell, its style or its value. Copy, paste and Format Painter ignore CF.

## 3. Evaluation

### 3.1 Combining rules

- Every true rule contributes its `dxf` properties.
- For each property (fill, font, border, number format), the **first true rule in priority order wins**. Lower-priority rules fill in properties the higher ones do not set.
- A true rule with **stop-if-true** stops lower-priority rules for that cell.
- **Manual formatting** shows through for every property CF does not set, and whenever no rule is true.
- CF fill hides gridlines; CF borders draw over manual borders.

### 3.2 Values and coercion

- Rules evaluate the cell's **current value** — the saved result for a formula cell. Evaluation never triggers recalculation.
- **Blank cells coerce to 0** in numeric comparisons, matching Excel: "less than 20" also matches blanks. Text comparisons treat a blank as an empty string.
- Text rules use the cell's **formatted display text** as the haystack, so a number formatted as `1,234` is matched as displayed. Matching is case-insensitive.
- Text rules support wildcards `*`, `?` and `~`, as Excel's "text that contains" rules do.
- Errors do not match cell-value, text or date rules. They can be matched by duplicate/unique rules.
- Dates are serial numbers. Date rules compare serials and use the workbook's 1900 or 1904 system.

### 3.3 Rule types

**Cell value (`cellIs`)**

- Operators: less than, less than or equal, equal, not equal, greater than or equal, greater than, between, not between.
- Numbers compare numerically; text compares case-insensitively; booleans compare `FALSE < TRUE`.
- A blank compares as 0.

**Text**

- Contains, does not contain, begins with, ends with.
- Case-insensitive, wildcard-aware, applied to the display text.
- A blank cell never matches a non-empty text rule.

**Date and time period**

- Comparisons: equals, not equals, before, after, between, not between.
- Presets: Excel's ten — yesterday, today, tomorrow, last 7 days, last week, this week, next week, last month, this month, next month. Excel has no year-based period, and neither do we.
- "Today" is the local calendar date. Periods use the workbook's date system for serial conversion.
- Re-evaluate on open, after edits, when the tab regains focus, and when the local date changes. An overnight workbook must not stay stale.

**Top/bottom (`top10`)**

- Top or bottom, N or N%, default 10.
- Ranks within each contiguous range in `sqref`.
- Blanks and errors are excluded.
- Ties are included, matching Excel: a top-10 rule may highlight more than ten cells.

**Duplicate / unique**

- Compare typed values; blanks are ignored.
- Errors can be duplicates.
- Each contiguous range in `sqref` is its own comparison domain.

**Colour scale**

- Two or three colour points.
- Threshold types: automatic min/max, number, percent, percentile.
- Interpolate linearly in sRGB between points. Automatic uses the range's min and max; the midpoint of a three-colour scale defaults to the 50th percentile.
- Blanks and errors are not painted and do not affect the range statistics.

**Data bar**

- Gradient or solid fill.
- Threshold types: automatic, number, percent, percentile for min and max.
- Positive and negative colours, axis at zero, optional border, left-to-right direction, show or hide the value.
- Automatic min is `min(0, range min)` and automatic max is `max(0, range max)`, so the axis at zero is meaningful for mixed signs.
- Blanks and errors are not painted and do not affect the statistics.
- The bar draws behind the cell text.

**Icon set**

- All Excel built-in sets: 3/4/5 arrows, grey arrows, traffic lights, signs, symbols, flags, ratings, red-to-black and quarters.
- Threshold types: number, percent, percentile.
- Reverse order, show-value and icon-only.
- Icons draw at the left of the cell; text shifts right by the icon width. Icon-only hides the text.
- Icons are bundled SVG approximations, near-identical to Excel rather than pixel-perfect.

**Unsupported**

- Formula-based, above-average, blank/error rules, formula thresholds and unknown operators are **preserved and not painted**.
- The cell shows manual formatting.
- The compatibility report lists them with sheet, range and count.
- An unsupported rule never stops lower-priority rules, because its truth cannot be known.

## 4. References

- Formula-based rules and formula thresholds are v2, so relative/absolute reference evaluation is not part of v1.
- The v2 contract, recorded for continuity: relative references anchor to the top-left cell of the rule's range and shift per evaluated cell; absolute references stay; cross-sheet references resolve through the formula engine; defined names resolve normally.
- In v1, those rules are preserved byte-for-byte and reported, never approximated.

## 5. Rendering

### 5.1 Paint order

Per cell:

1. Manual fill.
2. CF fill or colour-scale fill.
3. Data bar.
4. Manual border.
5. CF border.
6. Cell text, shifted right when an icon is shown.
7. CF font and CF number format.
8. Icon.
9. Selection overlay, drawn by the grid after all cell content.

CF fill covers gridlines. The formula bar always shows the raw value, not the CF-formatted display.

### 5.2 Merged cells

- Evaluate at the **anchor cell's value**.
- Paint fill, font and border across the whole merged rectangle.
- Data bars scale to the merged width; the axis is at zero.
- Icons draw at the left of the merged rectangle and shift the text right.
- A rule applies to a merged cell when its anchor is inside the rule's range; covered cells are ignored.

### 5.3 Freeze panes and clipping

- Each pane paints CF independently from the same evaluation, clipped to its own pane.
- Nothing bleeds across a freeze divider.
- A merged cell crossing a freeze boundary is clipped per pane; the text appears in the anchor's pane and the fill paints in each clipped portion.
- Every CF drawing is clipped to the cell or merged rectangle and to the pane. No overdraw.

### 5.4 Zoom

- Thresholds are value-based and never change with zoom.
- Bar widths, icon sizes and CF fonts scale with zoom.
- Below 25% zoom, icons are suppressed to avoid clutter; fills, bars and scales still paint.
- Zoom never re-evaluates rules.

### 5.5 Print, PDF and accessibility

- Print preview and PDF render CF exactly as on screen, including bars, scales and icons. The [print pipeline](print-pdf.md) must consume the same CF layer.
- Unsupported rules are not painted in print either; the export report notes them.
- The accessibility mirror exposes the effective display text after a CF number format. CF colours are not announced in v1; ticket 22 may add a description.
- Rich text displays as plain text; CF applies uniformly to the cell.

## 6. Performance and incremental evaluation

- Build a per-sheet **CF index**: rules sorted by priority, with a range index from cells to rules.
- **Per-cell rules** (cell value, text, date) evaluate on demand for visible cells and cache by rule plus cell value.
- **Range-statistic rules** (top/bottom, duplicate/unique, colour scales, data bars, icon sets) compute once per contiguous range and cache min/max, counts and thresholds.
- A cell edit invalidates that cell's per-cell results and every range statistic for ranges containing it. Recomputation is lazy, on the next paint, and the previous paint stays until the new result is ready.
- A format-only change repaints; it does not re-evaluate values.
- Hidden rows, filtered rows and off-screen cells are not painted, but still participate in range statistics.
- Budgets:
  - Visible-cell evaluation: under **4 ms** per frame.
  - 100k-cell range statistics: under **200 ms**, computed in row-block chunks across frames.
  - 250k-cell sheet: 60 fps scrolling, per [performance](performance.md).
  - Above the guaranteed envelope, the ticket 11 ladder may stop CF painting outside the viewport.

## 7. Round-trip

- Read both the base `conditionalFormatting` and the `x14:conditionalFormatting` extension; merge by priority. When both describe the same rule, the x14 version drives rendering and both XML forms are preserved.
- Preserve `dxfId`, `priority`, `stopIfTrue`, `pivot`, unknown operators, unknown attributes and unknown child elements.
- The `dxfs` table is append-only. Existing `dxfId` values are never renumbered.
- Structural edits transform `sqref` and rule references through the range-transform module. Rules whose ranges collapse are removed. Rule priority and order are preserved.
- A no-op save leaves the CF XML and `dxfs` byte-identical.
- One rule may cover several ranges; each contiguous range is a separate evaluation domain for ranking rules.

## 8. Surfacing

- There is no CF authoring UI in v1. The ribbon button is present but disabled, with a tooltip that authoring arrives in v2.
- The **compatibility report** lists unsupported CF rules by sheet, range and type, and says they are preserved but not rendered.
- The **export report** repeats the count on save if any unsupported rule is present.
- Manual formatting remains fully editable; CF is read-only.

## 9. Conformance fixtures

| ID | Fixture | Proves |
|---|---|---|
| CF1 | `cf-cell-value` | All comparison operators; blanks coerce to 0 |
| CF2 | `cf-text` | Contains, not contains, begins, ends; wildcards; numbers; case-insensitivity |
| CF3 | `cf-dates` | Date comparisons and every time-period preset; 1900 and 1904 |
| CF4 | `cf-top-bottom` | Top/bottom N and N%; ties included; blanks and errors excluded |
| CF5 | `cf-duplicates` | Duplicate and unique; blanks ignored; errors as values |
| CF6 | `cf-colour-scale-2` | Two-colour scale with automatic, number, percent and percentile points |
| CF7 | `cf-colour-scale-3` | Three-colour scale and sRGB interpolation |
| CF8 | `cf-data-bar-positive` | Gradient and solid bars, automatic and fixed thresholds, show/hide value |
| CF9 | `cf-data-bar-negative` | Negative bars, axis at zero, border, direction |
| CF10 | `cf-icon-sets` | All built-in sets, thresholds, reverse, show-value, icon-only |
| CF11 | `cf-priority` | Property merge, first writer wins, stop-if-true, manual formatting |
| CF12 | `cf-merged` | Anchor evaluation and merged painting |
| CF13 | `cf-panes-zoom` | Freeze panes, clipping and zoom scaling; icons suppressed below 25% |
| CF14 | `cf-unsupported` | Formula, above-average, blank/error and formula-threshold rules preserved and reported |
| CF15 | `cf-x14` | x14 data bars and icon sets merge with the base rules |
| CF16 | `cf-dxf-roundtrip` | dxf append-only; no renumbering; no-op save byte-identical |
| CF17 | `cf-structural` | sqref and references transform; collapsed ranges remove the rule |
| CF18 | `cf-100k` | 100k-cell range statistics under 200 ms; 4 ms visible evaluation |
| CF19 | `cf-print` | Print preview and PDF match the screen |
| CF20 | `cf-undo-structural` | Undo of a structural edit restores the transformed rules |

## 10. Acceptance criteria

1. The rendered rule set and the preserved-but-unpainted set match §1 exactly.
2. Rules evaluate in priority order; true rules merge property-by-property; first writer wins; stop-if-true is honoured.
3. CF overrides manual formatting only for the properties it sets; manual formatting shows through otherwise.
4. Blanks coerce to 0 in numeric comparisons; text rules use display text, are case-insensitive and support wildcards.
5. Date comparisons and all time-period presets work with both date systems and refresh at midnight.
6. Top/bottom includes ties and excludes blanks and errors; duplicate/unique ignores blanks.
7. Two- and three-colour scales, data bars and all built-in icon sets render as specified.
8. Merged cells evaluate at the anchor and paint the merged rectangle.
9. Freeze panes clip independently; zoom scales visuals but not thresholds; icons are suppressed below 25%.
10. Visible-cell evaluation stays under 4 ms per frame and 100k-cell statistics under 200 ms.
11. CF round-trips through base and x14 XML; dxf IDs are stable and append-only; no-op save is byte-identical.
12. Structural edits transform `sqref` and references; collapsed ranges remove the rule; undo restores it.
13. Unsupported rules are preserved, never painted, and listed in the compatibility and export reports.
14. CF never mutates the model; copy, paste and Format Painter ignore it.
15. Print, PDF and the accessibility display text follow §5.5.
16. All fixtures in §9 pass locally.

## 11. Deliberate deviations and v2 cuts

- Formula-based, above-average, blank/error and formula-threshold rules are preserved but not rendered. This is the feature matrix cut, kept deliberately; if the conformance corpus shows many such rules, it is the first v2 candidate.
- CF authoring, the rule manager and the stop-if-true UI are v2.
- Icon glyphs are bundled SVG approximations, near-identical rather than pixel-perfect.
- Colour-scale interpolation is linear sRGB; Excel's exact colour space may differ by a few levels.
- Icons are suppressed below 25% zoom.
- CF colours are not announced by the accessibility mirror in v1.
