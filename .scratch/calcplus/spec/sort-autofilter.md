# calcplus v1 — sort and AutoFilter semantics

Resolved by [19 — Sort and AutoFilter semantics](../issues/19-sort-autofilter.md) on 2026-11-13. It extends the [workbook model](workbook-model.md), the [command system](command-system.md) and the [file-compatibility contract](file-compat.md). Excel's documented sort order is the reference.

## 1. Scope

**In v1**

- Sort a selected range top-to-bottom, with one or more keys, ascending or descending.
- AutoFilter on a plain range: value checklist, text/number/date conditions, blanks, clear.

**Out of v1**

- Left-to-right sort, custom lists, cell/font-colour or icon sort, case-sensitive sort.
- Colour filters, Top 10, dynamic date filters (today, tomorrow, next week), advanced filter.
- Editing a table's own filter. Imported table filters are preserved but read-only.
- Sorting inside a table or across a legacy array formula.

## 2. Sort

### 2.1 Target range

- Sort operates on the **selected range**.
- If the selection is a single cell, the target is the **current contiguous region** around it.
- If the range has fewer than two data rows, sort is a no-op.
- The sort dialog shows the target range as a read-only label; the user changes it by changing the selection.
- Sorting a sub-range inside a larger block is allowed, but calcplus warns once, as Excel does, that the sorted rows will be disassociated from the surrounding data.

### 2.2 Header row

- The first row is treated as a header when the heuristic says so:
  - every cell in the first row is text, and
  - at least one key column has a non-text value below, or the first row has a distinct style (bold, fill or border).
- The Sort dialog always exposes **My data has headers** so the user can override the guess.
- With headers on, the first row is excluded from the sort and supplies the key labels.
- With headers off, the first row sorts as data and keys are labelled by column letter.

### 2.3 Keys and direction

- Up to **64 keys**, matching Excel.
- Each key is a column inside the range, a direction (ascending or descending), and **Cell Values** as the sort-on type.
- Keys apply in order: the first key decides, the second breaks ties, and so on.
- Ties after all keys preserve the original row order (**stable sort**).
- The key values are read from a snapshot before any row moves, so the result does not depend on move order.

### 2.4 Comparator

Ascending order:

| Rank | Type | Order |
|---:|---|---|
| 1 | Number | Smallest negative to largest positive. Dates and times are numbers. |
| 2 | Text | Excel's character order, case-insensitive. |
| 3 | Logical | `FALSE` before `TRUE`. |
| 4 | Error | All errors equal. |
| 5 | Blank | Always last. |

- **Descending** reverses ranks 1-4. **Blanks stay last in both directions.**
- Numbers stored as text sort as text, after all numbers.
- Text follows Excel's documented order: `0-9`, space, punctuation, then `A-Z`; apostrophes and hyphens are ignored except as a final tie-break; case is ignored. Non-ASCII characters fall back to code-point order. The order is the English locale; localization is v2.
- An empty cell is blank. A cell containing only spaces is text.
- Errors compare equal, so the stable tie-break keeps their original order.

### 2.5 What moves

Sort is a **row permutation**, not a value rewrite.

- Cell values, formulas, saved results, styles, comments, hyperlinks and rich-text payloads move with their row.
- **Formulas inside the range move with their row, and relative references shift by the row delta**, exactly as if the cell were moved. Absolute and mixed parts keep their anchors. Sheet names and defined names are preserved.
- Formulas **outside** the sorted range are not rewritten. If they reference a cell inside the range, they now read whatever value landed at that address. This is Excel's behaviour and the reason Excel advises absolute references.
- Cross-sheet references inside a moved formula shift by the same row delta; the sheet name is preserved.
- **Defined names are unchanged.** A name that points at an address inside the range still points at that address, whose content may have changed.
- Conditional-formatting and data-validation ranges are unchanged; they are address-based, not cell-attached.
- Row heights and manual hidden flags **stay with row positions**; column widths are unchanged.
- Uniform merged ranges move with their rows. A merge must be fully inside the sort range and exactly one row tall; if merges differ in size or a merge spans rows, the sort is blocked with Excel's message: "This operation requires the merged cells to be identically sized."
- Sorting is blocked if the range intersects a table or a legacy array formula, with a clear message. Protection is not enforced, so it does not block.

### 2.6 Hidden rows

Excel documents that hidden rows are not moved when sorting rows. calcplus matches that:

- Only **visible** rows are reordered, into the visible row slots.
- Manually hidden rows and filter-hidden rows stay in their positions.
- The user is warned once before a sort that crosses hidden rows, suggesting they unhide first.
- The warning is informational; the sort proceeds.

### 2.7 Transaction and undo

- One sort is **one transaction**. It stores the row permutation, the formula rewrites and the moved cell records.
- Undo restores the exact prior values, formulas, styles and positions.
- Redo reapplies the same permutation.
- The operation is recorded in the command log with `sort.apply`.

## 3. AutoFilter

### 3.1 State model

One AutoFilter per sheet, attached to a plain range:

```
AutoFilter {
  sheetId
  range: { r1, c1, r2, c2 }      // includes the header row
  columns: Map<columnOffset, ColumnFilter>
  raw?: unknown                  // unknown XML and attributes
}

ColumnFilter =
  | { kind: "values", selected: TypedValue[], includeBlanks: boolean }
  | { kind: "conditions", join: "and" | "or", conditions: Condition[] }
  | { kind: "unsupported", raw: unknown }
```

- `columnOffset` is relative to `range.c1`, so inserting a column shifts the criterion with its column.
- **Row 1 of the range is always the header row.** If it looks like data, the dropdown labels use `Column1`, `Column2`, … but the row is still excluded from filtering.
- The filter range is the current contiguous region when AutoFilter is switched on from a single cell, or the selected range when a range is selected.
- Filter criteria are **document state** and round-trip. The hidden-row effect is derived.

### 3.2 Criteria

**Value checklist**

- The distinct **typed values** in the column, sorted with the §2.4 comparator, with blanks last.
- `(Select All)`, `(Blanks)` and a search box that filters the checklist as the user types.
- Selecting a value includes only that value. The comparison is type-aware: number `10` is not text `"10"`.
- The checklist displays formatted values. If two underlying values format identically, one entry selects both.
- Errors appear as values and can be selected.

**Conditions**

- Up to **two conditions per column**, joined by **AND** or **OR**.
- Text: equals, not equals, begins with, ends with, contains, does not contain. Wildcards `*`, `?` and `~`; case-insensitive.
- Number: `=`, `<>`, `>`, `>=`, `<`, `<=`, between, not between.
- Date: equals, not equals, before, after, between, not between. Dates compare as serial numbers, so the 1900/1904 flag is irrelevant to comparison.
- Text operators apply only to text cells; number and date operators apply only to numbers. Other types do not match.
- Errors never match a condition; they can only be selected in a checklist.
- **Columns combine with AND**: a row is visible only if it passes every filtered column.
- The header row is always visible.
- Filtering evaluates the cell's current value (the saved result for a formula cell) and never triggers recalculation.

### 3.3 Visibility

Two independent layers:

- **Manual-hidden** — the row's hidden flag set by the user.
- **Filter-hidden** — derived: a data row that fails the active filter.

A row is hidden when either layer hides it.

- Applying or changing a filter recomputes only the filter layer.
- **Clearing a filter removes only the filter layer**; manually hidden rows stay hidden.
- With no active filter, every hidden row is manual.
- On save, write the **effective** hidden flag (`manual || filter`) plus the `autoFilter` XML.
- On open, rebuild the layers: rows that fail the saved criteria are filter-hidden; every other hidden row is manual.

### 3.4 Sort and filter together

- A sort reorders **visible** rows only, as in §2.6.
- Filter-hidden rows stay in place and are not sorted.
- The same filter criteria are reapplied after the sort; filter state is unchanged.
- Because all visible rows matched before the sort and keep their values, the visible set is unchanged.

### 3.5 Structural edits

Insert/delete rows or columns inside or above the filter range goes through the same range-transform module as every other structural edit.

- The `range` shifts or expands; column offsets move with their columns.
- Visibility is recomputed after the edit.
- If the header row is deleted, or the range collapses to fewer than two rows, the filter is **cleared** and filter-hidden rows are restored.
- Unsupported criteria are preserved through the transform where possible.

### 3.6 Round-trip

- Write `<autoFilter ref="A1:E5">` with one `<filterColumn colId="n">` per filtered column.
- Value lists use `<filters><filter val="…"/></filters>` and `blank="1"` for blanks.
- Conditions use `<customFilters and="1|0"><customFilter operator="…" val="…"/></customFilters>`. Between is two custom filters joined by AND.
- Update the hidden `_xlnm._FilterDatabase` defined name to match the range.
- Preserve unknown filter XML, operators and attributes verbatim.
- One `autoFilter` per sheet. A table's own filter lives in its table part and is **preserved but not editable** in v1.
- No-op save leaves the `autoFilter` part byte-identical.

### 3.7 Unsupported imported filters

Colour, icon, Top 10, dynamic date and unknown custom operators are read and preserved.

- Their hidden rows render correctly.
- The dropdown shows a read-only notice: "This filter uses a feature calcplus cannot edit."
- Only **Clear** is offered.
- Applying a new filter on that column replaces the unsupported criterion after a confirmation.
- Saving without touching it preserves the XML.

### 3.8 UI

- **Ctrl+Shift+L** toggles AutoFilter on the current range.
- Each header cell gets a dropdown button when the filter is on.
- The dropdown contains: Sort A→Z, Sort Z→A, Clear Filter, Text/Number/Date Filters, a search box, Select All, Blanks, the checklist, OK and Cancel.
- Filtered columns show a funnel icon in the header.
- The status bar shows `X of Y records` while a filter is active.
- **Alt+Down** opens the dropdown; Up/Down move, Space toggles, Enter applies, Esc cancels.
- **Reapply** and **Clear All** live in the Data tab.
- The dropdown is a DOM overlay; it is not drawn on the canvas.
- Opening the dropdown and typing in the search box are view-only and create no transaction.

### 3.9 Transaction and undo

- Toggling AutoFilter on/off, applying a criterion, changing one and clearing are each **one transaction**.
- A transaction stores the `autoFilter` before/after and the effective hidden-row changes.
- Undo restores the filter and the hidden rows together.
- Reapply is a no-op transaction when nothing changes.

## 4. Conformance fixtures

Sort:

| ID | Fixture | Proves |
|---|---|---|
| S1 | `sort-values` | Single key ascending and descending |
| S2 | `sort-multi` | Three keys, including a tie-break |
| S3 | `sort-header` | Header on and off |
| S4 | `sort-mixed-types` | Numbers, text, booleans, errors and blanks; blanks last both ways |
| S5 | `sort-formulas` | Relative, absolute, mixed and cross-sheet references inside the range |
| S6 | `sort-outside-refs` | Formulas outside the range are not rewritten |
| S7 | `sort-hidden` | Manual and filter-hidden rows stay put |
| S8 | `sort-merges-uniform` | One-row merges move with their rows |
| S9 | `sort-merges-mixed` | Mixed merges block with the Excel message |
| S10 | `sort-row-props` | Row heights and hidden flags stay; styles move |
| S11 | `sort-array` | Legacy array formula blocks the sort |
| S12 | `sort-table` | A range intersecting a table blocks the sort |
| S13 | `sort-64-keys` | The 64-key limit |
| S14 | `sort-stable` | Equal keys keep their original order |
| S15 | `sort-undo` | Undo and redo restore the exact prior state |

AutoFilter:

| ID | Fixture | Proves |
|---|---|---|
| F1 | `filter-values` | Value checklist, typed values and Select All |
| F2 | `filter-text` | Text conditions and wildcards |
| F3 | `filter-number` | Number conditions, including between |
| F4 | `filter-date` | Date conditions on serials, 1900 and 1904 |
| F5 | `filter-and-or` | Two conditions joined by AND and by OR |
| F6 | `filter-columns` | Multiple columns combine with AND |
| F7 | `filter-blanks` | Blanks and NonBlanks |
| F8 | `filter-errors` | Errors are checklist values and never match conditions |
| F9 | `filter-hidden-layers` | Clearing a filter keeps manually hidden rows hidden |
| F10 | `filter-sort` | Sort while filtered reorders visible rows and keeps the filter |
| F11 | `filter-structural` | Insert/delete shifts the range and criteria |
| F12 | `filter-header-delete` | Deleting the header clears the filter |
| F13 | `filter-roundtrip` | `autoFilter` XML, hidden rows and `_FilterDatabase` survive |
| F14 | `filter-unsupported` | Colour/dynamic filters are preserved and read-only |
| F15 | `filter-table` | A table filter is preserved and not editable |
| F16 | `filter-one-per-sheet` | A second filter replaces or is refused |
| F17 | `filter-100k` | 100k-row sort within 2 s and filter within 1 s |
| F18 | `filter-undo` | Apply, change and clear are undoable as one step each |

## 5. Acceptance criteria

1. Sort is a stable row permutation over the selected range; a single cell expands to the current region.
2. The comparator matches Excel's documented order, with blanks last in both directions.
3. Formulas inside the range move and shift relative references; absolute anchors and defined names stay; outside formulas are not rewritten.
4. Uniform one-row merges move; mixed merges block with Excel's message.
5. Cell styles, comments and hyperlinks move; row heights and hidden flags stay.
6. Hidden rows are not moved; the user is warned once.
7. Array formulas and tables block a sort.
8. One sort is one undo transaction and restores the exact prior state.
9. AutoFilter supports a value checklist and two AND/OR conditions per column.
10. Text, number and date conditions, wildcards and blanks behave as specified.
11. Columns combine with AND; the header row stays visible.
12. Manual-hidden and filter-hidden rows are independent; clearing a filter restores only filter-hidden rows.
13. Sorting while filtered reorders visible rows and keeps the filter.
14. Structural edits transform the filter range; an invalid range clears the filter.
15. `autoFilter` XML, `_FilterDatabase` and hidden rows round-trip; unsupported filters are preserved read-only.
16. Table filters are preserved but not editable.
17. The dropdown, funnel indicator, status-bar count and Alt+Down keyboard path work as specified.
18. Filter apply/change/clear are each one undo transaction.
19. A 100k-row sort completes within 2 s and a 100k-row filter within 1 s, per [performance](performance.md).

## 6. Deliberate deviations and v2 cuts

- Left-to-right sort is v2.
- Custom lists, cell/font-colour/icon sort and case-sensitive sort are v2.
- Colour, Top 10, dynamic date and advanced filters are v2.
- Table filters are preserved, not edited.
- Sorting a range that intersects a table or a legacy array formula is blocked rather than partially supported.
- Hidden rows stay in place; this is Excel's documented behaviour, even though it can surprise users.
- The comparator uses the English locale; localized sort orders are v2.
