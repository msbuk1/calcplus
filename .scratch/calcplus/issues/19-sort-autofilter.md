# 19 — Sort and AutoFilter semantics

Status: resolved
Type: grilling
Blocked by: none

## Question

Define exactly how sort and AutoFilter behave in v1.

**Sort**

- Single vs multi-key; header detection; ascending/descending; custom lists.
- What happens to formulas, mixed absolute/relative references, cross-sheet references, defined names, merged cells, hidden rows, errors and blanks.
- Whether sorting rewrites formulas as Excel does or keeps them as text; how the operation interacts with the undo model.
- Acceptance fixtures: values, formulas, merges, hidden rows, filtered rows.

**AutoFilter**

- State model: range, per-column criteria, value checklist, text/number/date conditions, blanks.
- Interaction with hidden rows, sort, recalculation, merged cells and structural edits.
- Round-trip of `autoFilter` XML; preservation when the range is not editable.
- UI behaviour: dropdown, clear, reapply, indicator state.

Output: the semantics contract and conformance fixtures for tickets 07, 12 and 13.

## Answer

The contract is in [sort-autofilter.md](../spec/sort-autofilter.md), with 15 sort fixtures and 18 filter fixtures.

**Sort.** Operates on the selected range; a single cell expands to the current region. Top-to-bottom only in v1. Up to 64 keys, stable ties, header heuristic plus a "My data has headers" checkbox. Excel's documented order: numbers, text, FALSE before TRUE, all errors equal, blanks always last in both directions; text case-insensitive in Excel's character order. Formulas move with their rows and relative references shift by the row delta; absolute anchors, sheet names and defined names stay; formulas outside the range are not rewritten. Uniform one-row merges move; mixed merges block. Cell styles, comments and hyperlinks move; row heights and hidden flags stay. **Hidden rows are not moved** (Excel's documented behaviour); the user is warned once. Sorting a range that intersects a table or a legacy array formula is blocked. One sort is one undo transaction.

**AutoFilter.** One filter per sheet on a plain range; row 1 is always the header. Each column is either a typed value checklist (search, Select All, Blanks, NonBlanks) or up to two conditions joined by AND/OR. Text conditions support wildcards and are case-insensitive; number and date conditions compare typed values; date filters compare serials. Columns combine with AND; the header row stays visible. Manual-hidden and filter-hidden rows are separate layers, so clearing a filter restores only filter-hidden rows. Sorting while filtered reorders visible rows and reapplies the same filter. Structural edits transform the range and clear it if the header is deleted. `autoFilter` XML, `_FilterDatabase` and hidden rows round-trip; colour, icon, Top 10 and dynamic date filters are preserved read-only. Apply, change and clear are each one undo transaction.

**V2 cuts:** left-to-right sort, custom lists, format/case-sensitive sort, colour/Top 10/dynamic/advanced filters, editable table filters.
