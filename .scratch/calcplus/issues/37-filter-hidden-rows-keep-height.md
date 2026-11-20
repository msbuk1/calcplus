# 37 — Filtering blanks rows instead of hiding them

Status: resolved
Type: defect (filter-hidden rows kept their height in the view)
Blocked by: none

## Question

Reported by the maintainer with a capture (2026-11-20): with an AutoFilter
applied, rows that fail the criteria keep their height and paint as empty
rows (cell borders and empty backgrounds visible); only their row numbers
disappear. Excel collapses filter-hidden rows to zero height so the visible
rows sit next to each other.

## Resolution (2026-11-20)

The row geometry in `src/view/projection.ts` was built from
`sheet.rowDimensions` only. Filter-hidden rows are derived from the
criteria (sort-autofilter §3.3), never stored: the projection folded them
into the effective hidden set for `isRowHidden` (so the gutter skipped the
numbers) but not into the geometry (so the heights stayed). Every
effective-hidden row — filter-hidden and manual — now gets a zero-height
override in one place; clearing the filter rebuilds the projection and the
original heights return exactly. Scrolling and hit tests already skip
zero-size rows, so they stay correct. The model layers
(`manualHidden`/`filterHidden`/`effectiveHidden`) were not touched.

Tests: `tests/m5-view-integration.test.ts` (heights, hit tests, clear-filter
restoration) and `tests/e2e/filter-dropdown.spec.ts` (visible filtered rows
adjacent, hidden rows zero-height, presses land on the moved rows, clearing
restores every rect).
