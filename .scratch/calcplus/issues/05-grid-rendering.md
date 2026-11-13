# 05 — Grid rendering approach

Status: claimed
Type: research
Blocked by: none

## Question

How should the grid render and interact at Excel scale?

Compare canvas, virtualized DOM, hybrid canvas + DOM overlays, and credible existing grid libraries. Judge against:

- Excel's sheet limits (1,048,576 rows × 16,384 columns) with sparse data
- Scroll and paint performance over large used ranges
- Selection, active cell, fill handle, and in-cell editing overlays
- Frozen panes/splits, merged cells, variable row heights and column widths
- Zoom, gridlines, headers, and custom cell rendering
- Accessibility: screen-reader and keyboard-only grid semantics
- Implementation and maintenance cost

Recommend an approach, name the library (if any), and sketch the component boundary between viewport, renderer, and editing layer.
