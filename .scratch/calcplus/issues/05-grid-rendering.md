# 05 — Grid rendering approach

Status: resolved
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

## Answer

Build a **custom Canvas 2D grid with DOM overlays**; use no third-party grid library as the foundation. Canvas draws cells, gridlines, headers, selection, fill handle, and up to four frozen/split panes; DOM provides the in-cell editor, the formula bar, and a transparent `role="grid"` accessibility mirror. The **viewport** owns pane geometry, scroll normalisation, row/column pixel mapping and hit testing; the **renderer** is a pure `RenderFrame` consumer; the **editing layer** is DOM and emits one command per commit.

Use **custom/normalised scrollbars**, never a full-height DOM spacer: 1,048,576 rows × 20 px is 20.97M px, beyond Gecko's historical 17.9M px CSS cap and beyond Blink's ~33.5M px cap at taller row heights. Canvas is viewport-sized, so browser canvas-size limits do not bind.

Trade-offs: this is the highest-cost route, and **canvas accessibility is the one unproven risk**. MDN says canvas content is not exposed to assistive technology; the ARIA `aria-rowcount` virtual-grid mechanism is specified but needs real NVDA/JAWS/VoiceOver testing, and AG Grid itself tells screen readers to disable virtualisation. Keep a DOM-cell renderer and an accessibility-mode fallback behind the same `SheetView`/`RenderFrame` seam.

Rejected: **AG Grid** (range selection, clipboard and fill handle are Enterprise-only); **Handsontable** (proprietary production licence); **Glide Data Grid** (React-only, no top-row freeze, no stable release since 2024); **canvas-datagrid** (no Excel pane/merge model); **Univer/SpreadJS** (commercial whole-SDKs with their own UI). Run the pane/scroll, accessibility, editing and fidelity spikes before freezing the design.

Full research: [05-grid-rendering.md](../research/05-grid-rendering.md)
