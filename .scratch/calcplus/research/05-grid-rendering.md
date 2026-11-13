# Grid rendering approach — research

- **Ticket:** 05 — Grid rendering approach
- **Date:** 2026-11-13
- **Status:** complete
- **Scope:** how calcplus renders and interacts with a worksheet at Excel scale in a desktop browser. Charts, images, and shapes are v2 and out of scope.

## Recommendation

**Build a custom canvas grid renderer with DOM overlays for editing, focus, and accessibility. Do not adopt a third-party grid library as the rendering foundation.**

Concretely:

1. **Canvas 2D** draws the sheet body: cell backgrounds, borders, gridlines, text, selection, active-cell outline, fill handle, and headers. The canvas is viewport-sized and split into up to four pane canvases (frozen/split panes), never a sheet-sized bitmap.
2. **DOM overlays** provide the in-cell editor, the formula-bar editor, and a transparent `role="grid"` accessibility mirror aligned to the visible cells. The canvas is `aria-hidden`.
3. **Custom scrollbars** and a virtual scroll model replace a full-height DOM spacer. This is required because browsers cap element height below what 1,048,576 rows need at common row heights.
4. **No grid library.** AG Grid's Excel-essential interactions are Enterprise-only; Handsontable is proprietary; Glide Data Grid is React-only, has no top-row freezing, and has no stable release since 2024; canvas-datagrid does not document an Excel pane/merge model; Univer and SpreadJS are whole spreadsheet SDKs with commercial licensing and their own UI, not a renderer we can shape into an Excel clone.

The main risk is accessibility: canvas content is not exposed to assistive technology.[6] The mitigation is a virtualized DOM mirror using the ARIA grid pattern and `aria-rowcount`/`aria-colcount`, but that must be prototyped and tested with NVDA, JAWS, and VoiceOver before the architecture is frozen. See “Risks and unknowns”.

## 1. What “Excel scale” actually requires

Two different limits are often conflated:

- **Addressable limit** — the grid must let the user navigate, select, and edit any cell in a 1,048,576 × 16,384 sheet, with sparse data. This is a correctness/UX requirement.
- **Performance envelope** — how many populated cells and how large a used range must scroll smoothly. Ticket 11 sets this. It should not be assumed that a fully dense 17-billion-cell sheet renders smoothly; Excel itself would not.

Microsoft documents the worksheet limits as 1,048,576 rows × 16,384 columns, a maximum of four panes in a window, zoom from 10% to 400%, column width up to 255 characters, and row height up to 409 points.[1] Those numbers drive the renderer: the scroll model cannot depend on materialising rows, and pane geometry must support four independent regions.

Excel's freeze and split behaviours are different and both must be supported:

- **Freeze Panes** freezes the rows above and the columns left of the selected cell.[2]
- **Split** creates separate regions of the same worksheet that scroll independently; it can split on rows, columns, or both.[3]

Merged cells keep only the upper-left value (upper-right for right-to-left languages) and delete the other values.[4] The fill handle is the small square at the selection corner used to drag or auto-fill a series; double-click and drag behaviours are part of Excel's editing model.[5]

## 2. Options compared

### 2.1 Rendering strategies

Ratings are relative to this product, not general-purpose grids. “High” cost means a larger calcplus workstream.

| Approach | Scale to 1,048,576 × 16,384 | Selection / active cell / fill handle | Frozen panes and splits | Merges, variable sizes | Accessibility | Implementation cost | Verdict |
|---|---|---|---|---|---|---|---|
| **Canvas body + DOM editing/a11y overlays** (recommended) | Yes. Paints only the viewport; virtual scroll, no giant DOM.[6][7] | Full control; all overlays are custom-drawn and hit-tested. | Four pane canvases map directly to Excel's four-pane limit.[1][2][3] | Full control; merge map + prefix-sum geometry. | Hardest. Requires a DOM mirror; canvas alone is not exposed.[6][10] | High | **Recommended** |
| **Virtualized DOM cells** (custom, e.g. TanStack Virtual for the math) | Possible only with a custom/normalised scrollbar; a full-height spacer exceeds some engines' limits.[15][38] | Easy CSS overlays; cross-pane ranges and merged hit-testing get complex. | Four scroll containers; duplicated cells in frozen panes complicate a11y. | CSS handles text/size naturally; merges need absolute positioning. | Best: real DOM cells. | Medium–high | Viable fallback, not primary |
| **Hybrid DOM cells + canvas decoration** | Same scroll-height problem as DOM; canvas adds overlays only. | Easy. | Same DOM pane duplication. | DOM + canvas overlays. | Good. | Medium | Rejected for performance risk; same DOM scroll ceiling |
| **AG Grid Community / Enterprise** | DOM-virtualised; AG Grid has a “stretching” workaround for browser height caps, at the cost of accelerated scrolling.[14][15] | Cell/range selection, fill handle, and clipboard are **Enterprise** features.[12][13] | Pinned rows/columns; no independent split-pane model documented; pinning is a row-model feature.[18] | Column/row spanning only; AG Grid warns cell selection does not work correctly with spanning.[16] | ARIA support, but docs recommend disabling row/column virtualisation for screen readers.[17] | Low integration + licence + fighting the data-grid model | Rejected |
| **Handsontable** | DOM-virtualised; claims hundreds of thousands of rows.[20] | Spreadsheet-like selection and editing built in. | Column freezing documented;[22] no split-pane model documented. | True Excel-like merges.[21] | Targets WCAG 2.2 AA, but virtualisation can make screen readers announce the wrong row total.[20][23] | Medium + licence | Rejected on licensing |
| **Glide Data Grid** | Canvas; advertises millions of rows and native scrolling.[24] | Editing built in; range selection exists, but no Excel-style fill handle is documented. | `freezeColumns` and `freezeTrailingRows` only; no top-row freeze in the public props.[25][27] | Variable row heights, merged cells, custom canvas cell renderers.[24] | Claims accessibility; maintainers say flaws are likely and no independent audit is referenced in the README.[24] | Medium; React-only and no stable release since 2024.[24][26] | Rejected |
| **canvas-datagrid** | Single canvas, immediate mode; advertises unlimited rows/columns.[28] | Basic grid selection/editing; no spreadsheet fill handle documented. | No Excel freeze/split pane model documented. | No Excel merge model documented. | No accessibility evidence found. | Low–medium | Rejected |
| **Whole spreadsheet SDK** (Univer, SpreadJS, Syncfusion) | Designed for large sheets; Univer uses a canvas render engine.[32] | Complete spreadsheet interactions out of the box. | Freeze panes supported; split-pane support varies and was not verified. | Supported. | Vendors claim ARIA/WCAG; no independent audit reviewed. | Low integration, high licence/lock-in | Considered as build-vs-buy, rejected for v1 |

### 2.2 Libraries and licences at a glance

| Option | Licence | Rendering model | Excel fit | Maintenance signal | Main blocker |
|---|---|---|---|---|---|
| AG Grid | Community MIT; Enterprise commercial EULA[11] | Virtualized DOM[14] | Data grid, not a sheet; Excel interactions paywalled | Active, frequent releases | Range selection + clipboard Enterprise[12][13] |
| Handsontable | Proprietary non-commercial/commercial; last MIT 6.2.2 in 2018[19] | Virtualized DOM[20] | Closest data-grid feature set to a sheet | Active (v18.1.2) | Commercial production licence required[19] |
| Glide Data Grid | MIT[24] | Canvas[24] | Data grid; frozen left columns + trailing bottom rows only[25] | Repo commits in 2026, but npm stable 6.0.3 is from 2024-02-03[26] | React-only, no top freeze, unstable release channel |
| canvas-datagrid | BSD-3-Clause[28] | Canvas[28] | Generic data grid | Active | No Excel pane/merge model |
| TanStack Virtual | MIT[30] | Headless virtualizer, no markup/styles[29] | Not a grid; can combine two axes[29] | Active | You still build the entire grid |
| Univer | Apache-2.0 core; commercial Pro[31] | Canvas render engine[32] | Full sheet SDK, own UI | Very active | XLSX conversion and advanced features commercial;[31] not Excel chrome |
| SpreadJS | Commercial[34] | Spreadsheet engine | Very Excel-like, XLSX, 500+ functions[33] | Active | $999/developer + $12,499/domain/year[34] |
| Syncfusion Spreadsheet | Commercial; community licence for small orgs[35][36] | Spreadsheet engine | Excel-like, XLSX | Active | Commercial unless eligibility criteria met[36] |

## 3. Evidence by axis

### 3.1 Scroll and paint performance over large used ranges

**Canvas.** The canvas is only as large as the viewport, so the browser's maximum canvas dimensions are not a constraint: MDN notes most browsers exceed 10,000 × 10,000 px but iOS caps at 4,096 × 4,096, and a canvas that exceeds the limit becomes unusable.[6] A 2,000 × 1,200 CSS-px viewport at DPR 2 is 4,000 × 2,400 backing pixels, comfortably inside desktop limits. Canvas drawing is immediate-mode; only the visible cells are drawn. MDN recommends layering canvases, batching calls, rendering differences only, and avoiding text rendering where possible.[7] Those recommendations map directly to a spreadsheet: a static/background layer, a content layer that can be blitted on scroll, and an overlay layer for selection.

**Virtualized DOM.** AG Grid's own docs describe the standard technique: render only visible rows/columns, with a row buffer, and cap rendered rows at 500 as a safety measure.[14] That works for data grids, but the scroll container still needs a height representing all rows. AG Grid documents that browsers impose a maximum `div` height — about 32,000,000 px on Chrome 118 — and that its “stretching” workaround makes scrolling appear faster when the row set is taller than the cap.[15] Chromium's tracker records the element-size limit as 33,554,428 px;[37] Gecko historically caps CSS lengths at 17,895,697 px.[38]

The arithmetic matters: 1,048,576 rows × 20 px = 20,971,520 px. That is under Blink's cap but **over Gecko's historical cap**, and it exceeds Blink's cap once the average row height passes roughly 32 px. A full-height spacer is therefore not portable. The renderer must either use a normalised scroll model (map scrollbar position to a virtual offset) or a custom scrollbar. A custom scrollbar is also the practical way to match Excel's scrollbar chrome and to avoid AG Grid's accelerated-scroll artefact.

**What the evidence does not establish.** There is no public, rigorous benchmark of canvas vs virtualized DOM at 1M × 16k with sparse data. Glide's FAQ says its earlier `react-virtualized` implementation became choppy once hundreds of DOM elements had to be loaded/unloaded per frame, which is why it moved to canvas.[24] That is a vendor engineering account, not an independent benchmark. The decision should be validated by the prototype benchmark described in §6 and by ticket 11's frame budget.

### 3.2 Selection, active cell, fill handle, and in-cell editing

These are the features that most clearly distinguish a spreadsheet from a data grid.

- **Selection** in Excel is a sheet-level range (or multiple ranges), not a row-model selection. It can include whole rows/columns and can span frozen panes. Drawing it is trivial on canvas: intersect the selection rectangle with each pane's visible sheet rectangle and paint each segment. Hit-testing must map a point to a pane, then to a sheet cell, then to a merge anchor.
- **Active cell** is separate from the selection: it is the anchor that receives typing, the formula bar, and the focus ring. In a canvas renderer it is just another drawn rectangle; in the DOM mirror it is the focused `gridcell`.
- **Fill handle** is a small square at the bottom-right of the selection; dragging previews a target range and commits an auto-fill.[5] Excel also supports double-clicking the fill handle to fill down adjacent rows.[41] That behaviour should be treated as a fidelity target for ticket 10. A canvas overlay can draw the handle, preview outline, and tooltip; a pointer controller owns the drag and autoscroll. Among the libraries reviewed, only AG Grid exposes a fill handle, and it is Enterprise.[12]
- **In-cell editing** is inherently a DOM problem. Contenteditable/`textarea`/`input` gives caret, selection, IME, accessibility, spellcheck, and native text editing for free. The canvas should hide the cell's painted text and place a DOM editor over the cell rectangle. The editor must match font, alignment, padding, zoom, and merged-cell size, and must grow for long text/multiline editing. The formula bar is a second editor bound to the same edit session.

A pure-canvas editor (drawing a caret and handling keystrokes) is a trap: it loses IME, screen-reader text editing, native selection, and platform input behaviour. The editing layer must be DOM.

### 3.3 Frozen panes, splits, merged cells, variable sizes

Excel's four-pane limit[1] is a gift to the renderer: the maximum pane layout is:

```
┌──────────┬────────────────────────────┐
│ top-left │ top-right                  │
│ (frozen  │ (frozen rows, scroll cols) │
│  rows &  │                            │
│  cols)   │                            │
├──────────┼────────────────────────────┤
│ bottom-  │ body                       │
│ left     │ (scroll rows & cols)       │
│ (scroll  │                            │
│  rows,   │                            │
│  frozen  │                            │
│  cols)   │                            │
└──────────┴────────────────────────────┘
```

- **Freeze**: top-right shares the body's horizontal offset; bottom-left shares the body's vertical offset; top-left is fixed.
- **Split**: each of the four panes has its own horizontal and vertical offset, matching Excel's independent scrolling.[3]
- **Headers**: a column-header strip and row-header strip are drawn across/along the panes; the exact header repetition behaviour for split panes should be verified against Excel in ticket 10 rather than guessed here.
- **Merged cells**: render only the anchor, clipped to the merged rectangle; covered cells render nothing. Hit-testing maps a covered cell to the anchor. Editing uses the merged rectangle. A merge that crosses a freeze/split boundary is an edge case: clip it per pane and decide whether content appears only in the anchor's pane; verify against Excel.
- **Variable row heights / column widths**: keep defaults plus sparse overrides. A Fenwick tree (binary indexed tree) or prefix-sum over overrides gives O(log n) mapping between index and pixel offset; 16,384 columns can even use a dense prefix array, while 1,048,576 rows should not allocate per-row objects.

The pane architecture is not speculative: Univer's canvas render engine creates viewports for main, top/bottom row headers, and left/right column headers, and uses separate document instances for cell editing and the formula bar.[32] That is the same separation recommended here.

### 3.4 Zoom, gridlines, headers, custom cell rendering

- **Zoom** must re-render at the new scale, not CSS-transform the canvas. MDN warns that changing a canvas's displayed size with CSS scales the rendered bitmap and can distort it; set the canvas backing dimensions and draw at the target scale.[6] `devicePixelRatio` changes with page zoom, so the canvas must resize and repaint when DPR changes.[8] Excel's zoom range is 10–400%.[1] Zoom should scale row heights, column widths, font sizes, and gridline positions together.
- **Gridlines** are a light gray 1-device-pixel grid under cell fills/borders; borders override gridlines. Draw at half-pixel offsets for crispness, and skip gridlines outside the used range if matching Excel's default.
- **Headers** need their own hit regions for selection, resize, and autofit. They can be drawn on canvas but must be represented in the DOM mirror as `columnheader`/`rowheader`.
- **Custom cell rendering** is a `CellPainter` interface: measure, paint, hit-test. V1 needs text, numbers, dates/times, booleans, errors, wrapped text, indentation, rotation, borders, fills, and number formats. Rich text, images, sparklines, and charts are v2. Canvas can draw all of these, but text layout (wrapping, ellipsis, bidi, shaping) is our code; browsers shape a single `fillText` run but do not lay out paragraphs for us.

### 3.5 Accessibility

This is the weakest part of every canvas approach, and the evidence should not be softened.

MDN is explicit: “The `<canvas>` element on its own is just a bitmap and does not provide information about any drawn objects. Canvas content is not exposed to accessibility tools as semantic HTML is. In general, you should avoid using canvas in an accessible website or app.”[6] That rules out a canvas-only grid. It does not forbid a canvas visual layer paired with an equivalent DOM representation, but it puts the entire accessibility burden on that mirror.

The ARIA grid pattern supplies the required semantics: `role="grid"` on the container, `role="row"` on rows, `role="gridcell"` on cells; `columnheader`/`rowheader` for headers; `aria-selected` for selection; `aria-readonly` where editing is disabled; and keyboard navigation with arrows, Home/End, Ctrl+Home/End, Page Up/Down, Shift+Arrow, Ctrl+Space, and Shift+Space.[10] The pattern explicitly supports rows/columns that are not all in the DOM: when only part of the rows is present, set `aria-rowcount` to the total and `aria-rowindex` to each row's absolute position; the spec gives an example of a 2,000-row grid that has only rows 100–102 in the DOM.[9][10] That is exactly the virtualized-grid case. `aria-colcount`/`aria-colindex` do the same horizontally.[10]

`aria-activedescendant` is an alternative focus model: keep DOM focus on the grid container and point at the active cell.[9] The spec says the active descendant should be visible and in view when focused.[9] For a spreadsheet, the simpler and more standard model is **roving tabindex**: one tab stop, DOM focus on the active `gridcell`, move focus as the active cell moves. That keeps the screen reader's application mode anchored to a real cell and avoids a container-level focus model that some screen readers handle inconsistently. Whichever model is chosen, the active cell element should be persistent and moved/reused; recreating it on every scroll can drop focus.

The DOM mirror design:

- Root `div role="grid"` with `aria-rowcount="1048576"`, `aria-colcount="16384"`, `aria-label` for the sheet, and one tab stop.
- Visible rows as `div role="row" aria-rowindex="<absolute 1-based>"`, visible cells as `div role="gridcell" aria-colindex="<absolute 1-based>"`, with `aria-selected` and `aria-readonly` as appropriate.
- The active cell is a real focusable element with `tabindex="0"`; all other mirrored cells use `tabindex="-1"`.
- The mirror is positioned exactly over the canvas, with transparent text/background, `pointer-events:none`, `user-select:none`, and `forced-color-adjust:none`; it must **not** use `display:none` or `visibility:hidden`, which remove it from the accessibility tree. Keep DOM order equal to visual order.
- The canvas is `aria-hidden="true"`.
- Header cells are mirrored as `columnheader`/`rowheader` with absolute indices.
- The formula bar is a labelled textbox; its value tracks the active cell's formula when not editing.
- Merged cells set `aria-rowspan`/`aria-colspan` on the anchor cell.[10]
- In split mode, mirror only the active pane. The duplicated visual panes are a presentation feature; exposing the same cells multiple times in the accessibility tree would confuse screen readers.

This is still a hypothesis at scale. AG Grid, which uses real DOM cells, tells users that screen readers assume all grid elements are loaded, and recommends `ensureDomOrder` plus disabling row and column virtualisation for accessibility — at the cost of memory and performance.[17] If a DOM-based grid with 1M rows punts on virtualization for screen readers, a canvas grid with a virtualized mirror is not obviously better. It must be tested, not assumed.

A pragmatic fallback should be specified now: an **accessibility mode** that renders the current used range (or a page of it) as a real DOM grid with pagination, using the same selection/edit commands. It is uglier and slower, but it gives assistive-technology users a fully semantic grid instead of a fragile mirror. This is analogous to AG Grid's guidance to disable virtualization for screen readers.[17] The decision to expose such a mode (or to invest further in the mirror) should be made after the AT prototype.

### 3.6 Licensing and maintenance

- **AG Grid**: Community is MIT; Enterprise requires a commercial EULA.[11] Cell selection is Enterprise,[12] clipboard is Enterprise,[13] and the fill handle lives under the Enterprise cell-selection module.[12] Those are not optional for an Excel clone; implementing them on top of Community means rebuilding the features the library is being adopted for.
- **Handsontable**: source is available under a proprietary non-commercial licence or a commercial licence; the last MIT release was 6.2.2 on 2018-12-19.[19] A commercial product needs the paid licence.
- **Glide Data Grid**: MIT, so commercial use is permitted.[24] But it is a React component (peer dependencies React 16–18 in the published 6.0.3), its last stable npm release is 6.0.3 from 2024-02-03 despite later repository commits,[26] it exposes frozen left columns and trailing bottom rows but not top-row freezing,[25][27] and its own FAQ concedes that the accessibility implementation likely has flaws because the primary developers are not accessibility users.[24]
- **canvas-datagrid**: BSD-3-Clause, a single canvas, immediate-mode, claims unlimited rows/columns.[28] It is an active project, but it is a generic data grid; Excel freeze/split panes and merge semantics are not part of its documented model.
- **TanStack Virtual**: MIT, headless, and can virtualize both axes to make a grid-like window.[29][30] It is a useful source of math and API ideas, but it renders nothing and provides no selection, editing, headers, panes, or accessibility. Adopting it would still leave the entire spreadsheet renderer to build.
- **Univer**: the open-source core is Apache-2.0 and includes basic sheet editing, formulas, formatting, filtering, sorting, and tables; xlsx conversion, charts, pivot tables, and collaboration are commercial, with Business starting at $99/month.[31] Its canvas engine already models multiple viewports and separate cell/formula editors.[32] It is a credible “buy” if the goal were any spreadsheet; it is a poor fit for an Excel-desktop clone because the product's UI and interaction model would be Univer's, not Excel's.
- **SpreadJS**: commercial, $999 per developer plus $12,499 per domain per year for commercial/SaaS deployment, with Excel-like UI, xlsx import/export, and 500+ functions.[33][34] It is the strongest “buy” option for fidelity, and the most expensive.
- **Syncfusion Spreadsheet**: commercial with a free community licence for organisations under $1M revenue, 5 or fewer developers, and 10 or fewer employees, plus other eligibility conditions.[35][36] It is a reasonable low-cost buy for a small team, but it is still a third-party spreadsheet UI rather than an Excel-shaped renderer.
- **Abandoned/stale traps**: Luckysheet is archived and its README points to Univer;[39] x-spreadsheet's repository says it migrated to `@wolf-table/table` and has not been pushed since 2024.[40] These are useful reference implementations, not dependencies.

### 3.7 Why not just use AG Grid Enterprise?

It is the closest data-grid option, and if the product were a data grid it would be the default. For this product:

1. **The interaction model is wrong.** AG Grid is a row/column data grid. Excel is a 2D sheet with an active cell, a formula bar, merged regions, frozen panes, and fill semantics. AG Grid's own docs warn that cell selection “will not work correctly” when column spanning is used, and that spanning cannot cross pinned regions.[16] That is a direct conflict between two Excel-essential features.
2. **The Excel features are the paid ones.** Range selection and clipboard are Enterprise.[12][13] So is the fill handle configuration.[12]
3. **Large-row scrolling is a workaround.** AG Grid's “stretching” makes scrolling faster than the finger when the row set exceeds the div height cap.[15] That is acceptable for a data grid; it is visibly wrong in an Excel clone.
4. **Accessibility at scale is unresolved.** AG Grid's own recommendation is to disable virtualization for screen readers.[17] A 1M-row sheet cannot do that.

AG Grid is a reasonable benchmark for features and a source of ideas (row buffer, pinned regions, ARIA wiring), but not the foundation.

## 4. Recommendation in detail

### 4.1 Component boundary

```
┌────────────────────────────── GridView (DOM root) ──────────────────────────────┐
│  owns focus, sizing, zoom, ARIA root, lifecycle                                  │
│                                                                                  │
│  ┌──────────────┐   ┌──────────────────┐   ┌───────────────────────────────┐     │
│  │ ScrollController │ │ InteractionController │ │ AccessibilityLayer        │     │
│  │ wheel/track/  │   │ pointer, keyboard, │   │ role=grid mirror,           │     │
│  │ keyboard/     │   │ drag, fill handle, │   │ active cell, aria-rowcount  │     │
│  │ scrollToCell  │   │ resize, clipboard  │   │                             │     │
│  └──────┬───────┘   └────────┬─────────┘   └──────────────┬────────────────┘     │
│         │                    │                            │                      │
│         ▼                    ▼                            ▼                      │
│  ┌──────────────────────────────────────────┐   ┌──────────────────────────┐    │
│  │ Viewport / PaneGeometry                  │   │ EditingLayer             │    │
│  │ pane rects, freeze/split, visible ranges,│   │ DOM textarea/input,      │    │
│  │ row/col ↔ px (Fenwick), hit testing,     │   │ IME, caret, commit/cancel│    │
│  │ merge map, scroll normalisation          │   │ formula-bar binding      │    │
│  └──────────────────┬───────────────────────┘   └──────────────┬───────────┘    │
│                     │                                          │                │
│                     ▼                                          ▼                │
│  ┌──────────────────────────────────────────┐   ┌──────────────────────────┐    │
│  │ CanvasRenderer                           │   │ SelectionModel           │    │
│  │ per-pane content + overlay layers;       │   │ active cell, ranges,     │    │
│  │ CellPainter registry; dirty rects        │   │ anchor/focus, modes      │    │
│  └──────────────────┬───────────────────────┘   └──────────────────────────┘    │
│                     │                                                           │
└─────────────────────┼───────────────────────────────────────────────────────────┘
                      │ reads
                      ▼
             SheetView (read-only projection from the workbook model)
```

`SheetView` is read by `GridView`/`Viewport` while building each `RenderFrame`; the renderer itself only consumes the frame, so it can be swapped without touching the workbook or the projection.

**Viewport layer** (no painting):

- Converts pane rects + scroll offsets + zoom into visible sheet ranges.
- Maps row/column index to pixel offset and back, using defaults + sparse overrides + a Fenwick tree for rows.
- Owns freeze/split configuration and pane decomposition.
- Hit-tests points to pane, cell, merge anchor, header, resize gutter, or fill handle.
- Exposes `scrollToCell`, `visibleRange(pane)`, `cellRect(cell, pane)`.

**Renderer layer** (no model access, no DOM reads):

- Receives a plain `RenderFrame` per pane: visible range, cell display records, styles, selection segments, merge clips, gridline/border settings, zoom, DPR.
- Draws content on a per-pane content canvas and overlays on a per-pane overlay canvas.
- Can be swapped: a DOM renderer could implement the same `RenderFrame` consumer if the canvas path fails accessibility or performance tests.
- Never mutates the workbook; never issues commands.

**Editing layer** (DOM):

- Owns the editor element, caret, IME composition, selection inside the editor, and formula-bar sync.
- Positions itself from `Viewport.cellRect`, including merged cells and zoom.
- Emits a single commit/cancel event to the command layer; one user edit becomes one undoable command.
- Does not paint cell values; it hides the underlying painted text while editing.

**Accessibility layer**:

- Maintains a DOM mirror of the visible range plus the persistent active cell.
- Owns `role=grid`, `aria-rowcount`, `aria-colcount`, header roles, `aria-selected`, and focus.
- Is always present, not a special mode, unless the accessibility-mode fallback is enabled.

**Interaction/controller layer**:

- Translates pointer, wheel, keyboard, and clipboard events into selection changes and commands.
- Owns drag state for selection, fill handle, column/row resize, and autoscroll.
- Does not draw.

### 4.2 Data flow

1. Workbook model (ticket 07) owns cell values, formulas, styles, merges, row heights, column widths, and freeze/split state.
2. A `SheetView` projection formats values (number formats, errors, booleans) and resolves effective styles/merges into `CellDisplay` records. The renderer never runs a number-format engine.
3. `GridView` subscribes to model change events. A change produces a dirty region (cell, range, row, column, or all).
4. On the next `requestAnimationFrame`, the renderer repaints only dirty panes/regions. Scroll uses a fast path that blits the existing canvas and repaints only the exposed strip.
5. Interactions mutate the `SelectionModel` immediately (cheap, view-only) and issue commands for workbook changes (edit, fill, resize, merge). Commands go through the undo stack (ticket 09).
6. The accessibility layer syncs after the same frame: active cell content/position, visible row/cell nodes, selection flags.

### 4.3 Rendering pipeline

**Panes and layers.** One `<canvas>` per pane for content, one for overlays. Four panes × two layers = eight viewport-sized canvases, plus header canvases. Each canvas is sized to its CSS rect × `devicePixelRatio`, with `ctx.setTransform(dpr, 0, 0, dpr, 0, 0)` so drawing code works in CSS pixels.[8] Overlay canvases are `pointer-events:none`.

**Scroll.** Do not create a 20-million-pixel-tall spacer. Keep a virtual scroll offset in sheet pixels; custom scrollbars map thumb position to offset. On wheel/trackpad, update the offset, clamp it, and schedule a paint. On scroll, for each pane, `drawImage` the previous content canvas shifted by the delta and repaint only the newly exposed strip; this avoids redrawing hundreds of text runs per frame. MDN explicitly recommends layered canvases and rendering differences only.[7] For large jumps (scrollbar drag, Ctrl+End, Find), do a full repaint.

**Invalidation.** Dirty rectangles from model changes. A cell edit repaints that cell (and its merged rect); a style change repaints the affected range; a row-height/column-width change repaints the pane; a zoom change repaints everything. Coalesce all invalidations into one rAF.

**Text.** Set `ctx.font` from the cell style at the zoomed size, then draw with clipping. Cache `TextMetrics` keyed by font + string where profiling shows it matters. Implement overflow (Excel lets text spill into adjacent empty cells), wrap, ellipsis, indent, and rotation in the painter. Use `Intl.Segmenter` for grapheme-safe truncation if needed. Rich text runs are a v2 extension of `CellPainter`.

**Borders and gridlines.** Draw gridlines first, then cell fills, then borders. Resolve shared edges deterministically (for example, right/bottom border wins) and align to device pixels. Excel's border-collision rules should be verified in ticket 10.

**Headers.** Draw column letters (A…Z, AA…) and row numbers; highlight the selected rows/columns; draw resize gutters and the freeze/split divider. Header double-click triggers autofit; header drag triggers resize; header click selects the row/column.

**DPR and browser zoom.** Listen for `ResizeObserver` and DPR changes via `matchMedia('(resolution: Xdppx)')`; resize backing stores and repaint.[8] Do not CSS-scale the canvas to implement Excel zoom; re-render at the target scale instead.[6]

### 4.4 Editing layer

- **Entry**: double-click, F2, or typing a character. Typing replaces the cell; F2/double-click places the caret.
- **Element**: a `<textarea>` for multiline/wrap and `<input>` for single-line, absolutely positioned over the cell or merged rect, with matching font/alignment/padding/zoom. It grows down/right as text exceeds the cell, up to a limit; it scrolls internally beyond that.
- **Formula vs value**: when editing, the editor shows the formula text; when not editing, the cell shows the formatted value. The formula bar mirrors the editor.
- **IME**: never commit on Enter during `compositionstart`…`compositionend`; let the composition finish first.
- **Commit/cancel**: Enter commits and moves down (Excel default), Tab commits and moves right, Shift+Enter/Tab reverse, Esc cancels, clicking another cell commits. Alt+Enter inserts a newline. Commit produces one command for undo/redo.
- **Selection while editing**: arrow keys move the caret inside the editor, not the active cell. This matches Excel and the ARIA grid pattern's “editing disables grid navigation” guidance.[10]
- **Merged cells**: editor covers the merged rectangle; content is stored in the anchor; covered cells stay empty.
- **Read-only/locked cells**: block editing and set `aria-readonly` on the mirror (protection is v2, but the seam belongs here).

### 4.5 Accessibility layer

- Root: `role="grid"`, `aria-rowcount="1048576"`, `aria-colcount="16384"`, `aria-label`, `aria-multiselectable="true"` if multiple ranges are supported.
- Rows/cells: only the visible range is mirrored, with absolute `aria-rowindex`/`aria-colindex`. This is the mechanism the ARIA spec provides for partially loaded grids.[9][10]
- Active cell: persistent, focusable (`tabindex="0"`), positioned over the canvas cell, transparent, with `aria-selected="true"`. All other mirrored cells are `tabindex="-1"`. Focus movement follows the active cell; the canvas draws the visible focus ring.
- Headers: mirrored as `columnheader`/`rowheader` with absolute indices; merged anchors get `aria-rowspan`/`aria-colspan`.[10]
- Formula bar: labelled textbox, with the active cell's formula when not editing.
- Keyboard: implement the APG grid keys[10] plus Excel-specific keys (Ctrl+Arrow to data edge, Ctrl+Home/End, Page Up/Down, Ctrl+Space, Shift+Space, F2, Delete, Alt+Enter). Where APG and Excel disagree, follow Excel and document the divergence.
- Canvas: `aria-hidden="true"` so the mirror is the only grid exposed.
- Fallback: if AT testing shows the mirror is unreliable, expose an accessibility mode that renders the used range as a paginated real DOM grid, reusing the same selection/edit commands. Do not pretend the mirror is guaranteed to work.

## 5. Risks and unknowns

1. **Canvas accessibility at scale is unproven.** MDN says canvas content is not exposed to accessibility tools and recommends avoiding canvas in accessible apps.[6] The ARIA virtual-grid mechanism is specified,[9][10] but I found no primary source that demonstrates it working well for a 1M-row spreadsheet across NVDA, JAWS, and VoiceOver. AG Grid's advice to disable virtualization for screen readers[17] is evidence that the industry has not solved the general case. **Mitigation:** build the a11y mirror prototype in the first grid milestone and test with real screen readers; keep the accessibility-mode fallback.
2. **No independent canvas-vs-DOM benchmark at Excel scale.** Vendor claims (Glide,[24] AG Grid,[14][15] Univer[32]) and browser limits[37][38] are the strongest evidence available. **Mitigation:** the ticket-11 benchmark corpus must include a 1M-row sparse sheet, a 100k-row dense sheet, tall rows, frozen panes, and merged cells; measure scroll frame time, paint time, and memory. Treat the renderer choice as provisional until that passes.
3. **Scroll-height portability.** A full-height spacer would exceed Gecko's historical cap at 1M rows × 20 px, and would exceed Blink's cap at taller average rows.[15][37][38] **Mitigation:** custom/normalised scrollbars from day one; test at 1,048,576 rows in Chrome, Firefox, and Safari.
4. **Text layout and IME fidelity.** Canvas handles a single shaped run, not paragraph layout; wrapping, ellipsis, bidi, and rich text are our code. The DOM editor handles IME, but the switch between painted text and editor can flicker or misalign at 10–400% zoom. **Mitigation:** a text-layout spike with long strings, CJK, Arabic/Hebrew, emoji, rotated text, and zoom extremes.
5. **Merge × freeze × selection edge cases.** Excel's exact behaviour for a merged cell crossing a freeze/split boundary, and for selection across panes, is not documented in the sources reviewed. **Mitigation:** capture Excel behaviour as golden fixtures in ticket 10/13; define clipping and hit-testing rules explicitly.
6. **Custom scrollbar fidelity.** Trackpad inertia, wheel delta modes, thumb sizing, and keyboard scrolling must feel native; a bad custom scrollbar is worse than a native one. **Mitigation:** prototype wheel/trackpad handling early and compare against Excel.
7. **Memory for large styled ranges.** Sparse cell maps plus interned style IDs should keep memory proportional to populated cells, not to 17 billion addresses. This is a workbook-model concern (ticket 07), but the renderer must not allocate per-cell objects for empty cells. **Mitigation:** measure with a 1M-cell styled fixture.
8. **Browser canvas limits on low-end/Safari hardware.** Desktop Safari has higher limits than iOS, but canvas memory is still finite.[6] Viewport-sized canvases mitigate this; do not create a sheet-sized canvas.

## 6. Spikes to run before the design is frozen

1. **Pane/scroll spike**: four pane canvases, custom scrollbar, 1,048,576 rows, sparse random cells, variable row heights. Measure frame time during wheel, thumb drag, and Ctrl+End. Verify no browser-height dependence.
2. **A11y spike**: mirror the visible range + active cell; test NVDA/Firefox, NVDA/Chrome, JAWS/Chrome, VoiceOver/Safari. Verify cell address, value, formula, selection, and edit announcements. Time-box a fallback decision.
3. **Editing spike**: textarea overlay over normal and merged cells at 10%, 100%, 400% zoom; IME (CJK), Alt+Enter, long text, formula bar sync, commit/cancel, undo coalescing.
4. **Fidelity spike**: screenshot comparison against Excel for gridlines, selection colours, active-cell border, fill handle, headers, freeze divider, merged cells, borders, and number alignment.

Each spike should be throwaway code outside the main app; the goal is to answer the question, not to become the product.

## 7. Consequences for the spec package

- The technical design should treat the grid as **one deep module** with a narrow `GridView` facade and explicit internal seams (`SheetView`, `Viewport`, `CanvasRenderer`, `EditingLayer`, `AccessibilityLayer`). The renderer must be replaceable without touching the workbook model.
- The PRD should state the **addressable** limit (all 1,048,576 × 16,384 cells reachable) separately from the **performance envelope** (ticket 11).
- The milestones should put the **a11y spike in the first grid milestone**, not at the end. It is the only risk that can invalidate the rendering choice.
- The design should explicitly reject a sheet-sized canvas and a full-height DOM spacer; both are traps at Excel scale.
- The design should reserve a **DOM-cell renderer** as a fallback behind the same `RenderFrame`/`SheetView` seam. That is cheap insurance and makes the accessibility-mode fallback credible.

## Sources

All sources accessed 2026-11-13.

1. Microsoft Support — *Excel specifications and limits* (rows/columns, panes, zoom, column width, row height, noncontiguous selection). https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
2. Microsoft Support — *Freeze panes to lock rows and columns*. https://support.microsoft.com/en-us/excel/get-started/freeze-panes-to-lock-rows-and-columns
3. Microsoft Support — *Split panes to lock rows or columns in separate worksheet areas*. https://support.microsoft.com/en-us/excel/split-panes-to-lock-rows-or-columns-in-separate-worksheet-areas
4. Microsoft Support — *Merge and unmerge cells in Excel*. https://support.microsoft.com/en-gb/excel/get-started/merge-and-unmerge-cells-in-excel
5. Microsoft Support — *Fill data automatically in worksheet cells* (fill handle). https://support.microsoft.com/en-us/excel/get-started/fill-data-automatically-in-worksheet-cells
6. MDN — *`<canvas>` HTML graphics canvas element* (maximum size, CSS scaling distortion, accessibility). https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Elements/canvas
7. MDN — *Optimizing canvas* (layered canvases, render differences only, high-DPI scaling, avoid text rendering). https://developer.mozilla.org/en-US/docs/Web/API/Canvas_API/Tutorial/Optimizing_canvas
8. MDN — *Window: devicePixelRatio property* (page zoom changes DPR; canvas resolution correction). https://developer.mozilla.org/en-US/docs/Web/API/Window/devicePixelRatio
9. W3C — *Accessible Rich Internet Applications (WAI-ARIA) 1.2* (`aria-rowcount`, `aria-colcount`, `aria-rowindex`, `aria-colindex`, `aria-activedescendant`). https://www.w3.org/TR/wai-aria-1.2/
10. W3C WAI — *ARIA Authoring Practices Guide: Grid (Interactive Tabular Data and Layout Containers) Pattern*. https://www.w3.org/WAI/ARIA/apg/patterns/grid/
11. AG Grid — *Community vs. Enterprise* (Community MIT; Enterprise commercial EULA). https://www.ag-grid.com/javascript-data-grid/community-vs-enterprise/
12. AG Grid — *Cell Selection* (Enterprise; range selection, fill handle configuration). https://www.ag-grid.com/javascript-data-grid/range-selection/
13. AG Grid — *Clipboard* (Enterprise; copy/cut/paste). https://www.ag-grid.com/javascript-data-grid/clipboard/
14. AG Grid — *DOM Virtualisation* (visible-only rendering, row buffer, 500-row safety cap). https://www.ag-grid.com/javascript-data-grid/dom-virtualisation/
15. AG Grid — *Massive Row Count* (browser div height cap, “stretching”, accelerated scrolling). https://www.ag-grid.com/javascript-data-grid/massive-row-count/
16. AG Grid — *Column Spanning* (spanning is not full merge; cell selection breaks with spanning; cannot cross pinned regions). https://www.ag-grid.com/javascript-data-grid/column-spanning/
17. AG Grid — *Accessibility* (ARIA wiring; recommends `ensureDomOrder` and disabling row/column virtualisation for screen readers; known limitations). https://www.ag-grid.com/javascript-data-grid/accessibility/
18. AG Grid — *Row Pinning* (pinned rows above/below normal rows). https://www.ag-grid.com/javascript-data-grid/row-pinning/
19. Handsontable — *Software license* (proprietary; last MIT 6.2.2, 2018-12-19). https://handsontable.com/docs/javascript-data-grid/software-license/
20. Handsontable — *Row virtualization* (visible-only rendering; screen readers may announce wrong total rows). https://handsontable.com/docs/javascript-data-grid/row-virtualization/
21. Handsontable — *Merge cells* (Excel-like merge; covered values cleared). https://handsontable.com/docs/javascript-data-grid/merge-cells/
22. Handsontable — *Column freezing* (`fixedColumnsStart`, manual freeze). https://handsontable.com/docs/javascript-data-grid/column-freezing/
23. Handsontable — *Accessibility* (WCAG 2.2 A/AA target; tested with VoiceOver/NVDA). https://handsontable.com/docs/javascript-data-grid/accessibility/
24. Glide Data Grid — GitHub repository (MIT; canvas; millions of rows; editing; merged cells; variable rows; canvas-only custom renderers; FAQ on virtualized DOM and accessibility). https://github.com/glideapps/glide-data-grid
25. Glide Data Grid — *Important Props* (`freezeColumns`, `freezeTrailingRows`, `rowHeight`, `provideEditor`). https://docs.grid.glideapps.com/api/dataeditor/important-props
26. npm registry — `@glideapps/glide-data-grid` metadata (latest 6.0.3, published 2024-02-03; peer dependencies React 16–18). https://registry.npmjs.org/@glideapps/glide-data-grid
27. Glide Data Grid — issue #653, *Freeze Rows functionality* (closed and deferred to the still-open group-rows work #823). https://github.com/glideapps/glide-data-grid/issues/653
28. canvas-datagrid — GitHub repository (BSD-3-Clause; single canvas; immediate mode; unlimited rows/columns). https://github.com/TonyGermaneri/canvas-datagrid
29. TanStack Virtual — *Introduction* (headless; no markup/styles; two-axis grid-like virtualization). https://tanstack.com/virtual/latest/docs/introduction
30. TanStack Virtual — GitHub repository (MIT). https://github.com/TanStack/virtual
31. Univer — *Office SDK pricing* (open-source Apache-2.0 core; commercial xlsx conversion/advanced features; Business from $99/month). https://univer.ai/pricing/
32. Univer — *This is Univer* (canvas render engine; viewports for main, top/bottom, left/right; separate cell and formula-bar editors). https://docs.univer.ai/blog/this-is-univer
33. MESCIUS SpreadJS — product page (Excel-like spreadsheet, xlsx import/export, 500+ functions). https://developer.mescius.com/spreadjs
34. MESCIUS SpreadJS — *Pricing* ($999/developer; $12,499/single domain/year for commercial/SaaS). https://developer.mescius.com/spreadjs/pricing
35. Syncfusion — *Spreadsheet Editor SDK* (Excel-like editing, xlsx support, virtualization, ARIA claims). https://www.syncfusion.com/spreadsheet-editor-sdk
36. Syncfusion — *Community License* (eligibility: <$1M revenue, ≤5 developers, ≤10 employees). https://www.syncfusion.com/products/communitylicense
37. Chromium issue tracker — *Maximum element size is too restrictive for virtual scrolling* (33,554,428 px; page requires sign-in, value corroborated by AG Grid's measured ~32,000,000 px on Chrome 118[15]). https://issues.chromium.org/issues/41441393
38. Mozilla Bugzilla — *Image height limited to 17895698 pixels* (Gecko CSS lengths capped at 17,895,697 px; comment by Boris Zbarsky). https://bugzilla.mozilla.org/show_bug.cgi?id=552412
39. Luckysheet — GitHub repository (archived; README says upgraded to Univer). https://github.com/dream-num/Luckysheet
40. x-spreadsheet — GitHub repository (migrated to `@wolf-table/table`; last push 2024). https://github.com/myliang/x-spreadsheet
41. Microsoft Support — *Enter a formula* (double-click the fill handle to fill the formula down). https://support.microsoft.com/en-us/excel/enter-a-formula
