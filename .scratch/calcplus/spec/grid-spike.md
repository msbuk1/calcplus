# calcplus v1 — grid spike: custom canvas vs Univer

Resolved by [17 — Stress-test the grid decision: custom vs Univer](../issues/17-grid-spike-univer.md) on 2026-11-13. It tests the provisional decision in [05 — Grid rendering approach](../issues/05-grid-rendering.md) and the engine choice in [03](../issues/03-formula-engine.md) and [24](../issues/24-ironcalc-vs-hyperformula.md).

## 1. Verdict

**Keep the custom canvas grid. Reject Univer as the v1 foundation. Keep HyperFormula as the formula engine.**

The spike does not overturn the grid research. It sharpens it:

- Univer is the faster route to **a spreadsheet**. It is not the faster route to **a faithful Excel alternative**. Its chrome, model, formula engine and interaction model are Univer's, and replacing them is the product.
- The custom grid reached a usable, Excel-shaped grid in a throwaway spike and met the frame budget; Univer missed it at p95.
- Univer's stock canvas exposes **zero grid cells** to assistive technology, so the mirror would still have to be built inside a renderer we do not own. Accessibility is v2 in either case, so this point no longer decides the choice.
- Univer's open-source preset has **no `.xlsx` import or export**. That is a commercial Pro feature. The hardest part of the product, the OOXML layer from tickets 04 and 12, is not bought by adopting Univer.

## 2. What was built

Both spikes are throwaway code on [branch `prototype/grid-spike`](https://github.com/msbuk1/calcplus/tree/prototype/grid-spike/prototype/grid-spike). Screenshots are in `screens/`.

**Custom canvas** — 696 lines of HTML, CSS and JavaScript, no dependencies:

- viewport-sized canvas with a frozen header row and first column (four panes)
- 1,048,576 × 16,384 addressable sheet with 100k sparse cells
- custom scrollbars and a virtual scroll offset (no sheet-sized canvas, no full-height spacer)
- selection, active cell, fill handle, merged range
- DOM in-cell editor with F2, typing, Enter/Tab/Esc and formula-bar commit
- keyboard navigation (arrows, Tab/Enter, Page Up/Down, Home/End, Delete)
- basic number formats, headers, gridlines and frozen dividers
- a DOM accessibility mirror with `role="grid"`, absolute row/column indices and 403 exposed grid cells

**Univer** — 75 lines of integration plus a 7-line worker entry, using the stock `@univerjs/preset-sheets-core` 1.4.0:

- full Univer chrome: tabs, toolbar, formula bar, sheet tabs, status bar
- frozen header row and first column, 200 rows of sample data, formulas, number formats
- context menu, clipboard, zoom and fill handle from the preset
- no `.xlsx` import: the open-source preset has no exchange package

## 3. Measurements

All figures are from one Linux machine, Chromium, 1440 × 900, warm cache unless stated. They are directional, not the contractual benchmark in [performance](performance.md).

| Metric | Custom canvas | Univer 1.4.0 |
|---|---:|---:|
| Production bundle, gzip | **7.9 KB** (27.8 KB raw) | **1.78 MB** main entry; 2.73 MB including every locale chunk |
| First visible grid, warm (3 runs) | 125 / 118 / 146 ms | 395 / 299 / 1,297 ms |
| First visible grid, cold | ~150 ms | ~1.3 s |
| Scroll, p95 frame | **16.8 ms** | 18.9-20.6 ms |
| Scroll, worst frame | **16.8 ms** | 20.9-72.1 ms |
| Heap after scroll | 93 MB with 100k sparse cells | 94 MB with 200 rows |
| Grid cells in the accessibility tree | **403** | **0** |
| Spike size | 696 lines | 82 lines |

Caveats: the custom figure is one file with no framework; Univer's includes React and its whole UI. Scroll was driven by synthetic wheel events over 200 rows, not a 250k-cell formatted sheet. The ticket 11 harness must repeat this on the reference machine and the full corpus.

## 4. Fidelity comparison

| Behaviour | Custom spike | Univer stock preset |
|---|---|---|
| Excel-shaped chrome | Owned; needs building | Univer's chrome; must be replaced for Excel fidelity |
| Selection and active cell | Yes | Yes |
| Fill handle | Drawn; drag not implemented | Present |
| In-cell editing | Yes, DOM editor | Yes |
| Formula bar | Yes | Yes |
| Frozen row and column | Yes | Yes |
| Merges | Render and clip | Supported |
| Number formats | Basic | Yes |
| Context menu | Not in spike | Yes |
| Clipboard | Not in spike | Yes |
| Zoom | Static label | Yes |
| Keyboard navigation | Yes | Yes |
| Accessibility grid | DOM mirror with 403 cells | None exposed |
| `.xlsx` open/save | Not in scope | Not in the open-source preset |

The custom column is a 696-line sketch, not the finished grid. The missing items — fill drag, context menu, clipboard, zoom, text layout, IME, autofit — are known work in the grid milestone. The Univer column is what the free preset gives on day one; making it look and behave like Excel is the patch layer.

## 5. Why custom wins for this product

1. **Chrome is the product.** The destination is near-identical Excel chrome. Univer ships its own tabs, toolbar, menus, formula bar, sheet tabs and status bar. Replacing them means rebuilding the chrome and fighting the framework's layout and command model.
2. **The model is ours.** The canonical workbook model (ticket 07), the command registry (09) and HyperFormula (03, 24) are the chosen architecture. Adopting Univer means either abandoning them or keeping two models in sync. The spike confirms Univer's model is the source of truth inside its own UI.
3. **`.xlsx` is not bought.** The open-source preset has no exchange package. The commercial Pro tier owns `.xlsx` conversion. calcplus would still build the preservation-first OOXML layer from tickets 04 and 12, then map it into Univer's model — the largest and hardest part, unchanged.
- **Accessibility is not bought** (v2 work, but it decides nothing today). Univer's canvas exposes zero grid cells, so the mirror would have to be built inside a renderer we do not own; in our own renderer it is simpler and testable.
5. **The patch surface is real.** GenOffice adopted Univer and still owns roughly 256 modules around the renderer, including named patches for merges, thick borders, number formats, the formula lexer, selection wrap, sheet rename, RTL text, rich-text bidi, clipboard anchors and filter buttons (ticket 16). Those patches are coupled to Univer internals and to its version.
6. **Bundle and startup headroom.** 7.9 KB versus 1.78 MB gzip, and roughly 125 ms versus 300-400 ms warm to a visible grid. Both fit the ticket 11 budget, but the custom path leaves room for the model, the engine and the file layer.
7. **Licence is not the blocker.** Univer's core is Apache-2.0 and compatible with GPLv3. The blocker is fit, not licence.

**The honest counterpoint:** if the product were "a good spreadsheet" rather than "a faithful Excel desktop alternative", Univer would be the rational buy. It is maintained, capable and visually close. The destination decides the answer.

## 6. What changes in the grid plan

The custom approach is confirmed, with four additions:

1. **The accessibility mirror is v2.** The spike shows the mirror can expose rows and cells; usability with real screen readers is unproven and is not a v1 gate. Ticket 22 owns the target and the test for the later release.
2. **Keep a DOM-cell renderer behind the same seam.** `RenderFrame` already isolates the renderer. A real-DOM renderer is the fallback if the canvas mirror fails accessibility or the 1M-cell benchmark. It is cheap insurance and makes the accessibility-mode fallback credible.
3. **Own the missing interactions explicitly.** Fill-handle drag and double-click, context menus, clipboard, zoom, text layout, IME and autofit are part of the grid milestone, not hidden in "polish".
4. **Keep the grid and the engine decoupled.** Univer couples its renderer to its formula engine; calcplus keeps the `FormulaEngine` port and HyperFormula. The spike did not find a reason to revisit tickets 03 or 24.

The rest of ticket 05 stands: viewport-sized panes, no sheet-sized canvas, no full-height DOM spacer, custom scrollbars, DOM editing overlay, ARIA mirror, and four-pane freeze/split geometry.

## 7. When to revisit

Re-open the buy decision only if one of these becomes true:

- The v2 accessibility work finds the canvas mirror unusable and the DOM-cell fallback cannot meet the ticket 11 frame budget.
- The custom grid cannot hold the 250k-cell guaranteed envelope in the ticket 11 harness after the first two milestones.
- Univer ships an Apache-2.0 `.xlsx` exchange layer **and** a supported way to replace its chrome with Excel's, without forking its internals.
- The product goal changes from a faithful Excel alternative to a general spreadsheet.

## 8. Acceptance criteria

1. The v1 architecture uses the custom canvas grid; no Univer dependency ships.
2. HyperFormula remains the engine behind the `FormulaEngine` port.
3. The DOM-cell renderer can carry the v2 accessibility mirror without changing the model or commands.
4. A DOM-cell renderer can replace the canvas renderer behind `RenderFrame` without changing the workbook model or commands.
5. The custom grid meets the ticket 11 scroll and memory budgets on the benchmark corpus.
6. Fill handle, context menu, clipboard, zoom, text layout, IME and autofit are scheduled in the grid milestone.
7. The spike remains throwaway: it is not merged into the product tree, and its code is not the production grid.
8. The measurements in §3 are reproducible from the branch and are superseded by the ticket 11 harness.

## 9. Evidence

- Custom spike: `prototype/grid-spike/custom/index.html` (696 lines).
- Univer spike: `prototype/grid-spike/univer/` (Vite app, 75-line `main.js`, 7-line `worker.js`).
- Screenshots: `screens/custom.png`, `screens/univer.png`, `screens/univer-scrolled.png`, `screens/univer-edit.png`.
- Accessibility snapshots: `screens/custom-a11y.yml` (403 grid cells, 33 rows), `screens/univer-a11y.yml` and `screens/univer-a11y-selected.yml` (0 grid cells).
- Bundle: `vite build` output; gzip measured per file at level 9.
- Startup: `requestAnimationFrame` poll until the largest canvas has more than 300 dark pixels in its top-left 320 × 220 region.
- Scroll: 2.2 s of synthetic wheel or scroll-offset changes, p95 of `requestAnimationFrame` deltas under 1 s.
