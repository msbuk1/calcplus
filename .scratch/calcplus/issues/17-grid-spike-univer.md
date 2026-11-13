# 17 — Stress-test the grid decision: custom vs Univer

Status: resolved
Type: prototype
Blocked by: none

## Question

The grid research (05) recommended a custom hybrid grid without evaluating Univer. GenOffice ships Univer 0.25 (Apache-2.0) in production and wraps it in a large compatibility patch layer. Which path gets calcplus to Excel fidelity faster?

Build a rough spike of both:

- **Custom**: a canvas grid rendering a large sparse sheet with selection, in-cell editing, and a frozen pane.
- **Univer**: the stock sheets preset with Excel-like chrome and a sample `.xlsx`.

Compare: time to a working grid, Excel behaviour fidelity, bundle size and startup, accessibility, licence obligations, and the patch surface calcplus would own. Because Univer bundles its own formula engine, this spike also tests whether adopting Univer should replace the HyperFormula recommendation (03) — treat grid and formula engine as one coupled build-vs-buy decision, not two. Decide whether to revise [Grid rendering approach](05-grid-rendering.md) before the workbook model (07) and UI chrome (10) are locked.

## Answer

**Verdict: keep the custom canvas grid; reject Univer as the v1 foundation; keep HyperFormula.** Full evidence and consequences in [grid-spike.md](../spec/grid-spike.md).

Both spikes were built on branch `prototype/grid-spike`:

- **Custom** — 696-line canvas grid with four panes, a frozen row and column, 1M-row addressable scroll, 100k sparse cells, selection, merges, DOM editing, formula bar, keyboard navigation and a 403-cell accessibility mirror.
- **Univer** — 75-line integration of the stock `@univerjs/preset-sheets-core` 1.4.0 with its full chrome, frozen panes, formulas and number formats. The open-source preset has no `.xlsx` exchange package.

**Measured:** custom 7.9 KB gzip, ~125 ms warm to visible, 16.8 ms p95 scroll; Univer 1.78 MB gzip main entry, ~300-400 ms warm (1.3 s cold), 18.9-20.6 ms p95 with a 72 ms worst frame; Univer exposes **0 grid cells** to assistive technology versus 403 in the custom mirror.

**Why custom:** Excel chrome is the product and Univer's chrome must be replaced; the canonical model, command registry and HyperFormula would be replaced or duplicated; `.xlsx` is a commercial Univer feature, so the OOXML layer is still ours; accessibility needs the same mirror inside someone else's renderer; GenOffice's 256-module patch layer shows the buy path still owns the fidelity tail. Univer would be the right buy for a general spreadsheet, not for an Excel clone.

**What changes:** the accessibility mirror becomes a first-milestone gate, a DOM-cell renderer stays behind the `RenderFrame` seam as the fallback, the missing interactions (fill drag, context menu, clipboard, zoom, text layout, IME, autofit) are scheduled explicitly, and the grid stays decoupled from the engine.
