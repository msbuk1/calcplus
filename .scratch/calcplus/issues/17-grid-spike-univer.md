# 17 — Stress-test the grid decision: custom vs Univer

Status: open
Type: prototype
Blocked by: none

## Question

The grid research (05) recommended a custom hybrid grid without evaluating Univer. GenOffice ships Univer 0.25 (Apache-2.0) in production and wraps it in a large compatibility patch layer. Which path gets calcplus to Excel fidelity faster?

Build a rough spike of both:

- **Custom**: a canvas grid rendering a large sparse sheet with selection, in-cell editing, and a frozen pane.
- **Univer**: the stock sheets preset with Excel-like chrome and a sample `.xlsx`.

Compare: time to a working grid, Excel behaviour fidelity, bundle size and startup, accessibility, licence obligations, and the patch surface calcplus would own. Because Univer bundles its own formula engine, this spike also tests whether adopting Univer should replace the HyperFormula recommendation (03) — treat grid and formula engine as one coupled build-vs-buy decision, not two. Decide whether to revise [Grid rendering approach](05-grid-rendering.md) before the workbook model (07) and UI chrome (10) are locked.
