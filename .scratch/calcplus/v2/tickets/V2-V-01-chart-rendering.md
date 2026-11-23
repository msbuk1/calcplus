# V2-V-01 — Chart rendering

- **Stream:** Visual
- **Status:** proposed
- **Effort:** 6–12 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §7; known
  deviation D-4

## Scope

Render preserved charts in the grid: column, bar, line, area, pie, doughnut and
scatter charts, with axes, gridlines, titles, legends, data labels, series
formats and theme colours. Honour anchors, scroll, zoom, print and PDF export.
This ticket does not add chart authoring.

## Acceptance

- [ ] The chart fixture corpus renders with the right geometry, colours and
      labels against Excel captures.
- [ ] Anchors, resize and z-order follow the drawing anchor.
- [ ] Scrolling and zooming keep charts aligned with their cells.
- [ ] Print and PDF include the charts as the preview shows them.
- [ ] The saved file keeps the chart parts byte-identical when untouched.
- [ ] The D-4 wording is narrowed to authoring only.

## Notes

The biggest single visual item. Start with the chart types the corpus already
contains and add types only with fixtures. Chart authoring stays out of the
roadmap until this ships.
