# Build the custom canvas grid; do not adopt Univer

calcplus renders the sheet with a viewport-sized Canvas 2D surface split into up to four panes, a DOM editing overlay and a DOM accessibility mirror. It does not adopt a third-party grid or spreadsheet SDK as the foundation.

**Considered options:** a custom canvas grid (chosen); a virtualized DOM grid (kept as a fallback behind the same `RenderFrame` seam); Univer (rejected for v1); AG Grid, Handsontable, Glide, canvas-datagrid and SpreadJS (rejected on licence, fit or missing Excel interactions).

**Why:** the destination is a faithful alternative to Excel desktop, not a generic spreadsheet. Univer ships its own chrome, model and formula engine, has no open-source `.xlsx` exchange, and its canvas exposed zero grid cells to assistive technology in the spike. The custom spike met the frame budget, exposed 403 grid cells through its mirror, and leaves the model, engine and file layer ours. GenOffice's 256-module patch layer shows that buying the grid does not buy the fidelity tail.

**Consequences:** a DOM-cell renderer stays behind `RenderFrame` as the performance fallback, and becomes the accessibility fallback when that work returns in v2; fill, context menu, clipboard, zoom, text layout, IME and autofit are owned work. See [grid spike](../../.scratch/calcplus/spec/grid-spike.md) and [ticket 17](../../.scratch/calcplus/issues/17-grid-spike-univer.md).
