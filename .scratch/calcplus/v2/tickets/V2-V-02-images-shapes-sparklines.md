# V2-V-02 — Images, shapes and sparklines rendering

- **Stream:** Visual
- **Status:** proposed
- **Effort:** 4–7 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §7; known
  deviation D-4

## Scope

Render preserved images, drawing shapes and sparklines: picture anchors and
cropping, shape fills and outlines, text boxes, z-order, and cell-anchored
sparklines with their markers. Honour scroll, zoom, print and PDF. This ticket
does not add insertion or editing.

## Acceptance

- [ ] The image, shape and sparkline fixture corpus renders against Excel
      captures.
- [ ] Anchors, cropping, z-order and transparency follow the file.
- [ ] Scroll and zoom keep the objects aligned with their cells.
- [ ] Print and PDF include the objects as the preview shows them.
- [ ] Untouched drawing parts stay byte-identical on save.
- [ ] The D-4 wording covers the remaining authoring gap.

## Notes

VML and DrawingML shapes are both preserved in the corpus. Render what the
file describes; do not normalise the parts.
