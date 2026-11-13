# 21 — Print and PDF pipeline

Status: open
Type: research
Blocked by: none

## Question

How should calcplus produce Excel-like printed pages and PDFs in a browser?

Evaluate:

- Browser print via a dedicated print stylesheet/DOM, versus rendering to canvas and emitting PDF with a library (`pdf-lib`, `jsPDF`, others). A headless server is not an option: no backend.
- Honouring print area, orientation, fit-to-width, margins, scaling, and pagination across pages.
- Fidelity of fonts, borders, fills, merged cells, row heights and images (images are v2, but layout should not break).
- Headers/footers and print titles as a v2 extension.
- Browser differences (Chrome/Firefox/Safari) and PDF output consistency.
- Memory and performance on large print areas.

Recommend an architecture and list the fidelity limits of the chosen path.
