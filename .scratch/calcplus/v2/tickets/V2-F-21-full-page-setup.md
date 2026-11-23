# V2-F-21 — Full page setup

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 3–5 pw (provisional)
- **Prerequisites:** none
- **Blocks:** V2-F-20
- **Contract:** [print and PDF](../../spec/print-pdf.md); known deviation D-3

## Scope

Complete the page setup the v1 cut deferred: headers and footers (including
the field codes), print titles (rows and columns), manual page breaks, page
order (down then over, over then down), and independent X and Y scale. Keep
the v1 print area, orientation, fit-to-pages, uniform scale and margins.

## Acceptance

- [ ] Headers and footers render, print and export to PDF.
- [ ] Print titles repeat on every page.
- [ ] Manual breaks and page order change the pagination as Excel does.
- [ ] Independent X and Y scale is honoured while fit-to-pages still works.
- [ ] Every setting round-trips and opens correctly in Excel.
- [ ] The D-3 deviation is removed and its pins updated.

## Notes

This is the largest print ticket. The v1 pagination planner and PDF pipeline
are reused; the work is the extra settings and their interaction.
