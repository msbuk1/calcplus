# 18 — Number-format engine: build vs buy

Status: open
Type: research
Blocked by: none

## Question

How should calcplus render Excel number formats, including the full custom grammar?

Evaluate `numfmt`, SheetJS's `ssf`, `Intl.NumberFormat` plus a custom Excel-format parser, and any other credible option. Compare on:

- Coverage of the Excel format grammar: sections, conditions, colours, currency, accounting, fractions, scientific, elapsed time, text, literals, padding, and locale decimal/thousands separators.
- Fidelity against Excel over a conformance corpus.
- Round-trip: preserving the authored format code and its `numFmtId`, including custom codes and locale-specific built-ins.
- Behaviour on the 1900/1904 date system and negative/zero/blank sections.
- Browser fit, bundle size, and performance inside a canvas renderer.
- Licensing and maintenance.

Recommend an approach and list the constructs that need custom work.
