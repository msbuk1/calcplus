# 18 — Number-format engine: build vs buy

Status: resolved
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

## Answer

**Buy the grammar, build the Excel seams.** Adopt `numfmt` 4.1.3 (MIT) as the value-to-string engine behind an internal interface; do not write the parser from scratch. Reject `ssf` as the production engine (strong date/time, weak numeric/conditional/colour/locale behaviour), `Intl.NumberFormat` alone (cannot express the grammar), and the Univer/numbro wrappers (coupling or wrong grammar). Keep `ssf` only as a differential test oracle.

Custom work is still required for: the full built-in `numFmtId` table (0–22, 37–49, and locale-specific 27–36, 50–58, 59–62, 67–81); the 1904 adapter (add 1462 for calendar formats, not `[h]`/`[m]`/`[s]`); cell-width layout (`*`, `_`, `?`, `General`, `#####`); colour/palette mapping; normalising lowercase `e+`/`a/p` and `[NatNum]`/`[DBNum]`; Indian 3/2 grouping; invalid-format fallback; and round-trip ownership in the style table (preserve code/ID, append-only, never materialise built-ins).

Trade-offs: `numfmt` has no public 1904 support and resolves built-in IDs above 58 incorrectly, so both must be wrapped; lowercase `e+` and `a/p` have fidelity bugs in 4.1.3; it is ~46.7 KB gzipped and actively maintained by one person, so pin the version and be ready to vendor/fork (MIT). These are bounded seam costs, not reasons to rebuild.

Full research: [18-number-format-engine.md](../research/18-number-format-engine.md)
