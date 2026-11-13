# 03 — Formula engine: build vs buy

Status: resolved
Type: research
Blocked by: none

## Question

Should calcplus build its formula engine or adopt one?

Evaluate at least HyperFormula, formulas.js, xlsx-calc, and a custom TypeScript engine; add any credible commercial option. Compare on:

- Excel semantic fidelity and function coverage
- Dependency graph and incremental recalculation
- Dynamic arrays and spill semantics
- Handling of names, references, and workbook-scoped state
- TypeScript/browser fit, bundle size, Web Worker compatibility
- Performance on large workbooks
- Licensing (copyleft vs permissive vs commercial)
- Maintenance risk and extensibility

Recommend an approach and list exactly which gaps would need custom work under that choice.

## Answer

Adopt **HyperFormula 3.5.x under a commercial licence** for v1, behind a narrow internal `FormulaEngine` port; do not build a custom engine. It is the only mature standalone TypeScript engine with a real dependency graph, incremental recalc, worker-capable runtime, custom functions and a 181 KB gzip bundle. Gate the choice on a written quote: GPLv3 is unusable for a closed product, and 3.5.0 can licence only part of the library.

Expect custom work around HyperFormula's gaps: ~130 missing Excel functions (notably `AVERAGEIFS`, `LOOKUP`, `INDIRECT`, `RANK`, `XMATCH`, `SORTBY`, `TEXTSPLIT`), documented semantic deviations (`INT(-8.9) = -8`), and engine-level limits (no 3D refs, tables, `A1#`, range spill, relative names or iteration). A parity plugin can fix many functions and deviations; parser/reference gaps need deferral or a fork.

Critical architectural rule: the workbook model must own **authored formula text and cached values**. HyperFormula rewrites `OFFSET` and cannot parse `_xlfn.` or `@`, so relying on `getCellFormula()` would break high-fidelity `.xlsx` round-trip.

If the quote is unacceptable, or ticket 02 requires full dynamic arrays/3D refs/`INDIRECT`/iteration in v1, use **Univer** (Apache-2.0, 528 functions, 1.75 MB gzip) as a platform bet — not SpreadJS and not a from-scratch engine. SpreadJS ($999/dev + $12,499/yr/domain) replaces the whole product.

Full research: [03-formula-engine.md](../research/03-formula-engine.md)

## Amendment (2026-11-13)

Ticket 02 fixed two things that change this decision:

- **calcplus is GPLv3.** Use HyperFormula's free GPLv3 edition, not the commercial licence. The written-quote gate is removed.
- **The published contract is 52 functions**, listed in [formula-semantics.md](../spec/formula-semantics.md). The engine's extra functions are not exposed or promised; the registry is gated to the 52.

The architectural rule stands: the workbook model owns authored formula text and cached values, and HyperFormula must never be the source of the saved formula text. Any supported construct the GPL edition cannot parse is a blocker for the engine choice, not a reason to drop the construct silently.
