# 03 — Formula engine: build vs buy

Status: claimed
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
