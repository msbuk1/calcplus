# HyperFormula behind a narrow formula-engine port

calcplus calculates with the GPL-3.0-only edition of HyperFormula, behind a narrow internal `FormulaEngine` port, in a Web Worker. The registry is gated to the 52 published functions, and a documented set of shims fixes the engine's deviations.

**Considered options:** HyperFormula (chosen); IronCalc 0.9.1 (rejected: no 1904 dates, the fake 29 February 1900 is wrong, and its `.xlsx` layer rewrites the package); Univer's engine (couples the grid, model and engine); a custom engine (too large for v1); Formula.js and xlsx-calc (no dependency graph, wrong semantics).

**Why:** HyperFormula is the only mature standalone TypeScript engine with a real dependency graph, incremental recalculation, worker capability and custom functions. The 52-function contract is small enough to gate and shim. The model, not the engine, owns the authored formula text and cached results.

**Consequences:** `INT`, `CHOOSE`, wildcard lookup and `#CYCLE!` → `#CIRC!` need shims; `OFFSET`, `_xlfn.` and structured references are preserved as text plus cached value; the engine is rebuilt on open and after structural edits; fan-out is a benchmark case. See [formula semantics](../../.scratch/calcplus/spec/formula-semantics.md) and [ticket 24](../../.scratch/calcplus/issues/24-ironcalc-vs-hyperformula.md).
