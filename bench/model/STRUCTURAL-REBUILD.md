# Structural-edit graph rebuild probe

Gate measurement for performance §3.2: a structural edit rebuilds the derived
dependency graph, and the guaranteed envelope allows at most 1.5 seconds for
250,000 touched cells.

The focused probe is `src/testing/structural-rebuild.ts`. It builds the same
shape a structural edit rebuilds — 5,000 rows × 50 columns = 250,000 touched
cells, with 50,000 row-local formulas — and times the same path
`CalculationController.rebuild()` uses:

1. `HyperFormulaEngine.start` with the workbook configuration.
2. `HyperFormulaEngine.load` with the workbook snapshot.
3. `HyperFormulaEngine.finishLoad` to complete the graph build.

Shape generation is outside the timed section. One warm-up run is discarded;
three measured runs are recorded. The contract budget is 1.5 seconds.

Run only the focused probe:

```bash
npm run test -- tests/m3c-structural-rebuild.test.ts
```

The test asserts the median run against the budget. The budget gates **only on
the reference machine** (performance §1, §6): the run must be declared with
`CALCPLUS_REFERENCE_MACHINE=1` and pass the hardware check, and a miss is
re-run once (performance §7). On any other machine the result is advisory and
1.5 s becomes a liveness ceiling, so a slow or busy laptop reports its real
number without failing the suite. The result object also returns every run, the
best run, the machine fingerprint, the machine-speed sample and the gate, so
the release gate can record the distribution rather than just pass or fail. The
App range-transform prototype stays read-only and does not claim the budget; the
focused probe is the graph-rebuild measurement. This probe is not a substitute
for the full local browser run in performance §7; it is the deterministic
Node-side gate for the M3C structural-edit budget.
