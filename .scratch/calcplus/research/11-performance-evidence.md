# Ticket 11 — performance evidence

Measurements taken 2026-11-13 to ground the targets in [performance.md](../spec/performance.md).

## Method

- **Environment:** Node 24.15.0 on Linux, one fresh process per run.
- **Package:** `hyperformula@3.5.0`, `licenseKey: "gpl-v3"`.
- **Memory:** `process.memoryUsage()` after `global.gc()` and a 200 ms settle. Node reports process-wide RSS, so the figure includes a 66-116 MB runtime baseline; the `hfRssMB` column subtracts it.
- **Caveat:** Node is not a browser. These are order-of-magnitude figures; the Playwright/DevTools harness in the spec produces the contractual numbers.

## Shapes

`bench.mjs` builds `rows × cols`, with a percentage of formula cells using `=SUM(A<row>:B<row>)`.
`hot.mjs` builds one value column and one formula column where every formula reads `$A$1`.
`steady.mjs` forces garbage collection before and after a build plus a full recalculation.
`columns.mjs` tests whole-column and whole-row references at the 1,048,576-row limit.

## Raw results

### Mixed sheet, row-local formulas

| Cells | Formulas | Build | Incremental recalc | RSS at build | RSS after full recalc |
|---:|---:|---:|---:|---:|---:|
| 100,000 | 0 | 195 ms | 2 ms | 160 MB | 217 MB |
| 100,000 | 20,000 | 358 ms | 2 ms | 175 MB | 267 MB |
| 250,000 | 50,000 | 746 ms | 2 ms | 293 MB | 385 MB |
| 500,000 | 100,000 | 1,301 ms | 2 ms | 407 MB | 634 MB |
| 1,000,000 | 200,000 | 2,540 ms | 2 ms | 491 MB | 969 MB |

### Steady state, garbage collection forced

| Cells | Formulas | HyperFormula RSS | JS heap |
|---:|---:|---:|---:|
| 250,000 | 50,000 | 250 MB | 70 MB |
| 500,000 | 100,000 | 385 MB | 128 MB |
| 1,000,000 | 200,000 | 621 MB | 245 MB |

### Fan-out: every formula reads one cell

| Formulas | Build | Incremental recalc |
|---:|---:|---:|
| 5,000 | 149 ms | 17 ms |
| 10,000 | 216 ms | 26 ms |
| 20,000 | 445 ms | 51 ms |
| 40,000 | 1,081 ms | 99 ms |
| 80,000 | 3,392 ms | 263 ms |
| 100,000 | 4,821 ms | 260 ms |
| 250,000 | 27,166 ms | 522 ms |

Build time grows roughly quadratically with fan-out; recalculation stays linear at about 2.6 µs per formula.

### Whole-column and whole-row references

| Formula | `maxRows` | Build | HyperFormula RSS | Result |
|---|---:|---:|---:|---|
| `=SUM(A:A)` | 1,048,576 | 45 ms | 5.7 MB | 5050 |
| 50 × `=SUM(A:A)` | 1,048,576 | 54 ms | 8.8 MB | 5050 |
| `=SUM(1:1)` in a spare row | 1,048,576 | 43 ms | 5.7 MB | — |

Whole-column references are lazy: they do not allocate the full grid.

## Consequence for the spec

- 250k cells / 50k formulas fits inside 500 MB once the model and shell are added.
- 1M cells / 200k formulas lands just under 1 GB, which is why that tier is degraded and the refusal point is 1.2 GB.
- Fan-out, not recalculation, is the performance cliff, so it is a named corpus case.
