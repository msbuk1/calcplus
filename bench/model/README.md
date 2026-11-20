# Node performance harness

The first layer of the measurement harness in
[performance §6](../../.scratch/calcplus/spec/performance.md): deterministic
Node microbenchmarks over the whole corpus, reporting build time, incremental
recalculation, steady heap and the machine fingerprint.

```bash
npm run bench:model
# optional: only some shapes
CALCPLUS_BENCH_SHAPES=tiny,fan-out npm run bench:model
```

Results are written under `bench/results/model/` (gitignored). The run prints a
table and records `corpus-definition.json`, one `latest-<shape>.json` per case,
and a `history-<machine>.json` rolling record.

## The corpus

`bench/corpus/generate.ts` holds the ten shapes of performance §5 with a fixed
seed, plus the two M2 scroll shapes:

| Shape | What it exercises |
|---|---|
| `tiny` | startup and smoke |
| `typical` | 3 sheets, 50k touched, 10k formulas, 200 styles, merges, comments, hyperlinks |
| `dense-formula` | 100k touched, 50k row-local formulas |
| `fan-out` | 100k formulas reading one cell |
| `wide` | 2k rows × 50 columns, 100k touched |
| `many-sheets` | 20 sheets × 5k touched |
| `styling-heavy` | 100k touched, 500 styles |
| `cf-heavy` | 100k touched, 50 conditional-formatting rules |
| `degraded-1m` | 1M touched, 200k mixed formulas |
| `pathological` | 200k formulas reading one cell, plus whole-column references |

`bench/model/corpus.ts` turns a shape into engine inputs. Shapes whose extra
content is model-only (styles, comments, hyperlinks, merges, conditional
formatting) say so in the result rather than pretending the engine measures it;
those extras are covered by the browser harness and the M5 fixtures.

## What is measured

- **Build:** `start`, `load` and `finishLoad` on a fresh engine. One warm-up run
  is discarded; three runs are recorded.
- **Incremental:** one edit of `A1`, repeated three times with alternating
  values. The fan-out and whole-column shapes make this the heavy case.
- **Memory:** the JS heap plus typed arrays (`process.memoryUsage`), and the
  resident set as a coarse upper bound, recorded after a forced GC (the runner
  is started with `--expose-gc`).

## Budgets and the reference machine

The absolute budgets are a contract with the reference machine of performance
§1. A run enforces them only when the maintainer declares the machine:

```bash
CALCPLUS_REFERENCE_MACHINE=1 npm run bench:model
```

The declaration also has to pass the hardware check (four physical cores, 16 GB
of RAM, Windows 11 or Ubuntu 22.04 or newer). Every other run records the
fingerprint and the measured machine speed and treats the number as advisory:
performance §6 says results from a different machine are not gating. A liveness
ceiling of ten times the budget still catches an engine that stopped working.
On the reference machine a miss is re-run once before it fails, as performance
§7 allows.

## The per-change regression gate

For changes on any machine, performance §7 blocks on a 20% regression against
the rolling median of the last ten local runs on the same machine. The runner
keeps that history keyed by machine fingerprint. A run whose calibration shows
the machine was competing for the CPU (median sample more than 1.5× the best)
is reported but is not recorded and is not compared, so a busy laptop cannot
poison the baseline or block an innocent change. A case whose rolling median is
below 50 ms is reported but never blocks: a 20% change there is measurement
noise, not a regression.
