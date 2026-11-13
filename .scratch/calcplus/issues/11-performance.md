# 11 — Performance targets and benchmarks

Status: resolved
Type: grilling
Blocked by: none

## Question

What performance must v1 hit, and how is it measured?

Set measurable targets for:

- Max workbook size and used range the app must stay usable with
- Cold open time for a representative workbook
- Full and incremental recalculation latency
- Scroll/paint frame budget and interaction latency
- Large paste, fill, copy, and find operations
- Memory ceiling in a browser tab
- Desktop-browser-only assumption (state the cut explicitly)

Define the benchmark corpus (synthetic and real-world shapes), the measurement harness, and the CI regression gate.

## Answer

Performance is a contract with a test behind it, not a hope. Full detail in [performance.md](../spec/performance.md).

**Reference machine:** mid-range 2023 laptop (4-core, 16 GB, current Chromium, 1440×900 at 100% zoom). A 2-core/8 GB floor tier must stay usable but may be 2× slower. Only desktop Chromium is covered; Firefox and Safari are best-effort with no performance claim, and mobile/touch is out.

**Size:** guaranteed 250,000 touched cells / 50,000 formulas; degraded up to 1,000,000 / 200,000; refuse above that. Memory 500 MB guaranteed, 1 GB degraded, warn at 1 GB, refuse new large operations at 1.2 GB.

**Budgets:** cold open 0.8 s / 2.5 s / 5 s / 15 s at 10k / 100k / 250k / 1M cells. Typical recalc ≤ 50 ms, 100k-formula fan-out ≤ 400 ms, full recalc 1.5 s / 5 s. 60 fps scrolling with no frame over 50 ms; 50 ms click-to-paint; 100 ms commit. 100k-cell paste/fill ≤ 1 s, find 250k ≤ 0.5 s, sort 100k rows ≤ 2 s.

**Harness:** Node microbenchmarks on every pull request; Playwright plus Chromium DevTools Protocol nightly on a fixed hardware class; an in-app panel behind `?perf=1`.

**Gate:** 20% regression against the rolling median of the last 10 same-runner runs blocks a pull request; absolute budgets are enforced nightly and at release.

**Degradation:** staged and visible — banner, then conditional-formatting painting outside the viewport, scroll overscan, idle recalculation, memory warning, refusal. Rendering degrades before correctness does.

**Measured evidence:** HyperFormula 3.5.0 steady state is ~250 MB at 250k cells, ~385 MB at 500k and ~620 MB at 1M. A single cell with 250k dependents takes 27 s to build the graph but only 522 ms to recalculate — so fan-out is a named corpus case. `=SUM(A:A)` at the full 1,048,576 rows costs 45 ms and 5.7 MB.

## Notes

- Benchmark HyperFormula's browser memory at the target sheet sizes. If memory is the constraint, move the engine to a Web Worker; do not switch to IronCalc (see ticket 24).
