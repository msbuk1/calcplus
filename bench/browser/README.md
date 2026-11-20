# Browser performance harness

Stand-up for performance contract §6, layer 2: a Playwright + Chromium
DevTools Protocol measurement that records the frame timings, the page heap,
the DOM count, the canvas size and the machine fingerprint.

```bash
npm run bench:browser
# optional: CALCPLUS_BENCH_OUTPUT=path/to/result.json
```

## What runs

- `memory.bench.ts` — the M1 stand-up: it proves the harness can read the CDP
  heap metrics and that the tiny fixture stays under the 500 MB guaranteed
  envelope. There is no formula engine worker yet, so the result records
  `engineWorker: null`.
- `scroll.bench.ts` — the M2 grid budgets:
  - the scroll frame budget on a 250,000 touched-cell formatted sheet
    (performance §3.3): p95 ≤ 16.7 ms and no frame over 50 ms;
  - a 1,000,000-row sparse sheet scrolled to its last row, which proves the
    grid never allocates the sheet itself.

The shapes come from the committed corpus generator (`bench/corpus/generate.ts`,
fixed seed) and are built in the page through the benchmark hook, so no large
workbook is committed to the repository. The same generator now carries the ten
contract shapes of performance §5; the formula-heavy shapes are measured by the
Node harness in `bench/model` (`npm run bench:model`), because the browser
benchmark hook builds literal values only.

## The reference machine

The absolute budgets are a contract with the reference machine of performance
§1: a mid-range 2023 laptop with a four-core CPU and 16 GB of RAM, running
Windows 11 or Ubuntu 22.04 or newer, with a recent Chromium, at 1440 × 900 CSS
pixels and one workbook per tab.

A run only claims the budget when **both** conditions hold:

1. the maintainer declares the run with `CALCPLUS_REFERENCE_MACHINE=1`, and
2. the hardware checks pass — Windows 11 (kernel build 22000 or newer) or
   Ubuntu 22.04 or newer, at least four **physical** CPU cores (hyper-threaded
   logical counts do not count), at least 16 GB of RAM, a 1440 × 900 viewport
   and a device pixel ratio of 1.

Every other run is advisory: the harness still measures and records the
frames, but it does not fail on the reference-only absolute budget. The JSON
result carries the fingerprint and the list of reasons the machine did not
qualify, so a report can say exactly what is unproven.
