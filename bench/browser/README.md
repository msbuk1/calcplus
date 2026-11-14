# Browser memory harness (M1)

Stand-up for performance contract §6, layer 2: a Playwright + Chromium
DevTools Protocol measurement that records the page heap, DOM count, canvas
size and machine fingerprint for an opened workbook.

```bash
npm run bench:browser
# optional: CALCPLUS_BENCH_OUTPUT=path/to/result.json
```

M1 scope only:

- Measures the page target; there is no formula engine worker yet, so the
  result records `engineWorker: null`.
- Uses the tiny `m1-chart-unknown.xlsx` fixture and asserts only that the
  harness can read CDP heap metrics and that it is below the 500 MB guaranteed
  envelope.
- Does not generate the 250k/50k formula corpus or enforce the full absolute
  budgets; those arrive with the formula engine and the nightly fixed-hardware
  runner in later milestones.

The JSON attachment contains the exact machine fingerprint required by the
performance contract (CPU model, core count, RAM and browser version). Results
from a non-reference machine are advisory.
