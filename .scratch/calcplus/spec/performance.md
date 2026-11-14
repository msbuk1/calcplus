# calcplus v1 — performance targets and benchmarks

Resolved by [11 — Performance targets and benchmarks](../issues/11-performance.md) on 2026-11-13. Every number here is a contract with a test behind it. Vocabulary lives in [CONTEXT.md](../../../CONTEXT.md).

## 1. Reference machine and the browser cut

Performance numbers are meaningless without a named machine. The contract applies only to:

- **Hardware:** a mid-range 2023 laptop — 4-core CPU (Intel i5 class or Apple M1), 16 GB RAM, integrated graphics, AC power.
- **Browser:** desktop Chromium, Chrome or Edge, the latest two stable versions.
- **OS:** Windows 11, macOS 13 or newer, Ubuntu 22.04 or newer.
- **Window:** 1440 × 900 CSS pixels at 100% zoom, one workbook per tab.

A second **floor tier** — 2-core CPU, 8 GB RAM, older Chromium — must stay usable but may be up to 2× slower. It is not the contract.

**Explicit cut:** Firefox and Safari open and edit workbooks using the download-based file fallback, but are **not performance-tested and carry no performance claim**. Mobile, tablet, touch and low-power devices are out of scope. A desktop wrapper is the only way to widen this later.

## 2. Size envelopes

A **touched cell** is a cell with a value, formula, style, comment, hyperlink or validation rule. Empty cells cost nothing: the grid is always Excel's 1,048,576 × 16,384, but memory is proportional to touched cells, not to the grid or its bounding box.

| Envelope | Touched cells | Formulas | What is promised |
|---|---:|---:|---|
| **Guaranteed** | 250,000 | 50,000 | Every budget in §3 holds. No banner, no degradation. |
| **Degraded** | 1,000,000 | 200,000 | The file opens within 15 s, stays under 1 GB, and remains editable. A visible "large workbook" banner explains the reduced extras. |
| **Refused** | > 1,000,000 | > 200,000 | A clear message; the file is not modified and not partially imported. |

Whole-column and whole-row references do **not** count as touched cells and must not allocate the full grid. Measured: `=SUM(A:A)` at the 1,048,576-row limit builds in 45 ms using 5.7 MB.

## 3. Budgets

All figures are measured on the reference machine. "p95" means 95% of runs or frames are at or under the number.

### 3.1 Opening

| Case | Budget |
|---|---|
| Shell first paint, warm cache | ≤ 1.0 s |
| Shell first paint, cold cache | ≤ 2.0 s |
| 10k touched cells to interactive | ≤ 0.8 s |
| 100k touched cells to interactive | ≤ 2.5 s |
| 250k touched cells to interactive (guaranteed) | ≤ 5.0 s |
| 1M touched cells to interactive (degraded) | ≤ 15.0 s |

**Interactive** means the grid shows the data, the selection responds to a click or arrow key, and the status bar is populated. A progress indicator must appear within 300 ms of a slow open.

### 3.2 Calculation

The engine runs in a dedicated Web Worker. No recalculation may block the main thread for more than 50 ms.

| Case | Guaranteed | Degraded |
|---|---|---|
| Typical edit, up to 1,000 dependent formulas (p95) | ≤ 50 ms | ≤ 100 ms |
| Heavy fan-out, 50,000 formulas reading one cell (incremental recalc after edit) | ≤ 400 ms | — |
| Heavy fan-out, 100,000 formulas reading one cell (incremental recalc after edit) | — | ≤ 800 ms |
| Full recalculation, 250k cells | ≤ 1.5 s | — |
| Full recalculation, 1M cells | — | ≤ 5.0 s |
| Structural edit (insert/delete row or column), graph rebuild, 250k cells | ≤ 1.5 s | — |
| Structural edit (insert/delete row or column), graph rebuild, 1M cells | — | ≤ 5.0 s |

A structural edit rebuilds the derived dependency graph, so its budget is the rebuild, not the delta. The grid repaints immediately and the 100 ms progress cue covers the convergence.

A recalculation longer than 100 ms shows a non-blocking progress cue. Typing and scrolling continue while it runs.

Budgets apply to incremental recalculation after an edit. Cold build on open is measured, not budgeted: 100k single-cell fan-out builds in about 4.8 s and 250k in about 27 s (§5 evidence), which is why the degraded envelope stops at 200,000 formulas.

**Open does not wait for the engine.** The grid paints from stored results, so "interactive" (§3.1) is reached before recalculation finishes. A cold build over 100 ms shows the non-blocking progress cue, and values converge while the user can already select, scroll and edit. Time to *all formulas current* is measured and reported, not gated, because the fan-out build is the reason. The only case that must wait is a shape with no stored results; it shows the progress cue for its whole build.

### 3.3 Scrolling and direct interaction

| Case | Budget |
|---|---|
| Scroll or drag, p95 frame | ≤ 16.7 ms (60 fps) |
| Scroll or drag, worst single frame | ≤ 50 ms |
| Click or arrow key to painted selection | ≤ 50 ms |
| Character typed in a cell to painted | ≤ 50 ms |
| Edit commit to visible result (excluding recalc beyond 100 ms) | ≤ 100 ms |
| Context menu or dialog open | ≤ 100 ms |

Frame timing is measured while scrolling a 250k-cell formatted sheet at 100% zoom.

### 3.4 Bulk operations

| Case | Budget |
|---|---|
| Paste or fill 100k cells | ≤ 1.0 s |
| Paste or fill 1M cells (degraded) | ≤ 3.0 s |
| Find across 250k cells | ≤ 0.5 s |
| Find across 1M cells (degraded) | ≤ 2.0 s |
| Sort 100k rows | ≤ 2.0 s |
| Filter apply, change or clear, 100k rows | ≤ 1.0 s |
| Undo or redo any bulk action | ≤ 1.0 s |

Each bulk action is one undo transaction (see [command system §4](command-system.md)).

### 3.5 Formatting and conditional formatting

| Case | Budget |
|---|---|
| Number formatting of the visible cells | inside the 16.7 ms frame budget |
| Conditional-formatting evaluation, visible cells | ≤ 4 ms per frame |
| Conditional-formatting range statistics, 100k cells | ≤ 200 ms, computed in row-block chunks |

### 3.6 Print and PDF

| Case | Budget |
|---|---|
| Print preview, 20-page A4 report | ≤ 2.0 s |
| PDF export, 20-page A4 report | ≤ 3.0 s |

## 4. Memory

**Working set** means the tab's JavaScript heap, typed arrays, the engine worker's heap, and grid canvases. It excludes the browser's shared baseline.

| Threshold | Behaviour |
|---|---|
| ≤ 500 MB | Guaranteed envelope must stay here. |
| ≤ 1 GB | Degraded envelope must stay here. |
| 1 GB | Visible warning; further large operations ask for confirmation. |
| 1.2 GB | Refuse new large operations; prompt to close the workbook. |
| Estimated > 1.2 GB at open | Refuse to open with a clear message. |

Measured HyperFormula steady state, garbage collection forced: 250k cells / 50k formulas ≈ 250 MB; 500k / 100k ≈ 385 MB; 1M / 200k ≈ 620 MB. The remaining budget covers the canonical model, style table and shell.

## 5. Benchmark corpus

A committed generator (`bench/corpus/generate.*`, fixed seed) produces ten synthetic shapes. Three anonymised real workbooks ship as small fixtures; anything larger is generated on demand and never committed.

| Name | Shape | What it catches |
|---|---|---|
| `tiny` | 1 sheet, 1k touched, no formulas | startup and smoke |
| `typical` | 3 sheets, 50k touched, 10k formulas, 200 styles, merges, comments, hyperlinks | everyday workbooks |
| `dense-formula` | 1 sheet, 100k touched, 50k row-local formulas | recalc throughput |
| `fan-out` | 1 sheet, 100k formulas all reading one input cell | dependency-graph build cliff |
| `wide` | 2k rows × 50 columns, 100k touched | horizontal scrolling and column widths |
| `many-sheets` | 20 sheets × 5k touched | sheet switching and per-sheet overhead |
| `styling-heavy` | 100k touched, 500 distinct styles, borders and fills | style-table lookup and paint |
| `cf-heavy` | 100k touched, 50 conditional-formatting rules | render-only rule evaluation |
| `degraded-1m` | 1M touched, 200k mixed formulas | degraded envelope |
| `pathological` | 200k formulas reading one cell, plus whole-column refs — the envelope ceiling; the 250k probe above the refusal line measured 27 s build and 522 ms recalc | must finish, must not hang |

Real workbooks: a budget/forecast model, a sales tracker and a gradebook, anonymised and cleared for test use. File sizes of 1 MB, 10 MB and 50 MB are covered by generated variants. The `.xlsx` fixtures from ticket 04 are reused for open-time coverage.

## 6. Measurement harness

Three layers:

1. **Node microbenchmarks** (`bench/engine`, `bench/model`, `bench/io`) — deterministic, fixed seed, 5 warm-up runs then 10 measured runs, reporting p50/p95/max plus memory after forced garbage collection (`--expose-gc`). Runs on every pull request.
2. **Browser end-to-end** (`bench/browser`, Playwright + Chromium DevTools Protocol) — cold open, scripted scroll, click/key-to-paint, paste/fill/find/sort and memory. Uses `performance.mark`/`measure`, `PerformanceObserver` long tasks, `requestAnimationFrame` deltas, and CDP heap metrics for both the page and the worker target. Runs nightly on a fixed hardware class.
3. **In-app performance panel** (`?perf=1`) — last-command timings, 5-second frame p95, working-set estimate and engine queue depth. For manual checks and bug reports.

Every result records a machine fingerprint (CPU model, core count, RAM, browser version). Results from a different machine are advisory, not gating.

## 7. CI regression gate

| Stage | Gate |
|---|---|
| Pull request | Node microbenchmarks. Block on a **20% regression** against the rolling median of the last 10 runs on the same runner. One automatic re-run filters noise; a second failure blocks. |
| Nightly | Browser end-to-end on the fixed runner. Enforces the absolute budgets in §3 and §4. A failure opens a tracked issue. |
| Release | Full corpus, all sizes. Absolute budgets block the release. |

A `perf-regression` label can override a pull-request block with a written justification. Overrides are listed in the release notes.

## 8. Degradation ladder

When a workbook exceeds the guaranteed envelope, calcplus degrades visibly and in this order:

1. A "large workbook" banner explains what is reduced.
2. Conditional-formatting painting stops outside the visible range.
3. Scroll overscan is reduced.
4. Volatile recalculation moves to idle time.
5. At 1 GB, a memory warning appears.
6. At 1.2 GB, new large operations are refused; the user is prompted to close the workbook.

Rendering degrades before correctness does. Values are never silently dropped, and the file on disk is never modified by degradation.

## 9. Acceptance criteria

1. Every budget in §3 and §4 has an automated test in the corpus.
2. The guaranteed envelope holds on the reference machine.
3. The degraded envelope opens within 15 s and stays under 1 GB.
4. A workbook above 1M touched cells is refused with a clear message; the file is untouched.
5. Working set never exceeds 1.2 GB; a warning appears at 1 GB.
6. Scrolling a 250k-cell formatted sheet holds 60 fps at p95, with no frame over 50 ms.
7. No recalculation blocks the main thread for more than 50 ms.
8. Whole-column and whole-row references do not allocate the full grid.
9. A 100k-cell paste is one undo step and undoes in under 1 s.
10. Pull requests block on a 20% regression; nightly and release builds enforce the absolute budgets.
11. The corpus is reproducible from a committed generator; real fixtures are anonymised and licensed for test use.
12. Every result carries a machine fingerprint; non-reference runs never gate.
13. Firefox, Safari, mobile and touch are documented as outside the performance contract.
14. The filter, conditional-formatting and print budgets in §3.4-§3.6 are enforced by the nightly browser suite.

## 10. Evidence

Measured on 2026-11-13 with HyperFormula 3.5.0, Node 24.15.0, one fresh process per run. Node memory is process-wide, so treat these as order-of-magnitude evidence; the browser harness produces the contractual numbers.

| Shape | Build | Incremental recalc | Steady memory |
|---|---|---:|---:|
| 250k cells / 50k formulas | 746 ms | 2 ms | ~250 MB |
| 500k cells / 100k formulas | 1.3 s | 2 ms | ~385 MB |
| 1M cells / 200k formulas | 2.5 s | 2 ms | ~620 MB |
| 100k formulas reading one cell | 4.8 s | 260 ms | ~279 MB |
| 250k formulas reading one cell | 27 s | 522 ms | ~352 MB |
| `=SUM(A:A)` at 1,048,576 rows | 45 ms | — | ~6 MB |

The fan-out rows are the reason `fan-out` is a named corpus case: build time grows roughly quadratically with the number of formulas sharing one input cell. Recalculation itself is linear at about 2.6 µs per formula. If the fan-out case breaches its budget, the fix is engine-side batching or a Web Worker queue — never a switch to IronCalc (ticket 24).
