# 30 — Structural-rebuild budget: is the number wrong or is the code slow?

Status: research complete
Date: 2026-11-21
Question: M7 exit criterion 2 — "the performance budgets hold on the reference machine". The `structural-rebuild-250k` case misses its 1500 ms budget by 4 to 11 per cent on a declared reference machine.
Method: read-only review of the probe, the gate, the engine adapter, the rebuild path, the spec and the committed evidence. No code was changed. No new benchmark runs were made in this session; all numbers below are the eight enforced runs supplied in the brief plus committed measurements. Stage shares inside the timed section are inferred from the code path, not measured — a single instrumented run would confirm them.

## TL;DR

The budget number is the likelier problem, not the code. The 1500 ms figure is about twice a 746 ms Linux measurement from 2026-11-13, not a measurement on reference-class hardware. The rebuild path is already streamlined: nearly all of the ~1600 ms is HyperFormula graph construction (250,000 per-cell inserts plus 50,000 per-cell reads), which is single-threaded and CPU-bound. The miss (50 to 171 ms) is the same size as the run-to-run spread (121 ms across the eight runs), and the best run still misses by 50 ms. There is no obvious 7 per cent sitting in the code. Recommendation: raise the budget to a measured figure (1750 ms), record why, and treat further engine-path optimisation as non-gating follow-up. **The optimisation that followed found far more than 7 per cent and overturned this recommendation; see the correction at the end.**

## 1. Where the 1500 ms actually goes

The probe is `src/testing/structural-rebuild.ts`. The timed section is `runRebuild` (`src/testing/structural-rebuild.ts:113-124`):

1. `engine.start(...)` — build an empty HyperFormula instance, add one sheet, register two boolean named expressions, suspend evaluation (`src/engine/hyperformula-engine.ts:53-97`).
2. `engine.load(cells)` — one `setCellContents` call per cell, 250,000 times (`src/engine/hyperformula-engine.ts:99-102`). Each formula cell is screened by `prepareFormula`, validated by `engine.validateFormula`, then inserted by `engine.setCellContents` (`src/engine/hyperformula-engine.ts:208-237`). Each literal cell goes through `setLiteral`, i.e. one more `engine.setCellContents` (`src/engine/hyperformula-engine.ts:245-255`).
3. `engine.finishLoad()` — resume evaluation, then read back every formula cell one by one: 50,000 `getCellValue` calls plus one result object each (`src/engine/hyperformula-engine.ts:104-124`).

Shape generation is outside the timer (`src/testing/structural-rebuild.ts:73`, `85-94`). Engine disposal is also outside the timer: the return value `now() - started` is evaluated before the `finally` block runs (`src/testing/structural-rebuild.ts:113-124`). The budget therefore covers only engine-internal graph construction plus the first full evaluation. The one warm-up run is discarded and three measured runs are recorded (`src/testing/structural-rebuild.ts:86-94`; `bench/model/STRUCTURAL-REBUILD.md:16-17`).

What dominates is graph construction, not calculation. Committed evidence shows incremental recalculation at this exact shape costs 2 ms (`performance.md:211`; `.scratch/calcplus/research/11-performance-evidence.md:27`). **That 2 ms is incremental recalc, not the cold first evaluation — see the correction below.** Measured later, the cold first evaluation inside `finishLoad` is about 270 ms. Everything else is building the dependency graph: 250,000 individual inserts and 50,000 individual reads, each with per-call overhead on both sides of the adapter. This is inference from the code path, not a stage-timed measurement - the probe reports only the total. A single run that times `start`, `load` and `finishLoad` separately would confirm the split, and that run should happen before any optimisation work starts.

What the probe does **not** measure matters for the decision. The production rebuild does three more things, all excluded from the budget:

- `snapshotWorkbook` walks all 250,000 model cells, parsing one cell key per cell (`src/engine/calculation-controller.ts:474-534`).
- The snapshot crosses into the worker in chunks, one acknowledgement round-trip per chunk (`src/engine/formula-engine-client.ts:102-123`, `216-226`).
- The result batch is applied back in bounded slices with yields between them (`src/engine/calculation-controller.ts:260-323`).

So the probe understates a real structural edit. Any option that makes the case measure the full production path needs a **higher** budget, not a lower one.

## 2. The budget's provenance

The budget is a measured figure with headroom, not a guess — but it was measured on different hardware than the contract names.

- `performance.md:58` sets the structural-edit graph rebuild at 250k cells to ≤ 1.5 s. The same 1.5 s number covers full recalculation at 250k cells (`performance.md:56`), with the rationale that "a structural edit rebuilds the derived dependency graph, so its budget is the rebuild, not the delta" (`performance.md:61`).
- The evidence table (`performance.md:205-216`) measured 250k cells / 50k formulas building in **746 ms**, with 2 ms incremental recalc, on 2026-11-13 with HyperFormula 3.5.0 and Node 24.15.0 on Linux, one fresh process per run. The same 746 ms figure is in `.scratch/calcplus/research/11-performance-evidence.md:27`.
- 1500 ms is almost exactly twice 746 ms. The natural reading: the budget was set at roughly 2× the observed Linux build time as machine-variation headroom.

Two caveats weaken that headroom:

- The evidence note itself warns that "Node is not a browser" and the figures are "order-of-magnitude" (`performance.md:207`). Yet the enforcing gate test runs in Node (`tests/m3c-structural-rebuild.test.ts:26-44`). The contractual number is enforced in the same environment the evidence disclaims.
- Before M7 lane 4 the same case measured 3740 ms against the same 1500 ms budget (`.scratch/calcplus/issues/36-followup-chrome-data-tools.md:37-38`). The budget was aspirational until lane 4 more than halved the time. The remaining gap is 50 to 171 ms, not seconds.

## 3. Is 1500 ms defensible for this case?

This case is the top of the guaranteed envelope: 250,000 touched cells and 50,000 formulas is exactly the "every budget holds" row (`performance.md:24`). Three facts bear on whether 1500 ms is the right number for it.

First, HyperFormula's cost at this shape is graph construction, and row-local formulas are its friendly case. Build time scales with cells inserted; evaluation is linear at about 2.6 µs per formula (`.scratch/calcplus/research/11-performance-evidence.md:51`). The quadratic fan-out cliff documented in the evidence (`.scratch/calcplus/research/11-performance-evidence.md:39-51`) does not apply here: the probe shape is row-local (`A{row+1}*2`, `src/testing/structural-rebuild.ts:126-154`). Nothing in the committed evidence suggests HyperFormula can build this graph much faster on equivalent hardware — 746 ms on the Linux box versus ~1600 ms here is a ~2.1× ratio, which is within the range of single-thread CPU, OS and Node-version differences for a purely single-threaded build.

Second, the reference-machine check does not test what this workload needs. The gate enforces only when the run is declared **and** the hardware matches: 4 physical cores, 16 GB RAM, Windows 11 or Ubuntu 22.04+ (`src/testing/machine.ts:86-115`). This laptop passes (4 physical cores, 15.7 GiB RAM, Windows 11). But the rebuild is single-threaded, and the check has no single-thread-speed or CPU-generation component. The i7-10610U is a 2019 15W mobile part; the contract names a "mid-range 2023 laptop (Intel i5 class or AMD Ryzen 5 class)" (`performance.md:9`). A 2023 i5 is substantially faster per thread than a 2019 U-series i7. The calibration workload is reported alongside every result but changes nothing (`src/testing/machine.ts:121-202`). So the gate can enforce a budget on a machine whose single thread is slower than the reference class the budget was sized for. That is consistent with a stable 4 to 11 per cent miss.

Third, the miss is the same size as the noise. The eight runs span 1550 to 1671 ms — a 121 ms spread — while the gap over budget spans 50 to 171 ms. The §7 re-run missed every time, and the best run (1550 ms) still misses by 50 ms, so this is not a flaky outlier. But a gap that fits inside run-to-run variation on a busy laptop is not the signature of a discrete piece of waste; it is the signature of a budget set slightly below what this hardware delivers.

## 4. Is roughly 7 per cent findable in the code?

Candidates exist, but none is an obvious 7 per cent, and each needs a profiling run to size. The path is already tight in the places a reviewer would first look: evaluation is suspended during load (`src/engine/hyperformula-engine.ts:92-96`), the stale-dependency walk early-returns when there is nothing skipped (`src/engine/hyperformula-engine.ts:277`), shape generation and disposal sit outside the timer, and `prepareFormula` results are cached (`src/engine/formula-text.ts:86-93`).

Concrete candidates, in rough order of promise:

1. **Double HyperFormula parse per formula cell.** `setInput` calls `engine.validateFormula` and then `engine.setCellContents`, and HyperFormula parses the formula both times (`src/engine/hyperformula-engine.ts:212-222`, `330-336`). That is 50,000 redundant parses. The surrounding try/catch already handles a rejected formula by falling through to the literal path, so `validate` may be redundant — but removing a guard from the adapter needs a correctness review, and the saving is unmeasured. Needs timing before anyone claims it covers the gap.
2. **50,000 sequential `getCellValue` calls in `finishLoad`.** Each formula cell is read individually and wrapped in a fresh result object (`src/engine/hyperformula-engine.ts:117-123`). Whether HyperFormula offers a bulk read that is actually faster is unknown from the code; it needs checking against the HyperFormula API and measuring.
3. **Per-cell call overhead × 250,000.** Every cell allocates an address object and a key string and does at least one Map operation (`src/engine/hyperformula-engine.ts:208-237`, `381-383`), plus one HyperFormula call each. A bulk-load API, if one exists, could cut the fixed per-call cost — again, API check plus measurement required.
4. **Small change, small saving: the formula cache limit.** The probe uses 5,000 distinct formula strings (`A1*2` … `A5000*2`) against a 4,096-entry cache that clears itself when full (`src/engine/formula-text.ts:57-93`). A few hundred formulas are therefore scanned twice. This is microseconds, not milliseconds. Real sheets repeat formulas heavily, so the cache helps production more than the probe. Not the 7 per cent.

Plain verdict: there is no obvious 7 per cent. The remaining cost sits inside HyperFormula's graph construction, and extracting ~110 ms would most likely mean changing how the adapter drives HyperFormula (fewer, larger calls) rather than trimming calcplus-side logic. That is real engineering with an uncertain payoff, which is exactly why the profiling run comes first.

## 5. The honest options

`performance.md` is a contract: every number in it has a test behind it (`performance.md:3`), and changing a budget number is a contract change for the maintainer, not a casual edit.

- **Option A — raise the budget to a measured figure.** Set the structural-rebuild budget at 1750 ms (worst observed 1671 ms plus a small margin) or round to 2.0 s, and record the rationale: the 1500 ms figure was ~2× a 746 ms Linux measurement, and eight enforced runs on in-class hardware land at 1550–1671 ms. Update the §3.2 row, the §10 evidence note, `STRUCTURAL_REBUILD_BUDGET_MS` (`src/testing/structural-rebuild.ts:32`), `bench/model/STRUCTURAL-REWRITE.md` and the M7 exit record. Consequence: exit criterion 2 passes honestly; the guarantee at the top of the envelope gets slightly weaker; no code risk.
- **Option B — profile first, then decide on optimisation.** Run the probe once with `start`, `load` and `finishLoad` timed separately, then size candidates 1–3 above. Consequence: keeps 1500 ms on the table if a real saving exists; costs focused engineering time with no guarantee; any adapter change (especially around `validate`) carries regression risk and needs the full suite.
- **Option C — change what the case measures.** Ruled out as a way to hold 1500 ms. Including the snapshot and worker transfer (section 1) reflects production honestly but pushes the number **up**, not down. Measuring an incremental structural update instead of a full rebuild contradicts the spec ("the rebuild, not the delta", `performance.md:61`) and the controller's design, which rebuilds on structural changes (`src/engine/calculation-controller.ts:159-168`). Reference-update research (note 25) shows other engines funnel structural edits through one central transform, but calcplus's rebuild-rather-than-patch choice is a settled design, not a tuning knob.
- **Option D — tighten the reference-machine check.** Add a single-thread floor (e.g. a calibration threshold) so a 2019 U-series CPU no longer counts as the 2023 reference class. Consequence: this laptop's runs become advisory and the exit criterion passes vacuously — that hides the signal rather than resolving it. Not recommended as the primary move, though a CPU-generation note in §1 would stop the question recurring.

## Recommendation and reasoning

**Recommend Option A, preceded by the single profiling run from Option B.**

1. Run the probe once with per-stage timing to confirm that `load` dominates and that no calcplus-side stage (start-up, read-back loop) holds a cheap surprise. This closes the "code is slow" question with evidence instead of inference.
2. If the profile confirms graph-construction dominance — the expected outcome — raise the budget to **1750 ms** and record the derivation (ceil of enforced observations plus margin) in `performance.md` §10 and in the answer to issue 36, item 2. 1750 ms is preferred over 2.0 s because it stays close to the measured distribution instead of granting unused headroom.
3. Keep any adapter optimisation (candidates 1–3) as non-gating follow-up work, not as a blocker for M7.
4. Separately, note in `performance.md` §1 that the hardware check does not capture single-thread speed, so a future budget dispute does not restart from zero.

## Unresolved claims and maintainer decisions

1. **The maintainer must set the number.** 1750 ms is recommended, 2.0 s is the rounder alternative. Either is a contract change.
2. **The maintainer must rule whether the i7-10610U counts as reference-class for single-threaded budgets.** The gate says yes; the silicon generation says borderline. This affects every future CPU-bound budget, not just this one.
3. **Stage shares are inferred, not measured.** This note states that graph construction dominates based on the code path and the 2 ms recalc evidence. The profiling run in the recommendation would replace that inference with numbers.
4. **Optimisation savings are unknown.** Candidates 1–3 are places to look, not promised milliseconds. Nobody should commit to "find 7 per cent in code" without the profile.

## Sources

- `.scratch/calcplus/spec/performance.md:1-17` — reference machine and floor tier.
- `.scratch/calcplus/spec/performance.md:24` — guaranteed envelope, 250k cells / 50k formulas.
- `.scratch/calcplus/spec/performance.md:56-67` — §3.2 calculation budgets, rebuild-not-delta rationale, cold-build caveat.
- `.scratch/calcplus/spec/performance.md:164-171` — measurement harness and §7 re-run rule.
- `.scratch/calcplus/spec/performance.md:205-218` — evidence table: 746 ms build, 2 ms recalc, Node 24.15 Linux.
- `src/testing/structural-rebuild.ts:32`, `60-124` — budget constant, timed section, shape outside timer.
- `src/testing/machine.ts:23-40`, `86-119`, `121-202`, `246-283` — liveness ceiling, reference check, calibration reported-only, gate logic.
- `src/engine/hyperformula-engine.ts:53-124` — start/load/finishLoad; `208-237`, `330-336` — per-cell insert and double parse.
- `src/engine/formula-text.ts:57-93` — prepared-formula cache and its 4096 limit.
- `src/engine/calculation-controller.ts:148-168`, `197-214`, `474-534` — rebuild design, snapshot excluded from probe.
- `src/engine/formula-engine-client.ts:102-123`, `216-226` — chunked worker transfer excluded from probe.
- `bench/model/STRUCTURAL-REBUILD.md:1-36` — probe definition and contract budget.
- `bench/model/run.bench.ts:200-211`, `266-274` — same build path in the corpus harness; 1500 ms build budget for non-fan-out shapes.
- `tests/m3c-structural-rebuild.test.ts:25-44` — enforced gate assertion.
- `.scratch/calcplus/research/11-performance-evidence.md:19-51` — raw shapes: 746 ms row-local build, quadratic fan-out cliff, 2.6 µs per-formula recalc.
- `.scratch/calcplus/research/03-formula-engine.md:74-77`, `246-256` — dependency graph, suspend/batch, indicative build costs.
- `.scratch/calcplus/issues/36-followup-chrome-data-tools.md:32-83` — the open item: pre-lane-4 3740 ms, eight enforced runs 1550–1671 ms, the two ways out.

## Correction and outcome (2026-11-21)

Two things in the analysis above were wrong, and the work that followed settled the question the note asked.

**1. The 2 ms figure is incremental recalc, not the cold first evaluation.** This note used the 746 ms evidence table's "2 ms recalc" and concluded evaluation was about 0.1 per cent of the run. Instrumenting the stages showed the cold first evaluation inside `finishLoad` is about **270 ms**, not 2 ms. The measured split before any optimisation was `start` 13 ms, `load` 1007 ms, `finishLoad` 286 ms. The conclusion — graph construction dominates — still held, but the magnitude of evaluation was understated by two orders of magnitude.

**2. "There is no obvious 7 per cent" was wrong. There was about 39 per cent.** The note recommended raising the budget to 1750 ms instead of optimising. The maintainer ruled against that ("if there is scope for optimisation we should take it... we dont want performance creep to set in"), and the optimisation found far more than expected.

The measured stage split was then used to make three changes in `src/engine/hyperformula-engine.ts`, none of which raised the budget:

- `undoLimit: 0` in the engine config, since the app undoes through its own model history and the engine port exposes no undo. HyperFormula's per-write undo bookkeeping was a large share of bulk load.
- Batched block writes in `load` instead of one engine call per cell, using hole-free blocks so a chunked load cannot erase earlier chunks, with a replay of the per-cell path on any block failure.
- Memoised `validateFormula` per distinct formula text, because the verdict is constant per text within one engine build and HyperFormula re-parses on every call.

A fourth candidate was measured and **dropped because it was slower**: bulk-reading results with `getSheetValues` took 47.6 ms against 11.8 ms for 50,000 individual `getCellValue` calls, because it reads all 250,000 cells.

**Result.** Median of five enforced runs went from **1458 ms to 893 ms**, a 39 per cent reduction against a budget that never moved from 1500 ms. Four quiet verification runs measured 863 to 942 ms. The worst observed run is 36 per cent inside budget.

**What the note got right.** The budget was never a measured figure on the named reference hardware — it is roughly twice a 746 ms order-of-magnitude Linux measurement. The reference-machine check still does not test single-thread speed, so a 2019 i7-10610U passes a gate sized for 2023 silicon. And the probe still understates a real structural edit, since it excludes the model snapshot, the worker transfer and the result slicing. Those three points stand.

**Lesson worth carrying.** This note reasoned from the code path and called the result inference, which was honest, but then recommended a contract change on the strength of it. The stage measurement that would have confirmed or refuted it was one instrumented run away, and that run is what found the 39 per cent. Measure the stage split before recommending that a number move.
