# 36 — Follow-up: open items after the chrome and data-tools sweep

Status: open
Type: follow-up (tracking; each item is small and self-contained)
Blocked by: nothing

## Question

The chrome and data-tools sweep landed as commit `4051f8b` on 2026-11-20, eight
lanes together. It closed tickets 29 and 31, filed 33, 34 and 35, and fixed the
merge, freeze, filter, ribbon and status-bar gaps the maintainer found in real
use. Four things came out of it unresolved, plus two decisions the lanes asked
for. None is big enough for its own milestone slice, so they are collected here
so none is forgotten.

## The open items

### 1. Merged-text rendering versus the visual baseline

`tests/e2e/formatted-visual.spec.ts` fails on the `m4a-formatted-sheet.png`
snapshot. The actual paint draws the full merged text; the checked-in baseline
shows it clipped. It is a 2 per cent pixel difference.

This has been left alone across every lane because it needs a ruling, not a fix:
either the paint is right and the baseline is stale, or the paint is wrong and
merged labels are drawing over content they should not. `.scratch/calcplus/spec/conditional-formatting.md`
section 5.2 and the merged-label renderer work in M5 both bear on it.

**Needs:** the maintainer to look at the diff and say which is correct. Only
then can the snapshot be refreshed or the renderer corrected.

### 2. Two performance budgets are missed on this machine

Both fail consistently on a quiet machine now, not just under load, and both
failed before any of the eight lanes:

- `tests/m3c-structural-rebuild.test.ts` — a 250k-cell, 50k-formula dependency
  graph rebuilds in about 3740 ms against a 1500 ms budget.
- `tests/m5-perf-fixtures.test.ts` — a 100k-row sort runs in about 2406 ms
  against a 2000 ms budget.

The machine is an Intel Core i7-10610U at 1.80 GHz with 8 logical cores. The
budgets are set in `.scratch/calcplus/spec/performance.md` section 3, which
names a reference machine. Either the budgets assume a faster reference machine
than this one, or these are real regressions worth finding.

**Needs:** decide whether the reference machine is the mismatch or the code is
slow. If it is the machine, the budgets or the reference-machine note should
say so. If it is the code, this is a real performance fix and belongs in M7.

**Evidence (2026-11-20).** M7 lane 4 answered the reference-machine half, and
the answer is that the machine was never the problem.

`performance.md` §1 now names the class: a mid-range 2023 laptop, 4-core CPU,
16 GB RAM, integrated graphics, AC power. `src/testing/machine.ts` gates on it
— `mode: 'enforced'` only when the run is **declared** with
`CALCPLUS_REFERENCE_MACHINE=1` and the hardware matches; every other run is
`advisory`, with a liveness ceiling of ten times the budget so a real
regression still fails. `tests/m7-reference-machine.test.ts` pins all of it,
including the one immediate re-run that performance §7 allows.

This maintainer's laptop **is a valid member of the class**: Intel Core
i7-10610U, 4 physical cores, 15.7 GiB RAM, Windows 11. So an enforced run is
possible here, and it was run. The result splits the two budgets this ticket
named:

- **`sort-100k` holds.** `tests/m5-perf-fixtures.test.ts` passes in enforced
  mode — sort, filter and both conditional-formatting budgets all inside their
  numbers against the 2000 ms, 1000 ms, 200 ms and 4 ms budgets.
- **`structural-rebuild-250k` misses.** Against its 1500 ms budget, five
  enforced runs measured 1550, 1550, 1575, 1612, 1616, 1625, 1657 and 1671 ms.
  It never passed. The §7 re-run missed every time. That is a consistent 4 to
  11 per cent over, not noise.

So it is the code, or the number — not the machine. Two ways out, and this is
the call: find roughly 7 per cent more in the rebuild path, or set the budget
at the measured 1.7 s and record why. Note the rebuild already improved a long
way under lane 4: it measured 3740 ms before that work.

M7 exit criterion 2 is "the performance budgets hold on the reference machine".
On this evidence it does **not** hold today, and the milestone record's note
that the budgets "ran advisory" is now out of date — the class exists, this
machine is in it, and enforcement is real.

### 3. Reconciliation pass over the command inventory

`src/commands/command-inventory.json` is the documented source of truth for the
ribbon map, but it contradicts what the code does. The `implemented` flag is
unreliable in three directions:

- **Not stale after all.** The whole milestone 6 file layer and read-only
  fidelity set say `implemented: false`, and the code is complete: `.xlsx` and
  `.xlsm` open and save, CSV import and export, encrypted rejection, external
  links, protection, hidden state, conditional formatting, themes, defined
  names, validation, tables and pivot caches. Selection navigation, in-cell
  editing, formula bar editing, freeze rendering, zoom and status-bar
  statistics are the same. Those flags are **correct** — see the correction at
  the end of this item. `implemented` records whether a command with that id is
  registered, and none of these is a command. Do not flip them.
- **Stale true or unverified.** `name.resolve` and `formula.evaluate` may claim
  more than works. Two entries were never checked: whether Show Comments
  toggles anything, and whether defined names resolve when typed into a formula.
- **Correctly false.** Print, export PDF, recent files, crash recovery are M7.
  Shrink-to-fit and text-rotation authoring, conditional-formatting authoring
  and manual F9 recalculation are deliberate v2 cuts.

The sweep already fixed the gaps it found: merge, Clear All, Close, function
autocomplete, freeze panes and the filter affordance are now real and marked
true. What is left is making the flag tell the truth everywhere else, and
settling the two unverified entries.

**Needs:** one pass that reads every `implemented: false`, checks the code, and
either flips the flag or records why it stays false. Run `npm run inventory`
afterwards and confirm the generated `docs/command-inventory.md` is stable.

**Progress (2026-11-20).** One of the two unverified entries is settled.
Defined names in formulas **work**: the engine registers every name as a
HyperFormula named expression (`src/engine/calculation-controller.ts`
`namedExpressionsFor`, `src/engine/hyperformula-engine.ts`), Go To and the name
box resolve names (`tests/m4c-find.test.ts`), and preservation round-trips
(`tests/m6-sheet-features.test.ts`). The missing piece was evidence, not code:
nothing ever put a defined name inside a formula. `tests/m3c-defined-names.test.ts`
now does, covering workbook and sheet scope, a name inside a function, a name
over a range, scope precedence, two names in one formula, and recomputation when
the named target changes.

There is no `name.resolve` command code and none is needed — resolution is
engine behaviour, not an invocable command. Excel's **Use in Formula**
drop-down does not exist either; the Defined Names group is only a placeholder
in `ribbon-model.ts` `GROUP_ORDER`. The second unverified entry — whether the
Show Comments ribbon button toggles anything — is still open.

**Correction to this item's premise.** `implemented` does not mean "the
capability works". `scripts/generate-command-inventory.mjs` enforces it in both
directions as "the command registry has a command with this id", and it fails
the run otherwise. So `name.resolve` and `name.preserve` are **correctly**
`false`: neither is a command. The same is true of the whole `fidelity.*` set
described above as stale — preservation is engine and file behaviour with no
invocable command, so those flags are right too. Flipping any of them breaks
the generator, which is the guard that caught this.

What the reconciliation pass therefore needs is not flag flipping. It is
deciding, per entry, whether the missing thing is **a command** or **a test**.
For defined names it was a test, and `tests/m3c-defined-names.test.ts` now
carries it. Keep the flag, and record the proven behaviour where a reader will
look for it.

### 4. Drag-and-drop features are still missing

Three version-one features confirmed missing after the sweep, all mouse drags:

- **Drag-and-drop move** (`edit.dragMove`) — drag a selected range to move it.
- **Shift-drag insert** (`edit.dragInsert`) — Shift-drag to insert and shift
  cells down or right.
- **Drag a file onto the app to open it** (`file.openDropped`) — no drop
  handler exists.

The grid's drag handling is the fill handle, the Format Painter and the range
autoscroll added in this sweep. None of these three exists. All are promised as
IN in `.scratch/calcplus/spec/feature-matrix.md`.

**Needs:** build all three. The autoscroll work in `src/grid/CanvasGrid.tsx`
and the shared pointer-drag pattern are the place to start.

## The two decisions

### 5. Select All placement in the filter dropdown

The lane pinned Select All above the scrolling value list so it stays reachable
while the values scroll. Excel scrolls it with the values. Recorded and left as
is.

**Needs:** the maintainer to say which layout is wanted. Small either way.

### 6. The spec line on selection scrolling

`.scratch/calcplus/spec/README.md` lists "selection scrolling edge cases" as
outside version one. The autoscroll fix in this sweep covers that ground. The
line is now wrong or the work landed early.

**Needs:** update the status line, or record the autoscroll work as a version
one bug fix.

## Ruled direction

Work the items in the order above, since 1 and 2 need a maintainer call before
anyone can act. Item 3 is a single mechanical pass. Item 4 is real feature work
and can be split into its own tickets if it grows.

## Acceptance shape

1. The merged-text snapshot is either refreshed or the renderer fixed, and
   `tests/e2e/formatted-visual.spec.ts` passes.
2. The two performance budgets pass on the stated reference machine, or the
   budgets and the reference-machine note are corrected to match reality.
3. Every `implemented` flag in `command-inventory.json` matches the code, the
   two unverified entries are settled, and the generated docs are stable.
4. The three drag features work and are covered by tests.
5. Select All placement is settled and the filter tests match.
6. The spec line is corrected.

## Placement

No milestone owns this ticket. Items 1, 2, 5 and 6 are small and can be done at
any time. Item 3 is a documentation and metadata pass. Item 4 is feature work
and should be split out before it starts.

## Answer

Pending.
