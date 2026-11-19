# 29 — Selection-aware command targeting (pin the cross-cutting rule)

Status: resolved
Type: hardening
Blocked by: none

## Question

Historically the app has been unreliable at selection-aware command targeting:
after a user highlights a range, executing a command from the ribbon, a menu or
the context menu did not consistently apply to the whole selected range (for
example formatting landing on the active cell only). This is cross-cutting —
every range-scoped command and every dispatch surface is involved — so it needs
a pinned rule and a systematic suite, not per-command fixes.

**Current-state evidence (2026-11-18, branch `m5/data-tools`):** the seams
exist — `CommandContext.selection` (`src/commands/types.ts`), `args.range ??
context.selection ?? rangeOfActiveCell(context)` resolution in `sort.ts`,
`filter.ts`, `clipboard.ts`, `painter.ts` and `format.ts` (which applies per
cell across `target.range`), and `dispatchSheetCommand` in `src/chrome/App.tsx`
passing `selectionRange(selection)` at dispatch time. What is missing is the
rule as a contract line and tests that fail when any surface or command
regresses.

**The rule to pin:**

1. Every range-scoped command resolves its target from `args.range` →
   `context.selection` → the active cell (or the current contiguous region
   where its contract says so, e.g. sort §2.1, filter §3.1).
2. Every dispatch surface — ribbon, compact bar, context menu, dialog confirm,
   keyboard shortcut — passes the live selection into the command context at
   dispatch time. No surface may dispatch a range-scoped command with a stale
   or missing selection.
3. Applying to a multi-cell selection is ONE transaction; undo restores the
   whole prior state (command-system.md transactions).
4. Enablement predicates in the command map read the selection ("a range is
   selected") and follow it as it changes.
5. Dialogs prefill from the active cell but APPLY to the whole selection (the
   Format Cells dialog and the M5 Sort dialog's read-only target label, sort
   §2.1).
6. Edge cases named in tests: hidden rows inside the selection (sort §2.6),
   merged ranges inside the selection, a single-cell selection where a contract
   expands it, and a collapsed selection on an empty cell.

**Output:** a targeting paragraph in [command-system](../spec/command-system.md),
and a systematic suite in the command layer: for every range-scoped command in
the command map, dispatch through each surface and assert the whole selection
is affected in one transaction and undo restores it. This belongs with the
testing-conformance §6 command suite and the M5 gate.

## Concrete instance (2026-11-18)

The maintainer reports the Format Cells dialog failing on range targets with
"Select a cell or a range of at most 100,000 cells before formatting."

Diagnosis: the selection seam is healthy — `applyFormatCells` passes
`selectionRange(selection)` as both `args.range` and the command context. The
error is the `MAX_FORMAT_CELLS = 100_000` budget in `resolveFormatTarget`
(`src/commands/format.ts`), which shares `noTarget()` with the genuinely
empty-selection case, and which **no contract backs**: the command map's
enablement says only "a sheet is active and a range is selected", and the
performance contract budgets 100k/250k touched cells but never caps formatting.
Column picks (1,048,576 cells), whole-sheet and large blocks all trip it. The
same guard family bounds painter (100k), clear and clipboard.

Ruled 2026-11-18 (maintainer): quick fix now, dimension styles next.

1. **Quick unblock:** split the message ("nothing selected" vs "range too
   large") and raise the cap to the guaranteed envelope (250,000 cells) across
   the guard family.
2. **Real fix — dimension styles:** full-column/row formatting must apply as a
   column/row-dimension style (`RowDimension`/`ColumnDimension.styleId`, OOXML
   `col style="n"`), not per-cell writes — O(1) and unbounded, matching Excel.
   Needs style resolution in the renderer (dimension style with cell override),
   round-trip in the patch writer, and its own tests. Lands as a dedicated lane.

Both quick-fix items are done (commit 98b25e0). A second concrete instance —
right clicking a selected range collapsing it before the context menu opened —
was fixed the same day in `src/grid/CanvasGrid.tsx`: a right click inside the
selection keeps it (the menu acts on the whole range), a right click outside
moves it, and header right clicks keep a header selection that covers the
clicked axis. Proven by `tests/e2e/context-menu.spec.ts`.

## Answer

Pending.
