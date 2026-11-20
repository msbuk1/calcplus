# calcplus v1 — technical design outline

This is the published outline of the architecture. Its source is
[`.scratch/calcplus/spec/technical-design.md`](../.scratch/calcplus/spec/technical-design.md),
which synthesises the contracts in
[`.scratch/calcplus/spec/`](../.scratch/calcplus/spec/README.md). If this page
and a contract disagree, the contract wins. The irreversible decisions are in
the [ADR index](adr/README.md).

## 1. Principles

1. **The workbook model is the source of truth.** `.xlsx` is an import/export
   format. Opening fills the model; saving patches only dirty parts
   ([ADR 0001](adr/0001-canonical-workbook-model.md)).
2. **No silent loss.** Anything not modelled is preserved and reported
   ([ADR 0003](adr/0003-preservation-first-ooxml.md)).
3. **One command registry.** Every surface dispatches a command ID; no surface
   mutates the model directly.
4. **One range-transform module.** Structural edits rewrite formulas, merges,
   names, validation, conditional formatting, comment/hyperlink positions and
   opaque anchors through one code path.
5. **The grid is a view.** It reads a read-only sheet projection; it never owns
   workbook state ([ADR 0007](adr/0007-sheet-projection.md)).
6. **The engine is a compute service.** HyperFormula calculates; it never
   supplies authored formula text or the saved result
   ([ADR 0004](adr/0004-hyperformula-engine.md)).
7. **Local-first.** No backend, no accounts, no telemetry.
8. **Measure the contract.** Performance, conformance and keyboard coverage
   have tests, not intentions.

## 2. Component map

```
UI shell (React)
  ribbon · compact bar · name box · formula bar · sheet tabs · status bar · dialogs
        │ command IDs, context, view state
        ▼
Command registry · transactions · undo/redo · focus · shortcuts
        │
        ├── Workbook model (canonical sparse cells, append-only styles, names,
        │   protection flags, opaque parts)
        ├── FormulaEngine port → HyperFormula in a Web Worker
        ├── Sheet projection → number formats, conditional-formatting overlay
        │        └── canvas grid renderer · DOM editor
        ├── OOXML layer (read, patch, preserve, write) → File System Access / download
        ├── Persistence (IndexedDB recovery copy, recent list, conflict prompts)
        └── Print planner → SVG preview · PDFKit · 300 DPI canvas fallback
```

The launcher sits outside the app: it serves the built static files and opens
the browser.

## 3. Main components

| Component | Responsibility | Contract |
|---|---|---|
| UI shell | Ribbon, bars, tabs and dialogs; every control dispatches a command and reads its enablement from the registry | [UI chrome](../.scratch/calcplus/spec/ui-chrome.md) |
| Grid | Custom canvas renderer with four-pane freeze geometry, viewport panes, DOM editor overlay; reads the projection only | [Grid spike](../.scratch/calcplus/spec/grid-spike.md), [ADR 0002](adr/0002-custom-canvas-grid.md) |
| Workbook model | Sheets, sparse typed cells, styles, names, theme, protection flags and opaque parts; stable IDs | [Workbook model](../.scratch/calcplus/spec/workbook-model.md) |
| Command system | One registry; one transaction per action; undo restores the exact prior state | [Command system](../.scratch/calcplus/spec/command-system.md) |
| Formula engine | HyperFormula behind a narrow port in a worker; 52-function registry; shims pinned by tests | [Formula semantics](../.scratch/calcplus/spec/formula-semantics.md) |
| Number formats | `numfmt` behind a `NumberFormatter` interface; calcplus owns the Excel seams | [Number formats](../.scratch/calcplus/spec/number-formats.md) |
| OOXML layer | Preservation-first package read, in-place part patch, append-only style and shared-string tables | [File compatibility](../.scratch/calcplus/spec/file-compat.md), [ADR 0006](adr/0006-patch-in-place-xml.md) |
| Persistence | Recovery copy in IndexedDB, recent list, external-edit conflict, close prompt; one workbook per tab | [Persistence](../.scratch/calcplus/spec/persistence.md) |
| Clipboard | HTML first, TSV second; `.xlsx` is the fidelity channel | [Clipboard](../.scratch/calcplus/spec/clipboard.md) |
| Print and PDF | Pure planner → `PrintPlan`; SVG preview and `window.print()`; PDFKit vector PDF; 300 DPI fallback | [Print and PDF](../.scratch/calcplus/spec/print-pdf.md) |
| Launcher | Go standard-library static server on `127.0.0.1:47823`, embeds `dist/`, opens the browser | [Distribution](../.scratch/calcplus/spec/distribution.md) |

## 4. Data flow

**Open.** Read the file, reject encrypted/corrupt/oversized input, parse the
package, map modelled parts into the canonical model while keeping original
bytes, configure the engine for the workbook's date system, project the sheet,
paint, and show the compatibility report only when something needs attention.

**Edit.** Input becomes a command. The transaction applies range-based changes,
the engine recalculates affected cells, the projection reformats them, the
renderer repaints dirty regions, and the recovery copy is scheduled.

**Save.** Collect dirty parts. Patch each against its source bytes; copy
untouched parts byte-for-byte; append new style and shared-string entries; apply
structural transforms; delete `calcChain` and set `fullCalcOnLoad` when the
contract requires it; write the package; show the export report.

**Undo.** Apply inverse changes in reverse order, recalculate, repaint. Sheet
delete is the exception: it is confirmed and terminal, with no undo entry, and
both history stacks are cleared
([ADR 0008](adr/0008-terminal-sheet-delete.md)).

## 5. Invariants

1. The model is the truth; `.xlsx` is at the edges.
2. Authored formula text and saved results never come from the engine.
3. Style and shared-string tables are append-only; IDs are stable.
4. One command registry; no surface mutates the model.
5. One transaction per user action; undo restores the exact prior state.
6. Structural edits use one range-transform module.
7. No silent loss; preserved parts are byte-identical unless a transform is
   reported.
8. View state never marks the document dirty.
9. The renderer reads the projection; it never reads the model.
10. The dependency graph is derived, not persisted.
11. Conditional formatting is an overlay; it never mutates the model.
12. A dirty part is patched, not regenerated; unmodelled children survive it.

## 6. Threading and scheduling

- **Main thread:** UI, input, selection, canvas paint.
- **Formula worker:** engine build and recalculation; deltas in, results out.
- **I/O worker:** package parse and write for large files.
- **Conditional formatting:** per-cell rules on demand, cached by row block.
- **Print and PDF:** planning on the main thread; PDF generation off the main
  thread where the library allows, otherwise chunked.
- **Persistence:** IndexedDB writes are async and never block input.
- No recalculation may block the main thread for more than 50 ms.

## 7. Performance and degradation

- Guaranteed envelope: 250k touched cells / 50k formulas. Degraded: 1M / 200k.
  Refuse above.
- Memory ladder: 500 MB guaranteed, 1 GB degraded, warning at 1 GB, refusal at
  1.2 GB.
- Degradation ladder: large-workbook banner → conditional formatting outside
  the viewport off → reduced overscan → volatile recalculation on idle → memory
  warning → refusal.
- Frame budget: 60 fps scroll, 50 ms direct interaction, 100 ms commit.

Budgets and the reference machine are in
[performance](../.scratch/calcplus/spec/performance.md).

## 8. Security and privacy

- No accounts, telemetry or workbook uploads; workbooks stay on the machine.
- Pasted HTML is sanitised; cell text is escaped.
- OOXML is treated as hostile: entity resolution off, DTDs rejected, package
  caps enforced, part paths checked.
- Macros are preserved but never executed; external links are preserved but
  never resolved.
- The launcher binds to `127.0.0.1` only and makes no outbound calls.

## 9. Distribution

- The static build deploys to any static host over HTTPS.
- The Go launcher embeds `dist/` and serves it on `http://127.0.0.1:47823/`.
  A second launch reuses the running instance; a busy port fails with a clear
  message; `--port N` uses a separate origin and says so.
- Windows x64 portable `.exe`, Linux x64 tarball, `SHA256SUMS`; no macOS
  artifact in v1. Builds are unsigned, and the launcher README explains the
  SmartScreen warning.
- No background updater and no outbound calls; updating means replacing the
  executable.

## 10. Build, test and release

- TypeScript + Vite; Vitest, Playwright and fast-check. Checks run locally
  before each commit; nothing runs on GitHub.
- Excel 365 on the maintainer's Windows PC is the conformance oracle.
- Launcher artifacts come from `node scripts/build-launcher.mjs` after
  `npm run build`; the script skips cleanly when Go is absent.
- Release gate:
  [testing and conformance](../.scratch/calcplus/spec/testing-conformance.md).
  Milestones: [milestones](../.scratch/calcplus/spec/milestones.md).

## 11. Deferred

Localization; formula-bar and name-box edge behaviours; fill-handle options,
merged-range selection, IME/bidi text and selection-scrolling edge cases;
co-authoring, cloud sync, charts, PivotTables, macro execution, mobile and dark
mode. The full list is the feature matrix's V2/NEVER set.
