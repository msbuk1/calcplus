# calcplus v1 — technical design

Assembled by [15 — Assemble the build-ready spec package](../issues/15-spec-assembly.md) on 2026-11-13. It synthesises the contracts in this directory; where this document and a contract disagree, the contract wins and this document is corrected.

## 1. Principles

1. **The workbook model is the source of truth.** `.xlsx` is an import/export format. Opening fills the model; saving maps it back and patches only dirty parts.
2. **No silent loss.** Anything not modelled is preserved and reported.
3. **One command registry.** Every surface dispatches a command ID; no surface mutates the model directly.
4. **One range-transform module.** Structural edits rewrite formulas, merges, names, validation, CF, comment/hyperlink positions and opaque anchors through one code path.
5. **The grid is a view.** It reads a read-only `SheetView`; it never owns workbook state.
6. **The engine is a compute service.** HyperFormula calculates; it never supplies the saved formula text.
7. **Local-first.** No backend, no accounts, no telemetry.
8. **Measure the contract.** Performance, accessibility and conformance have tests, not intentions.

## 2. Architecture

```
┌────────────────────────── UI shell (React 19 + TypeScript) ──────────────────────────┐
│ Ribbon / CompactBar · NameBox · FormulaBar · SheetTabs · StatusBar · Dialogs          │
│ every control dispatches a command ID; no business logic in chrome                     │
└───────────────────────────────────┬───────────────────────────────────────────────────┘
                                    │ commands, context, view state
┌───────────────────────────────────▼───────────────────────────────────────────────────┐
│ Command registry · transactions · undo/redo · focus manager · shortcut map             │
└───────┬───────────────────┬───────────────────┬───────────────────┬───────────────────┘
        │                   │                   │                   │
┌───────▼────────┐  ┌───────▼────────┐  ┌───────▼────────┐  ┌───────▼────────────────┐
│ Workbook model │  │ FormulaEngine  │  │ SheetView      │  │ OOXML layer            │
│ canonical      │◄─┤ HyperFormula   │  │ projection     │  │ package · parts ·      │
│ sparse cells   │  │ in a worker    │  │ numfmt · CF    │  │ preserve · dirty write │
└───────┬────────┘  └────────────────┘  └───────┬────────┘  └───────┬────────────────┘
        │                                       │                   │
┌───────▼──────────────┐              ┌─────────▼────────┐  ┌───────▼────────────────┐
│ Persistence          │              │ Grid renderer    │  │ File System Access     │
│ IndexedDB recovery   │              │ canvas + panes   │  │ + download fallback    │
│ recent list          │              │ DOM editor/a11y  │  │                        │
└──────────────────────┘              └──────────────────┘  └────────────────────────┘
```

Supporting modules: clipboard, print/PDF, licensing notices and the launcher. They sit at the edges and do not own workbook state.

## 3. Components

### 3.1 UI shell

- React 19 + TypeScript + Vite.
- Ribbon is the primary chrome; the compact bar takes over below about 900 px.
- Primitives: `Ribbon`, `RibbonTab`, `RibbonGroup`, `RibbonButton`, `RibbonToggle`, `RibbonSelect`, `CompactBar`, `NameBox`, `FormulaBar`, `SheetTabs`, `StatusBar`, `Menu`, `ContextMenu`, `Dialog`.
- Every control takes a `commandId`; enablement and checked state come from the registry.
- Only the grid scrolls vertically; the chrome has fixed height.
- Contract: [UI chrome](ui-chrome.md).

### 3.2 Grid

- Custom canvas renderer with viewport-sized panes; no sheet-sized canvas and no full-height DOM spacer.
- Four-pane freeze geometry, custom scrollbars, virtual scroll, Fenwick/prefix-sum row geometry.
- DOM editor overlay for in-cell editing and the formula bar.
- `RenderFrame` is the renderer seam; a DOM-cell renderer can replace the canvas without touching the model.
- ARIA mirror: `role="grid"`, absolute indices, roving tabindex, persistent active cell, live region.
- Selection, active cell, fill handle, headers, merges, gridlines, conditional formatting and selection overlay.
- Contract: [grid spike](grid-spike.md) and [accessibility](accessibility.md).

### 3.3 Workbook model

- Canonical workbook: sheets, sparse typed cells, shared append-only style table, names, theme, named styles, protection and opaque parts.
- Stable workbook, sheet and style IDs; cells addressed by `(row, column)`.
- Document state vs view state; view state never marks the document dirty.
- Cell: value, authored formula text, saved result, style ID, rich-text payload.
- JSON snapshot for recovery, tests and a future sync layer; `.xlsx` is not the snapshot format.
- Dependency graph derived from the engine, never persisted.
- Opaque parts keep original bytes and an anchor index.
- Contract: [workbook model](workbook-model.md); [ADR 0001](../../../docs/adr/0001-canonical-workbook-model.md).

### 3.4 Command system

- One registry for ribbon, menus, context menus, keyboard, formula bar and dialogs.
- Commands declare `id`, label, category, icon, keybindings, `isEnabled`, `isChecked` and `run`.
- `CommandResult` carries range-based changes, an optional selection and an optional UI request.
- One transaction per action; 200 transactions or 64 MB of undo, whichever comes first.
- Replayable log with `seq` and `actorId` for a future co-authoring layer; no CRDT in v1.
- Focus zones: grid, formula bar, name box, ribbon/menu, dialog, find/replace, sheet tabs.
- Contract: [command system](command-system.md).

### 3.5 Formula engine

- HyperFormula GPL-3.0-only behind a narrow `FormulaEngine` port, in a Web Worker.
- Configuration: `licenseKey: 'gpl-v3'`, `leapYear1900`, `nullDate` for 1900/1904, `useArrayArithmetic: false`, `ignoreWhiteSpace: 'any'`, `evaluateNullToZero`, `caseSensitive: false`, `useWildcards`, `matchWholeCell`, comma arguments, dot decimal.
- Registry gated to the 52 published functions; everything else is `#NAME?`.
- Shims: wildcard exact lookup, `INT` floors toward negative infinity, `CHOOSE` out-of-range → `#VALUE!`, `#CYCLE!` → `#CIRC!`, no `#ERROR!`/`#SPILL!`/`#CALC!`/`#LIC!` in the model or file.
- The model owns authored text and cached results; the engine's rewritten text is never saved.
- Volatile `TODAY`/`NOW`; circular-reference warning; no spill; legacy arrays read-only.
- Contract: [formula semantics](formula-semantics.md).

### 3.6 Number formats

- `numfmt` 4.1.3 behind a `NumberFormatter` interface.
- Owned seams: built-in ID table, 1904 adapter, width layout, colours, locale grouping, invalid fallback, style-table round-trip.
- `ssf` is a differential test oracle only.
- Contract: [number formats](number-formats.md).

### 3.7 OOXML layer

- Preservation-first in-house package layer: `fflate` for ZIP, `saxes` for SAX parsing.
- Three tiers: modelled and editable, preserved and not editable, rejected.
- No-op save returns the original bytes; untouched parts are copied; only dirty parts are rewritten.
- Style and shared-string tables are append-only; existing indices are never renumbered.
- Structural edits go through the range-transform module; unsafe anchors block the edit.
- Compatibility and export reports surface every preserved, stale, transformed, removed or warned item.
- Package caps and XXE-safe parsing.
- Contract: [file compatibility](file-compat.md).

### 3.8 Persistence

- Chrome/Edge: File System Access API for open and save in place.
- Safari/Firefox: file input and download; no in-place save, explained once.
- Recovery copy: canonical JSON snapshot in IndexedDB, debounced about 3 seconds after the last edit, plus on hide/pagehide.
- Recent list with handles where available; external-edit conflict prompts; one workbook per tab.
- Stable origin from the launcher; `--port N` creates a separate storage area.
- Contract: [persistence](persistence.md).

### 3.9 Clipboard

- Read `text/html` first, TSV second; write both.
- Excel dialect HTML with typed cells, `mso-number-format`, merges and hyperlinks.
- Formulas are best-effort; `.xlsx` is the fidelity channel.
- In-app paste uses a `data-calcplus-*` payload plus an in-memory snapshot.
- Contract: [clipboard](clipboard.md).

### 3.10 Print and PDF

- Pure `PrintPlanner` → `PrintPlan` (page boxes, bands, scale, vector display list).
- Inline SVG for preview and `window.print()`; PDFKit for vector PDF; 300 DPI canvas fallback.
- Print area, orientation, fit-to-width and margins are in v1.
- Contract: [print and PDF](print-pdf.md).

### 3.11 Distribution

- Static build served by any static host over HTTPS.
- Go stdlib launcher embeds the build and serves it on `127.0.0.1:47823`; opens the default browser.
- Windows signed portable `.exe`; macOS notarized universal `.app` in a DMG; Linux static tarball.
- No background updater, no outbound calls, no telemetry.
- Contract: [distribution](distribution.md).

## 4. Data flow

### 4.1 Open

1. Read the file as an `ArrayBuffer`; reject CFB/encrypted, corrupt, oversized and Strict inputs.
2. Parse the OPC package and relationships; build the part store.
3. Map workbook, styles, shared strings and worksheets into the canonical model; keep unmodelled parts opaque.
4. Configure and build the formula engine for the workbook's date system.
5. Project the active sheet into `SheetView`.
6. Paint the grid; build the accessibility mirror.
7. Show the compatibility report if anything is preserved, stale or unenforced.

### 4.2 Edit

1. Input becomes a command through the registry.
2. The command returns range-based changes; the transaction applies them to the model.
3. The engine recalculates affected cells.
4. `SheetView` reformats the changed cells; conditional formatting re-evaluates.
5. The renderer repaints dirty regions; the accessibility mirror updates.
6. The recovery copy is scheduled; the document is marked dirty.

### 4.3 Save

1. Collect dirty parts.
2. Serialize modelled parts; copy untouched parts byte-for-byte; append new style/shared-string entries.
3. Apply structural transforms; delete `calcChain` when formulas changed and set `fullCalcOnLoad`.
4. Write the package; produce the export report.
5. On success, clear the dirty flag and the recovery copy.

### 4.4 Undo

1. Undo applies inverse changes in reverse order.
2. The engine recalculates affected cells.
3. The view repaints; the accessibility mirror updates.
4. Undo of a sheet delete restores the sheet and its stable ID.

## 5. Invariants

1. The model is the truth; `.xlsx` is at the edges.
2. Authored formula text and saved results never come from the engine.
3. Style and shared-string tables are append-only; IDs are stable.
4. One command registry; no surface mutates the model.
5. One transaction per user action; undo restores the exact prior state.
6. Structural edits use one range-transform module.
7. No silent loss; preserved parts are byte-identical and reported.
8. View state never marks the document dirty.
9. The renderer reads `SheetView`; it never reads the model.
10. The dependency graph is derived, not persisted.
11. Conditional formatting is an overlay; it never mutates the model.
12. The accessibility mirror is the grid's accessible representation; the canvas is `aria-hidden`.
13. Performance and accessibility are contracts with tests.

## 6. Threading and scheduling

- **Main thread:** UI, input, selection, canvas paint, accessibility mirror.
- **Formula worker:** engine build and recalculation; deltas in, results out; request coalescing and backpressure.
- **I/O worker:** package parse and write for large files; the model stays on the main thread and receives a snapshot.
- **Conditional formatting:** per-cell rules on demand; range statistics cached and recomputed lazily in row-block chunks.
- **Print/PDF:** planning on the main thread; PDF generation off the main thread where the library allows, otherwise chunked.
- **Persistence:** IndexedDB writes are async and never block input.
- No recalc may block the main thread for more than 50 ms.

## 7. Performance and degradation

- Reference machine and budgets: [performance](performance.md).
- Guaranteed envelope: 250k touched cells / 50k formulas; degraded: 1M / 200k; refuse above.
- Memory ladder: 500 MB guaranteed, 1 GB degraded, warning at 1 GB, refusal at 1.2 GB.
- Degradation ladder: large-workbook banner → CF outside the viewport off → reduced overscan → volatile recalc on idle → memory warning → refusal.
- Frame budget: 60 fps scroll, 50 ms direct interaction, 100 ms commit.
- The engine runs in a worker; the grid never waits on recalc to paint selection.

## 8. Accessibility

- WCAG 2.2 AA for the chrome and the DOM mirror; the canvas is decorative.
- `role="grid"`, absolute counts/indices, headers, merged spans, roving tabindex, persistent active cell, `aria-selected`, live region.
- Every command keyboard-reachable; F6 zones; Esc returns to the grid.
- 2 px 3:1 focus ring; forced-colours canvas palette; 200% browser zoom; 24 × 24 chrome targets.
- AT sign-off before the grid milestone; DOM-cell fallback behind `RenderFrame`.
- Contract: [accessibility](accessibility.md).

## 9. Security and privacy

- No accounts, telemetry or workbook uploads.
- Paste HTML is sanitised; no scripts, event handlers or external resources.
- OOXML parsing is XXE-safe: entity resolution disabled, DTDs rejected, package caps enforced.
- Macros are preserved but never executed.
- External links are preserved but never resolved or refreshed.
- No outbound calls from the launcher or the app.
- The file stays on the machine; the recovery copy is local IndexedDB.

## 10. Failure handling

- Tier C inputs are rejected with a clear message and no partial import.
- A failed save keeps the recovery copy, shows an error and does not mark the document clean.
- A structural edit that cannot transform an opaque anchor is blocked with an explanation.
- Unsupported rules, parts and functions are preserved and reported, never approximated silently.
- A crashed tab offers Restore or Discard on the next start.
- An external file change offers Reload, Overwrite or Save As.
- Formula errors are cell values, not crashes.

## 11. ADR index

| ADR | Decision |
|---|---|
| [0001](../../../docs/adr/0001-canonical-workbook-model.md) | The workbook model, not the `.xlsx` file, is the source of truth |
| [0002](../../../docs/adr/0002-custom-canvas-grid.md) | Build the custom canvas grid; reject Univer |
| [0003](../../../docs/adr/0003-preservation-first-ooxml.md) | Preservation-first in-house OOXML layer |
| [0004](../../../docs/adr/0004-hyperformula-engine.md) | HyperFormula behind a narrow engine port |
| [0005](../../../docs/adr/0005-gplv3.md) | GPL-3.0-only and its obligations |

The decision tickets on the map remain the narrative record; the ADRs capture the irreversible choices.

## 12. Build, test and release

- TypeScript + Vite; Vitest, Playwright, fast-check and axe-core.
- Excel 365 on a Windows runner is the conformance oracle; stored goldens run on pull requests.
- Pull requests under 15 minutes; nightly full suite; scheduled oracle regeneration.
- Release gate: [testing and conformance](testing-conformance.md).
- Milestones: [milestones](milestones.md).
- Licence and notices: [licensing](licensing.md).

## 13. Deferred and open

- Localization of the UI and function names.
- Formula-bar and name-box edge behaviours.
- Fill-handle options, merged-range selection, IME/bidi text and selection scrolling edge cases.
- Co-authoring, cloud sync, charts, PivotTables, macros execution, mobile and dark mode.
- The full list is the feature matrix's V2/NEVER set.
