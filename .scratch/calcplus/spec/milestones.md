# calcplus v1 — milestones and phasing

Resolved by [14 — Milestones and phasing](../issues/14-milestones.md) on 2026-11-13. It sequences the contracts in [feature matrix](feature-matrix.md), [workbook model](workbook-model.md), [formula semantics](formula-semantics.md), [file compatibility](file-compat.md), [performance](performance.md), [accessibility](accessibility.md) and [testing and conformance](testing-conformance.md).

## 1. Principles

- **Hybrid vertical slices.** The first two milestones build the foundation; every later milestone ends in a demoable user journey.
- **Walking skeleton first.** The real OOXML package layer, the canonical model and a minimal grid are integrated in milestone 1. No interim lossy library.
- **A demo and a gate per milestone.** A milestone is not done because the code merges; it is done when its exit criteria pass.
- **Continuous quality.** Accessibility, performance and conformance tests are built with the features, not after them.
- **Risk spikes are scheduled.** The four risks that can invalidate an architecture have explicit spikes, not hopes.
- **The feature matrix is the scope fence.** Anything marked V2 or NEVER does not enter a milestone.

## 2. Team and estimates

- **Three engineers in parallel:** one on grid and rendering, one on model, engine and commands, one on file I/O and chrome. Part-time design and QA.
- **Person-weeks, not calendar dates.** The plan stays useful whatever the start date. A range is a planning estimate, not a commitment.
- **Total:** about **64-94 person-weeks**, plus a **20% contingency**, for roughly **77-113 person-weeks**. With three engineers and the dependencies below, that is about **26-38 working weeks** — six to nine months — if the tracks stay staffed.
- The conformance tail is the largest uncertainty. Milestone 8 carries an explicit reserve.

## 3. Overview

| # | Milestone | Demoable outcome | Effort (pw) | Main risk |
|---|---|---|---:|---|
| 1 | Walking skeleton | Open a small `.xlsx`, see it, edit a cell, undo, save with untouched parts byte-identical | 4-6 | Package parsing and model shape |
| 2 | Grid foundation | Keyboard navigation, editing, frozen panes, 1M-row scroll | 8-12 | Canvas rendering at scale |
| 3 | Formula engine and model depth | Formulas recalculate, 52-function subset passes, undo/redo, dates and names | 8-12 | HyperFormula shims and fan-out |
| 4 | Formatting and workbook | Format a report, number formats, merges, multiple sheets, find and Go To | 8-12 | Number-format seams |
| 5 | Data tools | Sort, AutoFilter, conditional formatting, clipboard, fill, insert/delete | 8-12 | Range transforms and clipboard |
| 6 | File fidelity | Charts, pivots and macros survive a round-trip; Excel opens without repair | 14-20 | Structural edits and preservation |
| 7 | Output and hardening | Print/PDF, 250k-cell performance, launcher | 8-12 | Print fidelity |
| 8 | Conformance and release | Release candidate with every gate green | 6-8 + reserve | The long tail |

## 4. Milestones

### M1 — Walking skeleton (4-6 pw)

**Build**

- Repo, build, lint, typecheck, Vitest and Playwright harnesses.
- Canonical workbook model: workbook, sheet, sparse cells, minimal styles, stable IDs, source-part skeletons, snapshot metadata plus binary part payloads.
- OOXML package layer: ZIP via `fflate`, SAX via `saxes`, part store, content types and relationships, one worksheet mapped into the model with its original bytes retained as the patch skeleton.
- No-op save returns the original bytes; untouched parts are copied; a dirty part is patched in place — cells only in M1, with row and column shifting arriving with M5 on the same writer — so the first edited save already preserves unmodelled sheet children.
- Minimal canvas grid: render values, click to select, type to edit, commit with Enter.
- Minimal command registry and one undoable `cell.set`.
- File open and save through the File System Access API with a download fallback.
- Conformance infrastructure stands up: the maintainer's own Excel 365 on their Windows PC with a COM/PowerShell harness, value and error golden generation run manually, and the browser memory harness from ticket 11 (testing-conformance §2). Sort-collation, CF-paint and print-pagination goldens are added by M5 and M7, when those features exist.

**Demo:** open a small workbook, see its values, edit a cell, undo, save, reopen in Excel; the untouched package parts are byte-identical, and the chart, drawing and extension children of the edited sheet survive.

**Exit criteria**

1. The model opens a minimal `.xlsx` and renders it.
2. No-op open → save is byte-identical.
3. A cell edit is one undoable command; undo restores the prior value.
4. An edited save of a workbook carrying a chart, a `<legacyDrawing>` and an unknown `<extLst>` keeps those elements and their order in the dirty worksheet part; a part inventory proves it, not inspection.
5. The app opens in Chromium and saves through the file API or a download.
6. Unit and smoke tests pass on the maintainer's PC; no GitHub test execution is required.

### M2 — Grid foundation (8-12 pw)

**Build**

- 1,048,576 × 16,384 addressable canvas grid with viewport-sized panes.
- Frozen rows and columns; four-pane geometry; custom scrollbars; virtual scroll.
- Selection, active cell, fill handle, headers, merged-range rendering.
- DOM in-cell editor with IME, formula bar, name box, commit/cancel.
- Keyboard navigation, focus zones and the shortcut map's navigation subset.
- Basic formatting rendering: font, fill, alignment and number formats from the style table.

**Demo:** a keyboard-only user navigates, selects, edits and scrolls; a 250k-touched-cell sheet scrolls at the frame budget and a 1M-row sparse sheet scrolls without allocating the grid.

**Exit criteria**

1. Selection, editing, formula bar, freeze panes and keyboard navigation work.
2. Scroll holds the ticket 11 frame budget on a 250k-cell sheet.
3. The focus model, Esc/F6 paths and the shortcut subset pass the keyboard suite.

### M3 — Formula engine and model depth (8-12 pw)

M3 runs as three sub-slices: engine wiring, function coverage and model depth. Together they keep M3's total scope and its six exit criteria; each slice carries its own demo and gate, and M3 is done when all three slice gates pass.

**Spikes:** HyperFormula fan-out, memory and worker behaviour at the ticket 11 sizes; a prototype of the OOXML range transform for insert/delete.

#### M3A — Engine wiring

**Build**

- `FormulaEngine` port and HyperFormula adapter with the required configuration.
- Dependency graph and incremental recalculation; volatile `TODAY`/`NOW`; circular-reference warning.

**Demo:** enter formulas and watch dependents recalculate.

**Gate:** recalculation works after edits, volatiles update, circular warning shows.

#### M3B — Function coverage

**Build**

- The 52-function registry gate; unsupported functions return `#NAME?` or a stale cached value.
- Shims: wildcard exact lookup, `INT`, `CHOOSE`, `#CYCLE!` → `#CIRC!`, error mapping, authored-text preservation.

**Demo:** the 52-function golden subset passes.

**Gate:** M3 exit criteria 1, 2, 3 and 4 pass.

1. The 52-function golden subset passes within tolerance.
2. The error/coercion/operator table passes.
3. Date, reference, lookup and circular-reference tests pass.
4. Every shim has a passing test.

#### M3C — Model depth

**Build**

- 1900 and 1904 date systems; serial 60.
- Defined names, cross-sheet references, whole-column and whole-row references.
- Full command registry, transaction boundaries and undo/redo for cell edits and formatting.
- Style table depth: fonts, fills, borders, alignment and number-format references.
- A prototype of the OOXML range transform for insert/delete.

**Demo:** switch between 1900 and 1904 workbooks, use names across sheets, undo and redo, and run the range-transform prototype rewrite.

**Gate:** M3 exit criteria 5 and 6 pass.

5. Undo/redo restores the exact prior state for cell and formatting commands.
6. The range-transform prototype rewrites formulas, merges and names correctly and meets the structural-edit rebuild budget on the 250k shape ([performance §3.2](performance.md)).

### M4 — Formatting and workbook (8-12 pw)

M4 runs as three sub-slices: cell formatting and number formats, workbook structure and view, and chrome and find. Together they keep M4's total scope and its six exit criteria; each slice carries its own demo and gate, and M4 is done when all three slice gates pass.

#### M4A — Cell formatting and number formats

**Build**

- Full cell formatting: font family/size/bold/italic/underline/strike/colour, fills and patterns, borders, alignment, indent, wrap, shrink-to-fit and rotation (imported).
- Number formats through `numfmt` plus the Excel seams: built-in IDs, 1904, width layout, colours, locale, invalid-format fallback.
- Merges, row heights, column widths, autofit, hidden rows/columns (imported).

**Demo:** format cells with fonts, fills, borders, currency and date formats.

**Gate:** M4 exit criteria 1, 2 and 6 pass.

1. Every formatting feature in the matrix renders and round-trips.
2. The number-format conformance corpus passes.
6. Visual baselines for the formatted sheet pass.

#### M4B — Workbook structure and view

**Build**

- Multi-sheet insert, delete and rename; sheet tabs; tab order and visibility.
- Zoom 10-400%, gridline and heading toggles, status-bar statistics.

**Demo:** three sheets with zoom and toggles.

**Gate:** M4 exit criterion 3 passes.

3. Multi-sheet insert/delete/rename and undo work.

#### M4C — Chrome and find

**Build**

- Chrome: ribbon primary, compact bar below 900 px, dialogs, context menu, status bar.
- Find & Replace and Go To; the Format Cells dialog; Format Painter.

**Demo:** ribbon and compact bar with find and dialog.

**Gate:** M4 exit criteria 4 and 5 pass.

4. Find/Replace, Go To and the Format Cells dialog pass the interaction suite.
5. The ribbon and compact bar meet the UI chrome contract.

### M5 — Data tools (8-12 pw)

**Build**

- Sort: selected range, header detection, up to 64 keys, Excel comparator, formula rewriting, hidden-row behaviour, merges and undo.
- AutoFilter: value checklist, two AND/OR conditions, text/number/date, blanks, hidden layers, dropdown UI, round-trip.
- Conditional formatting: cell-value, text, date/time-period, top/bottom, duplicate/unique, colour scales, data bars and icon sets; priority and stop-if-true; unsupported rules reported.
- Clipboard: HTML Excel dialect and TSV fallback; in-app copy/cut/paste; Excel interop where the OS allows.
- Fill handle drag and double-click; Ctrl+D/R; insert/delete cells, rows and columns with full reference rewriting, through the M1 part patch writer including row and column renumbering.
- Context menu, Clear, Format Painter.

**Spike:** Excel clipboard interop for values, formulas, formats and merges.

**Demo:** sort and filter a 100k-row table, show conditional formatting, copy a range to Excel and back, fill a series with the handle, insert a row and watch references update.

**Exit criteria**

1. The 15 sort and 18 filter fixtures pass.
2. The 20 conditional-formatting fixtures pass.
3. Clipboard round-trips values, formulas, formats and merges with Excel.
4. Fill and insert/delete are one undo transaction each.
5. 100k-row sort and filter meet the performance budgets.

**Exit note (added 2026-11-19).** The exit gate passed. Criteria 1, 2, 4 and 5
were proven by the local conformance and performance suites; criterion 3 was
verified by the maintainer on Windows against Excel 365 with the clipboard
interoperability spike and the oracle golden captures. The M4A win32 visual
golden was refreshed after the merged-label renderer fix. M5 is closed, and
the M6 exit gate is no longer blocked on it.

### M6 — File fidelity (14-20 pw)

**Build**

- Full `.xlsx` round-trip: styles and shared strings append-only; cached results; shared and array formulas; comments and threaded comments; hyperlinks; validation; conditional formatting; names; protection; print setup.
- Preserved parts: charts, images, shapes, pivots and caches, external links, VBA, signatures, tables, slicers, sparklines, custom XML and custom properties.
- Structural-edit range transforms across formulas, merges, names, validation, CF, comment/hyperlink positions, drawing anchors and chart series ranges; block only a non-range chart source, a pivot source range or an untransformable anchor.
- Compatibility report and export report; Tier C rejection for encrypted, corrupt, zip-bomb and Strict files.
- Producer matrix: Excel 365/2019, LibreOffice, Google Sheets, openpyxl, SheetJS, xlsx-populate.
- Hostile-input caps and XXE-safe parsing.

**Demo:** open a workbook with a chart, a pivot, an external link and a macro; edit cells and insert rows; save; Excel opens without repair; the reports list what was preserved and what went stale.

**Exit criteria**

1. All 24 file-compat fixtures pass.
2. No-op save is byte-identical; untouched parts survive with identical bytes.
3. Excel 365 opens every output without a repair prompt.
4. The producer matrix passes.
5. Structural edits transform or block correctly; no silent stale anchors.
6. `.xlsm` macros survive; `.xlsm` → `.xlsx` requires confirmation.
7. Hostile inputs fail safely within the caps.

**Handover entry notes (added 2026-11-18).** M6 is taken by a second engineer
once M5 closes. Fixture authoring (F1–F24) may start before the M5 exit — the
gate only requires M5 to have exited before the M6 exit gate runs. Read
[docs/m6-kickoff.md](../../../docs/m6-kickoff.md) before building: it lists
the work already banked by the M5 lanes (byte-identical saves, unknown-XML
preservation, the patch writer with row/column renumbering, the range-transform
module and the reporting hooks), the governing contracts, and the process
rules. The risk-register row that owns this milestone — structural transforms
corrupting references or anchors — is governed by one rule: transform or block,
never leave a silently stale anchor.

**Exit note (added 2026-11-19).** The exit gate passed. Criteria 1, 2, 5, 6
and 7 are proven by the local suites: the 24 fixtures, the identity harness
(no-op byte identity, dirty-part preservation, semantic equality), the
transform-or-block work including internal hyperlink locations, the macro
preservation and confirmation flow, and the hostile-input caps with XXE-safe
parsing. Criterion 3 was verified by the maintainer's demo run: a workbook
with a chart, macro, hyperlinks and formulas round-tripped through calcplus
and reopened intact in Excel 365, and the same output opens in Google Sheets
and LibreOffice. Criterion 4: the producer matrix is green — openpyxl,
LibreOffice, SheetJS and xlsx-populate generated; Google Sheets and Excel
365 maintainer-captured on 2026-11-19; Excel 2019 recorded as not available
(maintainer ruling). Two asterisks are recorded in
[docs/m6-kickoff.md](../../../docs/m6-kickoff.md): the demo workbook carried
no real PivotTable or external formula link (fixture F10 covers those), and
ticket 32 replaces the cross-sheet-link refusal with the break-and-mark
policy the research note recommends.

**Criterion 3 reopened, then met (2026-11-21).** The 2026-11-19 demo run
stood for the workbook it tested, but the criterion was **not met**: a workbook
saved by calcplus carried `<f>SUM()</f>` in its sheet part, Excel 365 showed a
repair prompt and stripped every Formula record from that part, which is data
loss in a file we produced. Two causes were found. The entry screen refused
only typed edits, so a formula Excel refuses could still arrive through a load
and be written straight back; and `dimension ref` under-declared the written
range in two fixtures.

The fix keeps such a formula visible in the model so it can be corrected and
never serialises it: the cell is written with its cached value alone, which is
what Excel's own repair leaves behind, and both reports name the cell instead of
the change passing silently. A file carrying one is no longer byte-identical on
a no-op save; every other file still is. The entry screen's remaining gaps are
ticket 48.

Verified on 2026-11-21: the maintainer's own workbook round-tripped through
calcplus — 29 formulas before, 28 after, `SUM()` gone and the other 28 intact —
and **Excel 365 opened it with no repair prompt**, which is also what the COM
check on the same output shows. **Exit criterion 3 is met.**

### M7 — Output and hardening (8-12 pw)

**Build**

- Print/PDF: `PrintPlanner`/`PrintPlan`, print area, orientation, fit-to-width, margins; SVG preview; PDFKit vector export; 300 DPI canvas fallback.
- Performance hardening: corpus, fan-out, memory, worker queue, degradation ladder.
- Persistence: autosave recovery copy, recent list, external-edit conflict, close prompt, stable origin.
- Launcher: Go static server on `127.0.0.1:47823`, Windows portable `.exe` and a Linux tarball, unsigned with the SmartScreen warning documented.
- Documentation skeleton: PRD, technical design, ADR index, known deviations.

**Spike:** vector print/PDF fidelity and font embedding.

**Demo:** print preview and PDF of a formatted report; NVDA and VoiceOver pass; a 250k-cell workbook scrolls at 60 fps; the launcher opens the app offline.

**Demo note (2026-11-20, maintainer).** The NVDA and VoiceOver pass is out
of the M7 demo: accessibility verification stays V2 with the rest of the
assistive-technology work (`accessibility.md`). The demo is the print/PDF
check, the 250k-cell scroll and the offline launcher.

**Testing status (2026-11-20, maintainer).** Verified by hand: the Linux
launcher artifact (the app opens offline at `127.0.0.1:47823`, current
build, icons and branding correct; the build-id guard refuses to reuse a
stale instance); the print preview, print and PDF export checked side by
side against Excel 365, which matches the oracle and honours the page
setup, so **exit criterion 1 is met**; and the Windows portable `.exe`,
which runs. Both launcher artifacts therefore build and run, and the
unsigned-build warning is documented as [known deviations](../../docs/known-deviations.md)
L-1, so **exit criterion 4 is met**. Recorded as **not tested for
this milestone**:

Verified by hand on 2026-11-21 (maintainer): the Windows portable `.exe`
started with `--port 45678` printed the separate-origin warning naming both
the chosen port and the default `47823`, then opened the app in Chrome at
that address. Also on 2026-11-21 (maintainer): **exit criterion 3 is met** —
the recovery copy was exercised by hand (edit, close the tab, restart,
Restore brings the edits back, a save clears the prompt), the external-edit
conflict offered Reload / Overwrite / Save As / Cancel on a file changed on
disk, and the close prompt appeared for a tab with unsaved changes.

Recorded as **not tested for this milestone**:

- A 250k-cell workbook scrolling at 60 fps in the browser (the demo item;
  the Node benchmark suite runs the corpus but the browser check was not
  watched).

Also recorded: **exit criterion 2 is met (2026-11-21).** The maintainer ruled
that this laptop — Intel Core i7-10610U, 4 physical cores, 15.7 GiB RAM,
Windows 11 — is the reference machine for v1, which settled the
reference-machine question (performance §1, ticket 36 item 2, research note 30
open decision 2). The budgets were then run in enforced mode, not advisory:
`sort-100k` and the conditional-formatting budgets pass, and the structural
rebuild of a 250k-cell, 50k-formula graph measured **1225.6 ms against the
1500 ms budget** after the optimisation in `9792fca` (down from the 1550-1671 ms
miss ticket 36 recorded, and from 3740 ms before lane 4). 6/6 enforced tests
green.

**Exit criteria**

1. Print and PDF match the preview and honour the page setup.
2. The performance budgets hold on the reference machine.
3. Recovery, conflict and close prompts work.
4. Launcher artifacts build and run on Windows and Linux; unsigned builds carry the documented SmartScreen note.

### M8 — Conformance and release (6-8 pw + reserve)

**Build**

- Full 52-function golden suite and generated cases.
- Error/coercion/operator, date, reference, lookup and circular-reference suites.
- Full file corpus, producer matrix, visual baselines and keyboard suites.
- Manual Excel repair check.
- Known-deviations list with a pinning test for each.
- PRD, technical design, milestone record and release notes.
- Release gate review.

**Demo:** a release candidate where every gate in [testing and conformance §13](testing-conformance.md) is green.

**Exit criteria**

1. The ticket 13 release gate passes in full.
2. No P0, P1 or data-loss bug remains.
3. Every known deviation is documented and pinned by a test.
4. The launcher artifacts and checksums are published.
5. The spec package is complete and internally consistent.

## 5. Dependency graph

M6 gates after M5: the file-fidelity exit criterion (all 24 file-compat fixtures, including F13 validation + conditional-formatting round-trip) needs M5’s CF engine. Fixture authoring may start early; the M6 exit gate cannot run before M5 exits. In the graph below M6 is drawn beside M2/M3 for staffing, not for gating.

```
M1 walking skeleton
 ├─ M2 grid foundation ──────────────┐
 ├─ M3 engine + model depth ─────────┤
 │    └─ range-transform prototype   │
 ├─ M4 formatting + workbook ────────┤
 │         └─ M5 data tools ─────────┤
 └─ M6 file fidelity ────────────────┤
              └─ M7 output + hardening
                        └─ M8 conformance + release
```

Parallel tracks:

- **Grid and rendering:** M1 minimal grid → M2 → M4 rendering → M5 CF and clipboard UI → M7 performance.
- **Model, engine and commands:** M1 model and commands → M3 engine → M4 styles and number formats → M5 sort/filter → M6 range transforms.
- **File I/O and chrome:** M1 package layer → M4 chrome → M6 full fidelity → M7 print and launcher.

Integration points are the end of every milestone. Tracks may overlap inside a milestone, but a milestone gate is not skipped. The model, command and file contracts are frozen before M4; changes after that need a spec ticket.

## 6. Milestone gates and the feature matrix

| Milestone | Feature-matrix sections |
|---|---|
| M1 | Workbook: new, open, save; grid basics |
| M2 | Grid and editing: selection, editing, formula bar, keyboard; view: freeze |
| M3 | Formulas and names; undo/redo; dates |
| M4 | Formatting; workbook and sheets; view; find and Go To |
| M5 | Data: sort and AutoFilter; conditional formatting; clipboard; fill; insert/delete |
| M6 | File I/O and read-only fidelity |
| M7 | Print/PDF; persistence; distribution; performance |
| M8 | The full release gate |

A milestone gate fails if any acceptance criterion in its section fails, even if the demo looks good.

## 7. Risk register

| Risk | Impact | Mitigation | Where |
|---|---|---|---|
| OOXML range transforms corrupt references or anchors | Data loss; broken charts/pivots | Prototype in M3; full transform in M6; block unsafe edits; fixtures | M3, M6 |
| HyperFormula shims or fan-out miss the contract | Wrong results or slow builds | Contract tests; 52-function gate; fan-out benchmark; worker | M3 |
| Number-format seams diverge from Excel | Visible formatting errors | `ssf` differential oracle; conformance corpus; built-in ID table | M4 |
| Clipboard interop with Excel is unreliable | Copy/paste loses formulas or formats | HTML-first dialect; TSV fallback; `.xlsx` is the fidelity channel | M5 |
| Print/PDF vector fidelity or fonts fail | Output does not match preview | Spike in M7; SVG preview; PDFKit; 300 DPI fallback | M7 |
| Performance at 250k/1M cells | Product misses its envelope | Continuous benchmarks; fan-out case; degradation ladder | M2-M7 |
| Unsigned Windows build | SmartScreen warns until reputation builds | Documented in the README and About dialog; signing stays optional | M7 |
| Conformance tail | Release slips | Continuous golden suite; M8 reserve; no late scope | M8 |

## 8. Definition of v1 done

v1 is done when **all** of these are true:

1. The [testing and conformance release gate](testing-conformance.md) passes in full.
2. Every feature-matrix **IN** item is present and meets its acceptance criterion.
3. The Windows and Linux launcher artifacts are built and published as [distribution](distribution.md) requires.
4. The PRD, technical design, ADRs and milestone record are complete and consistent.
5. Every known deviation is documented and has a pinning test.
6. No P0, P1 or data-loss bug remains.
7. The hosted static build and launcher checksums are published.

Store listings, auto-update and marketing are post-release.

## 9. Out of the milestones

- Everything marked V2 or NEVER in the feature matrix: charts, advanced data tools, tables, data validation authoring, CF authoring, comments/hyperlink authoring, split view, dark mode, localization, co-authoring, PivotTables, macros, Solver and the rest.
- Hosted deployment, store listings, auto-update and marketing.
- Mobile and touch.
- Native desktop wrappers.

## 10. Acceptance criteria

1. Eight milestones, each with a demo and a gate.
2. Milestone 1 is a walking skeleton with real OOXML, model, grid, edit and undo.
3. The dependency graph keeps the model, command and file contracts frozen before M4.
4. Every milestone's exit criteria are tied to the feature matrix.
5. The four risk spikes are scheduled in M3, M5 and M7.
6. Estimates are person-week ranges with a 20% contingency and no calendar dates.
7. The definition of v1 done includes the conformance gate, launcher artifacts and docs.
8. V2 and NEVER features are excluded from every milestone.
9. Each milestone produces a demoable state, not only internal plumbing.
10. The plan is build-ready: no milestone depends on an unresolved decision ticket.
