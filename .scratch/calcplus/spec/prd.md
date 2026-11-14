# calcplus v1 — product requirements document

Assembled by [15 — Assemble the build-ready spec package](../issues/15-spec-assembly.md) on 2026-11-13. The scope is fixed by the [feature matrix](feature-matrix.md); the technical design is in [technical-design.md](technical-design.md); the plan is in [milestones.md](milestones.md).

## 1. Product summary

**calcplus** is a browser-based, single-user clone of Excel desktop (Microsoft 365). It is a core credible spreadsheet: grid editing, formatting, formulas, multi-sheet workbooks and high-fidelity `.xlsx` round-trip, at functional, near-identical UI fidelity.

It runs as a static web app and as a bundled per-OS launcher that serves the same build from `127.0.0.1`. There is no backend, no account and no telemetry.

## 2. Problem and target user

People who work in Excel desktop want a spreadsheet they can open in a browser without uploading the workbook to a cloud service. Existing web spreadsheets either require an account, lose formatting on import, or replace the familiar Excel interaction model.

**Primary user:** an individual knowledge worker, analyst or small-team operator who already uses Excel and has `.xlsx` files. They want to open, edit, format, calculate, sort, filter, print and save those files with the same muscle memory, on a machine where installing a desktop app is inconvenient or unwanted.

**Jobs to be done**

- Open an existing `.xlsx` and see it faithfully.
- Edit values and formulas, format cells, and use the keyboard as in Excel.
- Sort and filter a plain range.
- Save back to the same file without losing anything calcplus does not understand.
- Print or export a clean PDF.
- Work offline and keep the workbook on the machine.

## 3. Goals

1. **Familiarity.** The ribbon, formula bar, grid, sheet tabs, dialogs and shortcuts feel like Excel desktop. When Excel and "clean design" disagree, match Excel.
2. **No silent loss.** A workbook that calcplus cannot fully model is preserved, reported and never silently dropped.
3. **Credible performance.** A 250k-cell workbook is guaranteed; up to 1M cells opens in a degraded mode; the app refuses rather than corrupts beyond that.
4. **Accessibility.** WCAG 2.2 AA for the chrome and the DOM grid mirror, with real screen-reader sign-off.
5. **Local-first.** No accounts, no uploads, no telemetry; the file stays on the machine.
6. **Open source.** GPL-3.0-only, with a build that anyone can reproduce.

## 4. Fidelity promise

- **Reference:** Excel desktop / Microsoft 365.
- **Near-identical, not pixel-perfect.** Chrome layout, colours, labels and interactions follow Excel; font and OS rendering differences are accepted.
- **Published formula contract:** the 52 functions in [formula semantics](formula-semantics.md), with Excel-exact errors, coercion and operator order. Full Excel function parity is not claimed.
- **Published file contract:** the three tiers in [file compatibility](file-compat.md). Modelled features round-trip; preserved features survive byte-for-byte; unsupported files are rejected with a clear message.
- **Published performance contract:** the budgets in [performance](performance.md).
- **Published accessibility contract:** the criteria in [accessibility](accessibility.md).

## 5. V1 scope

The authoritative list is the [feature matrix](feature-matrix.md). In summary:

| Area | In v1 |
|---|---|
| Grid and editing | Selection and keyboard navigation, in-cell and formula-bar editing, cut/copy/paste, fill handle, drag-and-drop move, insert/delete, resize/autofit, merges, undo/redo, find/replace, Go To, clear |
| Data | Sort a plain range; AutoFilter with value checklist and text/number/date conditions |
| Formatting | Fonts, fills, borders, alignment, number formats, Format Painter, Format Cells, render-only conditional formatting |
| Formulas and names | 52 functions, A1/mixed/cross-sheet/whole-column/names, 1900/1904 dates, no spill, read-only legacy arrays, volatile TODAY/NOW, `#CIRC!` |
| Workbook and sheets | New/open/save/save as/close, insert/delete/rename sheets, freeze panes, zoom, gridline/heading toggles, status-bar statistics |
| File I/O | `.xlsx` and `.xlsm` round-trip, CSV, encrypted rejection, external-link and macro preservation, drag-and-drop open, print/PDF |
| Read-only fidelity | Comments, hyperlinks, validation, tables, pivots, protection, hidden rows/columns/sheets, themes and named styles, defined names |
| Platform | Desktop Chromium contract; Firefox/Safari best-effort; static host and Go launcher; no backend |

## 6. Key user journeys

1. **Open and edit.** Open a sales workbook, navigate with the keyboard, edit cells, enter formulas, undo, and save in place.
2. **Format a report.** Apply fonts, fills, borders, currency and date formats, merge a header, resize columns, add a sheet, and print/PDF it.
3. **Analyse a list.** Sort a 100k-row range by two keys, filter by a value checklist and a text condition, and clear the filter.
4. **Open a rich workbook.** Open a file with a chart, a pivot, an external link and a macro; edit cells and insert a row; save; Excel opens it without repair; the reports explain what was preserved.
5. **Recover.** Close the tab mid-edit; reopen; restore the unsaved changes from the recovery copy.
6. **Work offline.** Run the launcher, open a local file, edit and save with no network.

## 7. Non-goals

Everything marked **V2** or **NEVER** in the feature matrix. Highlights:

- Real-time co-authoring, accounts, cloud sync and sharing.
- VBA/macro execution, Office Scripts and add-ins.
- PivotTables, Power Query, Solver/Goal Seek, the Data Analysis ToolPak and Consolidate.
- Charts, images, shapes and sparklines (parts are preserved, not rendered).
- Authoring or editing of tables, data validation, advanced/colour/Top-10 filters and conditional formatting; imported parts are preserved.
- Authoring of comments, hyperlinks, defined names, protection, themes and named styles.
- Paste Special, fill series, split view, full page setup, localization and dark mode.
- R1C1, 3D ranges, structured references, external-link refresh, dynamic arrays/spill, iterative/manual calculation and union/intersection operators.
- Native desktop wrappers, mobile/touch apps, `.xls`, `.xlsb` and ODS; encrypted workbooks are rejected with a clear message rather than opened.
- Store listings, auto-update and marketing as part of the build; the hosted static build is a release artifact.

## 8. Constraints and assumptions

- **Browser app.** Static files, no backend, no server-side logic.
- **Single user, one workbook per tab.** The model is sync-ready but co-authoring is not built.
- **GPL-3.0-only.** HyperFormula's GPLv3 edition is the binding dependency.
- **Desktop Chromium is the performance contract.** Firefox and Safari work with the download fallback but carry no performance claim; mobile and touch are out.
- **Distribution.** Static host plus a Go launcher on `127.0.0.1:47823`; Windows signed `.exe`, macOS notarized universal `.app`, Linux tarball.
- **Local-first.** No accounts, telemetry or workbook uploads.
- **English only.** Localization is v2.
- **No pixel-perfect requirement.** Near-identical chrome is enough.

## 9. Product acceptance criteria

1. A user can open an `.xlsx`, edit it, and save it back with modelled features intact and preserved parts byte-identical.
2. Excel 365 opens every output without a repair prompt.
3. The 52 published functions, error/coercion table and date systems pass the conformance suite.
4. Sort, AutoFilter and conditional formatting behave as specified and round-trip.
5. The grid supports the keyboard, editing, fill, resize, freeze, merges and selection journeys.
6. A 250k-cell workbook meets the performance budgets; 1M cells opens in degraded mode.
7. The accessibility contract passes, including NVDA/Chrome and VoiceOver/Safari sign-off.
8. The launcher runs offline on all three OSes and keeps a stable origin.
9. No P0, P1 or data-loss bug remains at release.
10. Every known deviation is documented and pinned by a test.

## 10. Release definition

v1 is done when the [release gate](testing-conformance.md) passes in full, every feature-matrix IN item is present, the launcher artifacts are signed or notarized, the spec package is complete, manual AT sign-off is recorded, and no P0/P1 or data-loss bug remains. The [milestone plan](milestones.md) carries the phase-by-phase gates.

## 11. Open questions

These are deliberately outside v1 and must not block the build:

- Localization of the UI and of function names.
- Formula-bar and name-box edge behaviours.
- Grid interaction edge cases not fixed by the matrix: fill-handle options, merged-range selection, complex bidi text layout and selection scrolling. IME input itself is in scope (M2’s DOM in-cell editor takes native composition); only bidi shaping and paragraph layout are deferred.

## 12. Risks

- **OOXML fidelity tail.** The range-transform and preservation work is the largest engineering risk; the walking skeleton and the M3 prototype are the mitigations.
- **Accessibility mirror.** A canvas grid is unproven with screen readers; the AT spike and DOM-cell fallback are the mitigations.
- **Formula shims.** HyperFormula deviations must be pinned; the contract tests and 52-function gate are the mitigation.
- **Performance at 1M cells.** The degraded envelope is close to the memory ceiling; the benchmark corpus and degradation ladder are the mitigation.
- **Conformance tail.** The last 10% of Excel edge cases can take as long as the first 90%; the M8 reserve and continuous golden suite are the mitigation.
- **Signing and notarization.** External dependencies can delay release; start early in M7.

## 13. Success measures

- **Conformance:** the release gate is green; no undocumented deviation.
- **Data safety:** zero data-loss bugs; no silent drops.
- **Performance:** all budgets on the reference machine.
- **Accessibility:** axe clean; AT sign-off.
- **Fidelity:** the core journeys complete without Excel.
- **Trust:** no workbook leaves the machine; no telemetry.

## 14. Dependencies

HyperFormula (GPL-3.0-only), React, `numfmt`, `fflate`, `saxes`, PDFKit, the Go standard library and the OFL fonts. The audit and obligations are in [licensing](licensing.md) and [THIRD_PARTY_NOTICES.md](../../../THIRD_PARTY_NOTICES.md).
