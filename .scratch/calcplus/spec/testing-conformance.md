# calcplus v1 — testing and Excel conformance

Resolved by [13 — Testing and Excel-conformance strategy](../issues/13-testing-conformance.md) on 2026-11-13. It ties together the semantics contracts in [formula semantics](formula-semantics.md), [file compatibility](file-compat.md), [sort and AutoFilter](sort-autofilter.md), [conditional formatting](conditional-formatting.md), [performance](performance.md) and [accessibility](accessibility.md).

## 1. What "faithful" means

**v1 is faithful when the full release gate in §13 passes.** Not a percentage, not a manual impression: every published function, every compatibility fixture, the grid interactions, accessibility and the performance budgets must pass, and every known deviation must be documented.

The claim is bounded by §12. calcplus does not claim full Excel parity.

## 2. The oracle

- **Real Excel 365 on the maintainer's local Windows PC is the authoritative oracle.** The COM/PowerShell harness opens generated workbooks, recalculates, and captures values, errors, number-format display text, formula text and saved files. The maintainer runs it manually.
- **No test or check runs on GitHub.** The local suite compares against **stored goldens**; GitHub is used only for source hosting and release downloads.
- Goldens regenerate manually when Excel updates or when a new case is added. Every golden manifest records the Excel version and build, OS, locale, date system and the date generated.
- **LibreOffice headless is a differential smoke test**, never the authority. It catches packaging and gross-semantic breaks cheaply.
- **`ssf` is a differential oracle for number formats only.** It never decides a formula result.
- Goldens cover sort collation, conditional-formatting paint and print pagination as well as values and files: the local Excel harness sorts fixture sheets, captures Excel-rendered CF ranges as images, and prints fixtures to PDF for page-break goldens. Every behavioural claim in §1 has a named oracle or contract-test path; accessibility and performance are contract tests with local harnesses, not Excel claims, and a claim without either cannot enter the release gate.
- **Interaction fidelity** is verified by the §6 interaction suite against the feature-matrix behaviour list (shortcuts, selection, editing semantics), with a manual Excel side-by-side checklist per release; there is no pixel oracle for behaviour.
- The oracle is release-gate infrastructure: the maintainer's licensed Excel 365 on their Windows PC, exercised manually in M1. The capture is local and manual; no hosted test runner is used.
- The oracle is a tool, not a dependency of the shipped app.

### 2.1 Tolerance

| Value kind | Comparison |
|---|---|
| Text | Exact, case-sensitive, no normalisation beyond the stored text |
| Boolean | Exact |
| Error | Exact, including `#N/A`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`, `#CIRC!` |
| Blank | Exact |
| Number | `abs(a - b) <= max(1e-12, 1e-9 * max(abs(a), abs(b)))` |
| Date / time serial | Exact serial; time fractions within `1e-9` of a day |
| Number-format display | Exact string |
| Formula text | Exact authored text, never normalised |
| Cached result | Same rules as the corresponding value kind |

The tolerance reflects Excel's 15-significant-digit arithmetic. It must not be widened to hide a real error.

## 3. Test layers and boundaries

| Layer | Tool | What it covers | Boundaries |
|---|---|---|---|
| Unit | Vitest | Pure modules: formula parser, coercion, comparator, number format, range transforms, CF evaluation, clipboard dialect, PrintPlan, filter predicates | No DOM, no I/O |
| Contract | Vitest | `FormulaEngine` port; HyperFormula adapter; engine configuration; the 52-function gate; shims; error mapping | Real engine |
| Integration | Vitest | Model + commands + engine + OOXML + `SheetView` | Fakes only for clock, file system, clipboard and Excel |
| Conformance | Vitest + Python | Golden formulas and files; independent readers | Stored goldens, no live Excel |
| End-to-end | Playwright | Real app, real Chromium, user journeys, keyboard, clipboard, file open/save | Real app; file picker shimmed |
| Visual | Playwright | Grid and chrome screenshots | Pinned local Chromium and fonts |
| Performance | Node + Playwright/CDP | Microbenchmarks and browser budgets | Fixed hardware class |
| Security | Vitest | Hostile packages, malformed XML, zip bombs, XXE | Time and size caps |

Rules:

- **Never mock the thing under test.** Fakes are allowed only at the clock, file system, clipboard and Excel boundaries.
- Every test is deterministic: frozen clock, fixed seed, fixed locale and timezone, fixed viewport.
- Integration tests use the real model, engine and OOXML layer together; a passing unit test does not excuse a broken integration.
- Contract tests pin the `FormulaEngine` port so HyperFormula can be upgraded or replaced without touching the model.

Coverage targets: **90% lines / 85% branches on pure modules**. Coverage is not a release gate and is not chased on UI code; conformance is the gate.

## 4. Formula conformance

### 4.1 Published contract

- One golden case per published function, plus edge cases: wrong arity, error arguments, blanks, text coercion, booleans, dates, wildcards and boundaries.
- The **error / coercion / operator table**: propagation, numeric-looking text, blank as 0, boolean arithmetic, type ordering, precedence, concatenation and comparison.
- **References:** A1, absolute, mixed, cross-sheet, quoted sheet names, whole-column, whole-row and defined names.
- **Dates:** 1900 and 1904 workbooks, serial 60 (the fake 29 February 1900), pre-1900-03-01, time fractions.
- **Lookups:** exact and approximate `VLOOKUP`/`HLOOKUP`, `MATCH` types, wildcards, text criteria, no-match.
- **Circular references:** detection, warning and `#CIRC!`.
- **Unsupported functions:** `#NAME?` with no cached value; cached value marked stale with one.
- **Volatile functions:** `TODAY`/`NOW` with a frozen clock.
- **Legacy arrays:** read-only, preserved text and cached result.
- **Engine shims:** exact-mode wildcard lookup returning the first match, `INT` flooring toward negative infinity, `CHOOSE` out-of-range → `#VALUE!`, `#CYCLE!` → `#CIRC!`, and no `#ERROR!`, `#SPILL!`, `#CALC!` or `#LIC!` reaching the model or file.
- **Authored text:** `OFFSET`, `_xlfn.` and structured references are preserved as authored text plus cached value; the engine's rewritten text is never saved.

### 4.2 Generated cases

- A seeded grammar generator produces formulas from the 52 functions, the supported reference forms and the supported literals.
- Generated cases are run through the Excel oracle, then stored as goldens with the same manifest.
- The generator is committed; the seed and case count are recorded. A new batch is reviewed before it becomes a gate.
- Property tests check invariants (no exceptions, deterministic output, error propagation) without an Excel oracle in the loop.

## 5. File conformance

Every fixture in [file compatibility §9](file-compat.md) is a test:

- **Identity:** open → save with no edits is byte-identical.
- **Repair-free:** every output opens in Excel 365 without a repair prompt. Local scripts use openpyxl, LibreOffice and a relationship verifier where installed; the Excel check is a manual local release step.
- **Part inventory:** for an edited file, every preserved part has identical uncompressed bytes and resolvable relationships. Compare a manifest of name, CRC, uncompressed size and SHA-256.
- **Dirty-part preservation:** a rewritten worksheet keeps every unmodelled child element, attribute, namespace declaration and its order, and every untouched row is byte-identical. The fixture asserts on the XML tree for the part as a whole and on bytes for the untouched rows.
- **Semantic equality:** styles, number formats, merges, dimensions, names, comments, hyperlinks, validation and conditional formats compare equal through an independent reader.
- **Formula fidelity:** authored text, shared/array attributes and cached values survive an edit to another cell.
- **Structural edits:** after insert/delete, every range-bearing construct points at the intended range; chart series ranges are rewritten, and pivot sources are reported stale or blocked.
- **Hostile inputs:** encrypted, truncated, non-ZIP, zip-bomb, XXE-style and oversized files fail safely within the time and size caps.
- **Producer matrix:** Excel 365, Excel 2019, LibreOffice, Google Sheets, openpyxl, SheetJS and xlsx-populate outputs.
- **Macros, external links, pivots, charts, threaded comments, x14 conditional formatting:** preserved and reported as specified.

Large files are generated on demand; only small fixtures live in git.

## 6. Grid, UI and interaction

- **Interaction suite:** selection and keyboard navigation, in-cell editing, formula bar, fill handle and Ctrl+D/R, row/column resize, freeze panes, merges, context menu, clipboard, zoom, sheet tabs, sort, AutoFilter, conditional formatting, undo/redo and the focus model.
- **Command suite:** every surface dispatches the same command ID; enablement and checked state follow the registry; one transaction per action; the shortcut map matches the registry.
- **Visual suite:** the grid at 100% and 200% zoom, with and without CF, freeze panes, merges, selection, editing, dialogs, the compact bar and forced colours.
- **Determinism:** a fixed Chromium version on the maintainer's PC with a pinned font set (Selawik, Carlito, Liberation Sans), fixed viewport and DPR 1 and 2, frozen clock, fixed locale and timezone, and animations disabled.
- **Pixel comparison:** about 0.1% of pixels may differ, with a per-channel tolerance of 2/255. Baselines are tied to the local Chromium version and machine and are updated deliberately.
- **Model assertions:** interaction tests also assert model snapshots and DOM state, so a screenshot that happens to look right cannot hide a wrong command.

## 7. Performance

The [performance contract](performance.md) is the source:

- Per change: Node microbenchmarks block on a 20% regression against the rolling median of the last 10 local runs on the same machine.
- Full local run: browser end-to-end on the fixed hardware class enforces the absolute budgets.
- Release: the full corpus at 1, 10 and 50 MB.
- Every result records a machine fingerprint; non-reference machines are advisory.

## 8. Accessibility

Assistive-technology support is **v2**. v1 has no accessibility layer, no axe job and no AT gate; keyboard operation and visible focus are covered by the interaction suite in §6. The [accessibility contract](accessibility.md) is retained for v2 and is not a v1 gate.

## 9. Supporting contracts

- **Number formats:** the conformance corpus from ticket 18, with `ssf` as a differential oracle; built-in IDs, 1904, width layout, colours and locale seams.
- **Print/PDF:** `PrintPlan` golden files, page count, page boxes, vector output and text extraction; preview and PDF must match.
- **Clipboard:** the HTML Excel dialect and TSV fallback, in-app round-trips and Excel interop where the OS allows.
- **Persistence:** autosave recovery, external-edit conflict, close prompt and origin-stable storage.
- **Sort/filter/CF:** the fixtures in [sort and AutoFilter §4](sort-autofilter.md) and [conditional formatting §9](conditional-formatting.md).

## 10. Property-based and fuzz testing

**fast-check invariants**

- Range transforms: applying an insert then the inverse delete restores the original model.
- Sort: the result is a permutation that preserves the multiset of rows and is stable on ties.
- Filter: every data row is either visible or filter-hidden; the header is always visible; clearing restores the manual layer.
- Undo/redo: applying a command then undo returns the exact prior snapshot; redo returns the post-state.
- Package: no-op save is byte-identical; an edited save changes only dirty parts.
- Part patch: a single-cell edit leaves every untouched row byte-identical and preserves the order of unknown elements.
- Number format: never throws, is deterministic, and returns a string.
- CF: data-bar lengths and colour-scale positions stay within bounds; icon thresholds are monotonic.
- Comparator: antisymmetry, transitivity and blank-last.

**Fuzzing**

- Hostile files: truncated ZIPs, bad central directories, zip bombs, oversized parts, malformed XML, DTD/XXE payloads, path traversal in entry names, huge shared strings.
- Every fuzz case must fail safely within the package caps and never execute external content.
- Fuzzing is an on-demand local job with a fixed corpus and seed; new crashes become regression fixtures.

## 11. Local test cadence and goldens

There is no GitHub Actions test workflow. Checks run on the maintainer's PC; GitHub hosts the source and releases only.

| Stage | What runs | Target |
|---|---|---|
| Per change | Typecheck, lint, format check, unit, integration, contract tests, a small formula golden set, round-trip identity, Playwright smoke, visual diff for changed areas | Under 15 minutes |
| Full local run | Full 52-function suite, generated cases, full file corpus and producer matrix, LibreOffice smoke, full visual suite, performance, accessibility | Under 60 minutes |
| Manual oracle | Excel 365 on the maintainer's PC regenerates formula and file goldens; manifest review | On Excel update or new case |
| Release | Full matrix, manual Excel repair check, performance on the reference machine, no P0/P1 or data-loss bugs | Blocks the release |

**Golden management**

- Small fixtures live in git. Large files are generated by a committed script.
- Every golden manifest records Excel version and build, OS, locale, date system and generation date.
- Golden changes are reviewed like code; a diff that only updates expectations is flagged.
- No secrets, personal data or unlicensed workbooks enter the corpus. Real fixtures are anonymised and cleared for test use.
- The `conformance` label marks every deviation and golden update.

## 12. What is not tested

Outside the v1 claim:

- Pixel-perfect Excel parity; fonts, DPI and OS rendering differ.
- The full Excel function set; only the 52 published functions are promised.
- Dynamic arrays and spill, R1C1 notation, 3D references, structured table references, external-link refresh, iterative calculation and manual calculation.
- Mac Excel, Excel Online and mobile apps.
- VBA execution, pivot refresh and chart rendering.
- Encrypted workbooks, `.xls`, `.xlsb`, ODS and Strict OOXML.
- Real-time co-authoring.
- Locales other than English.
- Excel versions other than 365 as the oracle and 2019 as a compatibility fixture.
- Workbooks beyond the performance envelope.
- Mobile screen readers.
- Assistive-technology support and WCAG conformance (v2).
- macOS and Safari.
- Tagged PDF.
- The contrast or colour-only meaning of user-authored formatting.
- Hardware outside the reference and floor tiers.
- Browsers other than Chromium for the performance contract.

## 13. Release gate

v1 is faithful only when all of these pass:

1. All 52 functions pass their golden cases within tolerance.
2. The error / coercion / operator table passes.
3. 1900/1904, serial 60 and the date edge cases pass.
4. Every HyperFormula shim has a passing test.
5. Every file-compatibility fixture passes; a no-op save is byte-identical.
6. Excel 365 opens every output without a repair prompt.
7. Sort, AutoFilter and conditional-formatting fixtures pass.
8. Grid interaction and visual baselines pass.
9. Performance budgets hold on the reference machine.
10. Command registry, undo/redo, transaction boundaries, shortcut map and focus model pass.
11. Number-format, print/PDF, clipboard and persistence contracts pass.
12. No P0, P1 or data-loss bug remains.
13. Every known deviation is documented and has a test that pins it.
14. No silent skips and no quarantined test older than one week.

## 14. Acceptance criteria

1. Excel 365 on the maintainer's local Windows PC is the authoritative oracle; the local suite uses stored goldens.
2. Goldens carry a manifest with Excel version, OS, locale and date system, and are reviewed.
3. The tolerance rules in §2.1 are enforced by the comparison helper, not by ad-hoc assertions.
4. The 52-function suite, error/coercion/operator table, date, reference, lookup and circular-reference tests exist and pass.
5. The shim tests exist and pass.
6. File fixtures cover identity, repair-free, part inventory, semantic equality, formula fidelity, structural edits, hostile inputs and the producer matrix.
7. Sort, filter and CF fixtures pass.
8. Grid interaction and visual tests run locally and assert model state as well as pixels.
9. The performance gate follows its contract.
10. Property and fuzz tests run locally on demand with fixed seeds.
11. The local cadence in §11 is implemented; per-change checks stay under 15 minutes.
12. The release gate in §13 blocks a release.
13. The non-goals in §12 are documented in the PRD.
14. No test mocks the model, engine, OOXML layer or app under test.
15. Coverage meets 90/85 on pure modules, and no coverage target is used to excuse a conformance gap.

## 15. Deliberate limitations

- The oracle is Excel 365, not every Excel version.
- LibreOffice and `ssf` are differential aids, not authorities.
- Visual tests cover one Chromium/font/OS combination; other platforms are manually checked.
- Property tests prove invariants, not Excel semantics.
- Fuzzing covers hostile packages, not formula semantics.
- The suite proves the published contract; it does not prove full Excel parity.
