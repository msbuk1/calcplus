# 13 — Testing and Excel-conformance strategy

Status: resolved
Type: grilling
Blocked by: 02, 03, 04, 18, 19, 20, 21, 22, 24

## Question

How do we prove calcplus is faithful?

Define:

- Formula conformance: golden workbooks, generated cases, comparison against real Excel, tolerance rules
- `.xlsx` round-trip fixtures and diffing strategy
- Grid UI/interaction regression testing
- Performance benchmarks in CI
- Unit/integration boundaries around engine, model, and view
- The minimum bar before v1 can be called "faithful", and what is explicitly not tested

Depends on the semantics contract (02), the engine choice (03), and the I/O approach (04).

## Notes

- Include a 1900/1904 fixture pair (date system, fake 29 February 1900) and conformance tests for the HyperFormula shims listed in [03 — Formula engine](03-formula-engine.md): wildcard lookup, `INT`, `CHOOSE`, and `#CYCLE!` → `#CIRC!`.
- Include tests for the command registry, undo/redo transaction boundaries, the shortcut map and the focus model in [command-system](spec/command-system.md).

## Answer

The strategy is in [testing-conformance.md](../spec/testing-conformance.md).

**Oracle:** real Excel 365 on a Windows runner is authoritative; a COM/PowerShell harness generates goldens, and pull requests run against stored goldens. LibreOffice is a differential smoke test, and `ssf` is a differential oracle for number formats only. Every golden manifest records Excel version, OS, locale and date system.

**Tolerance:** exact for text, booleans, errors, blanks, formula text and number-format display strings; numbers within `max(1e-12, 1e-9 * max(|a|, |b|))`; dates exact, time fractions within 1e-9 day.

**Layers:** Vitest unit and integration, contract tests for the `FormulaEngine` port, Playwright end-to-end and visual, fast-check invariants, axe-core, openpyxl/LibreOffice for independent file verification, Python for the oracle. Fakes only at the clock, file, clipboard and Excel boundaries. Coverage 90/85 on pure modules; conformance is the real gate.

**Formula:** one golden case per function plus edge cases, the error/coercion/operator table, 1900/1904 and serial 60, references, lookups, circular refs, unsupported functions, volatile functions, legacy arrays, and every HyperFormula shim (`INT`, `CHOOSE`, wildcard lookup, `#CYCLE!` → `#CIRC!`).

**Files:** every file-compat fixture, no-op identity, repair-free, part inventory, semantic equality, formula fidelity, structural edits, hostile inputs and the producer matrix.

**Grid/UI:** interaction suite, command/undo/shortcut/focus suite, and visual baselines on a pinned Linux Chromium runner with a pinned font set, fixed DPR and a 0.1% pixel threshold.

**CI:** pull requests run typecheck, lint, unit/integration, a small golden set, identity, smoke, axe and changed-area visuals under 15 minutes. Nightly runs the full formula/file/visual/performance/accessibility suite. The Excel oracle regenerates goldens on a schedule. The release gate requires all functions, files, UI, accessibility and performance to pass, with no P0/P1 or data-loss bugs.

**Not tested:** pixel-perfect parity, the full function set, dynamic arrays, R1C1, 3D/structured refs, external-link refresh, iterative/manual calc, Mac Excel/Excel Online/mobile, VBA, pivot refresh, chart rendering, encrypted/.xls/.xlsb/ODS/Strict OOXML, co-authoring, non-English locales, versions other than 365/2019, beyond-envelope workbooks, mobile AT, tagged PDF and user-authored formatting contrast.
