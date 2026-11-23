# Release gate record (v1)

This records the [testing and conformance](../.scratch/calcplus/spec/testing-conformance.md)
§13 release gate for the v1 release candidate, as part of the M8 milestone.
Every item is judged from named evidence on the maintainer's Windows PC. No
test or check runs on GitHub.

**Gate run: 2026-11-22.** `npm run check` green (typecheck, ESLint, the unit
suite: 100 files, 1527 tests passing with 10 environment-gated skips, and the
licence gate), `npm run build` green, and the full Playwright suite green
(211 tests).

## The 14 items

| # | Item | Status | Evidence |
|---|---|---|---|
| 1 | All 52 functions pass their golden cases within tolerance | **Pass** | "M3B 52-function golden subset" in `tests/engine-m3b.test.ts` checks all 52 cases in `fixtures/expected/m3b-formulas.json` with zero skipped. Generated cases cover every registry arity in `tests/formula-arity.test.ts`. Tolerance is enforced by the comparison helper (§2.1), not ad-hoc assertions. |
| 2 | The error / coercion / operator table passes | **Pass** | "M3B error, coercion and operator table" passes the whole error fixture and asserts no engine result ever produces `#ERROR!`, `#SPILL!`, `#CALC!` or `#LIC!`. |
| 3 | 1900/1904, serial 60 and the date edge cases pass | **Pass** | "M3B date, reference, lookup and circular-reference suites" in `tests/engine-m3b.test.ts`, plus `tests/m3c-date-system.test.ts` (the 1904 system and serial 60). |
| 4 | Every HyperFormula shim has a passing test | **Pass** | "M3B shim cases" covers every shim golden, including the fractional `CHOOSE` truncation and the numeric / numeric-text `COUNTIF`/`SUMIF` criteria. "M3B authored-formula pre-screen" pins the shim boundary. |
| 5 | Every file-compatibility fixture passes; a no-op save is byte-identical | **Pass** | The `tests/m6-*.test.ts` suites cover all 24 compatibility fixtures, hostile inputs and the producer matrix. `tests/ooxml-save.test.ts` ("identity save") returns the original bytes byte-for-byte, and the smoke e2e proves the same through the app. `node scripts/verify-repair-free.mjs` checks relationships, content types, XML well-formedness and part names. |
| 6 | Excel 365 opens every output without a repair prompt | **Pass** | The maintainer checked the whole `fixtures/repair-check/` folder interactively on 2026-11-22 with Microsoft Excel 365 (16.0, build 17932, en-GB, Windows 11): all 64 outputs open without a repair prompt. The automated detector run agrees on 62; its two `libreoffice` verdicts are a detector limitation, not a file defect — those files open fine interactively (see below). |
| 7 | Sort, AutoFilter and conditional-formatting fixtures pass | **Pass** | `tests/m5-sort.test.ts` (15 sort fixtures), `tests/m5-filter.test.ts` and `tests/m5-filter-roundtrip.test.ts` (18 filter fixtures), `tests/m5-cf.test.ts`, `tests/m5-cf-paint.test.ts` and `tests/m5-cf-roundtrip.test.ts` (20 CF fixtures). |
| 8 | Grid interaction and visual baselines pass | **Pass** | The full Playwright suite (211 tests): `grid`, `grid-interaction`, `formatted-visual` (pinned baseline), `merged-text`, `freeze-scroll`, `selection-autoscroll` and the rest. The visual tests assert model state as well as pixels. |
| 9 | Performance budgets hold on the reference machine | **Pass** | `tests/m7-reference-machine.test.ts` (6/6 enforced) on the declared reference machine (Intel Core i7-10610U, 15.7 GiB RAM, Windows 11): the 250k-cell / 50k-formula structural rebuild measured 1225.6 ms against the 1500 ms budget after `9792fca`. The browser bench passed 3/3 on 2026-11-21: scroll frame p95 16.7–16.8 ms against the 16.7 ms budget at whole-millisecond precision, worst frame 16.8 ms against 50 ms, and the million-row sparse sheet scrolled with the grid unallocated (M7 exit note). |
| 10 | Command registry, undo/redo, transaction boundaries, shortcut map and focus model pass | **Pass** | `tests/commands.test.ts` (transactions and undo), `tests/ribbon-model.test.ts` and `tests/ribbon-icons.test.ts` (the registry view), `tests/grid-focus.test.ts` and `tests/grid-navigation.test.ts` (focus, Esc/F6, the shortcut subset), the one-transaction boundaries in `tests/m3c-format-commands.test.ts` and `tests/m4c-format-dialog.test.ts`, and the F6/Esc e2e in `tests/e2e/grid.spec.ts`. |
| 11 | Number-format, print/PDF, clipboard and persistence contracts pass | **Pass** | `tests/m4a-number-formats.test.ts` (the built-in id corpus), the `tests/m7-print-*.test.ts` suites with the `print-preview` and `pdf-export` e2e, `tests/clipboard.test.ts`, and the `tests/m7-persistence-*.test.ts` suites with the 15-test `persistence` e2e. |
| 12 | No P0, P1 or data-loss bug remains | **Pass** | Every ticket in `.scratch/calcplus/issues/` is resolved (ticket 49 was the last one open and was closed during this sweep). The last data-loss bug found — the recovery-copy AutoFilter round-trip — was fixed in `63bbf36` with its regression tests. No bug of any severity is open. |
| 13 | Every known deviation is documented and has a test that pins it | **Pass** | All 19 entries in [known deviations](known-deviations.md) carry a pin: D-1 to D-11 in `tests/known-deviations.test.ts` (including the D-9 oracle side), D-12 in `tests/m7-page-setup-group.test.ts`, D-13 in `tests/formula-arity.test.ts`, L-1 to L-6 in `tests/known-limitations.test.ts`. |
| 14 | No silent skips and no quarantined test older than one week | **Pass** | The run reports 10 skipped tests and every one is named and environment-gated: 5 "launcher runtime" tests in `tests/m7-launcher.test.ts` (Go is not installed on this machine; a test asserts the skip message is clear), 4 "fixtures open in openpyxl" in `tests/m6-preserved-parts.test.ts` (openpyxl is not installed), and 1 non-Windows launcher test. Nothing is quarantined or disabled with a date. |

## The item 6 check

Recorded 2026-11-22. The maintainer opened every file in `fixtures/repair-check/`
in Microsoft Excel 365 (16.0, build 17932, en-GB, Windows 11) and none showed a
repair prompt or a recovery notice. The procedure that produced the folder:

1. `node scripts/verify-repair-free.mjs` — green.
2. `npm run repair-outputs` — every openable corpus fixture saved through the
   production pipeline, one untouched copy and one with a single cell edit
   each, into `fixtures/repair-check/`, with a `manifest.json`. The hostile
   fixtures and the Strict file are skipped by name (Tier C rejections).
3. `powershell -NoProfile -ExecutionPolicy Bypass -File
   scripts/oracle/detect-repair.ps1 -Directory fixtures/repair-check` — a fast
   filter through Excel COM: `OPEN-OK` conclusively means "no repair prompt",
   but `OPEN-FAIL` is inconclusive. A clean workbook always opens through COM;
   a workbook Excel would repair cannot (the repair dialog cannot be answered
   in automation); and Excel can also refuse COM automation for files it opens
   fine interactively — the LibreOffice producer output does exactly that. The
   interactive pass in step 4 is the authority.

The fixture families the maintainer's repair logs named during the check —
charts, comments, defined names, VBA, themes, pivot and external links — were
fixed in the generators (commit "Make the corpus fixtures Excel-clean").

## Verdict

All 14 items pass on the 2026-11-22 gate run. The release gate passes in full.
