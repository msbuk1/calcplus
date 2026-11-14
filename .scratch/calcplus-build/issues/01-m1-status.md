# 01 — M1 walking skeleton status

Status: accepted
Type: task
Blocked by: none

## Owner roles

- **core-coder** — implementation of the M1 slice.
- **quality-reviewer** — independent code review and local verification.
- **documenter** — this record and spec-package consistency.

## Code review

- M1 slice reviewer verdict: **PASS** at attempt **3/3**, with no P0/P1 findings.
- Oracle-harness fix reviewer verdict: **PASS** at attempt **2/3**, with no P0/P1 findings.
- Local-only testing change reviewer verdict: **PASS** at attempt **1/3**, with no P0/P1 findings after the workflow deletion was reviewed.
- Final Milestone 1 gate review: **PASS** on 2026-11-14, with all six exit criteria passing on local evidence.
- Remaining reviewer P2 items are deferred follow-ups, not blockers and not new scope.

## Testing plan

On **2026-11-14** the user approved a testing-plan change: **no test or check
runs on GitHub**. Both GitHub Actions workflows are deleted. Tests, fixtures and
stored goldens stay in the repository and run locally on the maintainer's PC.
The Excel 365 oracle is captured manually with `scripts/oracle/capture-goldens.ps1`.
The change is committed with DCO sign-off.

## Milestone gate: M1 is ACCEPTED

Exit criteria (per [milestones — M1](../../calcplus/spec/milestones.md)):

1. Model opens a minimal `.xlsx` and renders it — **PASS (local)**.
2. No-op open → save is byte-identical — **PASS (local)**.
3. A cell edit is one undoable command; undo restores the prior value — **PASS (local)**.
4. An edited save keeps chart, `<legacyDrawing>` and `<extLst>` children and their order, proven by part inventory — **PASS (local)**.
5. The app opens in Chromium and saves through the file API or a download — **PASS (local)**.
6. Unit and smoke tests pass on the maintainer's PC; no GitHub test execution is required — **PASS (local)**.

Additional demo item: **New blank workbook — locally pass.**

Exit criterion 6 evidence (local):

- Typecheck, lint, format check, unit/integration tests, production build,
  Chromium end-to-end smoke and the licence gate pass locally.
- Unit/integration suite: **85/85** tests pass.
- Chromium end-to-end smoke: **6/6** tests pass.
- GitHub hosts the source and release downloads only; no test or check executes
  on GitHub.

The earlier hosted CI run `43637929923` for commit `5b5230a` is retained only as
history. It is not the gate evidence and does not cover the local-only testing
change.

## Final gate review

Verdict: **PASS** on 2026-11-14.

- Criterion 1: the unit open tests map the minimal workbook and the Chromium
  smoke test renders its values.
- Criterion 2: the unit identity-save test and the browser no-op save test both
  compare the saved bytes with the original.
- Criterion 3: the command tests prove one transaction and exact undo; the
  browser test edits, commits and undoes.
- Criterion 4: the dirty-part test checks child order, the exact `<extLst>`,
  untouched row bytes and the part inventory; the browser test compares the
  chart bytes after an edited save.
- Criterion 5: the Chromium smoke test saves through the download fallback; the
  maintainer also exercised load and save manually.
- Criterion 6: typecheck, lint, format check, 85/85 unit tests, production
  build, 6/6 browser tests and the licence gate pass locally. No test or check
  runs on GitHub.

The gate was reviewed at code commit `13d0dcc6f04350b1cf328a26cc9fbce1466ac8b1`.
This acceptance record is a follow-up commit.

## Oracle harness — local capture evidence

The oracle capture was run manually on this PC:

- Windows PowerShell 5.1 with process-only `RemoteSigned` ran the capture cleanly.
- Excel **16.0, build 17932**; fixture `fixtures/oracle/m1-values-errors.xlsx`.
- Script exit code **0**; JSON is readable, with no BOM.
- Exact error names and the formula-versus-text distinctions are present.
- The fixture file was unchanged by the run, and the Excel process returned to
  its pre-run baseline.
- No self-hosted runner, scheduled GitHub job or `CALCPLUS_ORACLE_ENABLED`
  variable is used.

## Completion record

- The code, oracle harness, fixture, tests and documentation changes are
  committed with DCO sign-off.
- This acceptance record is committed with the same sign-off.
- Push when the maintainer is ready.

## Position

Milestone 1 is **ACCEPTED** as of 2026-11-14. All six exit criteria pass on
local evidence, and criterion 6 now means local tests, not GitHub. The
local-only testing and oracle-harness changes are committed. Work can move to
the next milestone.
