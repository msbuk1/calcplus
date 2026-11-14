# 01 — M1 walking skeleton status

Status: open
Type: task
Blocked by: none locally; external maintainer gates below

## Owner roles

- **core-coder** — implementation of the M1 slice.
- **quality-reviewer** — independent code review and local verification.
- **documenter** — this record and spec-package consistency.

## Code review

- Reviewer verdict: **PASS** at attempt **3/3**, with no P0/P1 findings.
- Remaining reviewer P2 items are deferred follow-ups, not blockers and not new scope.

## Milestone gate: M1 is NOT ACCEPTED

Exit criteria (per [milestones — M1](../../calcplus/spec/milestones.md)):

1. Model opens a minimal `.xlsx` and renders it — locally pass.
2. No-op open → save is byte-identical — locally pass.
3. A cell edit is one undoable command; undo restores the prior value — locally pass.
4. An edited save keeps chart, `<legacyDrawing>` and `<extLst>` children and their order, proven by part inventory — locally pass.
5. The app opens in Chromium and saves through the file API or a download — locally pass.
6. Unit and smoke tests run in CI — **not run**: no CI run against a DCO-signed maintainer commit/push.

Additional demo item: **New blank workbook — locally pass.**

Local verification summary (local checks, not CI):

- Vitest 83/83 and Playwright 6/6 pass locally.
- typecheck, lint, format, build and licence checks pass locally.
- Browser memory harness exists but was not executed.
- CI has not run.

## Unprovisioned external build item

- The self-hosted Excel 365 oracle runner is **unprovisioned**: the runner label, Excel licence, golden capture and `CALCPLUS_ORACLE_ENABLED` gate are not confirmed or configured; the oracle workflow stays skipped.

## Remaining maintainer decisions

1. Make a DCO-signed maintainer commit/push so CI runs and exit criterion 6 can be evaluated (DCO sign-off per [CONTRIBUTING](../../../CONTRIBUTING.md)).
2. Decide whether and how to provision the self-hosted Windows oracle runner with an activated Excel 365 licence.

## Position

M1 remains **pending external gates**. This record does not state that M1 is complete or accepted; no evidence beyond the local checks above is claimed.
