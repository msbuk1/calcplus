# 23 — Project licence: GPLv3 and dependency compliance

Status: open
Type: task
Blocked by: none

## Question

Make the GPLv3 decision real and keep the dependency set compatible.

- Add `LICENSE` (GPLv3) and any required `NOTICE` or copyright files to the repo.
- Audit every runtime dependency and candidate library for GPLv3 compatibility: HyperFormula (GPLv3 edition), Univer (Apache-2.0), `numfmt`, the OOXML layer, PDF/print libraries, and test tooling.
- Record the distribution obligations: source availability, licence notices, and what happens to the browser bundle.
- Flag any dependency that forces a different project licence or needs a commercial exception.
- Confirm the choice with all contributors before the spec is finalised.

This is a task, not a decision: it unblocks the spec package. The result is a licence file plus a dependency-compliance note.

## Notes

- **GenOffice prior art is Apache-2.0, not GPLv3.** Its top-level `LICENSE` is Apache-2.0; the `ee/` directory is under a separate proprietary GenOffice Enterprise License and must not be reused. Apache-2.0 is compatible with GPLv3, so the open-source parts may be reused with notices and attribution.
- **Univer** (the grid engine GenOffice uses) is Apache-2.0 and GPLv3-compatible.
- **HyperFormula** is used under its GPLv3 edition, which is what makes calcplus GPLv3.
