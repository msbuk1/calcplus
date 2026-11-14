# 23 — Project licence: GPLv3 and dependency compliance

Status: resolved
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

## Answer

> **Amended 2026-11-14:** the licence gate now runs locally; no test or check
> runs on GitHub. See the amendment at the end of this file.

The licence decision is now real in the repository:

- `LICENSE` — the unmodified GNU GPL version 3 text.
- `NOTICE` — copyright, no-warranty statement, source link and the Microsoft non-affiliation note.
- `THIRD_PARTY_NOTICES.md` — every runtime, build and test dependency with version, licence, use and obligations.
- `CONTRIBUTING.md` — GPL-3.0-only, DCO 1.1 sign-off, no CLA, third-party and asset rules.
- `CONTRIBUTORS.md` — the current contributor list and the contribution terms.
- `README.md` — project status, licence, source and privacy.
- Compliance note: [licensing.md](../spec/licensing.md).

**Audit result:** every distributed dependency is GPLv3-compatible. HyperFormula is GPL-3.0-only and is the binding constraint; React, `numfmt`, `fflate`, PDFKit, Vite, Vitest and fast-check are MIT; `saxes` is ISC; the Go launcher and TypeScript are BSD/Apache-2.0; the fonts are OFL-1.1; axe-core and LibreOffice are MPL-2.0 and test-only. Univer, GenOffice, IronCalc, SheetJS, ExcelJS and xlsx-populate are evaluated but not used; GenOffice's proprietary `ee/` directory must never be copied.

**Distribution obligations recorded:** corresponding source at each release, source tag and archive, notices in the About dialog and launcher package, no DRM or anti-circumvention, separate OFL font notices, original icons only, no Microsoft assets or logos, no store DRM that blocks modified builds, and an SBOM per release.

**Flags:** a proprietary calcplus would need a commercial HyperFormula licence and every contributor's consent; Apache-2.0 is compatible with GPLv3 but not GPLv2; MPL-2.0 test tools are compatible; OFL Reserved Font Names must be respected; store DRM is a conflict.

**Contributor confirmation:** the current contributor is the repository owner; every commit is signed off under the DCO, which certifies the right to submit under GPL-3.0-only. No CLA.

**Licence gate:** a local gate checks each change and the release against an allowlist and fails on unknown or disallowed licences; the SBOM is diffed between releases.

## Amendment — 2026-11-14

The testing-plan change removed GitHub execution. The licence gate runs locally before each commit and at release; there is no GitHub Actions licence job.
