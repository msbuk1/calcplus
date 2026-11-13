# calcplus v1 — licensing and dependency compliance

Resolved by [23 — Project licence: GPLv3 and dependency compliance](../issues/23-licence-compliance.md) on 2026-11-13. This is an engineering compliance note, not legal advice. Counsel should review the release package before the first public release.

## 1. Decision

**calcplus is licensed under GPL-3.0-only.**

- `LICENSE` contains the unmodified GNU GPL version 3 text.
- `NOTICE` carries the copyright, the warranty disclaimer and the source link.
- `THIRD_PARTY_NOTICES.md` is the dependency audit and obligation index.
- `CONTRIBUTING.md` records the DCO and the contribution licence.
- The source is <https://github.com/msbuk1/calcplus>.

The choice is forced by the formula engine: **HyperFormula is used under its GPL-3.0-only edition**. A proprietary calcplus would need a commercial HyperFormula licence and permission from every contributor to relicense; that is not planned.

## 2. Dependency audit

Every runtime, build and test dependency is listed in
[THIRD_PARTY_NOTICES.md](../../../THIRD_PARTY_NOTICES.md), with version, licence,
use and obligations.

Summary:

- **Runtime:** HyperFormula (GPL-3.0-only), React/React DOM (MIT), `numfmt`
  (MIT), `fflate` (MIT), `saxes` (ISC), PDFKit (MIT), Go standard library
  (BSD-3-Clause), Selawik/Carlito/Liberation fonts (OFL-1.1).
- **Build:** Vite (MIT), TypeScript (Apache-2.0), esbuild (MIT).
- **Test only:** Vitest (MIT), Playwright (Apache-2.0), fast-check (MIT),
  axe-core (MPL-2.0), openpyxl (MIT), LibreOffice (MPL-2.0), `ssf`
  (Apache-2.0), Python (PSF-2.0).
- **Evaluated, not used:** Univer (Apache-2.0), GenOffice (Apache-2.0 except
  proprietary `ee/`), IronCalc (MIT/Apache-2.0), SheetJS CE (Apache-2.0),
  ExcelJS (MIT), xlsx-populate (MIT).

All distributed licences are compatible with GPL-3.0-only. The test-only
dependencies are not distributed, and their licences are compatible in any
case.

## 3. Distribution obligations

### 3.1 Corresponding source

GPLv3 requires the **Corresponding Source** when object code is conveyed. For
calcplus that means:

- the application source at the released tag;
- the build scripts, lockfile and configuration needed to rebuild the bundle;
- the source of distributed dependencies, including HyperFormula;
- the launcher source and its embedded build inputs.

Every release publishes a source tag and a source archive next to the binaries.
The archive includes the lockfile and a reproducible build script.

### 3.2 Notices and legal notices

- Keep the copyright, licence and warranty notices intact in every copy.
- Mark modified versions and give a relevant date, as GPLv3 §5(a) requires.
- The app's **About** dialog is the Appropriate Legal Notices surface: it shows
  the copyright, the no-warranty statement, the licence, a link to the full
  text and the source link.
- The launcher package includes `LICENSE`, `NOTICE` and
  `THIRD_PARTY_NOTICES.md`.
- The About dialog also opens the third-party notices and the full licence
  texts.

### 3.3 Browser bundle

The minified JavaScript bundle is object code. The source archive is the
Corresponding Source. The bundle must not add restrictions that contradict
GPLv3, and it must not use DRM or anti-circumvention measures that prevent
users from running modified versions.

### 3.4 Launcher

The Go launcher is a separate program that embeds the static build. The
launcher's own code may use a permissive licence, but the combined distribution
still conveys the GPLv3 app and must include the app's source and notices. The
Go standard library notice is included. The launcher is not a "User Product"
and does not lock down the software, so no Installation Information is
required. It must not prevent a user from replacing the embedded build with a
modified one.

### 3.5 Hosted path

GPLv3 is not the Affero GPL: serving the app over a network does not by itself
add a source obligation. But downloading the bundle is conveying object code,
so the hosted download page and the About dialog link to the source tag and
archive. No additional terms are imposed on network users.

### 3.6 Fonts

Selawik, Carlito and Liberation are separate works under the SIL Open Font
License 1.1, not GPLv3. Their copyright and licence notices travel with them.
Reserved Font Names must be respected: do not modify or subset the fonts under
their reserved names. If a modified font is ever needed, rename it.

### 3.7 Icons and assets

The icon sets are original SVG artwork drawn for calcplus. No Microsoft icons,
logos, screenshots or other assets are copied. The UI may say that calcplus is
compatible with `.xlsx` and Excel-style workflows, but it must not use the
Excel logo or imply endorsement.

### 3.8 Test fixtures

Real workbooks used for conformance are anonymised and cleared for test use.
They are not distributed with the app and are not part of the Corresponding
Source unless their licence allows it.

## 4. Flags and constraints

| Item | Status | Action |
|---|---|---|
| HyperFormula GPL-3.0-only | Binding | The project is GPLv3. A proprietary edition needs a commercial licence and contributor permission. |
| Univer Apache-2.0 | Compatible, not used | If code is ever reused, add the Apache-2.0 notice. |
| GenOffice `ee/` | Proprietary | Never copy. Apache-2.0 parts need attribution if reused. |
| Apache-2.0 dependencies | Compatible with GPLv3 | Keep notices; do not assume GPLv2 compatibility. |
| MPL-2.0 test tools | Compatible | Test-only; no distribution obligation. |
| OFL fonts | Compatible, separate | Include notices; respect Reserved Font Names. |
| JSZip dual MIT/GPL-3.0-or-later | Compatible | Choose MIT if it is ever shipped. |
| Commercial store DRM | Conflict risk | Do not ship through a store whose DRM prevents modified versions. |
| Excel trademarks | Nominative use only | No logo, no product name containing "Excel", no endorsement claim. |

## 5. Contributor confirmation

- The current contributor list is the repository owner. The project has no
  separate CLA.
- Every commit is signed off under the Developer Certificate of Origin 1.1, as
  recorded in [CONTRIBUTING.md](../../../CONTRIBUTING.md).
- The DCO certifies the right to submit the contribution under GPL-3.0-only.
- Before the spec package is finalised, the maintainer confirms the licence in
  the release checklist. A future contributor who cannot certify the DCO cannot
  contribute code.

## 6. Build and release checklist

1. `LICENSE`, `NOTICE`, `THIRD_PARTY_NOTICES.md` and `CONTRIBUTING.md` are
   present in the repository.
2. The app's `package.json` declares `"license": "GPL-3.0-only"`.
3. The About dialog shows the copyright, no-warranty statement, licence link,
   source link and third-party notices.
4. The launcher package includes the same notices.
5. The release publishes a source tag and archive beside the binaries.
6. A machine-readable SBOM (SPDX or CycloneDX) is generated and attached.
7. The dependency audit is regenerated from the lockfile and reviewed.
8. No dependency with an incompatible or unknown licence is present.
9. No Microsoft asset, logo or proprietary GenOffice code is present.
10. The build contains no DRM or anti-circumvention measure.
11. The release notes state the licence and the source location.
12. Counsel reviews the first public release package.

## 7. CI licence gate

- A licence check runs on every pull request and on the release build.
- The allowlist is: GPL-3.0-only, MIT, ISC, BSD-2-Clause, BSD-3-Clause,
  Apache-2.0, MPL-2.0, OFL-1.1, PSF-2.0, 0BSD and CC0-1.0.
- An unknown or disallowed licence fails the build until it is reviewed and the
  allowlist or the dependency changes.
- The check covers runtime and build dependencies; test-only dependencies are
  reported but do not block.
- A dependency upgrade that changes a licence is a project-level event, not a
  routine bump.
- The SBOM is diffed against the previous release; new components are called
  out in the release notes.

## 8. Acceptance criteria

1. `LICENSE` contains the unmodified GPLv3 text.
2. `NOTICE` carries the copyright, warranty disclaimer and source link.
3. `THIRD_PARTY_NOTICES.md` lists every runtime, build and test dependency with
   its licence and obligations.
4. `CONTRIBUTING.md` records GPL-3.0-only, the DCO and the no-CLA policy.
5. Every distributed dependency is GPLv3-compatible.
6. HyperFormula's GPL-3.0-only licence is recorded as the binding constraint.
7. The distribution obligations for source, notices, bundle, launcher, hosted
   path, fonts and assets are documented.
8. The flags table covers every dependency that could force a different
   licence.
9. The release checklist includes the About dialog, source archive, SBOM and
   counsel review.
10. The CI licence gate runs and blocks unknown or disallowed licences.
11. No Microsoft asset or proprietary GenOffice code is present.
12. The maintainer confirms the licence before the spec is finalised.

## 9. Open risks

- **HyperFormula relicensing.** If the project ever needs a proprietary
  licence, the change is expensive and requires every contributor's consent.
- **Font subsetting.** Subsetting Selawik or Carlito may count as a modified
  version under the OFL; rename if that happens.
- **Store DRM.** A store that prevents modified builds conflicts with GPLv3.
- **Dependency drift.** A routine upgrade can change a licence; the CI gate and
  the SBOM diff are the guard.
- **Legal review.** The first public release should be reviewed by counsel in
  the project's jurisdiction.
