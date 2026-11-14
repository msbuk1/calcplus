# Third-party notices

calcplus is distributed under **GPL-3.0-only**. This file records every
runtime, build and test dependency, its licence and the obligations that
come with it. The release build must include the full licence texts for every
distributed component; this file is the human-readable index.

The audit was last reviewed on **2026-11-14** for the M1 walking skeleton.
Re-run it whenever a dependency is added, removed or upgraded. The local
licence gate (`scripts/licence-gate.mjs`) checks the lockfile before each
commit.

## Runtime dependencies (distributed)

| Component | Version | Licence | Use | Obligations |
|---|---|---|---|---|
| React / React DOM | 19.3.x | MIT | UI shell | Include the copyright and permission notice. |
| fflate | 0.8.x | MIT | ZIP read/write | Include the copyright and permission notice. |
| saxes | 6.0.x | ISC | XML parsing | Include the copyright and permission notice. |

Later milestones add HyperFormula (GPL-3.0-only), `numfmt` (MIT), PDFKit (MIT),
the Go standard library launcher (BSD-3-Clause) and the OFL-1.1 fonts. They are
not installed by the M1 walking skeleton; their notices and obligations are
recorded in [licensing §2](.scratch/calcplus/spec/licensing.md) and must
be added here when the dependency lands.

The bundled icon sets are original SVG artwork drawn for calcplus. They are
not copied from Microsoft or any other product.

## Build dependencies (not distributed)

| Component | Version | Licence | Obligations |
|---|---|---|---|
| Vite | 7.3.x | MIT | Notice in the source tree. |
| TypeScript | 5.9.x | Apache-2.0 | Keep the Apache-2.0 notice. |
| esbuild (via Vite) | 0.28.x | MIT | Notice in the source tree. |
| PostCSS and plugins (via Vite) | current | MIT | Notice in the source tree. |
| ESLint / `@eslint/js` | 9.39.x | MIT | Notice in the source tree. |
| typescript-eslint | 8.72.x | MIT | Notice in the source tree. |
| Prettier | 3.10.x | MIT | Notice in the source tree. |
| globals | 16.5.x | MIT | Notice in the source tree. |

Two transitive build/test licences are reviewed additions to the contract's
allowlist (`spec/licensing.md` §7). Both are GPL-3.0-only-compatible:

- **Python-2.0** — `argparse` (via ESLint's `js-yaml`); FSF classifies the
  Python 2.0 licence as GPL-compatible.
- **BlueOak-1.0.0** — `minimatch` 10 (via typescript-eslint); permissive,
  GPL-compatible.

They are not distributed in the browser bundle. Any further addition to this
list needs a review note and a change to the local licence gate.

## Test and conformance tooling (not distributed)

| Component | Version | Licence | Use | Note |
|---|---|---|---|---|
| Vitest | 4.1.x | MIT | Unit and integration tests | Compatible. |
| Playwright | 1.65.x | Apache-2.0 | End-to-end tests | Compatible with GPLv3. |
| fast-check | 4.10.x | MIT | Property tests | Compatible. |
| `@types/*` | current | MIT | TypeScript declarations | Build-only. |
| axe-core | MPL-2.0 | Accessibility tests (v2; unused in v1) | MPL-2.0 is GPLv3-compatible; test-only. |
| openpyxl | MIT | Independent file verification | Test-only; not yet wired. |
| LibreOffice | MPL-2.0 | Differential smoke oracle | Test-only; not yet wired. |
| ssf | Apache-2.0 | Number-format differential oracle | Test-only; later milestone. |
| Python | PSF-2.0 | Oracle and fixture scripts | Test-only. |
| JSZip | MIT or GPL-3.0-or-later | Prototypes and tests only | Choose MIT if it is ever shipped. |

## Evaluated but not used

| Component | Licence | Why it is not used |
|---|---|---|
| Univer | Apache-2.0 | Rejected as the grid foundation; compatible if code were reused, with notices. |
| GenOffice | Apache-2.0 except `ee/` | Prior art only. The `ee/` directory is proprietary and must never be copied. |
| IronCalc | MIT / Apache-2.0 | Rejected for v1 on date and `.xlsx` fidelity. |
| SheetJS CE | Apache-2.0 | Data extractor, not a preservation engine. |
| ExcelJS | MIT | Drops charts, external links, pivots and VBA. |
| xlsx-populate | MIT | Drops cached formula results. |
| SpreadJS | Commercial | Too expensive and replaces the product. |
| Handsontable | Proprietary | Commercial licence required. |

## Compatibility summary

- **GPL-3.0-only** is compatible with MIT, ISC, BSD-3-Clause, Apache-2.0,
  MPL-2.0, OFL-1.1, Python-2.0 and BlueOak-1.0.0.
- **HyperFormula GPL-3.0-only is the binding constraint** when the formula
  engine lands. It is why the project is GPLv3. A proprietary calcplus would
  require a commercial HyperFormula licence and permission from every
  contributor to relicense.
- **Apache-2.0** is compatible with GPLv3, but not with GPLv2. This project is
  GPLv3-only, so that distinction does not bite.
- **MPL-2.0** is file-level copyleft and GPLv3-compatible. axe-core and
  LibreOffice are test-only and not distributed. axe-core is unused until the
  v2 accessibility work.
- **OFL-1.1 fonts** are separate works. They are not relicensed under GPLv3;
  their notices travel with them. Reserved Font Names must be respected.
- **No Microsoft assets** are copied. Excel icons, logos and fonts other than
  the open-source Selawik are not used.

## Full licence texts

The release build must ship:

- `LICENSE` — GNU GPL version 3.
- `NOTICE` — calcplus copyright, warranty disclaimer and source link.
- `THIRD_PARTY_NOTICES.md` — this file.
- The full text of every distributed dependency licence, generated from the
  lockfile and included in the app's About > Licences panel and the launcher
  package.

A machine-readable SBOM (SPDX or CycloneDX) is generated for each release.
