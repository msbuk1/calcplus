# Third-party notices

calcplus is distributed under **GPL-3.0-only**. This file records every
runtime, build and test dependency, its licence and the obligations that
come with it. The release build must include the full licence texts for every
distributed component; this file is the human-readable index.

The audit was last reviewed on 2026-11-13. Re-run it whenever a dependency is
added, removed or upgraded.

## Runtime dependencies (distributed)

| Component | Version | Licence | Use | Obligations |
|---|---|---|---|---|
| HyperFormula | 3.5.x | GPL-3.0-only | Formula engine | Include the licence and copyright; the combined app is GPL-3.0-only; provide corresponding source. |
| React / React DOM | 19.x | MIT | UI shell | Include the copyright and permission notice. |
| numfmt | 4.1.3 | MIT | Number formats | Include the copyright and permission notice. |
| fflate | 0.8.x | MIT | ZIP read/write | Include the copyright and permission notice. |
| saxes | 6.x | ISC | XML parsing | Include the copyright and permission notice. |
| PDFKit | 0.20.x | MIT | Vector PDF export | Include the copyright and permission notice; PDFKit's font and image dependencies are bundled with their own notices. |
| Go standard library | Go 1.2x | BSD-3-Clause | Launcher | Include the Go copyright and licence notice in the launcher distribution. |
| Selawik | current | OFL-1.1 (Reserved Font Name) | UI font | Include the OFL text and copyright; do not modify or subset without renaming. |
| Carlito | current | OFL-1.1 (Reserved Font Name) | Document font | Include the OFL text and copyright; do not modify or subset without renaming. |
| Liberation Sans | current | OFL-1.1 | Document font | Include the OFL text and copyright. |

The bundled icon sets are original SVG artwork drawn for calcplus. They are
not copied from Microsoft or any other product.

## Build dependencies (not distributed)

| Component | Licence | Obligations |
|---|---|---|
| Vite | MIT | Notice in the source tree. |
| TypeScript | Apache-2.0 | Notice in the source tree. |
| esbuild (via Vite) | MIT | Notice in the source tree. |
| PostCSS and plugins | MIT | Notice in the source tree. |

Build tools are not part of the shipped bundle, but their licences must remain
compatible with GPL-3.0-only and their notices must be kept in the repository.

## Test and conformance tooling (not distributed)

| Component | Licence | Use | Note |
|---|---|---|---|
| Vitest | MIT | Unit and integration tests | Compatible. |
| Playwright | Apache-2.0 | End-to-end and visual tests | Compatible with GPLv3. |
| fast-check | MIT | Property tests | Compatible. |
| axe-core | MPL-2.0 | Accessibility tests | MPL-2.0 is GPLv3-compatible; test-only. |
| openpyxl | MIT | Independent file verification | Test-only. |
| LibreOffice | MPL-2.0 | Differential smoke oracle | Test-only; not distributed. |
| ssf | Apache-2.0 | Number-format differential oracle | Test-only; not distributed. |
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
  MPL-2.0 and OFL-1.1.
- **HyperFormula GPL-3.0-only is the binding constraint.** It is why the
  project is GPLv3. A proprietary calcplus would require a commercial
  HyperFormula licence and permission from every contributor to relicense.
- **Apache-2.0** is compatible with GPLv3, but not with GPLv2. This project is
  GPLv3-only, so that distinction does not bite.
- **MPL-2.0** is file-level copyleft and GPLv3-compatible. axe-core and
  LibreOffice are test-only and not distributed.
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
