# Third-party notices

calcplus is distributed under **GPL-3.0-only**. This file records every
runtime, build and test dependency, its licence and the obligations that
come with it. The release build must include the full licence texts for every
distributed component; this file is the human-readable index.

The audit was last reviewed on **2026-11-20** for the M7 PDF export
(PDFKit plus the OFL print fonts).
Re-run it whenever a dependency is added, removed or upgraded. The local
licence gate (`scripts/licence-gate.mjs`) checks the lockfile before each
commit.

## Runtime dependencies (distributed)

| Component | Version | Licence | Use | Obligations |
|---|---|---|---|---|
| HyperFormula | 3.5.0 | GPL-3.0-only | Formula engine (worker) | The project is GPL-3.0-only because of this dependency. Include the copyright and licence text; publish the corresponding source. |
| chevrotain | 6.5.0 | Apache-2.0 | Parser runtime (via HyperFormula) | Include the Apache-2.0 licence text; state any modified files. |
| tiny-emitter | 2.1.0 | MIT | Event emitter (via HyperFormula) | Include the copyright and permission notice. |
| React / React DOM | 19.3.x | MIT | UI shell | Include the copyright and permission notice. |
| fflate | 0.8.x | MIT | ZIP read/write; also PDF stream compression (shared with PDFKit) | Include the copyright and permission notice. |
| saxes | 6.0.x | ISC | XML parsing | Include the copyright and permission notice. |
| numfmt | 4.1.3 | MIT | Excel number-format grammar | Include the copyright and permission notice. The Excel seams (built-in ids, 1904 adapter, width layout, colours, grouping and invalid fallback) are calcplus code. |
| PDFKit | 0.21.x | MIT | Vector PDF export (File → Export PDF) | Include the copyright and permission notice. It is imported on demand, so the ~592 KB chunk is not in the startup bundle. **Dependency review (2026-11-20):** PDFKit 0.21 is MIT, actively maintained, and its browser build imports only `fflate`, `fontkit`, `linebreak` and `png-js` — no Node polyfills are shipped. The spike measured its subsetting and rename seam before the dependency was added (`.scratch/calcplus/research/26-print-pdf-fidelity.md`). |
| fontkit | 3.0.0 | MIT | Font parsing, subsetting and the plan-time glyph coverage scan (via PDFKit) | Include the copyright and permission notice. It is the same library PDFKit maps glyphs with, so the planner and the renderer cannot disagree. |
| linebreak | 2.0.0 | MIT | Unicode line breaking (via PDFKit) | Include the copyright and permission notice. |
| png-js | 3.0.0 | MIT | PNG decoding for the fallback page images (via PDFKit) | Include the copyright and permission notice. |
| unicode-trie, restructure, brotli | 3.0.0, 3.0.2, 2.1.0 | MIT | Fontkit/PDFKit support libraries | Include the copyright and permission notices. |
| Selawik, Carlito, Liberation (fonts) | see `assets/fonts` | OFL-1.1 | Embedded PDF text | Separate works, not relicensed under GPLv3. Ship the font files with the licence files in `assets/fonts`; respect the Reserved Font Names `Selawik` and `Liberation` (see below). |
| Go standard library | toolchain version (1.18 or newer) | BSD-3-Clause | Launcher static file server and embedded app | Include the Go BSD-3-Clause notice and the launcher source with the release. The build script copies the toolchain's licence as `GO-LICENSE`. The launcher uses no third-party Go modules. |

### The bundled print fonts

`assets/fonts` holds the font files every PDF export embeds, with their licence
files:

| Family | Files | Use | Reserved Font Name |
|---|---|---|---|
| Selawik | Regular, Bold | Segoe UI cells (metric clone) | **Selawik** |
| Carlito | Regular, Bold, Italic, Bold Italic | Calibri cells, the pinned Aptos deviation, and the Greek/Cyrillic supplement | none in the copyright record |
| Liberation Sans / Serif / Mono | four styles each | Arial, Times New Roman and Courier New cells | **Liberation** (the sources also reserve Arimo, Tinos and Cousine) |

The files ship unmodified. PDFKit subsets what it embeds, and every subset is
renamed `CALP+<family>` (for example `CALP+Selawik`) in all three font
descriptors, so no modified font is distributed under a Reserved Font Name
(licensing §3.6; research/26 §3.3). The subset carries no `name` table of its
own, so the PDF is the only place the name appears.


### The Go launcher

The launcher in [`launcher/`](launcher/) is a separate program that embeds the
static build. It imports only the Go standard library, so the only distributed
Go notice is the standard library's BSD-3-Clause notice, copied into each
package as `GO-LICENSE`. The release package carries `LICENSE`, `NOTICE` and
this file; the launcher source, its `go.mod` and the build script are part of the Corresponding Source ([licensing §3.1 and
§3.4](.scratch/calcplus/spec/licensing.md)). The launcher does not prevent a user
from replacing the embedded build with a modified one.

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
| openpyxl | MIT | Independent file verification and the F22 producer fixture | Test-only; not distributed. |
| LibreOffice | MPL-2.0 | Differential smoke oracle and the F22 producer fixture | Test-only; not distributed. |
| SheetJS CE (`xlsx`) | Apache-2.0 | Generated the committed F22 SheetJS fixture | Test-only; not distributed and no longer a dependency (0.18.5 advisory, maintainer review 2026-11-19): install it temporarily only to regenerate the fixture. |
| xlsx-populate | MIT | F22 producer fixture generator | Test-only; not distributed. |
| JSZip (via xlsx-populate) | MIT | Producer fixture archiving | Dual MIT/GPL-3.0-or-later; MIT is used. |
| Zlib (via pako, xlsx-populate) | Zlib | Producer fixture compression | Permissive, GPL-3.0-only compatible; reviewed addition to the gate. |
| ssf | Apache-2.0 | Number-format differential oracle | Test-only; later milestone. |
| Python | PSF-2.0 | Oracle and fixture scripts | Test-only. |
| JSZip | MIT or GPL-3.0-or-later | Prototypes and tests only | Choose MIT if it is ever shipped. |

## Evaluated but not used

| Component | Licence | Why it is not used |
|---|---|---|
| Univer | Apache-2.0 | Rejected as the grid foundation; compatible if code were reused, with notices. |
| GenOffice | Apache-2.0 except `ee/` | Prior art only. The `ee/` directory is proprietary and must never be copied. |
| IronCalc | MIT / Apache-2.0 | Rejected for v1 on date and `.xlsx` fidelity. |
| SheetJS CE | Apache-2.0 | Data extractor, not a preservation engine; generated the committed F22 fixture (dependency removed 2026-11-19). |
| ExcelJS | MIT | Drops charts, external links, pivots and VBA. |
| xlsx-populate | MIT | Drops cached formula results; used only to generate F22 producer fixtures. |
| SpreadJS | Commercial | Too expensive and replaces the product. |
| Handsontable | Proprietary | Commercial licence required. |

## Compatibility summary

- **GPL-3.0-only** is compatible with MIT, ISC, BSD-3-Clause, Apache-2.0,
  MPL-2.0, OFL-1.1, Python-2.0 and BlueOak-1.0.0.
- **HyperFormula GPL-3.0-only is the binding constraint.** The formula engine
  is distributed with the browser bundle. It is why the project is GPLv3. A
  proprietary calcplus would require a commercial HyperFormula licence and
  permission from every contributor to relicense.
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
