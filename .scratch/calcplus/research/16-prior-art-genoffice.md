# Prior art: GenOffice (genspark-ai/genoffice)

Researched 2026-11-13 against `main` at commit `d751696` (2026-11-13), https://github.com/genspark-ai/genoffice. Apache-2.0. ~12.3k stars.

## What it is

A local-first, AI-native office suite (Docs, Sheets, Slides, PDF, Markdown, HTML) shipped as an Electron desktop app, with a CLI and agent skill. Sheets is the relevant app: a spreadsheet that opens, edits and saves real `.xlsx`/`.xlsm` with live formulas, charts, pivots, slicers, filters, conditional formatting and data validation.

**Licence:** Apache-2.0, except the `ee/` directory, which is under a separate proprietary GenOffice Enterprise License and must not be reused. Apache-2.0 is compatible with GPLv3, so the open-source parts can be reused in calcplus with notices and attribution.

## Architecture snapshot

- **Shell**: Electron plus a separate React/Vite renderer per app (`apps/shell`, `apps/sheets`, ...). Not a single browser app.
- **Grid and formula engine**: Univer 0.25 (`@univerjs/core`, `engine-render`, `preset-sheets-core`, plus table/sort/filter/CF/DV/note presets), Apache-2.0. The Sheets renderer has ~256 modules around it, many of them compatibility patches.
- **`.xlsx` gateway**: `@genoffice/xlsx-gateway` reads and patches OOXML with JSZip + `fast-xml-parser` + `numfmt`. It does **not** use SheetJS or ExcelJS as the storage engine.
- **Native sidecars (Rust)**: `apps/sheets/native/xlsx-engine` (streaming package patch, `.xls`/ODF conversion, shared formulas, cached values, charts/images) and `native/win-clipboard-helper` (Windows clipboard formats).
- **Testing**: per-package Vitest, a large Playwright e2e suite with visual baselines, a byte-identical compatibility corpus, and benchmark scripts.

## Lessons for calcplus

### 1. Preservation-first OOXML is validated (tickets 04, 12, 13)

The gateway plans a small set of patched package entries and copies everything else verbatim. `apps/sheets/scripts/verify-compatibility.ts` states the rule outright: "One surgical cell edit per fixture; every other package entry must survive byte-identical, including binary and exotic parts." Their no-op save path returns the original archive bytes. This independently confirms ticket 04's recommendation and gives ticket 12 a concrete acceptance test.

### 2. Shared formulas and structural edits are the real cost centre (07, 12)

Evidence: `shared_formulas.rs`, `shared-formula-index.ts`, `shared-formula-journal.ts`, `xlsx-structure.ts`, `sort-formula-refs.ts`. Shared formulas (`<f t="shared" si="N"/>`) carry no body in follower cells, and inserting or deleting rows/columns forces reference rewriting across formulas, merges, names, tables, conditional formats, drawings and chart anchors. Ticket 04 called range/anchor transformation the largest hidden cost; GenOffice's code confirms it.

### 3. Cached formula values are a safety net (02, 12)

`formula-cached-fallback.ts` falls back to the cached result stored in the file when the live engine cannot compute. For a faithful clone the file's cached value is evidence; never render a blank or a wrong value just because the engine lacks a function. Mark it stale, but preserve it.

### 4. Fidelity is a long tail of named patches (01, 11, 13)

The Sheets renderer includes modules such as `merge-border-fix.ts`, `thick-border-fix.ts`, `numfmt-fix.ts`, `formula-lexer-fix.ts`, `selection-wrap-fix.ts`, `sheet-rename-fix.ts`, `rtl-text-fix.ts`, `rich-text-bidi-fix.ts`, `clipboard-anchor-tile.ts`, `filter-button-excel-style.ts`. Whatever the architecture, a credible Excel clone spends most of its effort on these edges. The spec and milestones must budget for them.

### 5. Buying the grid moves the cost into the seams (05, 17)

GenOffice did not build a grid; it adopted Univer. It still owns a large patch layer. That does not prove custom is wrong, but it disproves "adopt a framework and the fidelity work goes away". Ticket 05 did not evaluate Univer; that decision deserves a stress test before it is treated as final.

### 6. Browser-only changes the calculus (05, 11)

GenOffice's streaming save, `.xls` conversion and Windows clipboard helper are native Rust sidecars. calcplus is a browser app, so those specific wins are unavailable; the equivalent is WASM or a pure-TS design. Do not copy the Electron/sidecar architecture.

### 7. Their conformance harness is a template (13)

- Byte-identical untouched entries after a surgical edit.
- A fixture corpus including a "kitchen sink" workbook.
- Visual baselines for grid and chrome.
- Per-package unit tests and benchmark scripts.

This is close to the strategy ticket 13 needs to define.

## What not to copy

- The AI agent/CLI layer — out of scope for calcplus.
- Electron and native sidecars — our platform decision is browser-only.
- The full feature surface (pivots, slicers, charts, macros, AI) in v1.
- The Univer patch layer as a starting point: it encodes hard-won edge cases but couples you to Univer internals.

## Open questions

- Is Univer's formula engine good enough to be the whole answer, or only the fallback? Ticket 03 recommends HyperFormula with Univer as fallback.
- How much of GenOffice's patch layer is Univer-specific versus inherent Excel edge cases? That determines the true cost of the buy path.
- What is Univer's browser bundle/startup cost at calcplus's target workbook sizes?

## Sources

- Repository: https://github.com/genspark-ai/genoffice (commit `d751696`, 2026-11-13)
- `apps/sheets/package.json` — Univer 0.25 dependencies
- `packages/xlsx-gateway/package.json` — JSZip / fast-xml-parser / numfmt gateway
- `packages/xlsx-gateway/src/gateway/xlsx-package-io.ts` — streaming save, untouched entries
- `apps/sheets/scripts/verify-compatibility.ts` — byte-identical preservation corpus
- `apps/sheets/native/xlsx-engine/src/shared_formulas.rs` — shared formulas
- `apps/sheets/src/renderer/formula-cached-fallback.ts` — cached-value fallback
- Univer license: https://github.com/dream-num/univer/blob/dev/LICENSE (Apache-2.0)
- GenOffice license: https://github.com/genspark-ai/genoffice/blob/main/LICENSE (Apache-2.0)
