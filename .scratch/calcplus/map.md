# calcplus — Excel desktop clone spec

Type: map
Status: open

## Destination

A build-ready spec package for **calcplus**, a browser-based, single-user clone of Excel desktop (Microsoft 365): a core credible spreadsheet — grid and editing, formatting, formulas, multi-sheet workbooks, high-fidelity `.xlsx` round-trip — at functional, near-identical UI fidelity. The package is a PRD, a technical design (workbook model, architecture, ADRs) and phased milestones. No backend, but a data model that does not preclude co-authoring later.

## Notes

- **Effort slug**: `calcplus`; tracker is local markdown under `.scratch/calcplus/`.
- **Fidelity reference**: Excel desktop / Microsoft 365. When Excel and "clean design" disagree, match Excel — the product's promise is familiarity.
- **V1 scope cut**: core credible spreadsheet. Charts and data tools are deferred to v2; PivotTables and macros are never.
- **Constraints**: browser app; single-user; no backend; high-fidelity `.xlsx` import and export; no pixel-perfect requirement, but near-identical chrome.
- **Execution carried into the map**: because the destination is an artifact, ticket 15 assembles the spec package. Every other ticket resolves a decision.
- **Skills for sessions**: `grilling`, `domain-modeling`, `prototype`, `research`.
- **Research artifacts**: `.scratch/calcplus/research/<slug>.md` (files rather than throwaway branches, so they travel with the map).

## Decisions so far

- [Formula engine: build vs buy](issues/03-formula-engine.md) — adopt HyperFormula 3.5 under a commercial licence behind a narrow `FormulaEngine` port; Univer (Apache-2.0) is the fallback if the quote fails or dynamic arrays/3D refs are required. Do not build a custom engine.
- [`.xlsx` import/export approach](issues/04-xlsx-io.md) — preservation-first in-house OOXML layer (JSZip/fflate + saxes): model v1 features, copy unmodelled parts verbatim, patch only dirty parts, append-only style/shared-string tables; SheetJS/ExcelJS are extractors, not storage engines.
- [Grid rendering approach](issues/05-grid-rendering.md) — custom hybrid grid: Canvas 2D surface (cells, gridlines, headers, selection, fill handle, frozen/split panes) + DOM overlays for editing/focus and an ARIA grid mirror; no third-party grid foundation. **Provisional:** GenOffice's Univer experience reopened the buy option — stress-test in [Stress-test the grid decision](issues/17-grid-spike-univer.md).
- [Clipboard and Excel interop](issues/06-clipboard-interop.md) — HTML-first (`text/html` Excel dialect) with TSV fallback; formulas best-effort, `.xlsx` is the fidelity channel, and clipboard is never a dependency for correctness.
- [Prior art: what calcplus can learn from GenOffice](issues/16-prior-art-genoffice.md) — validates preservation-first `.xlsx`; shared formulas, structural-edit reference rewriting and the edge-case tail are the real cost; the Univer buy path deserves a stress test.

## Not yet specified

- Ribbon command inventory and grouping once the feature matrix lands.
- Detailed grid interaction semantics: fill handle, autofit, frozen panes/splits, zoom, merged-cell editing, selection behaviors.
- Number-format language coverage and whether to build or buy a format engine.
- Conditional-formatting scope (v1 or v2).
- Print and PDF output.
- Localization of the UI and of function names.
- Accessibility depth: screen-reader grid semantics and WCAG target.
- Theming, dark mode, and the visual token set.
- Formula-bar and name-box edge behaviors.
- Whether basic sort/filter belongs in v1 despite the data-tools deferral.

## Out of scope

- **Real-time co-authoring, accounts, cloud sync, sharing** — v1 is single-user. The model must not preclude it, but the feature is a later effort.
- **VBA/macros, Office Scripts, add-ins.**
- **PivotTables, Power Query, Solver/Goal Seek, Data Analysis ToolPak.**
- **Charts, images, shapes, sparklines** — v2.
- **Advanced data tools** (tables, sort/filter, data validation) — v2 by the chosen scope cut; ticket 01 fixes the exact boundary.
- **Native desktop and mobile/touch apps.**
- **Legacy `.xls` binary and ODS formats.**
