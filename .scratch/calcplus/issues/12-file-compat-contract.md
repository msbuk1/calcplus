# 12 — .xlsx file-compatibility contract

Status: resolved
Type: grilling
Blocked by: 01, 04

## Question

What exactly does "high-fidelity `.xlsx` round-trip" promise?

Using the chosen import/export approach (04) and the v1 feature matrix (01), define a testable compatibility contract:

- Features that must survive open → edit → save unchanged
- Features read-only but preserved verbatim
- Features that trigger a warning on open or save
- Features that block opening
- Behavior for formulas without cached values, external links, encrypted files, macros, and pivot caches
- How unsupported content is surfaced to the user

Output: a feature-by-feature matrix plus acceptance fixtures.

## Answer

"High-fidelity round-trip" is defined as three tiers with a named guarantee for each. Full contract in [file-compat.md](../spec/file-compat.md).

- **Tier A — modelled and editable:** cells, formulas plus cached results, styles and number formats, merges, dimensions, freeze panes, defined names, sheet metadata, comments, hyperlinks, validation, conditional formatting, AutoFilter, protection flags and print setup. Read, write and semantic round-trip.
- **Tier B — preserved, not editable:** charts, images, shapes, pivots and caches, external links, VBA, signatures, tables, slicers, sparklines, custom XML and custom properties. Parts are byte-identical; not rendered. Structural edits transform anchors where possible, otherwise the specific edit is blocked.
- **Tier C — rejected:** encrypted/IRM workbooks, `.xls`, `.xlsb`, ODS, corrupt or non-ZIP files, zip bombs, Strict OOXML and packages over the size caps. Clear message, no partial import, file untouched.

**Key behaviours:** a no-op save returns the original bytes; style and shared-string tables are append-only; formula text is never rewritten and `<v>` is written only when the engine evaluated the cell; shared and legacy array formulas are preserved (arrays read-only); external links are never refreshed; macros survive in `.xlsm` and never run, and `.xlsm` → `.xlsx` requires confirmation; pivot caches are never refreshed; digital signatures are preserved but any save invalidates them; protection is preserved but not enforced.

**Surfacing:** a non-modal compatibility report on open only when there is something to report, and an export report on save using five categories — unchanged, transformed, preserved but stale, removed, warning. No silent loss.

**Fixtures:** 24 named fixtures covering identity, append-only tables, cached and uncached formulas, shared/array formulas, charts, pivots, macros, threaded comments, validation and conditional formatting, protection, 1904 dates, print setup, rich text, encrypted/corrupt/zip-bomb/Strict inputs, the producer matrix (Excel 365/2019, LibreOffice, Google Sheets) and CSV.

## Notes

- The compatibility report and unsafe-anchor warnings described in [persistence](spec/persistence.md) depend on this ticket's matrix.
