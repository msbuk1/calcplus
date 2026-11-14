# Fixtures

Small committed workbooks. Regenerate with `npm run fixtures`
(`scripts/generate-fixtures.mjs`). The generator is deterministic: the ZIP
entry date is fixed and the XML payloads are handwritten.

| File | Purpose |
|---|---|
| `compat/identity-minimal.xlsx` | No-op open → save returns the original bytes (file-compat F1 shape). |
| `compat/m1-chart-unknown.xlsx` | M1 walking skeleton workbook: shared strings, inline string, number, boolean, error, formula with cached result, styles, merged range, freeze pane, a chart with its drawing/VML parts, a `<legacyDrawing>`, a `<phoneticPr>` and an unknown `<extLst>` with a custom-properties extension. The edited save test asserts the children and their order survive in the dirty worksheet part, and the part inventory proves the chart/drawing/VML bytes are untouched. This is a preservation fixture; Excel 365 refuses its deliberately minimal chart package, so the manual Excel capture uses `oracle/m1-values-errors.xlsx` instead. |
| `oracle/m1-values-errors.xlsx` | Excel-openable value and error oracle fixture: text, inline string, number, boolean, constant errors, a formula with a cached result, a formula with an error result and text beginning with `=` both as a constant and as a formula result. The manual local Excel capture uses this workbook; its narrow error column makes display text unusable as error identity. |
| `expected/m1-values.json` | **Hand-authored** M1 fixture expectations produced by the fixture generator, compared with `1e-9` relative / `1e-12` absolute tolerance. They are not Excel 365 oracle goldens. |
| `oracle-manifest.json` | Stub manifest for stored Excel-oracle goldens. Status is `not-configured`: no Excel version, locale or committed golden exists yet. |

No real or personal workbook data is stored here. The Excel 365 oracle goldens
are not committed yet; the capture runs manually on the maintainer's PC. See
`scripts/oracle/README.md`.
