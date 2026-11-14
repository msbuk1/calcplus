# Fixtures

Small committed workbooks. Regenerate with `npm run fixtures`
(`scripts/generate-fixtures.mjs`). The generator is deterministic: the ZIP
entry date is fixed and the XML payloads are handwritten.

| File | Purpose |
|---|---|
| `compat/identity-minimal.xlsx` | No-op open → save returns the original bytes (file-compat F1 shape). |
| `compat/m1-chart-unknown.xlsx` | M1 walking skeleton workbook: shared strings, inline string, number, boolean, error, formula with cached result, styles, merged range, freeze pane, a chart with its drawing/VML parts, a `<legacyDrawing>`, a `<phoneticPr>` and an unknown `<extLst>` with a custom-properties extension. The edited save test asserts the children and their order survive in the dirty worksheet part, and the part inventory proves the chart/drawing/VML bytes are untouched. |
| `expected/m1-values.json` | **Hand-authored** M1 fixture expectations produced by the fixture generator, compared with `1e-9` relative / `1e-12` absolute tolerance. They are not Excel 365 oracle goldens. |
| `oracle-manifest.json` | Stub manifest for stored Excel-oracle goldens. Status is `not-configured`: no oracle runner, Excel version, locale or generated golden exists yet. |

No real or personal workbook data is stored here. The Excel 365 oracle goldens
are not committed or configured yet; see `scripts/oracle/README.md`.
