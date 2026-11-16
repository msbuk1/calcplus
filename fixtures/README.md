# Fixtures

Small committed workbooks. Regenerate with `npm run fixtures`
(`scripts/generate-fixtures.mjs`). The generator is deterministic: the ZIP
entry date is fixed and the XML payloads are handwritten.

| File | Purpose |
|---|---|
| `compat/identity-minimal.xlsx` | No-op open → save returns the original bytes (file-compat F1 shape). |
| `compat/m1-chart-unknown.xlsx` | M1 walking skeleton workbook: shared strings, inline string, number, boolean, error, formula with cached result, styles, merged range, freeze pane, a chart with its drawing/VML parts, a `<legacyDrawing>`, a `<phoneticPr>` and an unknown `<extLst>` with a custom-properties extension. The edited save test asserts the children and their order survive in the dirty worksheet part, and the part inventory proves the chart/drawing/VML bytes are untouched. This is a preservation fixture; Excel 365 refuses its deliberately minimal chart package, so the manual Excel capture uses `oracle/m1-values-errors.xlsx` instead. |
| `compat/m2-grid.xlsx` | M2 grid-foundation workbook: one frozen row and one frozen column, a merged title range, an imported row height and column widths, and the font, fill, alignment and number-format styles the grid renders. The grid tests use it for the four-pane geometry, the merge anchor, the formatted cells and the scroll behaviour. |
| `manual/general-test.xlsx` | Manual smoke workbook. `Cover` has text, numbers, booleans, errors, dates, formulas with cached results, a merged header, a frozen top row and first column, fills and number formats. `Data` has 400 synthetic mixed-value rows under a frozen header for scrolling, selection, find and freeze-pane checks. |
| `oracle/m1-values-errors.xlsx` | Excel-openable value and error oracle fixture: text, inline string, number, boolean, constant errors, a formula with a cached result, a formula with an error result and text beginning with `=` both as a constant and as a formula result. The manual local Excel capture uses this workbook; its narrow error column makes display text unusable as error identity. |
| `expected/m1-values.json` | **Hand-authored** M1 fixture expectations produced by the fixture generator, compared with `1e-9` relative / `1e-12` absolute tolerance. They are not Excel 365 oracle goldens. |
| `expected/m3b-formulas.json` | **Hand-authored** M3B function, error/coercion/operator, date, reference, lookup and shim expectations, plus the data sheet those formulas read. `tests/engine-m3b.test.ts` compares them with `1e-9` relative / `1e-12` absolute tolerance. They are not Excel 365 oracle goldens. |
| `expected/m4a-number-formats.json` | **Hand-authored** M4A built-in number-format display corpus: one exact string for every ECMA-376 built-in id, including the pinned Gregorian en-US rendering of the locale-specific ids. Spaces in `expected` are significant. `tests/m4a-number-formats.test.ts` asserts the exact text. It is not an Excel 365 oracle golden. |
| `oracle-manifest.json` | Stub manifest for stored Excel-oracle goldens. Status is `not-configured`: no Excel version, locale or committed golden exists yet. |

No real or personal workbook data is stored here. The Excel 365 oracle goldens
are not committed yet; the capture runs manually on the maintainer's PC. See
`scripts/oracle/README.md`.

## Number-format corpus gate

Run the committed built-in corpus locally with:

```bash
npm test -- tests/m4a-number-formats.test.ts
```

`ssf` is a differential oracle only (number-formats §1 and
testing-conformance §2). It is not a dependency, is not installed by this
repository and the corpus test never reaches the network. If `ssf` is already
available on the maintainer's machine, compare the same cases manually; a
mismatch is evidence to investigate, but the committed corpus remains the local
gate because there is no Excel 365 golden yet.
