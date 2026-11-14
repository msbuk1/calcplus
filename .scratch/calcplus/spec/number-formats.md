# calcplus v1 — number formats

Resolved by [18 — Number-format engine: build vs buy](../issues/18-number-format-engine.md) on 2026-11-13. The grammar is bought; the Excel seams are owned.

## 1. Engine

- `numfmt` 4.1.3 (MIT) behind an internal `NumberFormatter` interface.
- `ssf` is a differential test oracle only; it never ships.
- `Intl.NumberFormat` alone cannot express the Excel grammar and is not used for cell formatting.

## 2. Owned Excel seams

| Seam | Behaviour |
|---|---|
| Built-in `numFmtId` table | 0-22, 37-49, and locale-specific 27-36, 50-58, 59-62, 67-81. Built-ins are never materialised; the ID is preserved. |
| 1904 adapter | Add 1462 days for calendar formats only; never for `[h]`, `[m]` or `[s]` elapsed formats. |
| Cell-width layout | Handle `*` fill, `_` skip, `?` digit alignment, `General`, and Excel's overflow fill — the cell is filled with `#` to its width when the value does not fit. |
| Colours and palette | Map `[Black]`, `[Blue]`, `[Cyan]`, `[Green]`, `[Magenta]`, `[Red]`, `[White]`, `[Yellow]` and indexed/theme colours to the workbook palette. Number-format colour codes are legacy **indexed** colours (research/18): `[Color n]` resolves against the workbook indexed palette. Theme-colour references inside number formats round-trip untouched but render through the indexed palette mapping — a pinned deviation. |
| Normalisation | Lowercase `e+` and `a/p`, and `[NatNum]`/`[DBNum]`, are normalised before formatting. |
| Grouping | Indian 3/2 grouping is supported. |
| Invalid format | Fall back to `General`; never throw and never show a raw code. |
| Round-trip | Preserve the authored code and `numFmtId` in the append-only style table; never renumber. |

## 3. Grammar coverage

- Sections: positive; negative; zero; text, separated by `;`.
- Conditions such as `[>100]` and `[<=0]`.
- Colours in brackets.
- Currency and accounting, including symbol placement and padding.
- Fractions, fixed and variable.
- Scientific notation.
- Elapsed time `[h]`, `[m]`, `[s]`.
- Text with `@` and literals.
- Padding with `*`, `_` and `?`.
- Locale decimal and thousands separators; English in v1.

## 4. Dates and times

- Dates and times are numbers with a format; there is no separate date type.
- 1900 and 1904 workbooks both work.
- Serial 60 is the fake 29 February 1900; the leap-year bug before 1 March 1900 is reproduced.
- Time fractions are formatted as fractions of a day; elapsed formats are not shifted by the 1904 adapter.

## 5. Rendering

- The formatter returns a string; the grid paints it.
- Alignment follows the value type and format: numbers right, text left, dates right, booleans centred as Excel does.
- If the string is wider than the cell, fill the cell with `#` to its width (Excel's overflow marker) unless the column is wide enough or the cell is wrapped. The number of hashes is width-dependent, so no golden may assert a fixed count.
- Colour from the format applies to the cell text unless a conditional-format font colour overrides it.
- The formula bar always shows the raw value or formula, not the formatted string.
- Formatting is cached per style ID and value; the cache is invalidated when the style or value changes.

## 6. Round-trip

- Built-in formats keep their `numFmtId` and are not written as custom codes.
- Custom codes are preserved exactly, including whitespace, colour brackets and conditions.
- The style table is append-only; existing IDs are never renumbered.
- A no-op save leaves `styles.xml` byte-identical.

## 7. Acceptance criteria

1. The conformance corpus of built-in and custom formats matches Excel's display text exactly.
2. Every built-in ID in the table renders as Excel does.
3. 1900 and 1904 workbooks render dates correctly, including serial 60.
4. Width layout produces Excel's `#` overflow fill when the value does not fit and never clips silently.
5. Colours, conditions and sections behave as Excel.
6. Invalid formats fall back to `General` without throwing.
7. `ssf` differential tests pass for the corpus.
8. Custom codes and built-in IDs round-trip through `.xlsx` without renumbering.
9. Formatting a 250k-cell sheet stays within the [performance](performance.md) budget.

## 8. Limits

- English locale semantics only in v1: locale-specific built-in IDs (59-62 and kin) render their **Gregorian en-US form**, and localized format codes and function names are shown but not re-rendered. numFmtId=59-62 Thai/Buddhist-era behaviour is therefore a **pinned v1 deviation** (PRD AC10), listed in the compatibility report. Localized formats and function names are v2.
- An invalid or unsupported imported format code falls back to `General`; Excel would not have accepted the code at entry, so there is no oracle for this case and it is pinned by a unit test.
- `numfmt`'s known lowercase `e+` and `a/p` bugs are handled by the seams and pinned by tests.
- Gradient and pattern fills are cell formatting, not number formats; they are handled by the style table.
