# 18 — Number-format engine: build vs buy

Researched 2026-11-13 against `numfmt` 4.1.3 (repo commit `32893df`, published 2026-11-10), SheetJS `ssf` 0.11.2, `Intl.NumberFormat` as implemented in Node 24.15.0 / ECMA-402, `@univerjs/sheets-numfmt` 1.4.0, `numbro` 2.5.0, and `excel-number-format` 1.2.11.

No Excel desktop installation was available in this environment, so the fidelity claims below come from (a) Microsoft/ECMA documentation, (b) the libraries' own Excel-derived test corpora, and (c) direct probes of the two libraries. §9 is explicit about what still needs an Excel-generated golden corpus.

## Recommendation

**Buy the grammar; build the Excel seams.**

Use **`numfmt` 4.1.3 (MIT)** as the value-to-string engine behind a small internal interface, and own the parts Excel does not put in the format-code string:

1. **Own the built-in `numFmtId` table.** Resolve IDs from a checked-in ECMA-376/ISO 29500 table, not from `numfmt.getFormatFromId()` for IDs above 58. Preserve the ID on save; never materialise a built-in ID into a custom `formatCode`.
2. **Own the 1904 adapter.** `numfmt` does not support the 1904 date system [S3]. Add 1462 to the serial for calendar/date-time formats, but not for elapsed-time formats (`[h]`, `[m]`, `[s]`). Set `dateSpanLarge:false`, `dateErrorNumber:false`, `dateErrorThrows:false` so out-of-range dates render the overflow string instead of a serial [S7].
3. **Own the cell-width layout.** `*` fill, `_` skip width, `?` alignment, `General`'s width-dependent truncation, and `#####` overflow are cell-rendering behaviours, not string formatting. `numfmt` deliberately behaves like Excel's `TEXT()` function for `*` and `_` (it emits no fill and a single space) [S3][S4]. The canvas `CellPainter` must measure and expand these.
4. **Own the locale policy.** `numfmt` bundles 77 locales and honours `[$-…]` separators/currency, but it does not automatically use Indian 3/2 grouping, and its `getFormatFromId()` collapses IDs 59–81 into IDs 1–23 [S5]. A small locale map and a built-in table fix both.
5. **Keep `ssf` only as a differential test oracle**, not as the production engine. It is Apache-2.0 and has strong date/time corpora, but its numeric custom-format path is incomplete and it has concrete fidelity bugs (§4.2).

Do **not** write an Excel-format parser from scratch for V1. The grammar is large, the edge cases are Excel-specific, and `numfmt` already carries a 53,082-case Excel-derived table suite [S11]. The work that remains is the seam layer above, which is substantial but bounded.

This also matches prior art: GenOffice's `.xlsx` gateway uses `numfmt`, and its Sheets renderer still carries a `numfmt-fix.ts` patch module [S48]. Univer's number-format package calls `numfmt.format()` and `numfmt.formatColor()` internally [S24]. Buying the engine does not eliminate the edge-case work; it moves it into a smaller, testable layer.

## 1. What the engine has to do

The v1 acceptance criterion is: "Rendering matches Excel over a conformance corpus of custom format strings; format codes round-trip" [S47]. The grammar includes, at minimum:

- four sections: positive; negative; zero; text, with omitted sections and blank sections [S32][S33][S38 §18.8.31]
- conditions (`[<=100]`, `[>100]`, `[<>0]`) and colours (`[Red]`, `[Color 3]`) [S32][S38 §18.8.31]
- currency and accounting (`$`, `[$€-407]`, `_)`, `* `) [S32][S33]
- fractions (`# ?/?`, `# ??/??`, fixed denominators) [S32]
- scientific (`0.00E+00`, `##0.0E+0`, `E-`/`e+`) [S32]
- elapsed time (`[h]:mm:ss`, `[mm]:ss`, `[ss].00`) [S32]
- text (`@`), literals (`"text"`, `\x`), padding (`?`, `_`, `*`) [S32][S33]
- locale modifiers (`[$-409]`, `[$€-407]`, `[$-F800]`, `[$-x-sysdate]`) and locale separators [S38 §18.8.30]
- built-in `numFmtId`s, including the locale-specific IDs 27–36, 50–58, 59–62, 67–81 [S37][S38 §18.8.30]
- 1900/1904 date systems, the fake 1900-02-29, and negative/zero/blank sections [S34][S35][S38 §18.17.4.1]
- round-trip of the authored format code and its `numFmtId`, including custom IDs and locale-specific built-ins [S38 §18.8.30–31][S40]

## 2. Candidates

### 2.1 `numfmt` 4.1.3 — MIT

Pure TypeScript/JavaScript, no runtime dependencies, ESM + CJS + types, `platform: neutral` [S2]. It exposes `format`, `formatColor`, `getFormatInfo`, `getFormatDateInfo`, `tokenize`, `getFormatFromId`, `isDateFormat`, `isPercentFormat`, `isTextFormat`, `isValidFormat`, and `parseValue` [S4]. It bundles locale tables for 77 locales and does not call `Intl` at runtime (source inspection; the locale data is in `src/locales.ts`).

It has a 1,000-entry LRU parse cache [S8] and a large test suite: the `getFormatFromId` spec alone contains 4,055 `expect()` calls [S10], the conditional spec 561 [S9], and the maintainer reports 53,082 passing Excel-table tests [S11]. Recent releases (Oct–Nov 2026) fixed Excel-fidelity bugs in `General`, fraction carry, and midnight carry [S11][S12][S49].

Known gaps, verified by probe or source:

- **No 1904 support.** The README says so explicitly [S3]; the public `dateFromSerial()` only takes `leap1900` [S6]. There is an internal `toYMD_1904()` helper [S6], but no public option.
- **Built-in IDs above 58 are wrong.** `getFormatFromId()` computes `1 + ((fmtId - 1) % 58)` [S5], so 59→1, 67→9, 81→23, and 59–62/67–80 land on unrelated numeric or date/time formats. For Thai (`th-TH`), ID 59 should be `t0` but returns `0`; ID 71 should be `ว/ด/ปปปป` but returns `# ??/??` [S5][S37].
- **Lowercase scientific `e+`/`e-` fails.** `format("0.00e+00", 12200000)` throws `Illegal format`; Excel accepts lowercase `e+`/`e-` [S32]. Lowercase `a/p` does **not** throw; it is silently treated as literal text (`"12:32 a/p"` instead of `"12:32 P"`). Both need normalisation before calling `numfmt`.
- **`[NatNum1]` throws** by default; `[DBNum2]` is accepted but is a no-op (digits are not converted) [S9]. These are East Asian modifiers, not part of the Microsoft custom-format guidance [S32]; they matter only if CJK numeral rendering is in scope.
- **Era names are not localised.** `ggge"年"m"月"d"日"` renders `2024年1月1日`, not the Japanese era (`令和6年…`); `[$-404]e/m/d` renders the Gregorian year, not the ROC year. `B1`/`B2` Buddhist/Hijri year conversion does work.
- **Indian grouping is not automatic.** `format("#,##0", 1234567, {locale:"en-IN"})` gives `1,234,567`; Excel's en-IN convention is `12,34,567`. `numfmt` can do it with `grouping:[3,2]` [S4], so this is a wrapper config issue, not a grammar gap.
- **`*`/`_` are `TEXT()` semantics.** `format("0*-", 3)` gives `"3"`; `fillChar` emits the fill character once, not to the cell width. `format("_(* #,##0_)…", 1234)` emits a single space, not the width of `(` [S3][S4].

### 2.2 SheetJS `ssf` 0.11.2 — Apache-2.0

Pure JavaScript, one runtime dependency (`frac`) [S16]. It is the formatter inside SheetJS; the standalone repository says the project has moved into the main SheetJS repo [S17]. The npm package is 0.11.2, published 2020-06-27 [S21]; the standalone GitHub repo's last push was 2023-02-23 [S15]. The current source is in the SheetJS monorepo on `git.sheetjs.com`, which required authentication to inspect; this evaluation is of the published 0.11.2 source [S18].

Strengths:

- Exact date/time corpora: `test/date.js` compares `SSF.format()` against 295,849 date rows and 1,048,576 time rows [S20].
- Built-in table for common IDs, `default_map`/`default_str` fallbacks, `SSF.load()`/`SSF.load_table()` for custom built-in tables [S18].
- `date1904` option is supported: `parse_date_code()` adds 1462 days [S18].
- Elapsed time (`[h]`, `[m]`, `[s]`) works.
- Apache-2.0, so it can be used as a test oracle or vendored.

Weaknesses, verified by probe or source:

- **Numeric custom-format coverage is narrow.** `write_num()` is a set of regex special cases and throws `"unsupported format"` for anything outside them [S18]. `???.???` throws; `# ???/???` works only because it matches a specific regex.
- **Conditions are only checked in the first two sections and do not implement Excel's "no match → #####" rule.** `choose_fmt()` only looks at `fmt[0]`/`fmt[1]` [S18]. `[>100]0;[<0]0` with 50 returns `"50"`; Excel shows `#####`. `[Red][<=100];[Blue][>100]` with 50 returns `""`; Excel shows `50` in red.
- **Colours are discarded.** `eval_fmt()` handles `[h]` and `[$…]` in brackets but has no branch for `[Red]`/`[Color n]` [S18]. There is no colour API.
- **Locale modifiers are only partially understood.** `[$-409]#,##0.00` returns `"$1,234.56"` (it treats `$-409` as a dollar currency marker); `[$€-407] #,##0.00` returns `"€ 1,234.56"` (en-US separators, not German) [S18]. There is no locale option on `format()` [S18].
- **`?` alignment and `*` fill are lost.** `0.0?` returns `"8.9"` without the alignment space; `0*-` returns `"3-"` (the next character once, not repeated) [S18].
- **Scientific exponent padding is missing.** `0.000E+000` returns `1.220E+7`, not `1.220E+007` [S18].
- **Rounding differs from Excel.** `SSF.format("0.00", 1.005)` returns `"1.00"`; Excel (and `numfmt`) return `"1.01"` because Excel rounds from 15 significant digits before applying the display places. `numfmt.round()` documents this 15-significant-digit behaviour [S4].
- **Negative date serials return `""`.** `parse_date_code()` returns `null` for `v < 0`, and `eval_fmt()` returns `""` [S18]. Excel shows `#####`.
- **`valid.tsv` only asserts "does not throw".** `test/valid.js` loops over 442 format strings and calls `SSF.format()` with a handful of values but never checks the output [S19]. That is much weaker than an exact-output conformance suite.
- **Built-in IDs above 58 are also approximate.** `format(59, 45292)` returns `"45292"`; `format(67, 45292)` returns `"4529200.00%"`; `format(81, 45292)` returns `"0000.0"` [S18].

### 2.3 `Intl.NumberFormat` plus a custom Excel-format parser

`Intl.NumberFormat` is native, zero bundle, and excellent at locale-aware decimal/group separators, currency, accounting parentheses, percent, compact notation, scientific/engineering notation, and rounding modes [S42][S43]. It does not parse Excel format codes, and it cannot express the Excel grammar directly:

- No sections, conditions, colours, literal placement, `_`/`*` padding, `?` alignment, fractions, elapsed time, or text sections [S42][S43].
- Scientific notation is `1.22E7`, not Excel's `1.22E+07`; there is no exponent-digit padding [S43].
- Accounting is `currencySign:"accounting"` and gives `($1,234.50)`, but cannot express Excel's `_($* #,##0.00_);…` padding and zero-dash section [S43].
- There is no fraction style; `style:"fraction"` is a `RangeError` (probe).
- Output can vary with the browser/ICU version; ECMA-402 marks some behaviour optional [S42].

A custom parser would still have to implement the entire tokenizer, section selection, conditions, colours, fractions, elapsed time, text, and padding. At that point `Intl` is only a locale-separator helper, not an engine. It is useful as a fallback for locales `numfmt` does not bundle, but it does not reduce the parser work.

### 2.4 Univer `@univerjs/sheets-numfmt` 1.4.0 — Apache-2.0

This is not an independent engine. The published package imports `numfmt` from `@univerjs/core` and calls `numfmt.format()`, `numfmt.formatColor()`, and `numfmt.getFormatInfo()` [S24]. It depends on `@univerjs/core`, `@univerjs/engine-formula`, and `@univerjs/sheets` [S23], so adopting it means adopting Univer's plugin/DJ architecture. Reject as a standalone option; it is evidence that `numfmt` is the de facto JS Excel formatter.

### 2.5 `numbro` 2.5.0 — MIT

`numbro` is a general number formatter with an options object (`{thousandSeparated, mantissa, average, output:"percent"}`), not an Excel format-code parser [S27]. It has no sections, conditions, colours, fractions, elapsed time, or `[$-…]` grammar. MIT, 1,141 stars, 270 open issues, last push 2025-10-17 [S26]. Reject.

### 2.6 `excel-number-format` 1.2.11 — MIT

A tokenizer/classifier for Excel's format mini-language; its README says explicitly that it is "a classifier, not a formatter" and that rendering needs locale data, fill/alignment geometry, conditional-section evaluation, and colour handling [S28]. Useful as a reference for tokenization and as a second opinion in tests; not a renderer.

### 2.7 HyperFormula's `TEXT()` formatter

HyperFormula is the chosen formula engine, but its documented format-code surface is the `TEXT()` function plus `stringifyDateTime`/`stringifyCurrency` callbacks [S30][S31]. It is not a cell number-format renderer, and its `TEXT()` supports only a subset of Excel's formats: the differences page lists only some date tokens (`hh`, `mm`, `ss`, `am`, `pm`, `a`, `p`, `dd`, `yy`, `yyyy`) and says currency formatting inside `TEXT` is not supported [S29]. Do not use it as the grid formatter. If calcplus needs Excel-compatible `TEXT()` results, the same `numfmt`-based function must be exposed to the formula layer (or `TEXT` must be overridden/aliased); otherwise `TEXT()` and the grid will disagree.

## 3. Comparison

| Criterion | `numfmt` 4.1.3 | SheetJS `ssf` 0.11.2 | `Intl.NumberFormat` + parser | Univer `sheets-numfmt` | `numbro` |
|---|---|---|---|---|---|
| Sections / conditions | Yes; 561 conditional assertions [S9] | Partial; first two sections only, no-match rule wrong [S18] | No | Via `numfmt` | No |
| Colours (`[Red]`, `[Color n]`) | Yes; `formatColor()` [S4] | No; brackets discarded [S18] | No | Via `numfmt` | No |
| Currency / accounting | Yes; `[$€-407]`, `_)`, `* ` | Basic `$`; locale/accounting wrong [S18] | Currency + accounting, no padding | Via `numfmt` | Currency options |
| Fractions | Yes; `# ?/?`, `# ??/??`, fixed denominators | Basic; regex-limited [S18] | No | Via `numfmt` | No |
| Scientific | Yes; exponent padding, `E+`/`E-`; lowercase `e+` fails | Basic; no exponent padding; lowercase `e+` wrong/empty | `notation:"scientific"`, output `1.22E7` | Via `numfmt` | No |
| Elapsed time | Yes; `[h]`, `[m]`, `[s]` | Yes | No (would need `Intl.DurationFormat` + custom) | Via `numfmt` | No |
| Text / literals / padding | Yes, but `*`/`_` are `TEXT()` semantics | Text/literals yes; `?`/`*` padding lost | No | Via `numfmt` | No |
| Locale separators / `[$-…]` | Yes, 77 locales; Indian grouping needs config | Only extracts a currency symbol; separators stay en-US [S18] | Yes, CLDR locales | Via `numfmt` | Yes for its own options |
| Built-in `numFmtId` table | 0–58 only; 59–81 wrong [S5] | Partial 0–58; 59–81 approximate [S18] | None | Via `numfmt` | None |
| 1900/1904 | 1900 + leap bug; **no 1904** [S3] | 1900 + leap bug; 1904 option [S18] | No serial-date model | Via `numfmt` | No |
| Negative/zero/blank sections | Yes | Basic yes; conditional no-match wrong | No | Via `numfmt` | No |
| Round-trip | Engine does not write xlsx; preserve code/ID in style table | Same | Same | Coupled to Univer | Same |
| Browser fit | Pure JS, ESM/CJS, no deps | Pure JS, CJS, `frac` dep | Native | Requires Univer core | Pure JS |
| Bundle (shipped, gzip) | ~46.7 KB (167.5 KB raw) | ~10.4 KB (38.3 KB raw) + `frac` | 0 KB | Large (Univer core) | Moderate |
| Performance (synthetic, Node) | ~0.40–0.55 M ops/s | ~0.60–0.76 M ops/s on mixed corpus | ~1.0 M ops/s for simple mixed | Same as `numfmt` | Not measured |
| Licence | MIT [S2] | Apache-2.0 [S16] | Native | Apache-2.0 [S23] | MIT [S26] |
| Maintenance | Active; 70 stars, 5 open issues, releases Oct–Nov 2026 [S1][S14] | Standalone stale (2020); moved into SheetJS [S17][S21] | Browser/ICU | Active | Active; 270 open issues [S26] |

`✅`/`❌` are avoided because several cells are "works for the common case, fails the long tail"; the notes are the point.

## 4. Findings

### 4.1 Grammar coverage

`numfmt` covers the full Microsoft-documented grammar in the common cases: four sections, conditions, colours, currency, accounting, fractions, scientific, elapsed time, text, literals, `?`/`_`/`*`, scaling commas, percent, and locale modifiers. The conditional test file alone exercises 561 combinations of operators, signs, and section counts [S9]. It also handles `B1`/`B2` Buddhist/Hijri calendars and `[$-F800]`/`[$-x-sysdate]` system date formats [S10].

The gaps are narrow but real:

- **Lowercase tokens.** Excel documents `E-`, `E+`, `e-`, `e+` and `AM/PM`, `am/pm`, `A/P`, `a/p` [S32]. `numfmt` handles `E+` and `A/P`, but throws on `e+` and silently treats `a/p` as literal. This is a pre-processing fix, not a reason to reject the library.
- **East Asian numeral modifiers.** `[DBNum2]` is a no-op; `[NatNum1]` throws [S9]. These are not in Microsoft's custom-format guidance [S32] and are outside the v1 localisation scope [S47]. The engine should strip them or render ASCII digits rather than crash.
- **Era names.** `g`/`gg`/`ggg`/`ggge` and `e` are tokenised, but the output is Gregorian year, not the Japanese/ROC era name. `B1`/`B2` year conversion is supported. If v1 does not localise, preserve the code and accept the Gregorian fallback; document it.
- **Indian grouping.** `numfmt` has a `grouping` option but does not derive 3/2 grouping from `en-IN`/`hi-IN`/`bn-IN`. A small locale map (`[3,2]` for those locales) fixes it [S4].

`ssf` covers the common numeric and date/time cases but is not a full grammar implementation. Its `write_num()` is a collection of regexes with explicit `throw new Error("unsupported format")` fallbacks [S18]. The probe found `???.???` throws, `?` padding is lost, `*` is not repeated, and conditional-only sections return `""`. Its date/time engine is much stronger than its numeric engine, because the date/time tests compare exact output over ~1.3 M rows [S20].

`Intl.NumberFormat` is not a grammar implementation at all. It is a locale-aware number formatter with a fixed option set; it cannot express per-section formats, conditions, colours, literal placement, fractions, or elapsed time [S42][S43].

### 4.2 Fidelity against Excel

The strongest evidence for `numfmt` is its Excel-derived test corpus. PR #110 reports "vitest run with the Excel tables passes (53082 tests)" [S11]. The repository contains `test/tables/cal-updated.tsv.gz` (12 MB), `ssf-dates.tsv.gz` (2.1 MB), and `ssf-times-full.tsv` (69 MB), plus a 259 KB `getFormatFromId.spec.ts` with 4,055 expectations [S10]. These are not independent of the library, but they encode Excel outputs, and the maintainer has been actively fixing mismatches found by comparing to Excel (e.g. `General` scientific notation [S12], fraction carry [S11], midnight carry [S49]).

The strongest evidence for `ssf` is its date/time corpora: `test/date.js` compares `SSF.format()` to `dates.tsv` (295,849 rows) and `times.tsv` (1,048,576 rows) for every date/time header [S20]. That is excellent coverage for date/time. Its numeric corpus `valid.tsv` only asserts that 442 formats do not throw [S19], so it says little about numeric output fidelity.

A direct probe against both libraries shows the practical differences. Values and formats are from Microsoft's documented examples where possible [S32][S33]; outputs are the actual strings from `numfmt` 4.1.3 and `ssf` 0.11.2.

| Format | Value | Excel (documented/expected) | `numfmt` | `ssf` |
|---|---:|---|---|---|
| `#,##0.00` | 1234.56 | `1,234.56` | `1,234.56` | `1,234.56` |
| `0.00E+00` | 12200000 | `1.22E+07` | `1.22E+07` | `1.22E+07` |
| `0.000E+000` | 12200000 | `1.220E+007` | `1.220E+007` | `1.220E+7` |
| `# ?/?` | 5.25 | `5 1/4` | `5 1/4` | `5 1/4` |
| `???.???` | 44.398 | aligned ` 44.398` | ` 44.398` | throws |
| `#.0,` | 12200000 | `12200.0` | `12200.0` | `12200.0` |
| `0.0,,"M"` | 12200000 | `12.2M` | `12.2M` | `12.2M` |
| `[Red][<=100];[Blue][>100]` | 50 | `50`, red | `50`, red | `""` |
| `[>100]0;[<0]0` | 50 | `#` fill | `######` | `50` |
| `[Blue]#,##0.00_);[Red](#,##0.00);0.00;"sales "@` | -1234.5 | `(1,234.50)`, red | `(1,234.50)`, red | `(1,234.50)` |
| `0;-0;;@` | 0 | empty | empty | empty |
| `;;;` | 123 | empty | empty | empty |
| `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | 1234 | padded accounting | ` 1,234 ` | ` 1,234 ` |
| `0*-` | 3 | fill to cell width | `3` | `3-` |
| `0.0?` | 8.9 | `8.9 ` (alignment space) | `8.9 ` | `8.9` |
| `0.00` | 1.005 | `1.01` | `1.01` | `1.00` |
| `[$-409]#,##0.00` | 1234.56 | `1,234.56` | `1,234.56` | `$1,234.56` |
| `[$€-407] #,##0.00` | 1234.56 | `€ 1.234,56` | `€ 1.234,56` | `€ 1,234.56` |
| `yyyy-mm-dd` | 45292 | `2024-01-01` | `2024-01-01` | `2024-01-01` |
| `[h]:mm:ss` | 1.5 | `36:00:00` | `36:00:00` | `36:00:00` |
| `[h]:mm:ss` | -1 | `#` fill | `######` | `""` |
| `yyyy-mm-dd` (1904 serial 0) | 0 | `1904-01-01` | `1900-01-00` (no 1904) | `1904-01-01` |
| `0.00e+00` | 12200000 | `1.22e+07` | throws | `""` |
| `h:mm a/p` | 0.5222625 | `12:32 P` | `12:32 a/p` | `12:32 P` |

The `numfmt` "no 1904" row is the only case where the recommended library is clearly behind `ssf`; the adapter in §5.2 fixes it. The lowercase `e+`/`a/p` rows are pre-processing fixes. The `ssf` rows show why it is a test oracle, not the production engine.

Caveat: Excel's exact output for a few of these (notably `##0.0E+0` at zero, error/text-section interaction, and `General` at narrow widths) was not checked against Excel desktop. Those are listed in §9.

### 4.3 Round-trip of format codes and `numFmtId`

Neither library writes `.xlsx`; round-trip belongs to the style table and the OOXML writer (tickets 04/12). The engine's job is to accept a format reference and never mutate it.

The file model should be:

```ts
type NumberFormatRef =
  | { kind: "builtin"; id: number }        // no formatCode stored
  | { kind: "custom"; id: number; code: string };
```

Rules:

1. **Preserve the authored `formatCode` exactly as decoded.** Do not round-trip through `numfmt.tokenize()`/`delocalize()` or `ssf` and re-emit a canonical string. ExcelJS already demonstrated this hazard: a fixture's `\$#,##0.00` became `$#,##0.00` in ticket 04's evaluation [S46]. Store the decoded XML string unchanged, and XML-escape it only on write.
2. **Preserve the `numFmtId`.** Built-in IDs must stay built-in. A cell with `numFmtId="14"` must not be rewritten as a custom `m/d/yyyy` code just because the UI displayed that string. If the user opens and closes Format Cells without editing, the style XML must not change.
3. **Append-only style table.** New `numFmt` entries go at the end; existing IDs are never renumbered. This matches ticket 04's append-only rule [S46]. Allocate new custom IDs above the current maximum, and at least 164 (Excel's BIFF8 custom-format base is `0x00A4`; xlsx writers conventionally start at 164) [S40]. The xlsx schema only constrains `numFmtId` to `unsignedInt` [S38 §18.18.47], so preserving an unusual existing ID is safer than "fixing" it.
4. **Do not deduplicate existing entries.** Two custom IDs may carry the same code; both are referenced by different `xf` records. Leave them. New entries may reuse an existing code/ID only if the user explicitly selects it; otherwise append.
5. **Preserve orphan/unreferenced `numFmt` entries** in untouched `styles.xml`. They are part of the byte-identical promise.
6. **Missing/unknown IDs.** If an `xf` references a `numFmtId` that is neither in the built-in table nor in `<numFmts>`, render `General` (or `#####` for a date-like unknown) and preserve the reference. Do not crash, and do not delete the `xf`.
7. **Locale-specific built-ins.** The file stores only the ID; the spec says these IDs "can be interpreted differently, depending on the UI language of the implementing application" [S38 §18.8.30]. Preserve the ID. Choose a rendering locale policy (v1: en-US, since localisation is out of scope [S47]) and document that a `numFmtId="59"` Thai cell will not look Thai in an en-US UI. If the app later adds a Thai UI, swap the table, not the file.

`numfmt.getFormatFromId()` is useful for IDs 0–58, but not sufficient:

- It normalises IDs modulo 58 [S5], so 59–81 are wrong.
- For 27–36 and 50–58 it returns a generic locale date/time (e.g. `getFormatFromId(27, "ja-JP")` → `yyyy/mm/dd`) rather than the exact ECMA code (`[$-411]ge.m.d`) [S5][S37].
- It has no way to represent the full Thai table (59–62, 67–81) [S37].

Therefore: **check in the ECMA-376/ISO 29500 built-in table**, including the locale columns, and use it to resolve IDs. It is about 80 entries plus locale variants, small enough to review and test. Use `numfmt.getFormatFromId()` only as a cross-check for 0–58, or not at all.

### 4.4 1900/1904 date systems

Excel's 1900 system starts at serial 1 = 1900-01-01, includes the fake 1900-02-29 at serial 60 (the Lotus 1-2-3 leap-year bug), and ends at 9999-12-31 = 2958465 [S34][S35]. The 1904 system starts at serial 0 = 1904-01-01; the same date is 1462 days lower than in the 1900 system [S34]. The workbook declares its system in `<workbookPr date1904="1"/>` [S38 §18.17.4.1][S41].

There is a spec/Excel discrepancy worth knowing: ECMA-376 says the 1900 base date is 1899-12-30 and serial 1 is 1899-12-31 [S38 §18.17.4.1], while Excel's UI and documentation say serial 1 is 1900-01-01 and serial 60 is 1900-02-29 [S34][S35]. For fidelity, follow Excel, not the ECMA arithmetic. `numfmt`'s `leap1900:true` default does this [S6][S7]; `ssf` also implements the fake day [S18].

`numfmt` has no public 1904 option [S3][S6]. The adapter:

```ts
const info = getFormatInfo(code);
const isDuration = tokenize(code).some(t => t.type === TOKEN_DURATION);
const isCalendar = info.type === "date" || info.type === "datetime" || info.type === "time";
const serial = (date1904 && typeof value === "number" && isCalendar && !isDuration)
  ? value + 1462
  : value;
```

Why the conditions:

- Date, datetime, and time-of-day formats: adding 1462 whole days shifts the calendar date without changing the time-of-day fraction (1462 is an integer number of days). This is correct for 1904 serials.
- Elapsed-time formats (`[h]`, `[m]`, `[s]`): adding 1462 would add 35,088 hours to the elapsed value. Do not shift. `getFormatInfo()` reports `type:"time"` for both `h:mm` and `[h]:mm:ss` [S4], so use the `TOKEN_DURATION` check, not `type`.
- `General`: no shift. A bare serial under General is a number; a date cell should carry a date format. (If the model tags a formula result as a date with General, that is a formula-layer concern, not the formatter's.)

Also set these `numfmt` options for Excel-like bounds and error behaviour [S7]:

```ts
{
  locale: "en-US",
  dateSpanLarge: false,      // Excel bounds, not Google Sheets' wider range
  dateErrorNumber: false,    // do not silently print the serial for out-of-range dates
  dateErrorThrows: false,    // do not throw; emit the overflow string
  throws: false,             // imported invalid formats must not crash the grid
  leap1900: true             // Excel's fake 1900-02-29
}
```

Probe results with those options:

| 1904 serial | Raw `numfmt` | Adapter (+1462) | Expected |
|---:|---|---|---|
| 0 | `1900-01-00` | `1904-01-01` | `1904-01-01` |
| 1 | `1900-01-01` | `1904-01-02` | `1904-01-02` |
| 43830 | `2019-12-31` | `2024-01-01` | `2024-01-01` |
| 45292 | `2024-01-01` | `2028-01-02` | `2028-01-02` |

`ssf` handles 1904 directly (`date1904:true`) and is a useful differential oracle for this adapter [S18]. But `ssf` returns `""` for negative dates where Excel shows `#####` [S18], so it is not the production path.

### 4.5 Browser fit, bundle size, and canvas-render performance

All three serious options are browser-compatible:

- `numfmt`: pure JS/TS, no dependencies, ESM + CJS + `.d.ts`, `platform: neutral` [S2]. It bundles locale data, so output is deterministic across browsers (no ICU dependency). That is a feature for a fidelity clone.
- `ssf`: pure JS, CJS, one dependency (`frac`) [S16]. Bundlers handle it; it is not ESM-first.
- `Intl`: native, zero bytes, but ICU-dependent.

Measured shipped sizes (this research, `gzip -c` of the package entry file, before app bundler minification):

| Package | Raw | Gzip |
|---|---:|---:|
| `numfmt` `dist/index.js` | 167,543 B | 46,730 B |
| `ssf` `ssf.js` | 38,297 B | 10,408 B |
| `ssf` + `frac` | ~94 KB | ~14 KB (estimate) |

`numfmt` is ~4.5× larger gzipped because it ships 77 locale tables. 46.7 KB is acceptable for a spreadsheet app; if the bundle budget is tight, the MIT licence permits a fork with only the locales calcplus supports. Do not assume tree-shaking removes the locale tables: the single entry imports them for side effects.

Synthetic performance (Node 24, 300k iterations, mixed formats, best of 5 after warm-up; **indicative only**, not a browser benchmark):

| Engine | Mixed corpus | Simple `#,##0.00` |
|---|---:|---:|
| `numfmt` (explicit locale) | ~0.40 M ops/s | ~0.45 M ops/s |
| `numfmt` (default locale) | ~0.55 M ops/s | — |
| `ssf` | ~0.76 M ops/s | ~0.62 M ops/s |
| `Intl` (reused formatters) | ~1.0 M ops/s | — |

`ssf` is faster on this corpus, partly because it takes simpler code paths and returns different (sometimes wrong) output; do not read the number as "ssf is more efficient at the same work". `numfmt`'s explicit locale lookup costs ~25%; pass a resolved locale/config object, or cache the formatter context, rather than a string on every cell.

For a canvas renderer, the absolute numbers matter less than the integration:

- A 10,000-cell visible viewport at 0.4 M ops/s is ~25 ms of formatting on a full repaint; at 0.55 M it is ~18 ms. That is too slow to redo every frame while scrolling, but fine for the initial paint.
- Cache the formatted `CellDisplay` per cell and invalidate on value/style/format change. The renderer already repaints only dirty regions (ticket 05) [S45]; formatting belongs in the `SheetView` projection, not in `CellPainter` [S45].
- `numfmt` caches parsed patterns (1,000-entry LRU) [S8], so the per-cell cost is mostly number/date formatting, not parsing. Reuse the same options object and resolved locale.
- `formatColor()` resolves the colour from the pattern (the same parse cache applies). Only call it when the pattern contains a `[` colour/condition token; most cells do not.
- Do not pass JS `Date` objects. Pass Excel serial numbers; `numfmt`'s `Date` handling applies local timezone/DST rules [S4], which is not what a workbook serial means.

The real canvas cost is not `numfmt`; it is the layout layer:

- `*x` must repeat `x` until the cell's remaining width is filled.
- `_x` must advance by the measured width of `x`, not one space.
- `?` already emits alignment spaces, but they must survive text measurement and not be trimmed.
- `General` is width-dependent: Excel displays up to 11 characters, switches to scientific notation, or shows `#####` if the value cannot fit [S38 §18.8.30]. `numfmt`'s `General` is width-independent. A first implementation can use `#####` when the formatted string exceeds the cell width, but that will not match Excel's "show fewer digits" behaviour in every case.
- If the formatted numeric/date string exceeds the cell width, Excel paints `#` across the cell, not the string [S38 §18.8.30]. The painter must implement this.

### 4.6 Licensing and maintenance

- **`numfmt`**: MIT [S2]. Active: 70 stars, 17 forks, 5 open issues, last push 2026-11-10, release 4.1.3 on 2026-11-10 [S1][S14]. Single maintainer; recent rapid changes. Pin the exact version, keep the wrapper thin, and be prepared to vendor/fork (MIT allows it). The open PR for negative-serial weekdays (#112) shows the long tail is still being found [S13].
- **`ssf`**: Apache-2.0 [S16]. The standalone package is stale (0.11.2, 2020-06-27) [S21], but the code moved into the active SheetJS repo [S17][S22]. For a test oracle, the published package is fine; for production, the maintenance path is less clear than `numfmt`'s.
- **Univer `sheets-numfmt`**: Apache-2.0 [S23], but it is a plugin over Univer's core, not a standalone formatter. Reject.
- **`numbro`**: MIT [S26], active, but not an Excel grammar engine.
- **`Intl`**: no licence, native; ICU version variability is the "maintenance" risk.

## 5. Constructs that need custom work

These are the items to put in the build plan. Each is small enough to test independently, and each is a place where a library-only solution will look right in demos and wrong in Excel.

### 5.1 Built-in `numFmtId` table

Check in the ECMA-376/ISO 29500 table:

- all-languages IDs: 0–22, 37–49 (with 23–26 reserved/implementation-defined; 5–8 follow `formatCode` if present) [S37][S38 §18.8.30]
- locale-specific IDs: 27–36, 50–58 (zh-TW/zh-CN/ja-JP/ko-KR), 59–62 and 67–81 (th-TH) [S37]
- a policy for IDs 5–8 when no `formatCode` is present (if a `numFmt` entry exists, use its `formatCode`; otherwise use the locale currency/accounting pattern)
- a policy for unknown IDs: render `General`, preserve the ID

Do not use `numfmt.getFormatFromId()` for IDs > 58. Use it, if at all, as a cross-check for 0–58.

### 5.2 1904 date-system adapter

As in §4.4. Detect `<workbookPr date1904="1"/>`; add 1462 for calendar/date-time formats; skip elapsed-time formats; set `dateSpanLarge:false`, `dateErrorNumber:false`, `dateErrorThrows:false`. Test serials 0, 1, 59, 60, 61, 1462, 2957003 (1904 max), and negative serials.

### 5.3 Cell-width layout

Implement in the `SheetView`/`CellPainter` boundary:

- `*x`: repeat `x` to fill the remaining cell width. Use `ctx.measureText(x).width`; cache metrics by font+char.
- `_x`: advance by `measureText(x).width`; do not draw `x`.
- `?`: keep the emitted spaces; do not trim before measuring.
- `#####`: if the formatted numeric/date string is wider than the cell's content box, paint `#` across the cell width. Excel uses this for numbers and dates, not for text (text overflows/clips instead) [S38 §18.8.30].
- `General`: start with `numfmt`'s General; if it does not fit, fall back to `#####` for v1 and record the divergence. A later refinement can re-run General with fewer significant digits, matching Excel's 11-character heuristic [S38 §18.8.30].
- Accounting alignment: `_` and `*` are how Excel aligns currency symbols and parentheses across a column. The layout layer, not the formatter, owns this.

A practical adapter: use `numfmt`'s `fillChar`/`skipChar` options with private-use sentinels, then replace `sentinel + char` in the output with the measured fill/skip. Test that the sentinel cannot collide with imported format text; if it can, use a token-aware pre-scan instead. Do not rely on `numfmt` to repeat to width; it does not know the cell width [S3][S4].

### 5.4 Colour and palette

- Use `formatColor(code, value, { indexColors: false })` to get either a named colour (`"red"`) or a numeric index for `[Color n]` [S4].
- Named colours map to the eight legacy names; canvas accepts CSS colour names directly.
- `[Color n]` indexes the legacy palette (1–56) [S38 §18.8.31]. If the workbook has a custom `<indexedColors>` palette in `styles.xml`, map the index through it; otherwise use the default palette. Do not use theme colours here; number-format colours are legacy indexed colours, not theme colours.
- `formatColor()` returns the colour of the selected section (including conditions). If no colour is specified, it returns `undefined`; use the cell's font colour.

### 5.5 Token normalisation and unsupported modifiers

Before calling `numfmt`, normalise only outside quoted strings/escapes:

- `e+`/`e-` → `E+`/`E-` when used as an exponent (the lowercase form currently throws). If the Excel golden shows that Excel preserves the authored case, restore `e` in the exponent output after formatting; the important part is not throwing.
- `a/p` → `A/P` (the lowercase form is currently treated as literal text, which is worse than throwing). Verify whether Excel preserves `a/p` case in the output; if so, restore it after formatting.
- `[NatNum…]` and `[DBNum…]`: either strip them and render ASCII digits, or implement a digit map. Do not let `[NatNum1]` throw into the grid. For v1, strip and log once per format; document that CJK numeral shaping is not supported.
- `[Color n]` is supported; do not strip.
- `[$-…]` is supported; do not strip. Unknown LCIDs fall back to en-US in `numfmt` (probe: `[$-9999]` → en-US), which is a reasonable default.

If normalisation changes the code, keep the original code for storage and use the normalised code only for rendering. The Format Cells dialog must show the authored code, not the normalised one.

### 5.6 Locale policy and grouping sizes

- Default v1 UI locale: en-US, since localisation is out of scope [S47]. Still honour explicit `[$-…]` modifiers in imported codes, because that is part of the authored format, not UI localisation.
- Map Indian locales (`en-IN`, `hi-IN`, `bn-IN`, and related) to `grouping:[3,2]`; `numfmt` does not derive this from the locale [S4].
- For locales `numfmt` does not bundle (it has 77), fall back to en-US separators but preserve the code. Do not use `Intl` to rewrite the code; at most use it to choose a separator for a locale the engine does not know.
- Decide whether to implement `[NatNum]`/`[DBNum]` digit shaping and era names (`g`/`e`). Recommendation: defer; they are localisation features, and the feature matrix puts localisation out of scope. Preserve the codes and render a Gregorian/ASCII fallback.

### 5.7 Invalid/unsupported format fallback

- Call `numfmt` with `throws:false` so an imported malformed code renders the overflow string (`######`) instead of throwing [S7].
- Wrap every call in a try/catch anyway; a formatter bug must not take down a repaint.
- If the code is invalid but the cell is text, show the text; if numeric/date, show `######`.
- Log invalid codes once per workbook (not per cell) for the conformance report.

### 5.8 Round-trip style table

As in §4.3. This is mostly ticket 04/12 work, but the number-format engine's API must not make it impossible:

- Accept `{id, code?}` and never mutate it.
- Expose a separate `formatCodeForDisplay(ref)` for the dialog; do not write its output back unless the user edits.
- Do not use `numfmt.delocalize()` for storage; it is a normaliser, not a round-trip function [S4].

### 5.9 `TEXT()` consistency

HyperFormula's `TEXT()` is not Excel-complete [S29][S30]. If calcplus exposes `TEXT()` and the grid, they must use the same engine or users will see `=TEXT(A1,"0.00E+00")` differ from the cell display. Options:

- override/alias HyperFormula's `TEXT` with a custom function that calls the calcplus formatter (verify HyperFormula's custom-function/alias mechanism supports replacing a built-in name before committing);
- or accept the divergence in v1 and document it in the formula-function matrix.

This is a ticket 02/03 decision, but it is a direct consequence of the number-format choice. Flag it now.

### 5.10 Conformance corpus and performance harness

See §7.

## 6. Proposed integration

Keep the engine behind an interface so `numfmt` is replaceable and so the Excel seams are testable:

```ts
type NumberFormatRef =
  | { kind: "builtin"; id: number }
  | { kind: "custom"; id: number; code: string };

interface FormatContext {
  locale: string;          // v1 default "en-US"
  date1904: boolean;
  cellWidthPx: number;     // content box, after padding/borders
  font: string;            // canvas font string
}

interface FormattedCell {
  text: string;            // after fill/skip expansion
  color?: string;          // CSS colour or undefined
  overflow: boolean;       // true if Excel would paint #####
}

interface NumberFormatEngine {
  format(value: number | string | boolean, ref: NumberFormatRef, ctx: FormatContext): FormattedCell;
  formatCodeForDisplay(ref: NumberFormatRef): string; // dialog only; never written back
  isValid(code: string): boolean;
}
```

Data flow, aligned with ticket 05:

1. Workbook model stores `NumberFormatRef` and the raw code/ID.
2. `SheetView` projects each visible cell: value + ref + style + width/font → `FormattedCell`. It caches the result and invalidates on value/style/format/width change.
3. `CellPainter` draws `text` with `color`; it does not call `numfmt`.
4. The Format Cells dialog uses `formatCodeForDisplay` and `getFormatInfo`-style classification; edits create a new custom ref (append-only), not a rewrite of the built-in ID.

Implementation notes:

- Resolve built-in IDs through the checked-in table; cache the resolved code per ID/locale.
- Cache a `numfmt` options object per `(locale, grouping, date1904)`; do not allocate per cell.
- Use `formatColor` only if the resolved code contains `[`; cache the colour decision per `(code, value section)` if profiling shows it matters.
- For `General`, call `numfmt` with `General`; apply the width fallback in the layout layer.
- For text/boolean/error values, bypass numeric formatting where Excel does; verify error/text-section behaviour against Excel (§9).
- For `TEXT()` in formulas, call the same `format` function with `cellWidthPx: Infinity` (no fill/skip/overflow), matching Excel's `TEXT()` semantics.

## 7. Conformance corpus

The v1 acceptance criterion is a conformance corpus, not a vibe [S47]. Build it from Excel desktop, not from `numfmt`'s tests (those are a starting point, not an independent oracle).

Minimum corpus:

1. **Built-in IDs:** every ID in the ECMA table (0–22, 37–49, and locale-specific 27–36, 50–58, 59–62, 67–81) × representative values (positive, negative, zero, text, date, time, duration, blank) × locales (en-US, en-GB, de-DE, fr-FR, ja-JP, zh-CN, th-TH). Capture the displayed string, font colour, and whether Excel shows `#####`.
2. **Custom grammar:** at least one case per construct: sections (1/2/3/4, skipped, blank), conditions (all six operators, no-match), colours (named, `[Color n]`, custom palette), currency/accounting, fractions (mixed, fixed denominator, padding), scientific (`E+`, `E-`, `e+`, exponent padding), elapsed time, text/`@`, literals/quotes/escapes, `?`/`_`/`*`, scaling commas, percent, `General`.
3. **Locales:** `[$-409]`, `[$-407]`, `[$-411]`, `[$-404]`, `[$-439]`, `[$-F800]`, `[$-x-sysdate]`, unknown LCID; separators, month/day names, currency position, Indian grouping.
4. **Date systems:** 1900 and 1904; serials 0, 1, 59, 60, 61, 1462, 2958465, 2957003; negative serials; `[h]` durations in 1904; leap-day display.
5. **Width sensitivity:** `*`, `_`, `?`, accounting, `General`, `#####` at narrow/medium/wide cell widths, in the actual canvas font and DPR.
6. **Round-trip:** fixtures with built-in IDs, custom IDs, duplicate codes, orphan `numFmt` entries, locale-specific IDs, and untouched `styles.xml`; assert byte-identical untouched parts and semantic equality for edited cells (ticket 12).
7. **Malformed/legacy:** lowercase `e+`, lowercase `a/p`, `[NatNum]`, `[DBNum]`, unterminated quotes/brackets, 5 sections, invalid codes, missing `formatCode`, unknown ID.

Process:

- Generate goldens once in Excel desktop (Windows and Mac if available), commit them as TSV/JSON with a stable schema, and run them in CI without Excel.
- Use `ssf` as a differential oracle for date/time and basic numeric; investigate every mismatch, but treat Excel goldens as truth.
- Use `numfmt`'s `test/tables` as a seed corpus (MIT), and `ssf`'s date/time corpora (Apache-2.0) as a second seed. Attribute both.
- Add a browser benchmark: 10k visible cells, mixed formats, cold and warm cache; target < 16 ms warm for scroll repaints, and record cold-paint time. Do not rely on the Node numbers in §4.5.
- Add a fuzz/property test: random valid format codes from the grammar, random values, assert no throw and no `undefined`; compare against a small independent reference for simple cases.

Acceptance threshold: zero mismatches on the curated corpus for IN constructs; documented, ticketed exceptions only for out-of-scope localisation (era names, `[NatNum]`/`[DBNum]` digit shaping, unknown locale digit systems). If a construct cannot be made to match, remove it from the IN claim rather than silently shipping a different rendering.

## 8. Evidence gaps and open questions

**Evidence gaps in this research**

- No Excel desktop oracle. The comparison table's "Excel expected" column is from Microsoft docs, ECMA examples, and common Excel behaviour, not a live Excel run. A few edge cases (error values with a text section, `General` at narrow widths, `##0.0E+0` at zero, exact Thai/Japanese built-in rendering) need Excel goldens.
- The browser benchmark is Node, not the target browser/canvas. Font metrics, DPR, and JIT behaviour will change the numbers.
- `ssf` was evaluated at 0.11.2, the published package. The current SheetJS source on `git.sheetjs.com` was not accessible without authentication; it may contain fixes not in 0.11.2.
- `numfmt`'s 53,082-case suite is the maintainer's report [S11], not an independently reproduced count in this environment.
- `numfmt` 4.1.3 is recent and changing quickly; the lowercase `e+`/`a/p` behaviour could change. Pin and re-run the corpus on upgrade.

**Open questions for the spec**

1. **Locale policy.** Does v1 render all built-in IDs in en-US, or does it use the browser locale? Recommendation: en-US for v1, honour explicit `[$-…]` codes, preserve IDs. This must be written down because it determines what a Thai/Japanese workbook looks like.
2. **`[NatNum]`/`[DBNum]` and era names.** Are these IN, V2, or never? Recommendation: V2 at the earliest; strip/no-op in v1 so they do not crash.
3. **`TEXT()`.** Does v1 need Excel-compatible `TEXT()`? If yes, the formatter must be exposed to HyperFormula (or `TEXT` overridden); if no, document the divergence.
4. **`General` width fidelity.** Is "show `#####` when it does not fit" acceptable for v1, or must Excel's "show fewer digits" behaviour be reproduced? The latter is a real layout algorithm, not a formatter option.
5. **Built-in ID table ownership.** Who checks in and reviews the ~80-entry table and its locale variants? It should be a reviewed data file with a test that every ID resolves.
6. **Colour palette.** Does v1 honour a workbook's custom `<indexedColors>` for `[Color n]`, or use the default palette? Recommendation: honour it if present; it is a small lookup.
7. **Unsupported format UX.** Should invalid imported formats be silently `######`, or surfaced in a workbook compatibility report? Ticket 12 owns the answer, but the engine needs a hook.

## 9. Sources

Primary sources (library repos, specs, first-party docs):

- [S1] `numfmt` repository — https://github.com/borgar/numfmt (commit `32893df`, 2026-11-10)
- [S2] `numfmt` `package.json` (MIT, version 4.1.3, ESM/CJS/types, `platform: neutral`) — https://github.com/borgar/numfmt/blob/master/package.json
- [S3] `numfmt` README (features, 1904 unsupported, `*`/`_` `TEXT()` semantics) — https://github.com/borgar/numfmt/blob/master/README.md
- [S4] `numfmt` API (`format`, `formatColor`, `getFormatInfo`, `tokenize`, `getFormatFromId`, `grouping`, `dateSpanLarge`, `throws`) — https://github.com/borgar/numfmt/blob/master/API.md
- [S5] `numfmt` `getFormatFromId.ts` (modulo-58 normalisation, 27–36/50–58 generic locale dates) — https://github.com/borgar/numfmt/blob/master/src/getFormatFromId.ts#L124-L186
- [S6] `numfmt` `serialDate.ts` (only `leap1900`; internal 1904 helper in `toYMD.ts`) — https://github.com/borgar/numfmt/blob/master/src/serialDate.ts#L90-L110
- [S7] `numfmt` `options.ts` (defaults: `dateSpanLarge:true`, `dateErrorNumber:true`, `leap1900:true`, `throws:true`) — https://github.com/borgar/numfmt/blob/master/src/options.ts#L110-L130
- [S8] `numfmt` `prepareFormatterData.ts` (1,000-entry LRU pattern cache) — https://github.com/borgar/numfmt/blob/master/src/prepareFormatterData.ts#L1-L35
- [S9] `numfmt` conditional tests (561 `expect()` calls; `[DBNum]`/`[NatNum]` behaviour at lines ~293–302) — https://github.com/borgar/numfmt/blob/master/test/conditionals.spec.ts and https://github.com/borgar/numfmt/blob/master/test/numfmt.spec.ts#L293-L302
- [S10] `numfmt` `getFormatFromId` tests (4,055 `expect()` calls) — https://github.com/borgar/numfmt/blob/master/test/getFormatFromId.spec.ts
- [S11] `numfmt` PR #110 ("vitest run with the Excel tables passes (53082 tests)"; fraction carry) — https://github.com/borgar/numfmt/pull/110
- [S12] `numfmt` issue #105 (General scientific-notation differences vs Excel) — https://github.com/borgar/numfmt/issues/105
- [S13] `numfmt` PR #112 (open; negative-serial weekday bug) — https://github.com/borgar/numfmt/pull/112
- [S14] `numfmt` npm package (4.1.3, published 2026-11-10, MIT, ~1.17 MB unpacked) — https://www.npmjs.com/package/numfmt
- [S15] SheetJS `ssf` repository (last push 2023-02-23) — https://github.com/SheetJS/ssf
- [S16] `ssf` `package.json` (0.11.2, Apache-2.0, `frac` dependency) — https://github.com/SheetJS/ssf/blob/master/package.json
- [S17] `ssf` README (project moved into SheetJS main repo) — https://github.com/SheetJS/ssf/blob/master/README.md
- [S18] `ssf` source (`init_table` 40–75, `default_map` 77–108, `default_str` 110–118, `parse_date_code` 149–160, `write_num` unsupported 498/613, `choose_fmt` 898–918, `eval_fmt` 720–731, `format` 919) — https://github.com/SheetJS/ssf/blob/master/ssf.js
- [S19] `ssf` `test/valid.js` (only asserts no throw; 442 formats) — https://github.com/SheetJS/ssf/blob/master/test/valid.js
- [S20] `ssf` `test/date.js` (exact date/time corpora: 295,849 dates; 1,048,576 times) — https://github.com/SheetJS/ssf/blob/master/test/date.js
- [S21] `ssf` npm package (0.11.2, published 2020-06-27, Apache-2.0) — https://www.npmjs.com/package/ssf
- [S22] SheetJS main repository (Apache-2.0) — https://github.com/SheetJS/sheetjs
- [S23] `@univerjs/sheets-numfmt` npm package (1.4.0, Apache-2.0, depends on `@univerjs/core`, `@univerjs/engine-formula`, `@univerjs/sheets`) — https://www.npmjs.com/package/@univerjs/sheets-numfmt
- [S24] Univer `sheets-numfmt` published code (calls `numfmt.format`, `numfmt.formatColor`, `numfmt.getFormatInfo`) — https://unpkg.com/@univerjs/sheets-numfmt@1.4.0/lib/es/index.js#L265-L288
- [S25] Univer repository — https://github.com/dream-num/univer
- [S26] `numbro` repository (MIT, 1,141 stars, 270 open issues, last push 2025-10-17) — https://github.com/BenjaminVanRyseghem/numbro
- [S27] `numbro` format documentation (option-object API, not Excel format codes) — http://numbrojs.com/format.html
- [S28] `excel-number-format` npm README ("classifier, not a formatter") — https://www.npmjs.com/package/excel-number-format
- [S29] HyperFormula "List of differences with other spreadsheets" (TEXT supports only a subset; no currency formatting inside TEXT) — https://hyperformula.handsontable.com/guide/list-of-differences.html
- [S30] HyperFormula built-in functions (`TEXT`, `stringifyDateTime`/`stringifyCurrency`) — https://hyperformula.handsontable.com/docs/guide/built-in-functions
- [S31] HyperFormula currency handling (TEXT + `stringifyCurrency`; limited built-in patterns) — https://handsontable.com/docs/hyperformula/guide/currency-handling.html
- [S32] Microsoft, "Review guidelines for customizing a number format" (sections, colours, conditions, `*`, `_`, `\`, `"`, `?`, scaling commas, percent, scientific, date/time, elapsed time) — https://support.microsoft.com/en-us/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5
- [S33] Microsoft, "Number format codes" (same grammar, Mac variant; `E-`/`E+`/`e-`/`e+`; `AM/PM`, `am/pm`, `A/P`, `a/p`) — https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
- [S34] Microsoft, "Date systems in Excel" (1900 vs 1904, 1462-day difference) — https://support.microsoft.com/en-US/Excel/date-systems-in-excel
- [S35] Microsoft, "Excel incorrectly assumes that the year 1900 is a leap year" — https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year
- [S36] Microsoft, "Change the date system, format, or two-digit year interpretation" — https://support.microsoft.com/en-US/Excel/change-the-date-system-format-or-two-digit-year-interpretation
- [S37] Microsoft Learn, `NumberingFormat` class / ISO/IEC 29500-1 §18.8.30 table (all-language IDs 0–22, 37–49; zh/ja/ko IDs 27–36, 50–58; th-th IDs 59–62, 67–81) — https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
- [S38] ECMA-376, Office Open XML file formats, 5th edition (Part 1: §18.8.30 `numFmt`, §18.8.31 `numFmts`, §18.17.4.1 date conversion, §18.18.47 `ST_NumFmtId`) — https://ecma-international.org/publications-and-standards/standards/ecma-376/ ; Part 1 PDF: https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
- [S39] ECMA-376 Part 1, §18.8.31 grammar table and §18.17.4.1 date-system text (used for the grammar/1900 base-date claims) — same as [S38]
- [S40] `[MS-XLS]` IFmt (custom format IDs `0x00A4`–`0x0188`; built-ins per ECMA-376 §18.8.30) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/9017e247-7995-4a9c-96e8-950df24735a2
- [S41] `[MS-XLS]` Date1904 (`f1904DateSystem`; 1900 serial 1 vs 1904 serial 0) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/4a5e900a-0eb0-4355-8fc1-81aab8f46e8b
- [S42] MDN, `Intl.NumberFormat` (locale-aware number formatting; options; browser/ICU caveat) — https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl/NumberFormat
- [S43] MDN, `Intl.NumberFormat()` constructor (style/currency/accounting/scientific/rounding options; no fraction style) — https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl/NumberFormat/NumberFormat
- [S44] ECMA-402, `NumberFormat` objects — https://tc39.es/ecma402/#numberformat-objects

Repo context:

- [S45] Ticket 05 research, grid rendering (`SheetView` formats; renderer never runs a number-format engine) — `../research/05-grid-rendering.md`
- [S46] Ticket 04 research, `.xlsx` I/O (append-only styles; observed `\$#,##0.00` → `$#,##0.00` normalisation; locale-specific formats preserved) — `../research/04-xlsx-io.md`
- [S47] Feature matrix (number formats IN; "rendering matches Excel over a conformance corpus"; localisation out of scope) — `../spec/feature-matrix.md`
- [S48] Ticket 16 prior art, GenOffice (`numfmt` gateway; `numfmt-fix.ts` patch layer) — `../research/16-prior-art-genoffice.md`
- [S49] `numfmt` PR #107 (midnight carry) — https://github.com/borgar/numfmt/pull/107

Probe method: `npm install numfmt@4.1.3 ssf@0.11.2`, Node 24.15.0, direct `format()`/`SSF.format()` calls; bundle sizes via `gzip -c`; performance via `process.hrtime.bigint()` over 300k iterations, best of 5 after 20k warm-up. These are indicative, not a browser benchmark.
