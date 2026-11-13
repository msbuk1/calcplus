# 24 — IronCalc 0.9.1 vs HyperFormula 3.5 for the 52-function contract

Researched 2026-11-13. Engine versions tested: **IronCalc 0.9.1** (`@ironcalc/wasm` and `@ironcalc/wasm-xlsx` 0.9.1, Rust crate `ironcalc_base` 0.9.1) and **HyperFormula 3.5.0** (GPLv3 key). Source read at `ironcalc/IronCalc` tag `v0.9.1` and `handsontable/hyperformula` tag `3.5.0`.

**Recommendation: keep HyperFormula.** Do not switch to IronCalc, and do not run a two-engine hybrid. Amend ticket 03 to record the re-evaluation and the Excel-compatibility profile HyperFormula needs. The decisive reason is dates: the v1 contract requires the 1900 system, the fake 29 February 1900 and the 1904 system, and IronCalc 0.9.1 gets the first wrong and has no support for the second. HyperFormula can meet all three through documented configuration. IronCalc's `.xlsx` layer is a from-scratch rewrite, not a preservation layer, so it cannot replace any part of ticket 04 either.

## 1. Method and evidence standard

Claims below are labelled:

- **Tested** — executed in Node 24 against the published packages on 2026-11-13. The harness initialises `@ironcalc/wasm` 0.9.1 with its WASM initialiser (`init({ module_or_path })`), calls `Model.setUserInput` / `evaluate` / `getFormattedCellValue` / `getCellContent`, and for `.xlsx` uses `Model.fromXlsx` / `toXlsx` from `@ironcalc/wasm-xlsx`. HyperFormula 3.5.0 is loaded from `dist/hyperformula.full.js` with `licenseKey: 'gpl-v3'`.
- **Source** — read in the tagged repository or the published package.
- **Vendor claim** — documentation or marketing, not independently reproduced.

The Node measurements are indicative, not a browser benchmark. Browser startup, peak WASM memory and worker behaviour are called out as thin evidence where they are not directly measured.

## 2. Comparison at a glance

| Dimension | IronCalc 0.9.1 | HyperFormula 3.5.0 | Edge |
|---|---|---|---|
| Function names in the 52-function contract | 52/52 present (497 registry variants in total) | 52/52 present (423 functions in total) | tie |
| Tested semantic deviations from Excel | 4 (`POWER`, `AND`, `DATE`/`DAY` around 1900-02-29, no 1904) | 3 (`INT`, wildcard exact lookups, `CHOOSE` out-of-range), plus config-dependent `TRUE`/`FALSE` and dates | HyperFormula |
| Error values | Excel's 7 plus non-standard `#ERROR!`, `#N/IMPL`, `#SPILL!`, `#CALC!`, `#CIRC!` | Excel's 7 plus `#CYCLE!`, `#ERROR!`, `#SPILL!`, `#CALC!`, `#LIC!` | both need mapping; IronCalc's `#CIRC!` already matches the contract |
| Type coercion | Mostly Excel-correct in the tested set | Mostly Excel-correct in the tested set | tie |
| Operator precedence | Matches Excel (`2^3^2 = 64`, `-2^2 = 4`) | Matches Excel (`2^3^2 = 64`, `-2^2 = 4`) | tie |
| Case-insensitive text | Yes | Yes (`caseSensitive: false`) | tie |
| Wildcards in `COUNTIF`/`SUMIF` | Yes | Yes (`useWildcards: true`) | tie |
| Wildcards in exact `VLOOKUP`/`MATCH` | Yes, first match | `VLOOKUP` yes but returns the **last** match; `MATCH` ignores wildcards | IronCalc |
| Exact and approximate lookups | Correct in the tested set | Correct in the tested set | tie |
| 1900 date system | `DATE(1900,2,29)` → 61, `DAY(59)` → 27, `DAY(60)` → 28 | Correct with `leapYear1900: true` + `nullDate: 1899-12-31` | **HyperFormula** |
| 1904 date system | **None** (no flag, no source reference) | Correct with `nullDate: 1904-01-01` (tested) | **HyperFormula** |
| Authored formula text | Stored as a parsed AST; regenerated on read/export | Stored as text but normalised; `OFFSET` rewritten to its target | neither is safe; model must own text |
| `.xlsx` import | Strict, partial, can panic | No `.xlsx` support at all | HyperFormula (no false promise) |
| `.xlsx` export | Rebuilds the package from a fixed part list | None | HyperFormula |
| Bundle (gzip) | 720 KB (calc WASM) / 915 KB (calc+xlsx WASM) | 171 KB min / 210 KB full | HyperFormula |
| Worker fit | WASM; no worker in IronCalc's own web app | Pure JS, no DOM in the engine path | HyperFormula |
| Licence | MIT OR Apache-2.0 | GPLv3 or proprietary | both GPLv3-compatible |
| Maturity | 0.9.1, pre-1.0 | 3.5.0, stable 3.x API | HyperFormula |
| Release cadence | 7 releases in 2026; last 2026-10-21 | 4 releases in 2026; last 2026-10-09 | tie |
| Contributors / open issues | 55 / 338 | 31 / 237 | tie (both single-vendor) |

## 3. Function-by-function against the 52

All 52 names are present in both engines: IronCalc's `Function` enum and lookup table in [`base/src/functions/mod.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/functions/mod.rs); HyperFormula's registry returns 423 names via `HyperFormula.getRegisteredFunctionNames('enGB')`, and the docs state 423 functions ([built-in functions](https://hyperformula.handsontable.com/guide/built-in-functions.html)). Presence is not the question; semantics are.

Tested with two seeds. Arithmetic/statistics/text: `A1:A3 = 10,20,30`, `B1:B3 = a,Ab,*b`, `C1:D3 = 1,a / 2,b / 3,c`, `E1:G2 = 1,2,3 / x,y,z`. Lookup/wildcard: `A1:B3 = alpha,10 / beta,20 / Alpine,30`, plus `alpha,alps,Alpine,alto` for the first-vs-last match check. HyperFormula was run with `leapYear1900: true`, `nullDate: 1899-12-31` and `TRUE`/`FALSE` named expressions (see §5.2 and §11.1). "DEV" = differs from Excel; "GAP" = differs and matters to the contract.

| # | Function | IronCalc 0.9.1 (tested) | HyperFormula 3.5 (tested) | Verdict |
|---|---|---|---|---|
| 1 | SUM | `SUM(A1:A2)=11`; `SUM("1",TRUE)=2` | same | OK both |
| 2 | PRODUCT | `PRODUCT(2,3)=6` | 6 | OK both |
| 3 | ABS | `ABS(-2)=2` | 2 | OK both |
| 4 | ROUND | `ROUND(2.5,0)=3`; `ROUND(-2.5,0)=-3` | same | OK both |
| 5 | ROUNDUP | `ROUNDUP(2.1,0)=3`; `ROUNDUP(-2.1,0)=-3` | same | OK both |
| 6 | ROUNDDOWN | `ROUNDDOWN(-2.9,0)=-2` | -2 | OK both |
| 7 | INT | `INT(-8.9)=-9` | **-8** | **HF DEV** (documented in ticket 03) |
| 8 | MOD | `MOD(-3,2)=1`; `MOD(3,-2)=-1` | same | OK both |
| 9 | POWER | `POWER(-8,1/3)=-2` | `#NUM!` | **IC DEV** (Excel returns `#NUM!`) |
| 10 | SQRT | `SQRT(-1)=#NUM!` | `#NUM!` | OK both |
| 11 | AVERAGE | `AVERAGE(A1:A3)=20`; `AVERAGE("1",2)=1.5` | same | OK both |
| 12 | MIN | `MIN(A1:A3)=10`; `MIN("1",0)=0` | same | OK both |
| 13 | MAX | `MAX(A1:A3)=30`; `MAX(TRUE,0)=1` | same (with `TRUE()`) | OK both |
| 14 | COUNT | `COUNT(A1:A3)=3`; `COUNT("1",TRUE)=2` | same | OK both |
| 15 | COUNTA | `COUNTA(B1:B3)=3` | 3 | OK both |
| 16 | COUNTBLANK | `COUNTBLANK(G1:G3)=3` | 3 | OK both |
| 17 | COUNTIF | case-insensitive; `a*`=2, `?b`=2, `~*b`=1 | same | OK both |
| 18 | SUMIF | `>10`=50, `<>10`=50, `a*`=30 | same | OK both |
| 19 | MEDIAN | `MEDIAN(A1:A3)=20` | 20 | OK both |
| 20 | IF | `IF(1,"a","b")=a` | a | OK both |
| 21 | AND | `AND(TRUE,"a")=TRUE` | `#VALUE!` | **IC DEV** (Excel `#VALUE!`) |
| 22 | OR | `OR(FALSE,0)=FALSE` | FALSE | OK both |
| 23 | NOT | `NOT(0)=TRUE` | TRUE | OK both |
| 24 | IFERROR | `IFERROR(1/0,"x")=x` | x | OK both |
| 25 | IFNA | `IFNA(NA(),"x")=x`; non-NA propagates | same | OK both |
| 26 | VLOOKUP | exact `b`; approx `b`; past-end `c`; below-first `#N/A`; wildcard `"al*"`→10 (first match) | same except wildcard `"al*"`→**30** (last match, not first) | **HF GAP** (wildcard occurrence) |
| 27 | HLOOKUP | `1`→x, `2`→y; exact/approx correct | same | OK both |
| 28 | INDEX | `INDEX(A1:A3,2)=20`; `INDEX(A1:A3,0)=10` | 20; `0`→`#VALUE!` | edge case; contract silent on `row=0` |
| 29 | MATCH | `25,type 1`→2; `type -1`→`#N/A`; wildcard `"be*"`→2 | same except wildcard→`#N/A` | **HF GAP** (wildcard) |
| 30 | CHOOSE | `CHOOSE(2,"a","b")=b`; out-of-range `#VALUE!` | b; out-of-range **`#NUM!`** | **HF DEV** (minor) |
| 31 | ROW | `ROW()` and `ROW(C5)` correct | correct | OK both |
| 32 | COLUMN | `COLUMN()` and `COLUMN(C5)` correct | correct | OK both |
| 33 | ROWS | `ROWS(A1:A3)=3` | 3 | OK both |
| 34 | CONCATENATE | `a1TRUE` | `a1TRUE` (with `TRUE()`) | OK both |
| 35 | LEFT | `LEFT("abc",2)=ab`; `0`→`""` | same | OK both |
| 36 | RIGHT | `RIGHT("abc",2)=bc` | bc | OK both |
| 37 | MID | `MID("abc",2,1)=b`; past end `""` | same | OK both |
| 38 | LEN | `LEN("")=0` | 0 | OK both |
| 39 | TRIM | `"  a  b  "`→`"a b"` | same | OK both |
| 40 | UPPER | `AB` | AB | OK both |
| 41 | LOWER | `ab` | ab | OK both |
| 42 | TODAY | correct local date | correct serial (1900 or 1904 per `nullDate`) | OK both |
| 43 | NOW | correct local date-time | correct serial | OK both |
| 44 | DATE | `DATE(1900,2,29)`→**61**; `DATE(2020,13,1)` rolls over; `DATE(99,1,1)`→1999 | `DATE(1900,2,29)`→**60**; same rollovers | **IC DEV** (fake leap day) |
| 45 | DAY | `DAY(60)`→**28**, `DAY(59)`→**27** | 29, 28 | **IC DEV** (off by one before serial 61) |
| 46 | MONTH | `MONTH(60)=2` | 2 | OK both |
| 47 | YEAR | `YEAR(60)=1900` | 1900 | OK both |
| 48 | ISBLANK | empty→TRUE, `""`→FALSE | same | OK both |
| 49 | ISNUMBER | `ISNUMBER("1")=FALSE` | FALSE | OK both |
| 50 | ISTEXT | `ISTEXT(1)=FALSE` | FALSE | OK both |
| 51 | ISERROR | `ISERROR(1/0)=TRUE` | TRUE | OK both |
| 52 | ISNA | `ISNA(NA())=TRUE` | TRUE | OK both |

Result: **IronCalc is missing no function but fails the date contract; HyperFormula is missing no function but needs shims for exact-lookup wildcards, `INT`, `CHOOSE` and bare `TRUE`/`FALSE`.** Both failures are in scope for v1's acceptance tests ([formula-semantics.md §10](../spec/formula-semantics.md)).

## 4. Excel semantics

### 4.1 Error values

Both engines add error codes Excel does not have, so the `FormulaEngine` port must map or suppress them.

- IronCalc's error enum is `REF, NAME, VALUE, DIV, NA, NUM, ERROR, NIMPL, SPILL, CALC, CIRC, NULL`, rendered as `#ERROR!`, `#N/IMPL`, `#SPILL!`, `#CALC!` and `#CIRC!` ([`base/src/expressions/token.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/expressions/token.rs)). `#ERROR!` is also what an unevaluated formula cell returns ([`base/src/cell.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/cell.rs)). Tested: `=SUM(A1:A3)+` → `#ERROR!`; `=1/0` → `#DIV/0!`; `=NA()` → `#N/A`.
- IronCalc already returns `#CIRC!` for cycles (tested: `=A1` in A1, and a two-cell cycle, both `#CIRC!`), which matches calcplus's deliberate deviation in §6 of the contract.
- HyperFormula returns `#CYCLE!` for cycles (tested), plus `#ERROR!` for parse failures and `#LIC!` for a licence-gated function ([3.5.0 release notes](https://github.com/handsontable/hyperformula/releases/tag/3.5.0)). It must be mapped to `#CIRC!` and never written to `.xlsx`.
- Both propagate errors correctly in the tested set; `IFERROR`/`IFNA`/`ISERROR`/`ISNA` handle them as Excel does.

### 4.2 Coercion and precedence

Tested, both engines: numeric-looking text converts in arithmetic (`"1"+1=2`), `TRUE` counts as 1, `1="1"` is FALSE, `TRUE=1` is FALSE, `"a"="A"` is TRUE, `"a"<"B"` is TRUE, `1+2&3` is `"33"`, `-2^2=4`, `2^3^2=64`, `50%*2=1`. No divergence found.

Two edge cases worth pinning in the conformance corpus rather than assuming:

- `=""+1` returned `#VALUE!` in IronCalc and `1` in HyperFormula. I did not verify Excel's answer with a real Excel fixture; treat this as **unverified** and add it to the error/coercion table in ticket 13.
- `AND(TRUE,"a")` is the one clear coercion divergence (§3, #21).

### 4.3 Text comparison and criteria

Both are case-insensitive by default and both support `*`, `?` and `~` escapes in `COUNTIF`/`SUMIF` (tested). HyperFormula's docs confirm `caseSensitive: false`, `useWildcards: true`, `matchWholeCell: true` are the defaults ([Excel compatibility](https://hyperformula.handsontable.com/guide/compatibility-with-microsoft-excel.html)).

### 4.4 Exact and approximate lookups

Approximate matching is correct in both (`VLOOKUP(2.5, …, TRUE)`→`b`, `VLOOKUP(3.5, …)`→`c`, `VLOOKUP(0.5, …)`→`#N/A`, `MATCH(25, …, 1)`→2, `MATCH(25, …, -1)`→`#N/A`). The divergence is exact-match wildcards:

- IronCalc: `VLOOKUP("al*", A1:B3, 2, FALSE)`→10 and `MATCH("be*", A1:A3, 0)`→2, both case-insensitive. Correct.
- HyperFormula: wildcard syntax works in `VLOOKUP`, but it returns the **last** match, not the first. With `alpha=10`, `alps=2`, `Alpine=3`, `alto=4`, `VLOOKUP("al*", …, FALSE)` returned 4 (`alto`); Excel returns 1 (`alpha`). `MATCH("be*", …, 0)` returned `#N/A` — `MATCH` has no wildcard path at all. The source explains both: `doVlookup` passes `ordering === 'none'` into `searchInRange`, which uses the wildcard matcher, while `doMatch` calls `searchStrategy.find` directly ([`LookupPlugin`](https://github.com/handsontable/hyperformula/blob/3.5.0/src/interpreter/plugin/LookupPlugin.ts)).

This is a contract gap (`formula-semantics.md §7`) and must be closed before v1: `VLOOKUP`/`HLOOKUP` must return the first wildcard match, and `MATCH` must support wildcards in type 0. It is a small, well-bounded TypeScript change (a custom function plugin, or an upstream patch), unlike patching a Rust engine.

## 5. Dates — the decisive section

The contract requires both date systems, the fake 29 February 1900 (serial 60), serial dates/times and Excel's pre-1900-03-01 leap bug ([formula-semantics.md §5, §10](../spec/formula-semantics.md)).

### 5.1 IronCalc 0.9.1

Tested:

| Formula | IronCalc | Excel |
|---|---|---|
| `DATE(1900,2,29)` | 61 (displays 3/1/1900) | 60 (2/29/1900) |
| `DAY(59)` | 27 | 28 |
| `DAY(60)` | 28 | 29 |
| `DATE(1900,3,1)` | 61 | 61 |
| `DAY(61)` | 1 | 1 |
| `YEAR(60)` | 1900 | 1900 |

So IronCalc's serial mapping is off by one for serials below 61 and it rolls the fake leap day forward to 1 March. The source has two conversions that disagree: `DAY`/`MONTH`/`YEAR` use the generic `from_excel_date`, while a separate helper `excel_serial_to_ymd` special-cases serial 60 only for `DAYS360` ([`base/src/constants.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/constants.rs), [`base/src/functions/date_and_time.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/functions/date_and_time.rs)).

**1904: absent.** There is no `1904`, `date1904` or date-system flag anywhere in `base/` or `xlsx/` at v0.9.1 (grep of the tagged source; the only matches were numeric literals in unrelated test values). `WorkbookSettings` contains only `tz` and `locale` ([`base/src/types.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/types.rs)). GenOffice's comment is accurate: "IronCalc 0.9.1 has no date1904 handling — nothing in the engine reads the flag — so every date result it computes is 1462 days away from the value Excel cached" ([`recalc.rs`](https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/src/recalc.rs)). Their audit explicitly skips 1904 workbooks.

This is a hard blocker: fixing it means forking the Rust engine and rebuilding the WASM, or pre-shifting every date serial in and out of the engine (including `TODAY`, `NOW`, `DATE`, `DAY`, `MONTH`, `YEAR` and any date arithmetic) — fragile and outside the "narrow port" design.

### 5.2 HyperFormula 3.5

Tested with the documented Excel profile (`leapYear1900: true`, `nullDate: {1899,12,31}`, `TRUE`/`FALSE` named expressions):

| Formula | HyperFormula | Excel |
|---|---|---|
| `DATE(1900,2,29)` | 60 | 60 |
| `DAY(59)` | 28 | 28 |
| `DAY(60)` | 29 | 29 |
| `DATE(1900,1,1)` | 1 | 1 |
| `DATE(1900,3,1)` | 61 | 61 |
| `DATE(2020,13,1)` | 44197 (rolls to 2021-01-01) | same |
| `DATE(99,1,1)` | 1999-01-01 | same |

With `nullDate: {1904,1,1}` (and no leap-year flag) it reproduces the 1904 system:

| Formula | HyperFormula | Expected 1904 |
|---|---|---|
| `DATE(1904,1,1)` | 0 | 0 |
| `DATE(2026,11,13)` | 44877 | 44877 |
| `TODAY()` | 44877 | 44877 |
| `DAY(0)` / `MONTH(0)` / `YEAR(0)` | 1 / 1 / 1904 | 1 / 1 / 1904 |
| `DATE(2020,1,1)` | 42369 | 42369 (43831 − 1462) |

`nullDate` is the documented epoch ("each date is represented as a number of days that passed since nullDate") and `leapYear1900` is explicitly "for compatibility with … Microsoft Excel" ([configuration reference](https://hyperformula.handsontable.com/api/interfaces/configparams.html)). The 1904 result is an emulation via a documented option, not a vendor-labelled "1904 mode"; it works because calcplus opens one workbook at a time and can build the engine with the workbook's epoch. It should be covered by a conformance test, not assumed.

**Verdict:** HyperFormula can satisfy §5 with configuration; IronCalc 0.9.1 cannot without a fork. This alone decides the ticket.

## 6. Formula storage — neither engine may own the text

The contract says: "Keep the exact formula text as authored. Never normalise, rewrite or re-serialise it" ([§9](../spec/formula-semantics.md)). Both engines rewrite.

### 6.1 IronCalc stores an AST and regenerates text

The source is explicit: "Formulas internally are stored in R1C1 format" ([`base/src/expressions/parser/stringify.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/expressions/parser/stringify.rs)); `get_cell_formula` looks up the parsed node and calls `to_localized_string` ([`base/src/model.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/model.rs)). The WASM binding does not even expose a formula getter — only `getCellContent`, which returns the regenerated string.

Tested round-trips of `setUserInput` → `getCellContent`:

| Authored | Read back |
|---|---|
| `=sum(A1:A2)` | `=SUM(A1:A2)` |
| `=SUM( A1 , A2 )` | `=SUM(A1,A2)` |
| `=SUM(A1:A2)+0.10` | `=SUM(A1:A2)+0.1` |
| `=2^3^2` | `=(2^3)^2` |
| `=_xlfn.CONCAT(A1,A2)` | `=CONCAT(A1,A2)` |
| `=FOOBAR(1)` | `=foobar(1)` |

The last two matter: the `_xlfn.` prefix is stripped from the model's view (it is re-added on export only for functions in IronCalc's own future-function list), and an unknown function's name is lower-cased. Whitespace, numeric literal spelling and parentheses are all normalised.

On `.xlsx` export IronCalc re-serialises every formula from the AST. In a fixture with a shared formula `A42:A43` (`<f t="shared" ref="A42:A43" si="99">1+1</f>` + follower `<f t="shared" si="99"/>`), the exported file contained two explicit `<f>1+1</f>` cells — the shared structure is flattened ([`xlsx/src/export/worksheets.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/xlsx/src/export/worksheets.rs)).

### 6.2 HyperFormula normalises text and rewrites `OFFSET`

Tested:

| Authored | Read back |
|---|---|
| `=sum(A1:A3)` | `=SUM(A1:A3)` |
| `=SUM( A1 , A2 )` | `=SUM( A1, A2 )` |
| `=SUM(A1:A3)+0.10` | `=SUM(A1:A3)+0.1` |
| `=OFFSET(A1,1,0)` | `=A2` |
| `=_xlfn.CONCAT(A1,B1)` | `#ERROR!` (parse) |
| `=@A1:A3` | `#ERROR!` (parse) |

The `OFFSET` rewrite is the one ticket 03 already called out; it resolves the reference instead of preserving the call. HyperFormula also cannot parse `_xlfn.` or the implicit-intersection `@`, so those must be stripped/normalised before the engine sees them.

**Conclusion:** the architectural rule in ticket 03 is correct and applies to both engines. The workbook model owns authored text and cached values; the engine is fed an internal form and its output is used only for values. IronCalc's rewriting is broader, which makes it *more* dangerous if anyone is tempted to use its `.xlsx` import/export as the storage path.

## 7. `.xlsx` import/export — IronCalc rebuilds, it does not preserve

This is the second decisive finding. IronCalc's export is not a patch of the input package; it writes a new zip from a fixed list of parts: `[Content_Types].xml`, `docProps/{app,core}.xml`, `_rels/.rels`, `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/workbook.xml`, `xl/metadata.xml`, `xl/theme/theme1.xml`, `xl/_rels/workbook.xml.rels`, worksheet parts and worksheet rels ([`xlsx/src/export/mod.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/xlsx/src/export/mod.rs)). Anything not in that list is gone.

Tested round-trips through `@ironcalc/wasm-xlsx` 0.9.1:

| Input | In → out bytes | Parts lost |
|---|---|---|
| `example.xlsx` (charts, pivot table, drawings, VML) | 41,197 → 14,118 | pivotTables, charts, drawings, VML, chart rels, pivot rels |
| `link_test.xlsx` (external links, webextensions) | 15,354 → 9,425 | externalLinks, webextensions |
| `invoice.xlsx` (calcChain) | 16,811 → 12,417 | calcChain |
| injected `customXml/item1.xml` + `xl/media/image1.png` | — | both dropped |

Cached values and unsupported formulas:

- On import, IronCalc keeps the `<v>` cache. A cell with `SUM(B1:B2)` and cached `999` read back as `999` without recalculation.
- After an edit elsewhere, the same cell recalculated to `12` (correct), and export wrote `<v>12</v>`.
- A cell with `_xlfn.FOO(1)` and cached `42` read back as `42` on import, but after any edit it became `#NAME?`, and export wrote `<c r="A41" t="e"><f>_xlfn.foo(1)</f><v>#NAME?</v></c>` — the cached value is destroyed. That violates the contract's "unsupported function with a cached result: show the cached result marked stale" rule.
- The importer is strict and can panic. GenOffice catches panics and says it "can panic on non-Excel producers (missing cellStyles, whitespace nodes)"; it also normalises OPC relationship targets before calling IronCalc ([`rels_normalize.rs`](https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/src/rels_normalize.rs)). A hand-built minimal package in this study failed with `RuntimeError: unreachable`.

### 7.1 Why GenOffice splits the two layers

GenOffice uses IronCalc **only for recalculation**, never for storage:

- `recalc.rs` imports the workbook into a session-resident IronCalc model (`load_model` → `ironcalc::import::load_from_xlsx`) and applies only the changed edits; it never calls `save_to_xlsx`.
- The `.xlsx` gateway (`@genoffice/xlsx-gateway`) plans a small set of patched entries and copies everything else verbatim; its compatibility script asserts "One surgical cell edit per fixture; every other package entry must survive byte-identical, including binary and exotic parts" ([`verify-compatibility.ts`](https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/scripts/verify-compatibility.ts), [`xlsx-package-io.ts`](https://github.com/genspark-ai/genoffice/blob/main/packages/xlsx-gateway/src/gateway/xlsx-package-io.ts)).
- The renderer keeps the file's cached value whenever the engine errors, so an engine gap never blanks a cell ([`formula-cached-fallback.ts`](https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/src/renderer/formula-cached-fallback.ts)).

That split exists **because** IronCalc's import/export cannot preserve a workbook. It is not evidence that a two-engine architecture is desirable for calcplus; it is evidence that IronCalc's `.xlsx` layer is unusable for a preservation-first product.

**Conclusion:** IronCalc cannot replace any part of ticket 04's OOXML layer. At most it can be a calculation engine behind the same kind of port HyperFormula would use — and it is a worse calculation engine for this contract because of dates.

## 8. Browser fit

### 8.1 Bundle size (measured from the published tarballs)

| Package | Raw | Gzip |
|---|---|---|
| `@ironcalc/wasm` 0.9.1 `wasm_bg.wasm` | 2,090,515 B | 719,964 B |
| `@ironcalc/wasm-xlsx` 0.9.1 `wasm_bg.wasm` | 2,546,925 B | 915,272 B |
| `hyperformula` 3.5.0 `dist/hyperformula.min.js` | 821,905 B | 170,719 B |
| `hyperformula` 3.5.0 `dist/hyperformula.full.min.js` | 988,809 B | 210,301 B |

Ticket 03's "~181 KB gzip" is close to the measured 171 KB minified build (the full build is 210 KB). IronCalc is **4.2×** the minified HyperFormula bundle for calculation only, and **5.4×** with `.xlsx` support. The WASM also needs instantiation and a separate memory space.

### 8.2 Startup, memory and worker

- **Startup:** WASM instantiation measured 28 ms in Node 24 for `@ironcalc/wasm` 0.9.1. HyperFormula is a normal JS import with no separate initialisation. Browser numbers will differ; treat as indicative.
- **Memory:** at 20,000 formula cells (`=A*2`), RSS grew by ~70 MB for HyperFormula and ~79 MB for IronCalc in Node. At small scale they are comparable. The real-world warning is GenOffice's: "a 31MB workbook's model holds ~1.2GB resident" ([`recalc.rs`](https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/src/recalc.rs)). A browser WASM heap has a 32-bit address space; a 1.2 GB model is possible but leaves little headroom for the grid, undo history and the rest of the app. Browser peak memory was **not** measured here — evidence is thin, but the direction is clear.
- **Worker:** HyperFormula's engine code is pure JS with no `window`/`document` API use in the published `es/` tree (the only occurrences are licence/error-message text). It can be hosted in a worker, though 3.5.0 has no dedicated worker guide — ticket 03's claim is architectural, not vendor-guaranteed. IronCalc's wasm-bindgen package can also run in a worker, but IronCalc's own web app contains no worker file (tree search at `v0.9.1` returns none), so its first-party path is main-thread WASM. If calcplus puts either engine in a worker, that is our integration work, not a vendor feature.
- **Performance:** indicative 10,000-row / 20,000-cell benchmark (A numbers, B `=A*2`, C `=SUM(B1:B10000)`):

| Engine | Build / set + evaluate | Edit + recalc |
|---|---|---|
| HyperFormula | 155 ms build | 3.84 ms |
| IronCalc WASM | 129 ms set + 38 ms evaluate (28 ms init) | 16.1 ms |

Both are adequate at this size. IronCalc's Rust/WASM is not dramatically faster in this test, and its incremental edit path was ~4× slower than HyperFormula's. This is a single Node run, not a benchmark suite; it should not be over-read.

## 9. Licensing and maintenance

- **IronCalc:** MIT OR Apache-2.0 ([crates.io 0.9.1](https://crates.io/api/v1/crates/ironcalc), npm package metadata). Compatible with GPLv3 and with any future relicensing.
- **HyperFormula:** dual-licensed GPLv3 or proprietary; the GPLv3 key is `gpl-v3` ([licensing](https://hyperformula.handsontable.com/guide/licensing.html), package `LICENSE.txt`). Compatible with calcplus's GPLv3.
- Because calcplus is GPLv3, IronCalc's permissive licence is not a decisive advantage; it only preserves optionality if the product licence ever changes.
- **Release cadence:** IronCalc shipped 0.6.0 (2025-10-19), 0.7.0/0.7.1 (2026-01), 0.8.0/0.8.2/0.8.3 (2026-07/08), 0.9.0/0.9.1 (2026-10). HyperFormula shipped 3.1.0 (2025-10), 3.2.0 (2026-02), 3.3.0 (2026-05), 3.4.0 (2026-08), 3.5.0 (2026-10-09). Both are active.
- **Maturity:** IronCalc is pre-1.0; 0.8→0.9 is a breaking series and the API is still moving. HyperFormula has a stable 3.x API and a documented migration path.
- **Bus factor:** both are single-vendor. IronCalc is a company (IronCalc GmbH) with EU/NLnet funding and 55 contributors; HyperFormula is a Handsontable product with 31 contributors. IronCalc has 338 open issues, HyperFormula 237. Neither is a community-owned project with a broad maintainer base. For calcplus, the mitigation is the same in both cases: keep the engine behind the `FormulaEngine` port so it can be replaced.

## 10. Known quirks

- **Strict IronCalc importer:** documented by GenOffice and reproduced here as a panic on a hand-built package. If IronCalc were used, every imported workbook would need to pass its parser before any calculation, and a failure would have to fall back to cached values. This is a real cost, not a theoretical one.
- **`#ERROR!`:** IronCalc's non-standard error for parse/unevaluated formulas; also appears in its formatter output. It must never reach the UI or `.xlsx`.
- **`mmss.0`:** verified in IronCalc's source. Built-in format id 47 is the correct ECMA-376 `mmss.0`, but the test comment says "IronCalc's formatter does not yet parse that pattern — a separate formatter gap" ([`base/src/number_format.rs`](https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/number_format.rs)). calcplus owns number formatting (ticket 18), so this is only a problem if we consume IronCalc's formatted strings — another reason to consume raw values only.
- **No 1904:** verified in §5.
- **HyperFormula quirks:** bare `TRUE`/`FALSE` are not literals (they are `TRUE()`/`FALSE()`; define named expressions to emulate Excel), `OFFSET` is rewritten, `_xlfn.`/`@` fail to parse, cycles are `#CYCLE!`, and `INT(-8.9)` is -8. All are documented or tested above.

## 11. Recommendation

**Keep HyperFormula 3.5.x (GPLv3) behind the existing `FormulaEngine` port.**

Why:

1. **Dates are a hard contract requirement and IronCalc fails them.** HyperFormula reproduces the 1900 fake leap day and the 1904 system with documented config; IronCalc 0.9.1 has no 1904 support at all and is off by one before serial 61. Fixing IronCalc means forking Rust and rebuilding WASM, which is a much larger and riskier commitment than the TypeScript shims HyperFormula needs.
2. **IronCalc's `.xlsx` layer is a rewrite, not a preservation layer.** It drops charts, pivots, drawings, external links, webextensions, custom parts and calcChain, flattens shared formulas, and destroys cached values for unsupported functions after a recalc. It cannot replace any part of ticket 04. HyperFormula has no `.xlsx` layer, which is the honest and safer shape for a preservation-first product.
3. **Browser weight and simplicity.** 171 KB gzip of pure JS versus 720–915 KB gzip of WASM, plus instantiation and a larger heap. The measured performance advantage of IronCalc is not there at 10k rows, and its real-world memory profile is worse.
4. **The remaining HyperFormula gaps are small and local.** Exact-lookup wildcards, `INT`, `CHOOSE`'s out-of-range error, bare `TRUE`/`FALSE` and `#CYCLE!`→`#CIRC!` are all fixable in a parity plugin/config layer. None requires a parser fork.

### 11.1 Required HyperFormula configuration for v1

Build the engine per workbook with:

```ts
{
  licenseKey: 'gpl-v3',
  leapYear1900: true,                         // 1900 workbooks
  nullDate: { year: 1899, month: 12, day: 31 }, // or 1904-01-01 for date1904
  useArrayArithmetic: false,                  // v1 has no dynamic arrays/spill; CSE is read-only
  ignoreWhiteSpace: 'any',
  evaluateNullToZero: true,
  caseSensitive: false,
  useWildcards: true,
  matchWholeCell: true,
  functionArgSeparator: ',',
  decimalSeparator: '.',
  // dateFormats/timeFormats per the v1 locale decision
}
```

plus `addNamedExpression('TRUE', '=TRUE()')` and `addNamedExpression('FALSE', '=FALSE()')` (the documented Excel-compatibility workaround).

### 11.2 Required custom work before v1

- Gate the function registry to the 52 published functions; everything else must be `#NAME?` at the port boundary.
- Fix exact-mode `VLOOKUP`/`HLOOKUP` to return the **first** wildcard match and add wildcard support to `MATCH` (contract §7). A custom function plugin can replace a built-in in 3.5.0 — tested with a `VLOOKUP` override returning a sentinel — but the docs require function IDs to be unique, so this relies on registry replacement behaviour: pin it with a test or send the fix upstream.
- Override `INT` to floor toward negative infinity, and map `CHOOSE` out-of-range to `#VALUE!`.
- Map `#CYCLE!` to `#CIRC!`; suppress `#ERROR!`, `#SPILL!`, `#CALC!`, `#LIC!`; never let them reach the model or the file.
- Feed the engine a normalised internal formula (strip `_xlfn.`, handle `@`/structured refs by deferring to the cached value); never read formula text back from the engine.
- Keep `OFFSET` and all unsupported constructs out of the engine path; the model preserves their authored text and cached value.

### 11.3 Why not a hybrid

A hybrid (IronCalc for authoritative recalc + HyperFormula for interactive) doubles the semantic surface: two sets of date rules, two error vocabularies, two function behaviours, and a reconciliation policy for disagreements. It also adds 720–915 KB of WASM for no capability HyperFormula lacks after the shims above. GenOffice's hybrid is justified by its Electron/Univer/native context and its need for a Rust sidecar; calcplus has neither. Reject the hybrid.

### 11.4 Should ticket 03 be amended?

**Yes.** Replace the closing paragraph "**Engine choice under review (2026-11-13):** IronCalc 0.9.1 … is an unevaluated alternative" with a resolution recording that IronCalc was evaluated in ticket 24 and rejected for v1, for the date and `.xlsx` reasons above. Add the configuration profile (§11.1) and the gap list (§11.2) to the architectural rule, so the "engine choice" is not reopened without new evidence. Do not change the existing rule that the model owns authored formula text and cached values.

### 11.5 Risks of the recommended path

- **Shim drift:** if the wildcard/`INT`/`CHOOSE` overrides are not shipped with v1, the published 52-function contract is false. They belong in the same milestone as the engine port, with conformance tests.
- **Date configuration is per workbook:** the engine must be rebuilt (or reconfigured) when a 1904 workbook is opened. Add a 1900/1904 fixture pair to ticket 13.
- **Array arithmetic and implicit intersection:** the config above leaves `useArrayArithmetic` off because v1 has no dynamic arrays/spill. Formulas that pass a range where Excel implicitly intersects need a decision in ticket 02/13 and a fixture; do not assume HyperFormula's scalar behaviour.
- **`TRUE`/`FALSE` named expressions are global:** they must be installed on every instance, and user-defined names must not shadow them.
- **Browser memory is unproven:** the 10k-row test is small. Ticket 11's large-sheet target needs a browser benchmark before the engine is declared fit; if HyperFormula's memory is the problem, the fallback is a worker, not IronCalc.
- **What would force a switch:** if v2 requires full dynamic arrays/spill, `INDIRECT`, 3D references or iteration at Excel fidelity, HyperFormula's limits become binding. At that point Univer (already evaluated as a platform bet in ticket 03) or a future IronCalc 1.x with date support are the candidates — not IronCalc 0.9.1.

## 12. Sources

**IronCalc (tag `v0.9.1`)**

- Repository: https://github.com/ironcalc/IronCalc/tree/v0.9.1
- Function registry: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/functions/mod.rs
- Date constants and base: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/constants.rs
- Date functions (fake leap helper, DAY/MONTH/YEAR): https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/functions/date_and_time.rs
- Workbook settings (only `tz`, `locale`): https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/types.rs
- Formula storage / R1C1: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/expressions/parser/stringify.rs
- Formula getter: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/model.rs
- Error enum (`#ERROR!`, `#N/IMPL`, `#CIRC!`): https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/expressions/token.rs
- Unevaluated cell → `#ERROR!`: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/cell.rs
- `mmss.0` formatter gap: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/number_format.rs
- XLSX export part list: https://github.com/ironcalc/IronCalc/blob/v0.9.1/xlsx/src/export/mod.rs
- XLSX formula re-serialisation: https://github.com/ironcalc/IronCalc/blob/v0.9.1/xlsx/src/export/worksheets.rs
- XLSX import: https://github.com/ironcalc/IronCalc/blob/v0.9.1/xlsx/src/import/mod.rs
- crates.io metadata (licence, size, date): https://crates.io/api/v1/crates/ironcalc
- npm `@ironcalc/wasm` 0.9.1: https://registry.npmjs.org/@ironcalc/wasm
- npm `@ironcalc/wasm-xlsx` 0.9.1: https://registry.npmjs.org/@ironcalc/wasm-xlsx
- Docs: serial numbers https://docs.ironcalc.com/features/serial-numbers.html · error types https://docs.ironcalc.com/features/error-types.html · unsupported features https://docs.ironcalc.com/features/unsupported-features.html
- Project site (function count claim, WASM): https://www.ironcalc.com/

**HyperFormula (tag `3.5.0`)**

- Repository: https://github.com/handsontable/hyperformula/tree/3.5.0
- Licensing: https://hyperformula.handsontable.com/guide/licensing.html
- Excel compatibility (TRUE/FALSE, leap year, `nullDate`, wildcards, whitespace): https://hyperformula.handsontable.com/guide/compatibility-with-microsoft-excel.html
- Configuration reference (`nullDate`, `leapYear1900`): https://hyperformula.handsontable.com/api/interfaces/configparams.html
- Built-in functions (423): https://hyperformula.handsontable.com/guide/built-in-functions.html
- Performance (dependency graph, incremental recalc): https://hyperformula.handsontable.com/guide/performance.html
- Known limitations: https://hyperformula.handsontable.com/guide/known-limitations.html
- Lookup plugin source (wildcard path in VLOOKUP vs MATCH): https://github.com/handsontable/hyperformula/blob/3.5.0/src/interpreter/plugin/LookupPlugin.ts
- Package licence and 3.5.0 release notes: https://github.com/handsontable/hyperformula/releases/tag/3.5.0

**GenOffice (real-world evidence, `main` at 2026-11-13)**

- IronCalc-backed recalc, 1904 skip, strict importer, 1.2 GB model note: https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/src/recalc.rs
- IronCalc dependency: https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/Cargo.toml
- OPC normalisation before IronCalc import: https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/native/xlsx-engine/src/rels_normalize.rs
- Cached-value fallback: https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/src/renderer/formula-cached-fallback.ts
- Preservation-first gateway: https://github.com/genspark-ai/genoffice/blob/main/packages/xlsx-gateway/src/gateway/xlsx-package-io.ts
- Byte-identical compatibility corpus: https://github.com/genspark-ai/genoffice/blob/main/apps/sheets/scripts/verify-compatibility.ts

**Tested harness versions**

- Node 24.15.0, `@ironcalc/wasm` 0.9.1, `@ironcalc/wasm-xlsx` 0.9.1, `hyperformula` 3.5.0 (`dist/hyperformula.full.js`), 2026-11-13. Measurements are single-run and indicative.
