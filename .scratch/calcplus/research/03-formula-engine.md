# 03 — Formula engine: build vs buy

Research for ticket `03-formula-engine`. Date: 2026-11-13.

## Recommendation

**Adopt HyperFormula 3.5.x under a commercial licence for v1, behind a narrow internal
`FormulaEngine` port. Do not build a formula engine from scratch for v1.** HyperFormula is the
most mature standalone, TypeScript, browser-first calculation engine with a real dependency
graph, incremental recalculation, named expressions, a worker-capable runtime and a
custom-function API. It is also small enough to ship (181 KB gzip in a minimal browser bundle)
and is actively maintained (3.5.0, 2026-10-09).

That recommendation has two hard gates, both of which must be closed before the engine is locked
into the technical design:

1. **Commercial licence quote.** HyperFormula is GPLv3-or-proprietary, and the proprietary
   licence is quote-only and priced by end users. A closed-source product cannot use the GPLv3
   option. Obtain a written quote and confirm that the key grants the full function set; since
   3.5.0 a commercial key may grant only part of the library, with excluded functions returning
   `#LIC!` [HF-KEY][HF-CHANGELOG].
2. **Scope fence from ticket 02.** HyperFormula does **not** support 3D references, structured
   table references, `INDIRECT`, relative named ranges, iterative calculation, `LAMBDA`/`LET`, or
   the Excel `A1#` spill operator. It has only partial dynamic-array behaviour: array functions
   spill and `#SPILL!` works, but a plain range reference such as `=A1:A3` returns `#VALUE!`, and
   array-producing functions need their output size known at parse time [HF-LIMITS][HF-ARRAYS].
   If ticket 02 makes any of those v1 requirements, HyperFormula is the wrong buy and the fallback
   is Univer (see below).

If gate 1 fails, or gate 2 cannot be fenced, the fallback is **Univer's Apache-2.0 formula engine
and workbook model** (`@univerjs/engine-formula` 1.4.0 plus `@univerjs/core`/`@univerjs/sheets`),
not a custom engine and not SpreadJS. Univer has materially better Excel coverage — 528 functions,
3D references, iterative calculation, range spill — but it is a platform commitment: a minimal
headless formula stack bundles to **1.75 MB gzip**, and its `.xlsx` import/export requires a
conversion **backend** through Univer Pro, which conflicts with calcplus's no-backend constraint
[UNI-DOCS-IE].

## Target architecture (if HyperFormula is adopted)

- **The workbook model owns the truth**: cell values, *authored* formula text, cached values,
  names/scopes, sheet metadata, styles and number formats. The engine is a compute service, not
  the store of record.
- **A narrow `FormulaEngine` port** exposes `load(snapshot)`, `setCells`, `recalc`, `addSheet`,
  `removeSheet`, `setName`, `getValues`, `getErrors`, `cancel`. Everything else in the app talks
  to the port, not to HyperFormula.
- **HyperFormula runs in a dedicated Web Worker.** The main thread sends deltas and receives
  value/error updates; snapshots are used for undo/redo and file load.
- **An Excel-parity plugin** (custom function plugins + overrides) carries the semantic fixes and
  the missing v1 functions, with regression tests pinned to the HyperFormula version.
- **The `.xlsx` layer (ticket 04)** maps formulas and cached values both ways and normalises
  `_xlfn.`/`@`; it never depends on `getCellFormula()` to reconstruct authored text.
- **A conformance corpus** generated from Excel is the release gate for engine upgrades.

## Why this is a real decision, not a formality

Excel's formula surface is large: Microsoft's alphabetical function list contains about 508
entries including compatibility aliases [MS-FUNCS], the grid is 1,048,576 × 16,384 [MS-LIMITS],
and modern Excel adds dynamic arrays and spill semantics [MS-SPILL], two date systems [MS-DATES],
structured table references [MS-TABLES], implicit intersection (`@`) [MS-AT], `_xlfn.` prefixes
for newer functions in stored files [MS-XLFN], and optional iterative calculation [MS-ITER].
A credible engine is a parser + evaluator + dependency graph + reference/name model + array model
+ hundreds of functions with Excel-specific coercion, error and date rules. Building that in v1
is a multi-engineer-month project with a long conformance tail; it is not a differentiator for a
spreadsheet whose promise is familiarity.

## Options evaluated

### 1. HyperFormula 3.5.0 (Handsontable / Handsoncode)

- Headless, TypeScript, no UI, one workbook with multiple worksheets per engine instance.
- 423 built-in functions in 3.5.0; vendor states 350 of 515 Excel functions (68%) as of 3.1.0
  [HF-FUNCS][HF-COMPAT]. Function families are uneven: 137 statistical, 75 math, 50 engineering,
  29 financial, 26 date, 25 text, 20 lookup/reference, 18 information, 12 database, 3 matrix.
- Real dependency graph with cell and range nodes, range-composition optimisation, and
  incremental recalculation of only affected cells; `suspendEvaluation`/`batch` exist for bulk
  edits [HF-GRAPH][HF-PERF]. Dependencies are collected at **parse** time, which produces
  `#CYCLE!` for formulas Excel evaluates as non-circular (e.g. `=IF(FALSE(),A1,0)`) [HF-DIFF].
- Array support is partial: `SEQUENCE`, `UNIQUE`, `SORT`, `FILTER`, `VSTACK`, `HSTACK` spill and
  `#SPILL!` is reported when blocked, but `=A1:A3` does not spill, `A1#` is not parsed, and
  runtime-sized arrays (`=SEQUENCE(A1)`) return `#VALUE!` [HF-LIMITS][HF-ARRAYS].
- Named expressions are global or sheet-scoped, but only absolute references are allowed and
  relative named ranges are unsupported [HF-NAMES][HF-LIMITS].
- No 3D references, no structured references/tables, no external workbook references, no
  multiple workbooks per instance, no asynchronous functions, no iterative calculation
  [HF-LIMITS][HF-SPEC].
- TypeScript source, ESM/CJS/UMD builds, no DOM requirement; runs in Node (verified). It is pure
  computation, so a Web Worker host is feasible, but no official worker adapter is documented,
  so the worker host is ours to write.
- Default sheet limits are 40,000 rows × 18,278 columns, configurable; Excel's grid is larger
  [HF-SPEC][MS-LIMITS]. Memory is the practical constraint.
- Dual-licensed GPLv3 or proprietary [HF-LICENSE]. Proprietary pricing is quote-only and, per
  Handsontable's FAQ, priced by end users [HF-PRICING]. 3.5.0 added per-function licence gating
  [HF-KEY][HF-CHANGELOG].
- Active maintenance: 3.5.0 released 2026-10-09, ~3.3k commits, 2.8k GitHub stars, 146 open
  issues, maintained by the Handsontable team [HF-REPO][HF-CHANGELOG][HF-NPM].

### 2. Univer formula engine 1.4.0 (DreamNum)

- Apache-2.0, TypeScript, part of the Univer office suite but usable headless
  [UNI-REPO][UNI-DOCS-NODE].
- Vendor claims 528 functions; the package exports 517 implemented function classes (counted
  from `ALL_IMPLEMENTED_FUNCTIONS` in the installed package) [UNI-DOCS-FORMULA][UNI-NPM].
- Dependency tree with dirty-range/incremental calculation, async calculation, and an official
  worker entry (`@univerjs/engine-formula/worker`, `notExecuteFormula`) [UNI-README][UNI-DOCS-WORKER].
- Supports dynamic arrays and spill, including range spill (`=A1:A3`), `#SPILL!`, legacy CSE
  arrays, 3D references, defined names, iterative calculation and `LAMBDA`/`LET` per its docs
  [UNI-DOCS-FORMULA]. One caveat from testing: the docs' comparison table claims `A1#` support,
  but `=SUM(C1#)` returned `#VALUE!` on 1.4.0 in a direct test (see empirical section).
- Engine is coupled to Univer's DI/plugin container and `IWorkbookData` model; using it means
  either adopting Univer's workbook model or writing an adapter/mirror. It is not a drop-in
  replacement for an application-owned workbook model.
- Minimal headless browser bundle (core + sheets + engine-formula + sheets-formula) measured at
  **1.75 MB gzip / 7.96 MB raw**; the engine package alone is 206 KB gzip. The whole stack is
  large relative to HyperFormula.
- `.xlsx` import/export is **not** in the Apache-2.0 core: the docs require a conversion backend
  and `@univerjs-pro/exchange-client`/`@univerjs-pro/sheets-exchange-client` [UNI-DOCS-IE]. This
  is a direct conflict with calcplus's no-backend constraint.
- Very active (1.4.0 on 2026-11-07, 23.7k stars, daily pushes) but still 1.x with a large, fast-
  moving API surface [UNI-REPO][UNI-NPM].

### 3. SpreadJS (MESCIUS) — commercial whole-spreadsheet option

- Closed-source JavaScript spreadsheet component with a calculation engine of 500+ functions,
  dynamic arrays, iterative calculation, formula auditing, `.xlsx` import/export and a Calc
  Worker [SJS-PRODUCT][SJS-CALC].
- Licensing is per developer and per deployment hostname: **$999/developer** plus
  **$12,499/year per commercial domain** (or up to five named hostnames), with add-ons for
  PivotTables, collaboration, etc. [SJS-PRICING].
- This is not a formula-engine purchase; it replaces the grid, workbook model, UI layer and I/O.
  It would collapse much of calcplus into a vendor component and leave little of the product to
  own. It is the right comparison point for "buy the whole spreadsheet", not for "buy the engine".

### 4. xlsx-calc 0.9.2

- MIT, tiny (9.8 KB gzip in a minimal bundle), and designed to recalculate an already-parsed
  SheetJS workbook object [XC-NPM][XC-README].
- Only **46 registry entries** in `src/formulas.js` (including aliases), with optional import of
  Formula.js functions [XC-SRC]. No dependency graph: it walks formulas and relies on lazy value
  objects, so there is no incremental recalculation and no reliable cycle/volatile handling
  [XC-CALC].
- No dynamic arrays (only a `FILTER` name that deliberately throws), no TS types, no named-range
  scope model of consequence, one maintainer, last release 2024-08 [XC-NPM][XC-SRC].
- Not credible as the core engine for an Excel clone. Useful only as a reference implementation
  of a small evaluator.

### 5. Formula.js / "formulas.js"

There is no single canonical package called `formulas.js` today. The maintained project is
**Formula.js** (`@formulajs/formulajs` 4.6.1, MIT), a library of Excel function implementations;
the older `formulas` npm package (0.7.0, 2014) is a different, dormant project [FJS-NPM][FORMULAS-NPM].

- Formula.js exports **399 functions** (counted from the installed package), MIT, active
  (4.6.1, 2026-07-28) [FJS-NPM][FJS-README].
- It has **no parser, no references, no workbook, no dependency graph, no spill semantics** — it
  is a function library, not an engine. It also deliberately diverges from Excel: `DATE`,
  `DATEVALUE`, `EDATE`, `EOMONTH`, `NOW`, `TODAY` return JavaScript `Date` objects instead of
  Excel serial numbers, and several text/math functions were removed [FJS-README].
- It is a plausible source of function bodies for a custom engine, but the adapter work to fix
  date/coercion/error semantics would be substantial. Not an engine choice on its own.

### 6. fast-formula-parser 1.0.19 (parser-only comparison)

- MIT, TypeScript, LL(1), 280 functions, ~91 KB gzip; parses and evaluates a formula against
  `onCell`/`onRange` callbacks [FFP-README][FFP-NPM].
- No workbook model, no dependency graph, no incremental recalculation, no arrays/spill; external
  references, 3D references and some sheet-name syntax are explicitly unsupported. Last npm
  release was 2020, so it is effectively dormant [FFP-README][FFP-NPM].
- Useful as a parser reference; not a v1 engine.

### 7. Custom TypeScript engine

- Full control over Excel semantics, licensing, bundle size, worker architecture and the
  workbook model. No vendor lock-in and no per-user fee.
- Requires: lexer/parser for Excel's formula grammar (including `_xlfn.`, `@`, unions,
  intersections, structured references), AST, dependency graph with range nodes, incremental
  recalculation, cycle/iteration handling, coercion/error/date systems, names with scope, array
  and spill semantics, 150–300 functions for a credible v1, and a conformance corpus.
- Realistic effort for a *credible* v1 is in the multiple-engineer-month range (judgement, not a
  measured estimate); ongoing conformance and Excel-version churn are permanent. Formula.js can
  supply some function bodies, but its date/error semantics diverge [FJS-README].
- Not recommended for v1. It is a reasonable long-term hedge only after product-market fit and
  only if engine licensing becomes a strategic constraint.

## Comparison

### A. Excel fidelity and calculation semantics

| Criterion | HyperFormula 3.5 | Univer 1.4 | SpreadJS | xlsx-calc 0.9.2 | Formula.js | Custom |
|---|---|---|---|---|---|---|
| Excel function coverage | 423 built-ins; vendor says 350/515 (68%) [HF-FUNCS][HF-COMPAT] | 528 claimed; 517 classes exported [UNI-DOCS-FORMULA] | 500+ claimed [SJS-PRODUCT] | 46 [XC-SRC] | 399 exports, functions only [FJS-NPM] | Whatever you build |
| Dependency graph / incremental recalc | Yes; cell + range nodes; parse-time dependency collection [HF-GRAPH] | Yes; dependency tree, dirty ranges [UNI-DOCS-FORMULA] | Yes (vendor) | No; lazy whole-sheet walk [XC-CALC] | No engine | Build |
| Dynamic arrays / spill | Partial: array functions spill, `#SPILL!`; no `A1#`, range refs don't spill, runtime sizes fail [HF-LIMITS][HF-ARRAYS] | Yes: range spill, `#SPILL!`, CSE; `A1#` docs claim but test returned `#VALUE!` [UNI-DOCS-FORMULA] | Yes [SJS-PRODUCT] | No | No | Build |
| Names / workbook-scoped state | Global + sheet-scoped names; absolute refs only, no relative names [HF-NAMES] | Defined names; 3D refs; per-workbook settings [UNI-DOCS-FORMULA] | Full Excel-like model | Basic workbook names only [XC-CALC] | None | Build |
| 3D refs / tables / `INDIRECT` | No / no / no [HF-LIMITS] | Yes / yes / yes per docs | Yes (vendor) | No | No | Build |
| Iterative calculation | No [HF-SPEC] | Yes; max 32,767 iterations, default 100/0.001 [UNI-DOCS-FORMULA] | Yes [SJS-PRODUCT] | No | No | Build |
| Formula text round-trip | `OFFSET` is rewritten to a plain reference in `getCellFormula`; `_xlfn.` and `@` fail to parse [HF-LIMITS] | Stores formula text in cell model; import/export via Pro backend | Full `.xlsx` I/O | Operates on SheetJS objects | n/a | Build |

### B. Engineering, performance, licensing

| Criterion | HyperFormula 3.5 | Univer 1.4 | SpreadJS | xlsx-calc 0.9.2 | Formula.js | Custom |
|---|---|---|---|---|---|---|
| TypeScript / browser fit | TS, ESM/CJS/UMD, no DOM; 181 KB gzip minimal bundle | TS, ESM/UMD; requires core+sheets; 1.75 MB gzip minimal stack | JS component, closed source; full grid | JS, no types; 9.8 KB gzip | JS/TS functions; 25 KB gzip | Full control |
| Web Worker | Engine is pure computation and worker-capable; **no official adapter documented** | Official worker entry; docs recommend worker [UNI-README][UNI-DOCS-WORKER] | Calc Worker [SJS-PRODUCT] | No | Trivially usable | Native |
| Large-workbook performance | 100k-cell chain: 588 ms build / 196 ms recalc; 400k cells: 1.37 s build (Node, one run) | 20k chain completed within a 3 s wait; not a controlled benchmark | Vendor benchmark; no independent test | No graph; not viable at scale | n/a | Depends |
| Licence | GPLv3 or proprietary; quote-only, priced by end users [HF-LICENSE][HF-PRICING] | Apache-2.0 core; Pro features/commercial terms for exchange, external refs [UNI-REPO][UNI-DOCS-IE] | $999/dev + $12,499/yr per commercial domain [SJS-PRICING] | MIT | MIT | None |
| Maintenance risk | Active; 3.5.0 2026-10; commercial vendor | Very active; 1.4.0 2026-11; 1.x API churn | Active; vendor product | Low activity; 0.9.2 2024-08 [XC-NPM] | Active function library [FJS-NPM] | You own it |
| Extensibility | Custom function plugins; parser/graph not extensible; undocumented ability to override built-ins (tested) [HF-CUSTOM] | Plugin/DI; custom functions; dependency APIs; large API surface | Custom functions/API; closed internals | `import_functions` only | Add functions easily | Total |

## Empirical checks

All tests were run on 2026-11-13 with Node v24.15.0 on Linux, using the published npm packages
(`hyperformula@3.5.0`, `@univerjs/engine-formula@1.4.0` + core/sheets, `xlsx-calc@0.9.2`,
`fast-formula-parser@1.0.19`, `@formulajs/formulajs@4.6.1`). They are smoke tests, not a
conformance suite.

| Formula / behaviour | HyperFormula 3.5 | Univer 1.4 |
|---|---|---|
| `=INT(-8.9)` | `-8` (Excel: `-9`) [HF-DIFF] | `-9` |
| `=CEILING.MATH(-4.3,2,2)` | `-4` (Excel: `-6`) [HF-DIFF] | `-6` |
| `=IF(FALSE(),A1,0)` in A1 | `#CYCLE!` [HF-DIFF] | `0` |
| `=AVERAGEIFS(A1:A3,B1:B3,1)` | `#NAME?` (function absent) | not tested |
| `=SEQUENCE(3)` | spills 1,2,3; `#SPILL!` when blocked | spills 1,2,3 |
| `=A1:A3` in a single cell | `#VALUE!` | spills A1:A3 |
| `=SUM(C1#)` where C1 is a spill | `#ERROR!` (no `A1#` operator) | `#VALUE!` (docs claim support) |
| `=SUM(Sheet1:Sheet2!A1)` | `#NAME?` | `110` (3D refs work) |
| Circular `=A1+1`, iterative on/off | not supported [HF-SPEC] | `10` after 10 iterations / `0` when off |
| `=OFFSET(A1,0,1)` then read formula | returns `=B1`, original lost [HF-LIMITS] | not tested |
| `=_xlfn.SEQUENCE(3)` | parse error | not tested |
| `=@A1:A3` | parse error (no `@`) | not tested |
| Override built-in `INT` with a custom plugin | works (returned `-9`), though docs say IDs must be unique [HF-CUSTOM] | n/a |

### Bundle sizes (minified + gzip -9, esbuild, browser ESM)

| Entry | Raw | Gzip |
|---|---:|---:|
| HyperFormula minimal app | 817,547 B | **181,200 B** |
| Univer core + sheets + engine-formula + sheets-formula | 7,961,448 B | **1,751,431 B** |
| Univer engine package alone (UMD) | 1,194,598 B | 206,423 B |
| xlsx-calc | 26,969 B | 9,759 B |
| fast-formula-parser | 327,305 B | 91,082 B |
| Formula.js | 72,419 B | 24,795 B |

These are single-entry bundles, not full applications; code splitting can reduce initial Univer
cost, but the dependency graph is still much larger than HyperFormula's.

### HyperFormula performance (indicative)

- 100,000-cell chain (`A1=1`, `A2=A1+1`, …): **588 ms** to build, **196 ms** to recalculate after
  editing `A1`.
- 50,000 rows × 9 columns (400,000 cells, 50,000 `SUM` formulas): **1,374 ms** to build; editing
  `A1` recalculated in ~0 ms because only one formula depended on it.
- 20,000 volatile `=RAND()` cells: **109 ms** build, **71 ms** recalc after an edit.
- Process RSS was ~600 MB after the combined run, with `maxRows` raised to 200,000.

Caveats: Node, not a browser; one run, no warm-up control; RSS is process-wide rather than
per-workbook. Treat these as order-of-magnitude evidence, not a benchmark.

## Licensing and commercial risk

### HyperFormula

HyperFormula is dual-licensed under GPLv3 or a proprietary licence. If calcplus itself were
released under GPLv3-compatible terms, the GPLv3 option would be free; assuming a closed-source
product, that option is not available. The GPLv3 option is copyleft: shipping a proprietary
frontend bundle that includes HyperFormula is the kind of distribution that triggers
source-availability obligations; this needs counsel, but it is not a viable path for a closed
product. The proprietary option is quote-only, and Handsontable's FAQ describes pricing by the
number of end users [HF-LICENSE][HF-PRICING]. Two specific risks:

- **Price is unknown.** There is no public HyperFormula price. A per-end-user model on a
  consumer spreadsheet could be material. This is the single largest unknown in the
  recommendation.
- **Function gating.** Since 3.5.0 a proprietary key may grant only part of the library; excluded
  functions evaluate to `#LIC!`, and editing/named-expression APIs can throw
  `LicenseCapabilityMissingError` [HF-KEY][HF-CHANGELOG]. The quote must specify that the key
  grants every function and API calcplus needs.

### Univer

The core is Apache-2.0, which is permissive and includes a patent grant [UNI-REPO]. That is
materially better than GPLv3 or per-user pricing. But the parts calcplus most needs beyond the
engine — `.xlsx` exchange, stable external-reference bindings — are Pro/commercial and, for
`.xlsx`, require a server conversion backend [UNI-DOCS-IE]. Using only the Apache core leaves the
xlsx layer to be built anyway. The commercial terms for Pro are not public on the pages reviewed.

### SpreadJS

Public pricing, but expensive and per-domain: $999 per developer plus $12,499/year per commercial
domain, with add-ons [SJS-PRICING]. It is a closed component, so the product's differentiation and
roadmap are constrained by the vendor.

### MIT options

xlsx-calc, Formula.js and fast-formula-parser are MIT, so licensing is a non-issue. Their problem
is capability, not licence.

## Gaps that need custom work under HyperFormula

This is the "exactly what do we build" list if the recommendation is accepted. Items marked
**v1** are likely required for a credible core spreadsheet; the rest can be deferred by ticket 02.

### A. Missing functions (add via custom function plugins)

HyperFormula has broad family coverage but misses specific, commonly used functions. Confirmed
absent via `getAvailableFunctions()`: `AVERAGEIFS`, `LOOKUP`, `INDIRECT`, `XMATCH`, `SORTBY`,
`CONCAT`, `DOLLAR`, `ERROR.TYPE`, `TYPE`, `NUMBERVALUE`, `AGGREGATE`, `RANK`/`RANK.EQ`/`RANK.AVG`,
`PERCENTRANK` family, `FREQUENCY`, `TREND`, `GROWTH`, `LINEST`, `LOGEST`, `INTERCEPT`,
`MDETERM`, `MINVERSE`, `CELL`, `INFO`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, `TAKE`, `DROP`,
`CHOOSEROWS`, `CHOOSECOLS`, `RANDARRAY`, `LAMBDA`, `LET`, `MAP`, `REDUCE`, `SCAN`.

- **v1 candidates:** `AVERAGEIFS`, `LOOKUP`, `RANK`/`PERCENTRANK`, `FREQUENCY`, `CONCAT`,
  `DOLLAR`, `ERROR.TYPE`, `TYPE`, `NUMBERVALUE`, `MDETERM`, `MINVERSE`, `TREND`/`GROWTH`/
  `LINEST`/`LOGEST`/`INTERCEPT`. These are implementable as custom functions with Excel semantics
  and tests.
- **Defer unless ticket 02 says otherwise:** `XMATCH`, `SORTBY`, `TEXTBEFORE`/`TEXTAFTER`/
  `TEXTSPLIT`, `TAKE`/`DROP`/`CHOOSEROWS`/`CHOOSECOLS`, `RANDARRAY`, `AGGREGATE`, `CELL`/`INFO`.
  Some are possible as custom array functions, but array-size handling is constrained (next
  section).
- **Not realistically fixable as custom functions:** `LAMBDA`/`LET`/`MAP`/`REDUCE`/`SCAN` need
  evaluator scoping/closures; `INDIRECT` needs dynamic reference resolution in the dependency
  graph.

### B. Built-in semantic deviations (patch via an "Excel parity" plugin)

HyperFormula's own difference table documents deviations from Excel, including `INT(-8.9) = -8`
instead of `-9`, `CEILING.MATH`/`FLOOR.MATH` honouring only `mode = 1`, and `ISEVEN`/`ISODD` not
truncating fractional values [HF-DIFF]. These are visible to users and easy to hit in real
workbooks.

- **v1:** override `INT`, `CEILING.MATH`, `FLOOR.MATH`, `ISEVEN`, `ISODD` and any other confirmed
  deviations with a parity plugin. We verified that a custom plugin can override a built-in
  (`INT(-8.9)` became `-9`), but the docs say custom function IDs must be unique, so this is
  undocumented behaviour [HF-CUSTOM]. Treat it as a supported-by-test lever, not a stable
  contract; pin the HyperFormula version and add regression tests.
- This is a maintenance liability: every HyperFormula upgrade can change the deviations or the
  override behaviour.

### C. Reference and parser gaps (engine-level; fork or accept)

These cannot be fixed by a function plugin because they change how references are parsed and how
dependencies are tracked:

- **v1 impact, hard:** `INDIRECT` (absent; common in imported workbooks), dynamic `OFFSET`
  arguments (HyperFormula resolves `OFFSET` at parse time and only accepts static integers), and
  `A1#` / range-reference spill (`=A1:A3` → `#VALUE!`) [HF-LIMITS].
- **v1 impact, policy:** `@` implicit intersection and `_xlfn.`-prefixed formulas appear in
  real `.xlsx` files; HyperFormula rejects both [MS-AT][MS-XLFN]. The xlsx adapter can strip
  `_xlfn.` where the function is supported and translate simple `@` cases, but semantics are not
  identical.
- **v1 impact, defer:** 3D references (`Sheet1:Sheet3!A1`), structured table references,
  union/intersection operators, relative named ranges, external workbook references, and
  `LAMBDA`/`LET` [HF-LIMITS].
- **Defer to v2:** tables (already v2 in the map), external links, multi-workbook calculations.

**Architectural mitigation:** make the workbook model own the *authored formula text* and the
*cached value*, and treat HyperFormula as a compute service. Do not rely on
`getCellFormula()` as the source of truth: HyperFormula rewrites `OFFSET` to a plain reference
and normalises other formulas [HF-LIMITS]. A sidecar preserves round-trip fidelity even when the
engine cannot compute a formula.

### D. Engine features (fork or defer)

- Iterative calculation is unsupported [HF-SPEC]. If ticket 02 requires intentional circular
  models, this is a fork or a different engine.
- One workbook per engine instance [HF-LIMITS]: fine for single-user calcplus; external workbook
  links must be preserved as text/cached values, not calculated.
- Array-producing custom functions cannot resize when dependencies change; the size must be
  derivable at parse time [HF-LIMITS]. This limits custom `TEXTSPLIT`, `FILTER`-like and
  `SEQUENCE`-like functions.
- Default `maxRows = 40,000` / `maxColumns = 18,278` must be raised to Excel's
  1,048,576 × 16,384, with a memory budget and a large-workbook test [HF-SPEC][MS-LIMITS].

### E. Workbook and `.xlsx` integration (custom regardless of engine)

- **v1:** an xlsx adapter that maps cell values, formulas, shared formulas, array formulas,
  cached values, error values, date system (1900/1904), named ranges and sheet metadata into the
  engine and back. HyperFormula has no xlsx reader; ticket 04 owns the file layer.
- **v1:** normalise Excel's stored formula conventions (`_xlfn.` prefixes, `@`, `_xlfn.SINGLE`)
  before handing formulas to the engine [MS-XLFN][MS-AT].
- **v1:** preserve authored formula text and cached values in the workbook model so an unsupported
  formula can round-trip unchanged and display its cached value rather than `#NAME?`.
- **v1:** decide how recalc-on-load works: use cached values when the engine cannot compute, and
  mark the workbook "partially calculated" rather than silently writing wrong values.
- **v1:** error-value and date/number type mapping between xlsx, the workbook model and
  HyperFormula (`DetailedCellError`, date subtypes) [HF-COMPAT].

### F. Runtime integration (custom)

- **v1:** a dedicated Web Worker host with a delta protocol (`setCells`, `addSheet`, `setName`,
  `recalc`, `cancel`), request coalescing, backpressure and error propagation. HyperFormula is
  worker-capable but ships no adapter; the protocol and snapshot strategy are ours.
- **v1:** batch/suspend evaluation for paste, fill, undo/redo and structural edits; do not use
  HyperFormula's undo stack as the product undo stack (ticket 09 owns command undo). HyperFormula
  exposes `batch`, `suspendEvaluation`, `resumeEvaluation` and its own `undo`/`redo`
  [HF-PERF][HF-CHANGELOG].
- **v1:** memory limits and eviction for large workbooks; measure browser RSS with the grid and
  formula engine in the same tab.

### G. Conformance harness (custom, non-negotiable)

- **v1:** an Excel-generated golden corpus covering coercion, errors, date systems, names,
  cross-sheet refs, arrays, and the functions calcplus claims to support. Run it against the
  engine on every upgrade and against the parity plugin. Vendor documentation is not a
  conformance suite: HyperFormula's own docs disagree with observed behaviour in places (for
  example, constant arrays are listed as unsupported in "Known limitations" but `SUM({1,2,3})`
  worked in testing), and Univer's docs claim `A1#` support that our test did not reproduce.

## Decision rules and what would change the recommendation

1. **Licence gate.** If the HyperFormula quote is acceptable and grants the full function set,
   proceed with HyperFormula. If it is not, move to Univer (Apache-2.0) rather than starting a
   custom engine.
2. **Scope gate.** If ticket 02 requires `INDIRECT`, 3D references, structured tables,
   `LAMBDA`/`LET`, full dynamic arrays including `A1#`, or iterative calculation in v1, choose
   Univer. HyperFormula would need parser/graph forks that are more expensive than switching.
3. **Round-trip gate.** If the `.xlsx` layer (ticket 04) cannot preserve authored formulas and
   cached values around an engine with partial coverage, the engine choice is secondary: fix the
   workbook model first. A cached-value sidecar makes HyperFormula viable for import fidelity
   even where it cannot compute.
4. **Performance gate.** Re-run the benchmark in a browser with the real grid and a 1M-cell
   workbook. If HyperFormula's memory or recalc latency is unacceptable, Univer's worker model is
   the alternative; a custom engine is still not the v1 answer.
5. **Do not choose SpreadJS** unless the product decision is to stop building a spreadsheet and
   ship a licensed one. It is a credible commercial product but not an engine purchase.

## Evidence thin / open questions

- **HyperFormula commercial price and terms** are not public; the quote is the largest unknown.
  The per-end-user model is described in Handsontable's FAQ, but the actual HyperFormula contract
  may differ [HF-PRICING].
- **Univer conformance** is vendor-claimed. We verified core behaviours (INT, 3D refs, iteration,
  range spill) but not 528 functions, and found an `A1#` discrepancy. No independent conformance
  suite was found.
- **Univer performance** was only smoke-tested (20k chain within a 3 s wait); no controlled
  large-workbook benchmark.
- **SpreadJS performance and bundle size** are vendor claims; we did not license or measure it.
- **Custom engine effort** is a judgement based on the feature surface, not a measured estimate.
- **The `.xlsx` layer** is ticket 04's decision; this ticket only states the engine-facing
  requirements (authored formula text, cached values, `_xlfn.`/`@` normalisation, date/error
  mapping).

## Sources

All URLs accessed 2026-11-13.

| ID | Source |
|---|---|
| HF-REPO | HyperFormula repository — https://github.com/handsontable/hyperformula |
| HF-LICENSE | HyperFormula licence text (GPLv3/proprietary) — https://raw.githubusercontent.com/handsontable/hyperformula/master/LICENSE.txt |
| HF-CHANGELOG | HyperFormula changelog (3.5.0, 2026-10-09; licence gating) — https://github.com/handsontable/hyperformula/blob/master/CHANGELOG.md |
| HF-NPM | npm registry metadata for `hyperformula` 3.5.0 (2026-10-09) — https://registry.npmjs.org/hyperformula |
| HF-FUNCS | Built-in functions (423; categories) — https://handsontable.com/docs/hyperformula/guide/built-in-functions.html |
| HF-CUSTOM | Custom functions (plugin API; unique IDs; array-size limits) — https://handsontable.com/docs/hyperformula/guide/custom-functions.html |
| HF-COMPAT | Compatibility with Microsoft Excel (350/515 functions, 68%) — https://handsontable.com/docs/hyperformula/guide/compatibility-with-microsoft-excel.html |
| HF-LIMITS | Known limitations (no 3D refs, tables, dynamic arrays, iteration, relative names, multiple workbooks) — https://handsontable.com/docs/hyperformula/guide/known-limitations.html |
| HF-DIFF | Runtime differences vs Excel/Google Sheets (INT, CEILING.MATH, IF cycle, etc.) — https://handsontable.com/docs/hyperformula/guide/list-of-differences.html |
| HF-GRAPH | Dependency graph (cell and range nodes, range composition) — https://handsontable.com/docs/hyperformula/guide/dependency-graph.html |
| HF-PERF | Performance (incremental recalc, batch/suspend) — https://handsontable.com/docs/hyperformula/guide/performance.html |
| HF-SPEC | Specifications and limits (40,000 rows default; iterations not supported) — https://handsontable.com/docs/hyperformula/guide/specifications-and-limits.html |
| HF-NAMES | Named expressions (global/sheet scope, absolute refs only) — https://handsontable.com/docs/hyperformula/guide/named-expressions.html |
| HF-ARRAYS | Array formulas and array arithmetic — https://handsontable.com/docs/hyperformula/guide/arrays.html |
| HF-LICENSING | Licensing overview — https://handsontable.com/docs/hyperformula/guide/licensing.html |
| HF-KEY | Licence key behaviour (feature packages, `#LIC!`, gated APIs) — https://handsontable.com/docs/hyperformula/guide/license-key.html |
| HF-PRICING | Handsontable pricing/FAQ (HyperFormula commercial, priced by end users, quote-only) — https://handsontable.com/pricing |
| UNI-REPO | Univer repository (Apache-2.0; 23.7k stars; active) — https://github.com/dream-num/univer |
| UNI-NPM | npm registry metadata for `@univerjs/engine-formula` 1.4.0 (2026-11-07, Apache-2.0) — https://registry.npmjs.org/@univerjs/engine-formula |
| UNI-README | `@univerjs/engine-formula` README (worker entry, `notExecuteFormula`) — https://github.com/dream-num/univer/blob/dev/packages/engine-formula/README.md |
| UNI-DOCS-FORMULA | Univer formula guide (528 functions, 3D refs, iterative calculation, worker, A1# claim) — https://docs.univer.ai/guides/sheets/features/core/formula |
| UNI-DOCS-WORKER | Univer Web Workers guide — https://docs.univer.ai/guides/sheets/getting-started/worker |
| UNI-DOCS-NODE | Headless Univer guide — https://docs.univer.ai/guides/sheets/getting-started/node |
| UNI-DOCS-IE | Univer import/export guide (requires conversion backend and Pro exchange packages) — https://docs.univer.ai/guides/sheets/features/import-export |
| SJS-PRODUCT | SpreadJS product page (500+ functions, dynamic arrays, Calc Worker, xlsx I/O) — https://developer.mescius.com/spreadjs |
| SJS-CALC | SpreadJS calculation engine page (500+ functions) — https://developer.mescius.com/spreadjs/javascript-excel-calculations |
| SJS-PRICING | SpreadJS pricing ($999/developer; $12,499/year per commercial domain) — https://developer.mescius.com/spreadjs/pricing |
| XC-REPO | xlsx-calc repository — https://github.com/fabiooshiro/xlsx-calc |
| XC-NPM | npm registry metadata for `xlsx-calc` 0.9.2 (2024-08-16, MIT) — https://registry.npmjs.org/xlsx-calc |
| XC-README | xlsx-calc README (recalculates a SheetJS workbook; `import_functions`) — https://github.com/fabiooshiro/xlsx-calc/blob/master/README.md |
| XC-SRC | xlsx-calc `src/formulas.js` (46 registered functions) — https://github.com/fabiooshiro/xlsx-calc/blob/master/src/formulas.js |
| XC-CALC | xlsx-calc `src/Calculator.js` (formula walk, no graph) — https://github.com/fabiooshiro/xlsx-calc/blob/master/src/Calculator.js |
| FJS-README | Formula.js README (differences from Excel; Date objects; removed functions) — https://github.com/formulajs/formulajs/blob/master/README.md |
| FJS-NPM | npm registry metadata for `@formulajs/formulajs` 4.6.1 (2026-07-28, MIT) — https://registry.npmjs.org/@formulajs/formulajs |
| FORMULAS-NPM | npm registry metadata for the older `formulas` package 0.7.0 (2014, MIT) — https://registry.npmjs.org/formulas |
| FFP-README | fast-formula-parser README (280 functions; unsupported external/3D refs) — https://github.com/LesterLyu/fast-formula-parser |
| FFP-NPM | npm registry metadata for `fast-formula-parser` 1.0.19 (2020-11-26, MIT) — https://registry.npmjs.org/fast-formula-parser |
| MS-FUNCS | Microsoft Excel functions (alphabetical), ~508 entries — https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188 |
| MS-LIMITS | Excel specifications and limits (1,048,576 × 16,384; 32,767 iterations) — https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 |
| MS-SPILL | Dynamic array formulas and spilled array behavior (`#SPILL!`) — https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531 |
| MS-ITER | Remove or allow a circular reference (iterative calculation, default 100/0.001) — https://support.microsoft.com/en-us/office/remove-or-allow-a-circular-reference-8540bd0f-6e97-4483-bcf7-1b49cd50d123 |
| MS-DATES | Date systems in Excel (1900 vs 1904) — https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487 |
| MS-XLFN | `_xlfn.` prefix in formulas — https://support.microsoft.com/en-us/office/issue-an-xlfn-prefix-is-displayed-in-front-of-a-formula-882f1ef7-68fb-4fcd-8d54-9fbb77fd5025 |
| MS-AT | Implicit intersection operator (`@`) — https://support.microsoft.com/en-us/excel/implicit-intersection-operator |
| MS-TABLES | Using structured references with Excel tables — https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e |

### Local test artifacts

The test scripts and measured bundles used for this note were written to `/tmp/opencode/fe/`
(`hf-test*.mjs`, `univer-*.mjs`, `bundles/`); they are intentionally outside the repo. The repo
is not yet a git repository, so this Markdown file is the durable artifact.
