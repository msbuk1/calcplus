# calcplus v1 — formula semantics

Resolved by [02 — Define formula semantics and function coverage](../issues/02-formula-semantics.md) on 2026-11-13.

## 1. Fidelity contract

V1 supports a **published set of 52 functions**. A supported function must return the same result as Excel desktop / Microsoft 365, including error values, type conversion and edge cases. A function outside the set returns `#NAME?`. If the workbook already contains a cached result for that cell, calcplus displays it and marks it "not recalculated" (see §9). Full Excel function parity is not claimed.

**Engine:** HyperFormula, GPLv3 edition, behind a narrow internal `FormulaEngine` port (see the amendment to ticket 03). The engine's function registry is gated to the 52 published functions.

## 2. Supported functions (52)

| Group | Functions |
|---|---|
| Maths (10) | SUM, PRODUCT, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT |
| Statistics (9) | AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF, MEDIAN |
| Logic (6) | IF, AND, OR, NOT, IFERROR, IFNA |
| Lookup (5) | VLOOKUP, HLOOKUP, INDEX, MATCH, CHOOSE |
| Position (3) | ROW, COLUMN, ROWS |
| Text (8) | CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER |
| Dates (6) | TODAY, NOW, DATE, DAY, MONTH, YEAR |
| Checks (5) | ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA |

## 3. Errors and type conversion

Match Excel exactly.

- **Error values:** `#DIV/0!`, `#VALUE!`, `#N/A`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`.
- **Propagation:** an error argument propagates unless the function handles it (`IFERROR`, `IFNA`, `ISERROR`, `ISNA`).
- **Coercion:** numeric-looking text converts in arithmetic; blank counts as 0; `TRUE`/`FALSE` count as 1/0 in arithmetic; comparisons follow Excel's type ordering.
- **Operator precedence:** reference operators, unary minus, percent, exponent, multiply/divide, add/subtract, concatenation, then comparison — as Excel.
- **Text comparison:** case-insensitive by default.
- **Criteria:** wildcards `*` (any run), `?` (one character), `~` (escape), and text conditions such as `">10"`, `"<>"`, `"<=5"`; case-insensitive.
- **Locale:** English function names and comma argument separators in v1. Localized names and separators are v2.

## 4. References

**Supported:** `A1`, `$A$1`, `$A1`, `A$1`, `Sheet2!A1`, `'Sheet name'!A1`, whole-column `A:A`, whole-row `1:1`, and defined names (workbook- and sheet-scoped, read-only).

**V2:** R1C1 notation, 3D ranges (`Sheet1:Sheet3!A1`), structured table references, union `(A1:A10,C1:C10)` and intersection `A1:A10 B1:B10` operators, and refresh of external-workbook links. External links are preserved and their cached values shown, but never refreshed.

## 5. Dates and times

- Support both Excel date systems: **1900** and **1904**. New workbooks use 1900. The workbook-level flag decides how serial numbers are interpreted.
- Dates are whole-day serial numbers. Times are fractions of a day; a date-time is an integer plus a fraction.
- Reproduce Excel's **fake 29 February 1900** (serial 60) and its leap-year bug for dates before 1 March 1900.
- `TODAY` and `NOW` read the local system clock and are volatile.
- Text is not silently converted to a date; `DATE`, `DAY`, `MONTH` and `YEAR` operate on serials.

## 6. Arrays, volatility and circular references

- **No dynamic arrays or spill.** A formula returns one scalar value. `#SPILL!` is never produced.
- **Legacy array formulas** (Ctrl+Shift+Enter) are preserved from imported files as authored text plus cached result, and are **read-only** in v1: editing one shows a message rather than silently breaking it. Creating new array formulas is v2.
- **Volatile functions:** `TODAY` and `NOW` recalculate on open and after any edit.
- **Circular references:** detect the cycle, show a warning, and display an in-app `#CIRC!` error. Do not iterate. This is a deliberate deviation: Excel shows `0` plus a warning. `#CIRC!` is never written to `.xlsx`; the authored formula and any cached value are preserved.
- **Iterative calculation** is v2.

## 7. Lookup behaviour

- `VLOOKUP` / `HLOOKUP`: exact (`FALSE` or `0`) and approximate (`TRUE` or omitted); approximate matching requires sorted data and follows Excel. Wildcards apply to exact text lookups.
- `MATCH`: match types `1`, `0`, `-1`; wildcards apply to type `0`.
- `INDEX`: array form `INDEX(array, row, column)` is v1; the reference form is v2.
- `CHOOSE`: index plus values.
- `COUNTIF` / `SUMIF`: criteria support wildcards and text conditions; matching is case-insensitive.
- No match returns `#N/A`; invalid arguments return `#VALUE!`.

## 8. Calculation

- Automatic calculation after every edit. Manual calculation mode is v2.
- Recalculate only the cells affected by a change (dependency graph).
- Recalculate volatile functions on open and on every edit.
- Store every result as a cached value in the model and write it to `.xlsx`.

## 9. Storage and round-trip

- Keep the **exact formula text as authored**. Never normalise, rewrite or re-serialise it for storage. Reference updates from model commands (sheet rename, row/column transforms) are edits in the command sense, not rewrites: they change the authored text deliberately and are undoable like any other edit.
- Keep the **cached result** from the file.
- Recalculate when the formula uses only supported functions and references and the inputs are known. Otherwise display the cached result and mark it "not recalculated".
- Unsupported function with no cached result: show `#NAME?`. With a cached result: show the cached result marked stale.
- Preserve `_xlfn.` prefixes and shared-formula structures in the file; the model stores the authored text.
- The formula bar shows the authored text; the engine parses an internal form for calculation only.
- A formula the engine **cannot parse at all** — `@` implicit intersection, unknown constructs, malformed text — takes the unsupported-construct path above: authored text and cached result kept, cell marked stale, never recalculated. The text is never handed to the engine, so nothing can rewrite it. Excel-365-saved files commonly carry `@` forms; they are expected, not errors.
- Dependants recalculate against the cached result and are marked stale in turn.

## 10. Acceptance and conformance

- One test per published function, compared with Excel.
- An error/coercion/operator table test.
- Date-system tests: 1900, 1904, serial 60, pre-1900-03-01.
- Reference tests: absolute, mixed, cross-sheet, quoted sheet names, defined names, whole-column and whole-row.
- Lookup tests: exact, approximate, wildcards, text conditions, no-match.
- Circular-reference detection and warning test.
- Round-trip test: formula text unchanged, cached values preserved, unsupported functions marked stale.

## 11. Deliberate deviations

- Circular references show a warning plus `#CIRC!` instead of Excel's `0` plus a warning.
- Unsupported functions are never approximated; they show `#NAME?` or a stale cached result.
- Legacy array formulas are read-only.
- `@` implicit intersection and other unparseable constructs keep the cached result and are never recalculated (§9). calcplus never strips `@` to force a recalculation, because implicit-intersection edge semantics would silently diverge from Excel.
- A self-reference in a statically dead `IF` branch (for example `=IF(FALSE(),A1,0)` in A1) is flagged `#CIRC!` although Excel evaluates it to a number.
- `INDEX` with `row=0` or `column=0` — Excel returns the whole row or column — returns `#VALUE!` in v1.
- Empty-string coercion edge cases (`""+1` and kin) follow the error/coercion table (ticket 02); unlisted cases are tested against Excel and added to the table before release.
- No dynamic arrays/spill, R1C1, 3D ranges, structured table references, external-link refresh, iterative calculation, manual calculation mode, or union/intersection operators in v1.
