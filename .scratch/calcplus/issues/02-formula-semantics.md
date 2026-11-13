# 02 — Define formula semantics and function coverage

Status: resolved
Type: grilling
Blocked by: none

## Question

What must "Excel-compatible formulas" mean in calcplus v1?

Decide, with a defensible line for each:

- **Function coverage**: target count and named function families (math, text, lookup/reference, date/time, logical, statistical, financial, engineering, information). Name must-have functions and the families deferred.
- **Errors**: error values and propagation (`#VALUE!`, `#REF!`, `#DIV/0!`, `#N/A`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`, and the rest).
- **Coercion**: text vs number, blank vs zero, booleans, and comparison rules.
- **Evaluation**: operator precedence and evaluation order.
- **References**: A1 vs R1C1, relative/absolute/mixed, cross-sheet and 3D references, structured table references, defined names.
- **Dates**: 1900 vs 1904 system, serial dates, time values.
- **Arrays**: legacy CSE arrays vs dynamic arrays and spill — in or out for v1?
- **Volatility**: volatile functions and recalculation triggers.
- **Iteration**: iterative calculation and circular references.
- **Storage**: keep formulas as authored Excel syntax, or an internal AST — and how cached values are treated.

Output: a semantics contract the formula engine and the conformance tests can be built against.

## Answer

The v1 formula contract is in [formula-semantics.md](../spec/formula-semantics.md).

- **52 published functions** (maths 10, statistics 9, logic 6, lookup 5, position 3, text 8, dates 6, checks 5). Anything else shows `#NAME?`, or a stale cached result if the file already had one.
- **Errors, type conversion, operator order and text comparison** match Excel exactly.
- **References:** A1, absolute/mixed, cross-sheet, quoted sheet names, whole-column/row, and defined names. R1C1, 3D ranges, table references, union/intersection and external-link refresh are v2.
- **Dates:** both the 1900 and 1904 systems, including Excel's fake 29 February 1900; serial dates and times.
- **Arrays and loops:** no spill; legacy Ctrl+Shift+Enter arrays are preserved read-only; TODAY/NOW are volatile; circular references warn and show `#CIRC!` instead of iterating.
- **Lookups:** exact and approximate matching, wildcards, text conditions, case-insensitive.
- **Saving:** keep the formula exactly as written plus the cached result; recalculate when possible, otherwise mark it "not recalculated".
- **Licence:** calcplus is GPLv3, so the free HyperFormula edition is used; this amends ticket 03.
