# 02 — Define formula semantics and function coverage

Status: claimed
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
