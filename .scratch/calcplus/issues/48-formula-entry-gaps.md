# 48 — Follow-up: two formula gaps left after the arity work

Status: open
Type: follow-up (tracking; two small self-contained items)
Blocked by: nothing

## Question

The arity work (2026-11-21) closed the `SUM()` data-loss bug: all 52 functions
carry an argument range, entry refuses a call outside it, and the cell keeps
its previous value. Excel's own parser verified the whole table over COM.

Two gaps in the same family were found while building it and left out of scope.
Both are "a formula we accept that Excel refuses", which is the shape that
costs data when the file is opened.

### 1. Empty parenthesised operands

`prepareFormula` accepts an empty pair of parentheses used as an operand.
Excel refuses these at entry (COM-verified):

- `=SUM(())` — an empty tuple as an argument.
- `=1+()` — an empty right-hand operand.
- `=(1)` variants where the empty group is not a function call.

The scanner counts top-level commas and checks function arity, but an empty
group is an operand problem, not an argument-count problem, so it passes the
screen, reaches HyperFormula and gets written to the file. Excel then strips
the formula on open — the same repair prompt the `SUM()` bug caused.

**Needs:** the scanner rejects an empty parenthesised group that is not a
function call's empty argument list, at entry, with the same refusal message
shape as the arity check. `SUM()` (an empty call) stays refused by arity;
`SUM(())` and `1+()` become refused by this rule.

### 2. Arity is not checked for functions outside the 52

`functionArity()` returns `null` for any name not in the registry, so a typed
call to an unsupported function is screened only for being unsupported, never
for its argument count. `=XLOOKUP()` with a wrong argument count is refused as
an unsupported function (correct), but `=FOO(1)` style calls with no registry
entry carry no arity information at all.

This is lower priority: the unsupported-function screen already keeps the
authored text and a stale cached value, and the file-compat contract says such
formulas are written verbatim (fixture F7 depends on it). The gap matters only
if Excel also refuses the shape — for the functions we do not model, we cannot
check without importing their signatures.

**Needs:** a decision — either record it as a known deviation with a pinning
test, or add arity data for the small set of Excel functions users are most
likely to type that fall outside the 52.

## Acceptance

1. `=SUM(())` and `=1+()` are refused at entry with a plain message; the cell
   keeps its previous value; both are covered by a test in
   `tests/formula-arity.test.ts`.
2. `=SUM(1,2)` and `=1+(2)` still commit — the rule narrows only the empty
   group case.
3. Item 2 has either a recorded deviation with a pinning test or arity data;
   whichever the maintainer picks is written down here as the ruling.
4. A saved workbook still contains no formula Excel refuses, so
   `tests/formula-arity.test.ts`'s round-trip property stays green.

## Notes

- Evidence for item 1 came from Excel 365's parser over COM automation, the
  same harness the arity table was verified with.
- The write-side guard (a formula Excel refuses is never serialised) is being
  handled separately; this ticket is only about the entry screen.
