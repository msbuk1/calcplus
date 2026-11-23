# V2-S-03 — Reference model extensions

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 3–5 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [formula semantics](../../spec/formula-semantics.md) §4

## Scope

Extend the parser and evaluator to R1C1 notation, 3D ranges
(`Sheet1:Sheet3!A1`), and the union `(A1:A10,C1:C10)` and intersection
`A1:A10 B1:B10` operators. Keep authored formula text as stored. Structural
edits rewrite these references the same way they rewrite A1 references.

## Acceptance

- [ ] Each reference form parses, evaluates and returns Excel's result.
- [ ] Row and column edits rewrite the reference forms correctly and undoably.
- [ ] Invalid or ambiguous references return the right Excel error.
- [ ] The formula text and cached values round-trip unchanged.
- [ ] Excel 365 opens the saved file without a repair prompt.

## Notes

Structured table references are V2-S-01, not this ticket. External-workbook
references stay preserved; refreshing them is V2-S-04.
