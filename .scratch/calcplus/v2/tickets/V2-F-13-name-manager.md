# V2-F-13 — Name Manager

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §4

## Scope

Create, edit and delete defined names with workbook or sheet scope, validate
the refers-to formula, and list conflicts. Names already evaluate in formulas
and Go To; this ticket adds the manager and authoring.

## Acceptance

- [ ] The Name Manager lists, filters and edits names and their scopes.
- [ ] A new name works in formulas and in Go To.
- [ ] Invalid or duplicate names are refused with a clear message.
- [ ] Row, column and sheet edits update name references and undo restores.
- [ ] Names round-trip and open correctly in Excel.

## Notes

Sheet-scoped names stay distinct from workbook-scoped names with the same
text. Names feed the function wizard (V2-F-16) and conditional-formatting
formulas (V2-F-14) but do not block them.
