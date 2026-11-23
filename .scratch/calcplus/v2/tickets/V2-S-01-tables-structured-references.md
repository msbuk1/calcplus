# V2-S-01 — Tables and structured references

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 4–6 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §9 (tables), §2

## Scope

Add the table (ListObject) model: name, range, header row, totals row, banding
and styles. Evaluate structured references in formulas (`Table1[Amount]`,
`Table1[[#Headers],[Amount]]`, `[@Amount]`, `[#All]`). Create, rename and delete
tables from the UI. Imported table parts already round-trip; the new work is
the model, evaluation and authoring.

## Acceptance

- [ ] An imported table renders and its structured references evaluate.
- [ ] Create, rename and delete a table; references update and undo restores.
- [ ] Table styles render and round-trip.
- [ ] AutoFilter on a table range works with the table's own headers.
- [ ] A no-op save stays byte-identical; an edit patches only the dirty parts.
- [ ] Excel 365 opens the saved file without a repair prompt.

## Notes

Structured references are a formula-syntax feature; the table model is the
prerequisite. Keep the existing table XML untouched when the table is not
edited.
