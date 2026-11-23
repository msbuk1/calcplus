# V2-F-17 — Formula auditing

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 2–4 pw (provisional)
- **Prerequisites:** none (soft: V2-S-02)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §4

## Scope

Add Trace Precedents, Trace Dependents, Remove Arrows, the Watch Window, error
checking and Show Formulas. The dependency graph already exists in the engine;
this ticket exposes it in the UI.

## Acceptance

- [ ] Trace arrows list the correct precedents and dependents, including
      cross-sheet references.
- [ ] Remove Arrows clears the traces; navigation follows the traces.
- [ ] The Watch Window tracks chosen cells through edits.
- [ ] Error checking lists the workbook's errors and jumps to each cell.
- [ ] Show Formulas displays formula text without changing the file.

## Notes

Spill ranges and dynamic arrays add trace cases; they wait for V2-S-02 but do
not block the scalar paths.
