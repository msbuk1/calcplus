# V2-S-02 — Dynamic arrays and spill

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 4–7 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none (soft: V2-F-01 array-return form, V2-F-17 spill auditing)
- **Contract:** [formula semantics](../../spec/formula-semantics.md) §6

## Scope

Add an array model and spill ranges to the engine: an array result occupies its
anchor cell plus a spill range, paints as a block, and blocks writes inside it.
`#SPILL!` is produced when the target is occupied. Editing or deleting the
anchor clears the spill; undo restores it. Legacy CSE arrays stay read-only as
in v1.

## Acceptance

- [ ] An array-returning formula spills and paints its whole range.
- [ ] A blocked spill shows `#SPILL!` and changes nothing.
- [ ] Editing inside a spill range is refused with a clear message.
- [ ] Undo and redo restore the spill range and its contents.
- [ ] Formula text survives a round-trip; spill cells carry cached values.
- [ ] Excel 365 opens the saved file without a repair prompt.

## Notes

This ticket adds the capability, not the dynamic-array function set. Functions
such as `UNIQUE`, `SORT` and `FILTER` remain outside the published registry
until a later scope decision. `XLOOKUP`'s array-return form waits on this
ticket; its scalar form can ship first.
