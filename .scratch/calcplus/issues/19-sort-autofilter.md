# 19 — Sort and AutoFilter semantics

Status: claimed
Type: grilling
Blocked by: none

## Question

Define exactly how sort and AutoFilter behave in v1.

**Sort**

- Single vs multi-key; header detection; ascending/descending; custom lists.
- What happens to formulas, mixed absolute/relative references, cross-sheet references, defined names, merged cells, hidden rows, errors and blanks.
- Whether sorting rewrites formulas as Excel does or keeps them as text; how the operation interacts with the undo model.
- Acceptance fixtures: values, formulas, merges, hidden rows, filtered rows.

**AutoFilter**

- State model: range, per-column criteria, value checklist, text/number/date conditions, blanks.
- Interaction with hidden rows, sort, recalculation, merged cells and structural edits.
- Round-trip of `autoFilter` XML; preservation when the range is not editable.
- UI behaviour: dropdown, clear, reapply, indicator state.

Output: the semantics contract and conformance fixtures for tickets 07, 12 and 13.
