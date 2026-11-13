# 07 — Workbook data model

Status: open
Type: grilling
Blocked by: none

## Question

What is calcplus's canonical in-memory workbook model?

Define:

- Workbook → sheet → cell/range hierarchy, and sparse storage of cells
- Cell content: literal value, formula, cached result, error, rich text?
- Styles and number formats: shared style table vs per-cell
- Row/column dimensions, hidden state, merged ranges
- Defined names (workbook- and sheet-scoped)
- Tables (if v1), comments/notes, hyperlinks, data validation
- Calculation dependency graph: stored, derived, or rebuilt
- View/selection state vs document state
- The serialization boundary: what is canonical JSON, what is derived, what maps to OOXML
- An operation-log shape that would let a future co-authoring layer replay edits

Use domain-modeling vocabulary. Output: the model the rest of the technical design references.
