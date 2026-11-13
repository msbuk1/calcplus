# 07 — Workbook data model

Status: resolved
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

## Answer

The model is in [workbook-model.md](../spec/workbook-model.md); the canonical-model decision is recorded in [ADR 0001](../../../docs/adr/0001-canonical-workbook-model.md).

- **Workbook is the truth;** `.xlsx` is an import/export format.
- **Sparse, typed cells:** value, authored formula, saved result, style ID. Dates are numbers plus a format.
- **Shared, append-only style table;** number format is part of the style.
- **Stable IDs** for workbook, sheets and styles; sheet names are labels; cells are row/column.
- **Document/view split:** selection and scroll never mark the file dirty.
- **Dependency graph is derived** by HyperFormula, never persisted.
- **Canonical JSON snapshot** plus original bytes for unmodelled parts.
- **Command log** drives undo/redo and future co-authoring.
- **Opaque parts keep an anchor index** so structural edits move charts, drawings and pivots.
- **Rich text** is plain in v1, preserved unless edited.
