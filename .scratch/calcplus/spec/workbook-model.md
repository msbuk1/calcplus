# calcplus v1 — workbook model

Resolved by [07 — Workbook data model](../issues/07-workbook-model.md) on 2026-11-13. Vocabulary lives in [CONTEXT.md](../../../CONTEXT.md); this document is the implementation-facing model.

## 1. Shape

**Workbook**

- `id` — stable identifier.
- `sheets` — ordered list of sheets; order is the tab order.
- `names` — workbook-scoped defined names. System names owned by an IN feature (`_xlnm._FilterDatabase`, `_xlnm.Print_Area`) live here too and are maintained by that feature's command, not by user name editing.
- `styles` — the shared style table.
- `theme` — theme colours and fonts.
- `namedStyles` — named style definitions (read-only in v1).
- `protection` — workbook protection flags (preserved, not enforced).
- `opaqueParts` — path → opaque part.
- `sourceParts` — path → the original bytes of every part the model reads (workbook, styles, shared strings, worksheets). A dirty save patches into them; see §6.

**Sheet**

- `id` — stable identifier.
- `name` — mutable label.
- `cells` — sparse map `(row, column)` → cell.
- `rowDimensions` — sparse map `row` → `{ height, hidden, styleId? }`.
- `columnDimensions` — sparse map `column` → `{ width, hidden, styleId? }`.
- `merges` — list of rectangular ranges.
- `names` — sheet-scoped defined names.
- `comments` — cell → comment (legacy note or threaded comment).
- `hyperlinks` — cell → hyperlink.
- `validation` — range → validation rule.
- `conditionalFormats` — rules with their ranges.
- `protection` — sheet protection flags (preserved, not enforced).
- `view` — freeze panes, gridline/heading visibility, saved zoom.
- `opaqueParts` — path → opaque part.

**Cell**

- `value` — typed cell value.
- `formula` — authored formula text, or `null`.
- `formulaDerived` — `true` when `formula` was synthesised from a shared-formula group at import rather than authored in the file. Derived text is never written back as a plain formula; the group's `t="shared"`, `si` and `ref` are preserved instead.
- `savedResult` — the result stored in the file, or `null`.
- `styleId` — reference into the workbook style table.
- `richText` — opaque payload, or `null`.

**Cell value**

- number (float64), text, boolean, error (`#DIV/0!`, `#VALUE!`, `#N/A`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`, `#CIRC!`) or blank.
- Dates and times are numbers with a number format; there is no separate date type.

## 2. Styles and number formats

- One shared table of immutable style records: font, fill, border, alignment, number format and (preserved) protection flags.
- Cells reference a style by ID.
- Number format is part of the style: the authored format code plus an optional built-in `numFmtId`. Built-in formats are never materialised; the ID is preserved.
- The table is append-only. Editing a cell's formatting creates or reuses a record; existing IDs are never renumbered, so `.xlsx` style IDs survive a round-trip.
- Theme colours are resolved for display at import; theme references are preserved.

## 3. Identity and addressing

- Workbook, sheet and style IDs are stable.
- Sheet names are mutable labels; the tab order is the ordered sheet list, not an ID.
- Cells are addressed by `(row, column)`, zero-based internally; A1 notation appears only at the edges.
- Formulas display sheet names, but resolve to stable sheet IDs internally, so renaming a sheet does not break references.
- Row heights are stored in points; column widths in Excel character units. Both convert at the render edge.

## 4. Document state versus view state

**Document state** — everything in §1, plus sheet view settings that round-trip: freeze panes, gridline and heading visibility, saved zoom.

**View state** — active sheet, selection, scroll position, focus, open dialogs and other transient UI. Kept per session (optionally persisted locally); never written to `.xlsx` and never marks the document dirty.

## 5. Calculation

- The model owns the authored formula text and the saved result. The engine never becomes the source of the saved text.
- HyperFormula builds the dependency graph. The graph is derived, not persisted; it is rebuilt on open and after structural edits.
- Excel's `calcChain` is preserved as an opaque part but is not trusted.
- Volatile functions and recalc triggers follow [formula semantics](formula-semantics.md).

## 6. Serialization

- The canonical snapshot is versioned JSON metadata plus binary part payloads: workbook, sheets, cells, styles, names, side tables and view settings are JSON; every part payload — modelled source parts and opaque parts, the original bytes — is a binary record, never base64 inside the JSON (persistence.md §2).
- `.xlsx` import maps the modelled parts into the model, keeps their original bytes in `sourceParts`, and stores everything else as original bytes with an anchor index where the part is positioned on a sheet.
- `.xlsx` export patches only dirty parts. Untouched parts are byte-identical; a dirty part is patched in place — its original bytes are the skeleton, and only changed cells, rows and range-bearing elements are rewritten — so unmodelled children, attributes and namespace declarations survive. Style and shared-string tables are append-only.
- The snapshot is used for autosave, crash recovery, tests and a future sync layer. It is not the interchange format. Because it carries every part payload, restore → save is lossless (persistence.md §2).

## 7. Edits and undo

- Every document mutation is a command: `{ id, type, sheetId, range, before, after, payload }`.
- Commands cover cell edits, formatting, structural changes, merges, names, comments, hyperlinks, validation, conditional formatting and protection.
- Undo/redo applies inverse commands. The log is bounded and compacted; it is not the stored document.
- The command log is the replay surface for a future co-authoring layer. View-state changes are not commands.

## 8. Structural edits

Insert/delete rows, columns and cells goes through one range-transform module, which rewrites:

- formulas and their relative references
- merged ranges
- defined names
- validation and conditional-formatting ranges
- hyperlink and comment positions
- row/column dimension records
- anchor indexes for opaque charts, drawings and pivots
- chart series references expressed as plain A1 ranges; charts with a non-range source and pivot source ranges are reported stale and the edit is blocked

An opaque part's bytes change only where the transform requires it: its anchor index, or a chart's series ranges. Everything else stays untouched. Shared-formula masters and followers are handled at the `.xlsx` I/O edge; the model stores authored text per cell.

## 9. Rich text

- Cells store plain text.
- Imported rich text is stored as an opaque payload, displayed as plain text, and preserved on save if untouched.
- Editing a rich-text cell replaces it with plain text after a warning. Authoring rich text is v2.

## 10. Scope

**Modelled:** cells, styles, dimensions and hidden state, merges, defined names, comments, hyperlinks, validation, conditional formatting, protection, freeze panes, theme and named styles.

**Opaque:** charts, pivots, drawings, macros, external links, custom parts and webextensions.

**Deferred:** tables (imported parts preserved; structured references per ticket 02), data-validation enforcement, protection enforcement, rich-text authoring.

## 11. Acceptance criteria

1. A sheet with 100 populated cells in a 1,048,576-row grid uses memory proportional to the populated cells, not the grid.
2. Open → save with no edits returns byte-identical package parts.
3. A round-trip preserves cells, formulas and saved results, styles, merges, dimensions, names, comments, hyperlinks, validation, conditional formatting, protection and freeze panes.
4. Insert/delete rows or columns shifts formulas, merges, names, validation, conditional formatting, comment/hyperlink positions and opaque anchors correctly.
5. Undo/redo restores the exact prior document state.
6. A 1904 workbook keeps its date system; a 1900 workbook keeps the fake 29 February 1900.
7. Renaming a sheet updates formulas and references without breaking.
8. Imported rich text is preserved if untouched; editing it warns and stores plain text.
9. View-state changes never mark the document dirty.
10. Opaque parts survive byte-identical unless their anchor must shift.
11. A dirty save preserves unmodelled children, attributes and namespace declarations inside the part it rewrites.
