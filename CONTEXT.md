# calcplus

The spreadsheet application this repo specifies: a browser-based, single-user clone of Excel desktop.

## Language

**Workbook**:
The document a user opens, edits and saves. It owns an ordered set of sheets, workbook-scoped names, and shared resources such as styles.
_Avoid_: File, document, spreadsheet

**Sheet**:
One grid inside a workbook, with its own cells, dimensions, merges and sheet-scoped names.
_Avoid_: Tab, page, worksheet

**Cell**:
A position in a sheet, identified by row and column. It holds a value, an optional formula, an optional saved result, and a style reference.
_Avoid_: Box, field

**Cell value**:
The typed content of a cell: number, text, boolean, error or blank. Dates and times are numbers with a number format.
_Avoid_: Data, content

**Formula**:
The calculation text a user authored for a cell, kept exactly as written. It is not the cell's value.
_Avoid_: Expression, calculation

**Saved result**:
The value a workbook last stored for a formula cell. It is shown when calcplus cannot or has not recalculated the formula.
_Avoid_: Cached result, computed value

**Style**:
A reusable set of visual properties — font, fill, border, alignment and number format — referenced by cells.
_Avoid_: Format, formatting (when referring to the shared record)

**Opaque part**:
A part of an `.xlsx` package that calcplus does not model, such as a chart, pivot cache or macro project. It is preserved byte-for-byte.
_Avoid_: Blob, attachment, unsupported feature

**Workbook model**:
The canonical in-memory representation of a workbook. Opening a file fills it; saving maps it back to `.xlsx`. It is the source of truth, not the file.
_Avoid_: State, store, document model

**Document state**:
The part of a workbook that belongs to the document and is saved — cells, styles, structure, names and sheet view settings.
_Avoid_: App state, UI state

**View state**:
The transient screen state — active sheet, selection, scroll, focus and open dialogs — that is not part of the document.
_Avoid_: UI state, session state

**Style table**:
The workbook-wide, append-only collection of Style records. Cells reference a style by ID.
_Avoid_: Style sheet, format table

**Defined name**:
A user-facing name for a cell, range or formula, scoped to the workbook or to one sheet.
_Avoid_: Named range (when it names a formula), alias

**Merged range**:
A rectangular range presented as one cell, whose value lives in the top-left cell.
_Avoid_: Merge, merged cell

**Command**:
A recorded document edit with enough before/after information to undo, redo or replay it.
_Avoid_: Action, mutation, event

**Anchor index**:
The cell range or drawing anchor recorded for an opaque part so structural edits can move it.
_Avoid_: Position, bounds
