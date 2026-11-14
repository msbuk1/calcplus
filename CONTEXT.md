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
A part of an `.xlsx` package that calcplus does not model, such as a chart, pivot cache or macro project. It is preserved byte-for-byte, except where a structural edit must transform its anchor index or a chart's series ranges.
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

**Command registry**:
The single catalogue of commands that every UI surface dispatches through.
_Avoid_: Action list, handler map

**Transaction**:
One undoable unit, made of one or more changes.
_Avoid_: Batch, operation

**Focus zone**:
A region that owns keyboard input: the grid, formula bar, name box, ribbon, dialog or panel.
_Avoid_: Context, widget

**Chrome**:
The non-grid UI around the sheet: app header, ribbon or compact bar, formula row, sheet tabs and status bar.
_Avoid_: Shell, frame

**Ribbon**:
The tabbed command surface used as the primary chrome.
_Avoid_: Toolbar, command bar

**Compact bar**:
The one-row responsive command surface used below roughly 900 px.
_Avoid_: Mobile bar, mini ribbon

**Anchor index**:
The cell range or drawing anchor recorded for an opaque part so structural edits can move it.
_Avoid_: Position, bounds

**Recovery copy**:
A local JSON snapshot of unsaved work, kept in browser storage. It is never written over the user's file.
_Avoid_: Autosave file, backup

**File handle**:
The browser's permissioned reference to a file on disk, available through the File System Access API.
_Avoid_: Path, file reference

**Recent list**:
The local list of previously opened workbooks, with file handles where the browser allows them.
_Avoid_: History, recents

**Touched cell**:
A cell that holds a value, formula, style, comment, hyperlink or validation rule. Empty cells cost nothing, so memory scales with touched cells rather than the grid.
_Avoid_: Used cell, populated cell, used range

**Guaranteed envelope**:
The workbook size at which every performance budget holds without a warning: 250,000 touched cells and 50,000 formulas.
_Avoid_: Target size, supported size

**Degraded envelope**:
The larger workbook size that still opens and stays editable with reduced extras: up to 1,000,000 touched cells and 200,000 formulas.
_Avoid_: Oversize, fallback mode

**Working set**:
The memory a calcplus tab actually uses: JavaScript heap, typed arrays, the engine worker's heap and grid canvases, excluding the browser's shared baseline.
_Avoid_: Memory usage, footprint

**Fan-out**:
The number of formulas that depend on one cell. High fan-out is a build-time cliff in the dependency graph, not a recalculation cost.
_Avoid_: Dependency breadth, spread

**Benchmark corpus**:
The fixed set of synthetic shapes and anonymised real workbooks that every performance budget is measured against.
_Avoid_: Test files, sample data

**Conformance oracle**:
The authoritative source of expected Excel behaviour: real Excel 365 on a Windows runner, captured into stored goldens.
_Avoid_: Reference implementation, baseline

**Golden file**:
A committed expected result — a value, error, display string, saved workbook or screenshot — produced by the oracle or a pinned runner and compared in CI.
_Avoid_: Baseline, snapshot

**Release gate**:
The fixed set of conformance, accessibility, performance and bug conditions that must pass before v1 is called faithful.
_Avoid_: Quality bar, definition of done

**Walking skeleton**:
The first milestone: the real OOXML layer, canonical model, minimal grid, one edit and undo, integrated end to end.
_Avoid_: Prototype, MVP

**Milestone gate**:
The exit criteria that must pass before a milestone is accepted; a good demo does not substitute for them.
_Avoid_: Checkpoint, sign-off

**Conformance reserve**:
The contingency inside the final milestone for the long tail of Excel edge cases.
_Avoid_: Buffer, slack

**Corresponding source**:
All source, build scripts and dependency source needed to rebuild and modify a released calcplus binary, as GPLv3 requires.
_Avoid_: Source dump, code drop

**Appropriate legal notices**:
The copyright, no-warranty statement, licence and source link that the app's About dialog must display.
_Avoid_: Licence page, legal footer

**DCO**:
The Developer Certificate of Origin sign-off on every commit, certifying the right to submit the contribution under GPL-3.0-only.
_Avoid_: CLA, contributor agreement

**SBOM**:
The machine-readable software bill of materials attached to each release, listing every component and its licence.
_Avoid_: Dependency list, manifest

**Spec package**:
The PRD, technical design, milestones, contracts and ADRs that together define v1; its index is `.scratch/calcplus/spec/README.md`.
_Avoid_: Docs, binder

**Contract**:
A normative spec file with acceptance criteria and fixtures; where a summary disagrees with a contract, the contract wins.
_Avoid_: Guide, notes

**Architecture decision record**:
A short record of an irreversible choice, the options considered and its consequences, stored in `docs/adr/`.
_Avoid_: Design doc, decision log

**Tolerance**:
The allowed numerical difference when comparing a calculated value to its golden, currently 1e-9 relative or 1e-12 absolute.
_Avoid_: Epsilon, margin

**Preserved feature**:
A construct calcplus keeps byte-for-byte but does not model, render or edit, such as a chart, pivot cache or macro project.
_Avoid_: Unsupported feature, opaque blob

**Compatibility report**:
The read-only list shown on open of preserved, stale or unenforced content, with what calcplus will and will not do with it.
_Avoid_: Warning dialog, import log

**Export report**:
The list shown on save of what changed, what was preserved, what went stale and what needs a warning.
_Avoid_: Save log, error list

**Accessibility mirror**:
The transparent DOM copy of the visible cells that exposes the canvas grid to assistive technology, using absolute row and column indices.
_Avoid_: Screen-reader layer, a11y overlay

**DOM-cell renderer**:
The fallback grid renderer that draws real DOM cells behind the same frame seam, used if the canvas plus mirror fails accessibility or performance testing.
_Avoid_: HTML grid, fallback table

**Live region**:
The visually hidden `aria-live` element that announces edits, structural changes, sort, filter, save and errors without moving focus.
_Avoid_: Announcer, screen-reader toast

**Roving tabindex**:
The focus model where one cell has `tabindex="0"` and every other mirrored cell has `tabindex="-1"`, so focus stays on a real cell as the active cell moves.
_Avoid_: Focus grid, active descendant

**Forced colours**:
The operating-system high-contrast mode exposed through `forced-colors: active`, where the canvas must repaint with system colours because it cannot inherit them.
_Avoid_: High-contrast mode, dark mode

**Sort key**:
A column whose values determine the row order, with a direction; multi-key sorts apply keys in order.
_Avoid_: Sort column, sort level

**AutoFilter**:
The per-column filter attached to a plain range, stored as `autoFilter` XML and evaluated to hide non-matching rows.
_Avoid_: Filter, data filter

**Value checklist**:
The list of distinct typed values in a filter column, with Select All, Blanks and a search box.
_Avoid_: Filter list, value list

**Filter-hidden row**:
A row hidden because it fails the active filter, as distinct from a row the user hid manually.
_Avoid_: Filtered row, invisible row

**Conditional formatting rule**:
A rule attached to one or more ranges that overrides cell formatting while its condition is true; it never changes stored values or styles.
_Avoid_: CF, conditional format

**dxf**:
The differential format record a conditional-formatting rule points to, holding only the properties that rule changes.
_Avoid_: CF style, rule format

**Rule priority**:
The evaluation order of conditional-formatting rules; priority 1 is highest and lower numbers win.
_Avoid_: Rule order, z-index

**Range-statistic rule**:
A rule whose result depends on a whole range — top/bottom, duplicate/unique, colour scale, data bar or icon set — as opposed to a per-cell rule.
_Avoid_: Aggregate rule, range rule
