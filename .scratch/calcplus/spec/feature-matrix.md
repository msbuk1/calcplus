# calcplus v1 — feature matrix

Resolved by [01 — Define the v1 feature matrix and acceptance criteria](../issues/01-feature-matrix.md) on 2026-11-13.

**Legend:** **IN** = v1 · **V2** = next effort · **NEVER** = beyond the product.

Criteria are written against Excel desktop / Microsoft 365. "Round-trips" means open → edit → save preserves the feature, and untouched package parts stay byte-identical (ticket 04).

## 1. Grid and editing

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Rectangular range selection and keyboard navigation (arrows, Tab/Enter, Shift-extend, Ctrl+arrow, Page Up/Down, Ctrl+Home/End) | IN | A keyboard-only user can reach and select any cell or range; the listed keys move focus and selection as Excel does. |
| Discontiguous (Ctrl) multi-range selection | V2 | — |
| In-cell editing (F2, double-click, direct typing; Enter/Tab commit, Esc cancel) | IN | Enter commits and moves in Excel's default direction; Esc restores the previous value. |
| Formula bar editing | IN | The bar mirrors the active cell; edits commit identically; long formulas remain editable. |
| Entry AutoComplete from existing values in the column | V2 | — |
| Cut/copy/paste (single range; in-app and OS clipboard) | IN | Values, formulas, formats and merges paste as Excel does; clipboard interop follows ticket 06; pasting from Excel works for the supported formats. |
| Paste Special (values, formulas, formats, transpose, arithmetic) | V2 | — |
| Fill handle: copy, relative-reference adjustment, double-click to adjacent data | IN | Dragging copies content and formatting with Excel's relative-reference adjustment; double-click fills to the last contiguous row. |
| Fill series (linear, growth, date, custom lists) | V2 | — |
| Drag-and-drop move; Shift-drag insert | IN | Dropping moves cells and rewrites references; Shift-drag inserts and shifts cells. |
| Insert/delete rows, columns and cells (shift right/down options) | IN | Structural edits rewrite formulas, merges, names and references as Excel does; undo restores the exact prior state. |
| Row/column resize by drag; autofit by double-click and menu | IN | Dragging sets width/height; autofit sizes to content; imported sizes render. |
| Hide/unhide rows and columns (authoring) | V2 | Imported hidden state is respected and preserved. |
| Merge & Center, Merge Across, Merge Cells, Unmerge | IN | Merged ranges render, select and edit through the anchor cell; merge/unmerge round-trips in `.xlsx`. |
| Undo/redo for every user mutation | IN | Each mutation is undoable and redoable; history is bounded per ticket 09. |
| Find & Replace (values and formulas; sheet or workbook; match case / whole cell) | IN | Finds and replaces in the chosen scope with the chosen options; the operation is undoable as one action. |
| Go To (Ctrl+G and name box) | IN | Accepts A1 addresses, ranges and defined names; invalid input shows an error and leaves selection unchanged. |
| Clear contents / formats / all | IN | Delete clears contents; the Clear menu clears the chosen aspect; hyperlinks and comments are preserved unless "all" is chosen. |

## 2. Data

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Sort a range (one or more keys; header row; ascending/descending) | IN | Rows reorder and relative references rewrite as in Excel; the header row is excluded when flagged; undo restores the original order. |
| AutoFilter on a plain range (value checklist; text/number/date conditions; clear) | IN | Non-matching rows hide; filter state renders and round-trips as `autoFilter`; clearing restores all rows. |
| Filter by colour, Top 10, advanced/custom filter | V2 | — |
| Tables / ListObjects (structured references, table styles) | V2 | Imported table parts are preserved on save. Structured-reference evaluation is decided by ticket 02. |
| Data validation (lists, ranges, input/error messages) | V2 | Imported validation parts are preserved; dropdowns and enforcement are v2. |
| Text to Columns, Remove Duplicates, Subtotals, Group/Outline | V2 | — |
| Flash Fill | V2 | — |
| What-If / Goal Seek / Scenario Manager | NEVER | — |
| Consolidate | NEVER | — |
| PivotTables / Power Query (as features) | NEVER | Pivot parts are preserved on save but not displayed or edited. |
| Data Analysis ToolPak | NEVER | — |

## 3. Formatting

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Font family, size, bold, italic, underline (single/double), strikethrough, colour | IN | Renders and round-trips for the supported properties. |
| Fill: solid and pattern | IN | Renders and round-trips. |
| Borders: all line styles, weights and colours; outline and inside | IN | Presets and custom edges render and round-trip; borders follow merged ranges. |
| Alignment: horizontal/vertical, indent, wrap text | IN | Renders and round-trips; wrapped text expands row height on autofit. |
| Shrink to fit and text rotation | IN (imported) | Imported values render and round-trip; the authoring controls are V2. |
| Format Painter | IN | Picks up all supported formatting and applies it to the target range; Esc cancels. |
| Format Cells dialog | IN | Every supported property can be set from the dialog; the Protection tab is V2. |
| Number formats: built-in categories + full custom grammar | IN | Rendering matches Excel over a conformance corpus of custom format strings; format codes round-trip. |
| Named cell styles and theme colours/fonts (authoring) | V2 | Imported styles and theme colours resolve and preserve. |
| Conditional formatting: render cell-value, text, date, top/bottom, duplicate/unique, data-bar, colour-scale and icon-set rules | IN | Imported rules of these types evaluate and paint as Excel; rule XML is preserved unless a structural edit requires range transformation. |
| Conditional-formatting authoring, rule manager, formula-based custom rules, stop-if-true UI | V2 | — |

## 4. Formulas and names

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Formula bar, in-cell entry, automatic calculation | IN | Covered by the grid/editing criteria above. |
| AutoSum (`Alt+=`, command `formula.autosum`) | IN | Inserts `=SUM(...)` over the adjacent range; Home → Editing group. |
| Function autocomplete and signature hints | IN | Typing `=` and a function name suggests matching functions; choosing one shows argument signatures; keyboard navigation works. |
| Defined names: evaluate, preserve, use in formulas and Go To | IN | Names in imported workbooks resolve and round-trip; no editor in v1. |
| Name Manager / name creation and editing | V2 | — |
| Insert Function dialog | V2 | — |
| Formula auditing: trace precedents/dependents, watch window, error checking, show formulas | V2 | — |
| Manual calculation mode / calculation options | V2 | — |
| Formula semantics, function coverage, errors, arrays, dates, references | IN | Owned by [02 — Formula semantics and function coverage](../issues/02-formula-semantics.md). |

## 5. Workbook and sheets

| Feature | v1 | Acceptance criterion |
|---|---|---|
| New blank workbook | IN | Creates a one-sheet workbook with Excel's default grid and sheet settings. |
| Open, save, save as, close | IN | Mechanism owned by [08 — File handling and persistence](../issues/08-file-persistence.md); fidelity owned by tickets 04 and 12. |
| Multi-sheet: insert, delete, rename | IN | All three work; deleting a sheet with content prompts; undo restores. |
| Reorder, move/copy, hide/unhide sheets, tab colours | V2 | Imported order, hidden state and tab colours render and preserve. |
| Grouping/outline | V2 | Imported outline parts are preserved. |
| Shell protection / workbook protection | IN (preserve) | XML round-trips untouched; cells remain editable in calcplus (deliberate deviation, see below). |
| Protection authoring, enforcement, allow-edit ranges | V2 | — |

## 6. View and navigation

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Freeze panes (rows and columns) | IN | Frozen rows/columns stay fixed while scrolling; state round-trips. |
| Split view | V2 | — |
| Zoom 10–400% (menu, slider, Ctrl+wheel) | IN | Scales grid and content; imported zoom state renders and preserves. |
| Gridlines and headings show/hide | IN | Display-only toggles; no document mutation. |
| Formula bar and name box | IN | Covered by the editing criteria above. |
| Status bar: sum/average/count for the selection; zoom slider | IN | Stats update for numeric selections; slider and menu stay in sync. |
| Page-break preview / page-layout view | V2 | — |
| Custom views / sheet views | V2 | — |
| Assistive-technology support (screen readers, WCAG 2.2 AA) | V2 | v1 keeps keyboard operation and visible focus as ordinary usability and makes no accessibility claim. Contract: [accessibility](accessibility.md). |

## 7. File I/O and printing

| Feature | v1 | Acceptance criterion |
|---|---|---|
| `.xlsx` open and save | IN | Values, formulas, cached results, styles, merges, dimensions, names, comments, hyperlinks, validation, conditional formatting and freeze panes round-trip; untouched parts byte-identical per ticket 12. |
| `.xlsm` open and save; macro parts preserved, never executed | IN | The VBA project survives a round-trip byte-identically; there is no macro UI or execution. |
| CSV import and export | IN | RFC 4180-style quoting and delimiter defaults; values and text round-trip; formulas export as displayed values. |
| Encrypted / password-protected workbooks | IN (reject) | Opening shows a clear unsupported message; the file is not modified. |
| External links | IN (preserve) | Links are preserved; values are not refreshed; link editing is V2. |
| Charts, images, shapes, sparklines | V2 | Parts are preserved on save but not rendered or authored. |
| Drag-and-drop a supported file onto the app to open it | IN | Dropping `.xlsx`, `.xlsm` or `.csv` opens the file; unsupported types show an error. |
| Print and PDF export: print area, orientation, fit-to-width, margins | IN | Printed/PDF output paginates and honours those settings; PDF matches the print preview. |
| Headers/footers, manual page breaks, print titles, multi-page scaling, page order | V2 | — |
| Recent files, autosave, crash recovery | IN | Behaviour owned by [08 — File handling and persistence](../issues/08-file-persistence.md). |

## 8. Read-only fidelity (render/preserve, no authoring)

| Feature | v1 | Acceptance criterion |
|---|---|---|
| Legacy notes and threaded comments | IN | Indicators and popups render; XML round-trips; add/edit/delete and replies are V2. |
| Hyperlinks | IN | Rendered and followable (Ctrl+click); round-trip; insert/edit dialog is V2. |
| Sheet and workbook protection | IN | Preserved on save; not enforced (deliberate deviation). |
| Hidden rows, columns and sheets | IN | Respected on open and preserved on save; authoring is V2. |
| Conditional formatting | IN | Rendered as described in section 3. |
| Theme colours and named styles | IN | Resolved for display and preserved; authoring is V2. |
| Defined names | IN | Evaluated and preserved; no editor. |
| Data validation | IN (preserve) | Preserved; dropdowns and enforcement are V2. |
| Tables | IN (preserve) | Preserved; structured-reference evaluation decided by ticket 02. |
| Pivot caches / PivotTables | IN (preserve) | Preserved but not displayed or edited. |

## 9. Deferred to v2

Paste Special · fill series · discontiguous selection · entry AutoComplete · hide/unhide rows and columns · reorder/move/copy/hide sheets · tab colours · grouping/outline · tables · data validation · filter by colour/Top 10/advanced filter · conditional-formatting authoring · theme and named-style authoring · Name Manager · Insert Function · formula auditing · manual calculation · split view · page-break preview · custom views · full page setup · hyperlink authoring · comment authoring · protection authoring and enforcement · charts/images/shapes/sparklines · assistive-technology support and WCAG conformance · localization · dark mode/theming.

## 10. Never / out of scope

PivotTables and Power Query (as features) · Solver / Goal Seek / Scenario Manager · Data Analysis ToolPak · Consolidate · macro execution / Office Scripts / add-ins · native desktop and mobile/touch apps (the distribution launcher is in scope; a Tauri/Electron desktop app is not) · legacy `.xls` binary and ODS.

## 11. Deliberate deviations and open questions

**Deviations from Excel:**

- Protection is preserved but **not enforced**; calcplus lets the user edit protected cells and saves the protection XML unchanged.
- Comments, hyperlinks and defined names are **read-only** in v1.
- Sort and AutoFilter operate on **plain ranges**; tables and structured references are v2 (evaluation of imported structured refs is pending ticket 02).
- Conditional formatting **renders but cannot be authored**.
- Charts, images, shapes and PivotTables are **preserved but not rendered**.
- Chart series references that are plain A1 ranges are **rewritten** by structural edits; pivot source ranges are not, and an edit that would move one is blocked.
- Data validation is **preserved but not enforced**.

**Open questions graduated to tickets:**

- Number-format engine choice — [18 — Number-format engine: build vs buy](../issues/18-number-format-engine.md)
- Sort and AutoFilter semantics — [19 — Sort and AutoFilter semantics](../issues/19-sort-autofilter.md)
- Conditional-formatting evaluation and rendering — [20 — Conditional-formatting rendering](../issues/20-conditional-formatting.md)
- Print/PDF pipeline — [21 — Print and PDF pipeline](../issues/21-print-pdf.md)
- Accessibility target — [22 — Accessibility target for the grid](../issues/22-accessibility.md)
