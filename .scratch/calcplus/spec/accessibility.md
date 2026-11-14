# calcplus v1 — accessibility contract

Resolved by [22 — Accessibility target for the grid](../issues/22-accessibility.md) on 2026-11-13. It builds on the [grid spike](grid-spike.md), the [UI chrome](ui-chrome.md) and the [command system](command-system.md).

## 1. Target

**calcplus targets WCAG 2.2 Level AA for its own interface and for the accessible representation of the grid.**

- The canvas is a **visual layer** and is `aria-hidden`. The DOM accessibility mirror is the conformant representation of the sheet.
- Workbook content — cell values, fonts, fills and conditional formatting — is **user-authored**. calcplus guarantees that the value is always exposed to assistive technology, but cannot guarantee the contrast or colour-only meaning of formatting the user chose.
- The default chrome theme must meet 4.5:1 for text and 3:1 for non-text controls. Author formatting may not.
- The app is English-only in v1; announcements are English.

## 2. Applicable criteria

| Criterion | How calcplus meets it |
|---|---|
| 1.1.1 Non-text content | Every icon has an accessible name or adjacent text; the canvas is decorative. |
| 1.3.1 Info and relationships | ARIA grid, rows, cells, headers; labelled dialogs, menus and ribbon tabs. |
| 1.3.2 Meaningful sequence | Mirror DOM order equals visual row-major order; chrome order is logical. |
| 1.4.1 Use of colour | App UI never uses colour alone for state; CF is author content and the value is always exposed. |
| 1.4.3 Contrast (minimum) | Default chrome and grid text must meet 4.5:1; author formatting is a content caveat. |
| 1.4.4 Resize text | 200% browser zoom works; app zoom is 10-400%. |
| 1.4.10 Reflow | Chrome reflows to the compact bar; the grid uses the two-dimensional data-table exception and scrolls. |
| 1.4.11 Non-text contrast | Focus rings, control borders and icons meet 3:1 against their background. |
| 1.4.12 Text spacing | Formula bar, editor and dialogs survive increased spacing; grid cells may clip but the full value is in the formula bar. |
| 1.4.13 Content on hover or focus | Tooltips are dismissible with Esc, hoverable and persistent; no essential information lives only in a tooltip. |
| 2.1.1 Keyboard | Every command is reachable by keyboard through the ribbon, menus, context menu or a shortcut. |
| 2.1.2 No keyboard trap | The grid uses Tab for cells, so **F6** is the documented exit; dialogs trap focus intentionally and Esc closes them. |
| 2.4.1 Bypass blocks | A "Skip to grid" link is the first focusable element. |
| 2.4.2 Page title | The title shows the workbook name and dirty state. |
| 2.4.3 Focus order | Skip link → ribbon → name box → formula bar → grid → sheet tabs → status bar; F6 cycles zones. |
| 2.4.7 Focus visible | A 2 px focus ring with 3:1 contrast on chrome and on the active cell; never removed. |
| 2.4.11 Focus not obscured | Sticky headers and the ribbon never cover the focused cell; keyboard navigation scrolls with padding. |
| 2.5.1 Pointer gestures | No multipoint or path gestures; everything works with a single pointer. |
| 2.5.2 Pointer cancellation | Actions fire on pointer-up; Esc cancels a drag; no down-event activation. |
| 2.5.7 Dragging movements | Fill, resize and move have keyboard alternatives: Ctrl+D/R, Format > Row Height / Column Width, cut/paste. |
| 2.5.8 Target size | Chrome controls are at least 24 × 24 CSS px, with invisible padding where the visual is smaller. Grid cells and resize gutters use the essential/two-dimensional exceptions and have keyboard alternatives. |
| 3.1.1 Language | `lang="en"` on the document. |
| 3.2.1 On focus | Focus never triggers a context change. |
| 3.2.2 On input | Filters apply on OK, not on every keystroke; no unexpected navigation. |
| 3.3.1 Error identification | Invalid name-box input, formula errors and file errors are announced and focused. |
| 3.3.2 Labels | Every input has a visible or programmatic label. |
| 3.3.3 Error suggestion | Invalid addresses suggest the A1 form; unsupported files suggest Save As. |
| 3.3.4 Error prevention | Destructive actions confirm; undo is available. |
| 3.2.6 Consistent Help | The About/help entry and the shortcut map sit in the same place in every view; the app has no embedded help widget. |
| 4.1.2 Name, role, value | Custom controls expose role, name and state; canvas content is represented by the mirror. |
| 4.1.3 Status messages | Save, sort, filter, recalc, print and error messages go through a polite live region. |

**Not applicable:** 1.2.x media, 1.3.4 orientation, 1.3.5 input purpose, 1.4.2 audio control, 2.2.2 pause/stop/hide, 2.3.1 flashing, 2.5.4 motion actuation, 3.3.7 redundant entry, 3.3.8 accessible authentication. **AAA criteria** such as 1.4.6 and 2.4.12 are not claimed.

## 3. Screen-reader semantics

### 3.1 Grid structure

- Root: `role="grid"`, `aria-rowcount="1048576"`, `aria-colcount="16384"`, `aria-label` set to the sheet name.
- Rows: `role="row"` with absolute 1-based `aria-rowindex`.
- Cells: `role="gridcell"` with absolute 1-based `aria-colindex`.
- Headers: `role="columnheader"` and `role="rowheader"` with absolute indices.
- The active cell is a persistent, focusable element with `tabindex="0"`; every other mirrored cell has `tabindex="-1"` (roving tabindex).
- `aria-selected="true"` marks the active cell and every visible cell in the selection.
- Merged anchors carry `aria-rowspan` and `aria-colspan`; covered cells are omitted from the mirror.
- `aria-readonly` is set where editing is blocked. Protection is not enforced in v1, so ordinary cells stay editable.

### 3.2 Focus and navigation

- Focus stays on a real `gridcell`, not on a container using `aria-activedescendant`. This keeps screen-reader application mode anchored to a real cell.
- The active-cell element is **persistent**: it is moved and updated, never recreated on scroll, so focus is not dropped.
- If the active cell scrolls out of view, the mirror still contains it at its absolute position; keyboard navigation scrolls it back into view with padding.
- Arrow keys move the active cell; Shift+Arrow extends the selection; Ctrl+Arrow jumps to the data edge; Home/End, Ctrl+Home/End and Page Up/Down follow the [command system](command-system.md).
- Tab and Enter move cells inside the grid, as Excel does. **F6** moves to the next zone; Shift+F6 moves back; Esc returns to the grid.
- Only the visible range is mirrored, plus the persistent active cell. `aria-rowcount` and `aria-colcount` convey the full sheet.
- Freeze panes mirror each cell once. Split panes (v2) mirror only the active pane.

### 3.3 Names and values

- A grid cell's accessible name is its **formatted display text** after number formatting and conditional-format number formats. Position comes from `aria-rowindex` and `aria-colindex`.
- The name box and formula bar are labelled textboxes. The formula bar shows the authored formula when not editing.
- The status bar's selection statistics are real text in a `role="status"` region.

### 3.4 Announcements

One visually hidden `aria-live="polite"` region carries change announcements. Blocking errors use `aria-live="assertive"`. Announcements are queued and debounced by about 150 ms to avoid chatter.

Examples:

- `B4, 1,234`
- `B4 to D10 selected, 27 cells`
- `B4, formula =SUM(A1:A3), result 6`
- `Edit committed, B4, 1,234`
- `Edit cancelled`
- `Row 5 inserted`
- `Column C deleted`
- `Sheet renamed to Q1`
- `Filter applied, 42 of 100 rows visible`
- `Sorted ascending by Region`
- `Cannot sort: merged cells are not the same size`
- `Saved`
- `Restored unsaved changes`

Navigation itself is announced by the screen reader from the focused cell; the live region is for changes that are not a direct focus move.

## 4. Editing

- The in-cell editor is a labelled `<textarea>` or `<input>` positioned over the cell. Focus moves into it on F2, double-click or typing.
- The editor exposes the formula text, not the formatted result, matching Excel.
- IME composition is respected; Enter during composition does not commit.
- Commit returns focus to the active cell and announces the new value. Cancel returns focus and announces `Edit cancelled`.
- The formula bar and the in-cell editor stay in sync; editing in either updates the other.
- Alt+Enter inserts a line break; Ctrl+Enter fills the selection.
- Read-only cells set `aria-readonly` and refuse the edit with an announcement.

## 5. Mirror and canvas sync

- The canvas is `aria-hidden="true"`; the mirror is the only grid exposed to assistive technology.
- The mirror sits exactly over the canvas, is transparent, `pointer-events:none`, `user-select:none` and `forced-color-adjust:none`. It must never use `display:none` or `visibility:hidden`.
- Mirror DOM order equals visual row-major order.
- After each painted frame, the mirror updates the visible range, the selection flags and the active cell. Nodes are reused; the active cell is persistent.
- Scrolling updates the mirror after the paint. Mouse-wheel scrolling does not move focus; keyboard navigation does.
- Editing overlays the mirror and hides the painted text underneath, so the same text is not read twice.
- The canvas focus ring and the mirror's focused cell are the same logical state; both draw a 2 px, 3:1 ring.
- Conditional formatting never changes the accessible name except through its number format; colours are not announced in v1.

## 6. Chrome, dialogs and menus

- **Ribbon:** `role="tablist"` with `role="tab"` and `role="tabpanel"`; arrow keys move tabs; Enter/Space activates; `aria-selected` and `aria-expanded` reflect state.
- **Menus and context menus:** `role="menu"` with `menuitem`, `menuitemcheckbox` and `menuitemradio`; arrow keys, Home/End, Enter/Space; Esc closes and returns focus. Shift+F10 or the Menu key opens the cell context menu.
- **Dialogs:** `role="dialog"`, `aria-modal="true"`, labelled by the title and described by the body; focus is trapped; Esc cancels; Enter confirms; focus returns to the invoker. Errors use `aria-invalid` and `aria-describedby`.
- **Find/Replace:** labelled fields, a live result count, Esc closes and returns to the grid.
- **Sheet tabs:** `role="tablist"`; arrow keys move; Enter activates. Ctrl+Page Up/Down is browser-reserved, so visible tab buttons are the alternative.
- **Tooltips:** shown on hover and focus, dismissible with Esc, hoverable, persistent while focus remains.
- **Toasts:** `role="status"`; never the only channel for an error.

## 7. Keyboard-only operation

Every command in the registry is reachable without a pointer:

- Ribbon, menus and context menus expose every command.
- The shortcut map from the [command system](command-system.md) covers navigation, editing, formatting, data, file, find and sheets.
- Browser-reserved shortcuts have visible alternatives: sheet-tab buttons, the ribbon, the name box, the File menu.
- The fill handle has Ctrl+D / Ctrl+R and copy/paste; resize has Format > Row Height / Column Width; drag-and-drop move has cut/paste.
- The filter dropdown opens with Alt+Down; Space toggles a value; Enter applies; Esc cancels.
- No action requires a drag, a right-click or a hover.

**No keyboard trap:** the grid's Tab moves cells, so F6 is the documented exit; dialogs trap focus intentionally and close with Esc.

## 8. Visible focus

- A 2 px ring with at least 3:1 contrast against adjacent colours.
- The active cell ring is drawn on the canvas and mirrored by the focused DOM cell.
- Focus is never removed for aesthetics; `:focus-visible` is used only to suppress the ring for pointer clicks on chrome, never on the grid.
- Sticky headers and the ribbon must not obscure the focused cell; keyboard navigation scrolls with at least 24 px of padding.
- In forced-colours mode the ring uses the system `Highlight` colour.

## 9. High contrast, forced colours, zoom and reflow

- **Forced colours:** detect `(forced-colors: active)`. Chrome uses CSS system colours. The canvas cannot inherit them, so it repaints with `Canvas`, `CanvasText`, `Highlight`, `HighlightText`, `GrayText` and `LinkText` read from a probe element. The mirror stays transparent. Repaint on the media-query change.
- **Browser zoom:** 200% and 400% are tested. The mirror geometry follows the canvas.
- **App zoom:** 10-400%; fewer cells are mirrored as zoom increases; the active cell stays in view.
- **Reflow:** below 900 px the ribbon collapses to the compact bar; the grid uses two-dimensional scrolling under the data-table exception.
- **Text spacing:** the formula bar, editor and dialogs reflow without loss; grid cells may clip, but the full value is in the formula bar.
- **Reduced motion:** `prefers-reduced-motion: reduce` disables non-essential transitions.
- **Target size:** chrome controls are at least 24 × 24 CSS px, using invisible padding where the visual is smaller. Grid cells and resize gutters are exempt as essential two-dimensional targets and have keyboard alternatives. The custom scrollbar's hit area is at least 24 px wide even though the visible thumb is narrower.

## 10. Testing

**Automated, in CI**

- axe-core with no serious or critical violations on the chrome, dialogs, menus and the mirror.
- Playwright keyboard suite: every command reachable, focus order, F6 zones, Esc paths, no traps.
- Live-region assertions for navigation, edit commit/cancel, structural changes, sort, filter, save and errors.
- Accessibility-tree assertions: `grid`, `row`, `gridcell`, `columnheader`, `rowheader`, counts and indices.
- Forced-colours emulation and 200%/400% zoom screenshots.
- Mirror/canvas sync: active cell, selection, merged spans, frozen panes.

**Manual**

NVDA/Chrome and VoiceOver/Safari gate the grid milestone; NVDA/Firefox and JAWS/Chrome are added at the M7 accessibility pass and gate the release.

- NVDA with Chrome and Firefox on Windows 11.
- JAWS with Chrome on Windows 11.
- VoiceOver with Safari on macOS 14+.
- Script: navigate, read address/value/formula, select a range, edit, commit, cancel, insert a row, delete a column, sort, filter, undo, open a dialog, open a context menu, switch sheets, find, trigger an error, use forced colours at 200% zoom.
- Sign-off is required. If the mirror fails, the milestone is blocked and the DOM-cell renderer fallback is built behind `RenderFrame`.

## 11. Out of scope for v1

- Mobile and touch screen readers (TalkBack, iOS VoiceOver).
- Braille-specific layouts and sonification.
- WCAG AAA, dark mode and user-authored high-contrast themes.
- Tagged PDF export; the PDF is visual output.
- Charts, images, shapes and PivotTables, which are not rendered in v1.
- Split panes and the accessibility of duplicated panes.
- Localized announcements.
- Announcing every cell during a large range selection; the live region summarises the range.
- Guaranteeing the contrast or colour-only meaning of user-authored cell and conditional formatting.

## 12. Acceptance criteria

1. The canvas is `aria-hidden`; the DOM mirror exposes a `grid` with absolute row and column counts.
2. Rows, cells and headers expose correct roles and absolute indices; merged anchors expose spans.
3. The active cell is persistent and focusable; roving tabindex works; focus is never dropped on scroll.
4. Selection state is exposed with `aria-selected`; the live region summarises ranges larger than the mirror.
5. Navigation, edit commit/cancel, structural changes, sort, filter, save and errors are announced.
6. Every command is reachable by keyboard; F6 cycles zones; Esc returns to the grid; no keyboard trap exists.
7. The focus ring is visible at 3:1, is never obscured, and uses `Highlight` in forced colours.
8. Forced-colours mode repaints the canvas with system colours.
9. 200% browser zoom and 10-400% app zoom work; the mirror follows the canvas.
10. The chrome reflows to the compact bar below 900 px; the grid uses the two-dimensional exception.
11. Chrome controls meet 24 × 24 target size; grid and resize targets have keyboard alternatives.
12. axe-core reports no serious or critical violations on the chrome, dialogs, menus and the mirror.
13. The manual AT script passes with NVDA/Chrome and VoiceOver/Safari before the grid milestone is accepted, and with the full four-combination matrix — adding NVDA/Firefox and JAWS/Chrome — before release.
14. If the mirror fails AT testing, the DOM-cell fallback is built behind `RenderFrame` and the milestone is not accepted without it.
15. Author formatting is documented as a content caveat; the cell value is always exposed.

## 13. Deliberate deviations

- The grid is a two-dimensional data table; the 1.4.10 reflow and 2.5.8 target-size exceptions apply to cells and resize gutters.
- Grid cells use Tab for cell movement, so F6 is the exit; this is a documented, non-trapping exception to the usual Tab order.
- Conditional-formatting colours are not announced; the value and the formatted display text are.
- PDF export is not tagged.
- The AT matrix is desktop screen readers only; mobile assistive technology is v2.
