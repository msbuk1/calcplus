# calcplus v1 — command system, keyboard and undo

Resolved by [09 — Command system, keyboard parity and undo/redo](../issues/09-command-undo.md) on 2026-11-13. The command-log shape extends [workbook model §7](workbook-model.md).

## 1. Command registry

- Every user action dispatches a command by **stable ID**. Surfaces: ribbon, menus, context menus, keyboard, formula bar and dialogs.
- No surface mutates the workbook model directly.
- A command declares:
  - `id` — stable string, e.g. `edit.copy`, `format.bold`, `sheet.insert`
  - `label`, `category`, `icon`
  - `keybindings` — platform-aware
  - `isEnabled(context)`, `isChecked(context)`
  - `run(context, args)` → `CommandResult`
- `CommandResult`:
  - `changes` — range-based before/after patches
  - `selectionAfter` — optional range to restore
  - `ui` — optional request to open a dialog, show a toast or move focus
- Commands are deterministic, use stable sheet and cell IDs, and read no view state except the selection and active cell passed in.
- Dialogs are commands: `run` opens the dialog; the dialog submits a second command with the payload.
- Command IDs are the seam for future scripting; scripting itself is not v1.

## 2. Change sets

A change is a range-based patch:

| Change | Payload |
|---|---|
| `cell.set` | sheet, range, before/after values, formulas and saved results |
| `style.set` | sheet, range, before/after style IDs |
| `dimension.set` | sheet, axis, index, before/after size or hidden flag |
| `merge.set` | sheet, before/after merged ranges |
| `structure.insert` / `structure.delete` | sheet, axis, index, count, plus all reference rewrites in the same transaction |
| `name.set` | scope, before/after defined name |
| `comment.set`, `hyperlink.set`, `validation.set`, `cf.set`, `protection.set` | sheet, target, before/after |
| `sheet.add` / `delete` / `rename` / `reorder` / `hide` | before/after sheet metadata |
| `sort.apply`, `filter.set` | sheet, range, before/after order or filter state |

Changes reference stable sheet IDs and row/column coordinates. Recalculated results are derived, not changes; the engine recalculates after applying.

## 3. Undo and redo

- One log per document. Each transaction stores its changes and inverse changes.
- Undo applies inverses in reverse order; redo reapplies.
- A new edit clears redo.
- Undo is global per document. Switching sheets to reveal a change is view state and is not logged.
- Undo history is not saved in the recovery copy; closing the tab or reloading a file discards it.
- Structural edits are single transactions with full reference rewrites.
- A sheet delete is terminal (ADR 0008): its change set applies the sheet
  removal plus the reference rewrites, no inverse transaction is recorded, and
  both stacks clear.

## 4. Transaction boundaries

| Action | Transaction |
|---|---|
| Cell edit commit | One per Enter/Tab/blur; typing is coalesced |
| Paste | One |
| Fill-handle drag or double-click fill | One |
| Format range or toolbar toggle | One |
| Insert/delete rows, columns or cells | One, including reference rewrites |
| Sort | One |
| AutoFilter change or clear | One |
| Merge/unmerge | One |
| Sheet add/delete/rename/reorder/hide | One each |
| Name, comment, hyperlink, validation, conditional format, protection | One each |
| Selection or view change | None |

## 5. Memory bounds

- Keep the last **200 transactions or 64 MB**, whichever comes first; drop the oldest.
- Range-based changes keep large pastes compact.
- A single transaction larger than the budget is kept; older entries are dropped to make room.
- Redo entries count toward the budget.
- The bound is configurable in a later version.

## 6. Keyboard map

Platform mapping: `Ctrl` on Windows and Linux. macOS is not supported in v1, so there is no `Cmd` variant.

**Navigation**

Arrows · Ctrl+arrows · Home · Ctrl+Home · End · Ctrl+End · Page Up/Down · Shift+arrows · Ctrl+Shift+arrows · Ctrl+A · Ctrl+Space (column) · Shift+Space (row) · Ctrl+Shift+Space (all) · Ctrl+G (Go To) · Ctrl+Page Up/Down (sheet tabs; browser-reserved, see below).

**Editing**

F2 · Enter/Shift+Enter · Tab/Shift+Tab · Esc · Delete/Backspace · Ctrl+C/X/V · Ctrl+Z/Y · Ctrl+D (fill down) · Ctrl+R (fill right) · Ctrl+Enter (fill selection) · Alt+Enter (line break) · F4 (toggle reference) · Ctrl+; (date) · Ctrl+Shift+; (time).

**Formatting**

Ctrl+B/I/U · Ctrl+1 (Format Cells; browser-reserved) · Ctrl+Shift+~ (General) · `$` (currency) · `%` (percent) · `#` (date) · `@` (time) · `!` (number) · `^` (scientific) · `&` (outline border) · `_` (no border).

**Data**

Ctrl+Shift+L (AutoFilter). Sort uses the ribbon or menu.

**File**

Ctrl+S (save) · Ctrl+Shift+S (Save As; may be reserved) · Ctrl+O (open; may be reserved) · Ctrl+P (print).

**Find**

Ctrl+F · Ctrl+H.

**Sheets**

Shift+F11 (new sheet).

**Formula**

Alt+= (AutoSum) · F9 (recalculate).

**Browser-reserved shortcuts**

Ctrl+1–8 (tab switching) · Ctrl+Page Up/Down (tab switching) · Ctrl+Tab · Ctrl+N/T/W · F5 (reload) · F11 (fullscreen) · F12 (devtools).

Policy:

- Intercept with `preventDefault` only when the grid, formula bar or app owns focus.
- Where the browser wins, document the exception and provide a visible alternative: ribbon/menu button, name box, sheet-tab buttons, or an alternate binding.
- Never use bare-letter shortcuts outside the grid; typing always goes into the cell.
- Ctrl+S saves the workbook, not the page. Ctrl+F/H open in-app find/replace. Ctrl+P opens print.
- The app ships a searchable shortcut map; tests assert it matches the registry.
- Full Excel keyboard parity is impossible in a browser. A desktop wrapper would be needed for the reserved set.

## 7. Focus model

- Zones: grid, formula bar, name box, ribbon/menu, dialog, find/replace panel, sheet tabs.
- One focus owner; a focus manager routes keyboard events by zone.
- **Grid:** active cell; typing starts an edit; F2 edits; arrows navigate; Tab/Enter commit and move.
- **Formula bar:** edits the active cell; Esc cancels and returns to the grid; Enter commits and moves down; Ctrl+Enter commits and stays.
- **Name box:** type an address or defined name; Enter selects; Esc returns to the grid.
- **Ribbon/menu:** arrow navigation; Enter activates; Esc returns to the grid.
- **Dialogs:** modal, focus trapped; Esc cancels; Enter confirms; focus returns to the invoking control.
- **Find/Replace:** panel; Esc closes and returns to the grid.
- After any command, focus returns to the grid unless the command opened a dialog or panel.
- Focus is always visible. Screen-reader announcements are v2 ([accessibility](accessibility.md)).

## 8. Co-authoring log

Each entry:

```
{ seq, actorId, commandId, args, changes, inverseChanges, timestamp }
```

- `seq` is monotonic per document.
- `actorId` is a local device ID, not an account.
- View state and recalculated results are never in the log.
- Undo/redo is expressed as inverse transactions, so a future sync layer can merge rather than rewind.
- No CRDT/OT in v1; the log shape is the seam.

## 9. Acceptance criteria

1. Every surface dispatches the same command ID; no surface mutates the model directly.
2. Enablement and checked state follow the current context.
3. Undo/redo restores the exact prior state for cell edits, formatting, paste, fill, structural edits, sort, filter, merge, names and sheet operations.
4. Typing produces one undo entry per commit, not one per keystroke.
5. A large paste is one transaction and stays within the memory budget.
6. The shortcut map matches the registry; platform mapping works; browser conflicts are documented with alternatives.
7. Focus moves predictably between zones; Esc returns to the grid.
8. Undo of a structural edit restores references; a sheet delete prompts first, then runs terminal and clears the history.
9. The log is replayable with stable IDs; view changes are absent.
10. The memory bound is enforced.
