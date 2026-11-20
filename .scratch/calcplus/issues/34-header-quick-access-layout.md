# 34 — App header places Save, Save As and Undo outside the oracle layout

Status: resolved
Type: defect (visual completeness gap against the PRD fidelity target)
Blocked by: none

## Question

Save, Undo and Redo sit as text buttons in a `.quick-access` toolbar on the
right of the app title strip. The Excel 365 desktop oracle puts those commands
in the **Quick Access Toolbar above the ribbon on the top left**, as small icon
buttons, not as labelled buttons on the right of the window. Save As is not in
the oracle's quick access toolbar at all. It lives in the File menu and in the
Save button's split dropdown. Here it is only in the File menu.

Current-state evidence (2026-11-19, main at `a58d1ba`):

- `src/chrome/AppHeader.tsx` lines 42-67 render `.quick-access` with three
  `<button>` elements: Save (`file.save`), Undo (`edit.undo`), Redo
  (`edit.redo`). The container sits last inside the flex header, so it lands on
  the right.
- `src/styles.css` line 86 defines `.app-header` as a flex row with the brand
  first, `.file-meta` in the middle and the quick-access block last.
- `src/commands/command-inventory.json` line 1408 gives `file.saveAs` a File
  menu placement and a shortcut but no quick-access placement.
- `src/chrome/ribbon-model.ts` line 343 maps `file.saveAs` to a
  `save-as-button` test id, but nothing in the header renders it.

The [PRD](../spec/prd.md) fidelity rule is "When Excel and 'clean design'
disagree, match Excel." The [ui-chrome](../spec/ui-chrome.md) contract lists an
**App header** carrying "document name, dirty indicator, quick access (save,
undo, redo)" but does not say where the quick access lives or what it looks
like. That silence let a non-oracle layout ship. The contract needs the oracle
position pinned the same way the ribbon tab set is pinned.

Two more oracle mismatches ride in the same header and belong to this ticket:

1. **The dirty marker is a separate dot.** Excel shows unsaved state on the
   Save button itself (title text and tooltip), not as a marker beside the
   filename. A separate dot is a second signal where the oracle has one.
2. **The quick access buttons are text labels.** Excel uses 16px monochrome
   icons with tooltips. This depends on ticket 33, which owns the icon set.

## Ruled direction

Match the Excel 365 desktop layout: a Quick Access Toolbar above the ribbon,
anchored top left, holding Save, Undo and Redo as small icon buttons with
tooltips. Save As stays out of the quick access toolbar. Unsaved state shows
on the Save button. The filename stays in the title area.

## Acceptance shape

1. The Quick Access Toolbar sits above the ribbon at the top left of the
   window, not on the right of the title strip.
2. It carries Save, Undo and Redo only, as icon buttons with tooltips naming
   the command and its shortcut. Save As is not in it.
3. The unsaved-changes state shows on the Save button. The separate dirty dot
   beside the filename is removed.
4. Keyboard reach is unchanged: Ctrl+S, Ctrl+Shift+S, Ctrl+Z and Ctrl+Y still
   work, and the commands stay reachable through the File menu and the command
   registry. Enablement predicates are unchanged.
5. Every quick-access control keeps an accessible name. Focus order stays
   skip link, ribbon, name box, formula bar, grid, sheet tabs, status bar as
   [accessibility](../spec/accessibility.md) 2.4.3 requires, with the quick
   access in the ribbon zone.
6. `command-inventory.json` and the generated `docs/command-inventory.md` gain
   a quick-access placement for Save, Undo and Redo, and drop any header
   placement. The generator stays the source of truth, per ui-chrome §4.1.
7. The narrow-width rules hold: below 1100 px the quick access keeps its
   icons and tooltips while group labels hide.
8. `ui-chrome.md` §1 is updated to name the oracle position and the icon form,
   so this cannot drift again.

## Placement

Shares ticket 33's lane inside milestone 7, since both are chrome surface work
and 33's icon set is a dependency. Neither touches the model, engine or file
layers. The Format Cells border-preset defect stays a separate ticket.

## Answer

Pending.

## Resolution (2026-11-20)

The sweep (`4051f8b`) had already moved most of this ticket: `AppHeader` now
renders `QuickAccessToolbar` first, so it sits above the ribbon at the top
left; Save, Undo and Redo are 24 px icon buttons with command-and-shortcut
tooltips; the filename's separate dirty dot was deleted and unsaved state moved
onto the Save button; `quickAccess` placements entered the map and the
generator; `ui-chrome` §1 and §2 were rewritten to pin the oracle position.

The one piece still missing was the oracle's Save split drop-down. Save As was
reachable from the File menu and Ctrl+Shift+S only, not from the Save button as
Excel shows it.

- `QuickAccessPlacement` (`src/commands/inventory.ts`) gains an optional `menu`
  of `RibbonMenuItem`s. The generator accepts a quick-access drop-down only for
  `file.saveAs`, checks every row against the map, and prints the rows in the
  generated Quick access column.
- `file.save`'s map entry now carries `{ label: "Save As", commandId:
  "file.saveAs", shortcut: "Ctrl+Shift+S" }` on its `quickAccess.menu`.
- `QuickAccessToolbar` renders that as the Excel split button: the icon button
  plus a 12 px arrow, opening the shared `Menu`. Save As is still not a
  quick-access button of its own; the row dispatches `file.saveAs`.
- `ui-chrome` §1 and §2 now name the split arrow.

Acceptance: all eight items hold. Keyboard reach and enablement are unchanged
(the row and the button both go through the command registry).

Tests: `tests/ribbon-model.test.ts` (drop-down shape), `tests/e2e/
quick-access.spec.ts` (position, icons, tooltips, Save As in the File menu and
on the arrow, arrow geometry, focus return on Esc).
