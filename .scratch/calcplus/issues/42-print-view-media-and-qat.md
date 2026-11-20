# 42 — Print view: paper selection and a Print icon in the Quick Access Toolbar

Status: resolved
Type: feature parity (maintainer-ruled additions to the print view)
Blocked by: none

## Question

Maintainer review of the print view (2026-11-20) named three gaps:

1. **No media/paper selection** — the preview prints Letter (or the sheet's
   paper size) with no way to choose. Excel's print view offers paper
   selection.
2. **No Print icon next to Save** — the Quick Access Toolbar has Save but
   not Print.
3. **No printer control** — answered, not built: the browser's own print
   dialog (opened by the preview's Print button) is the printer control —
   printer, copies and duplex live there. A browser app cannot enumerate
   printers. The preview's hint already says this; make sure it is clear.

## Ruled behaviour (2026-11-20, maintainer)

- **Paper selection** is added to the print preview: a media list (the
  papers the planner's `paperSize` table knows) that updates the plan live
  and **writes the sheet's page setup** as one undoable transaction — Excel's
  print view persists the setting the same way. The PDF export follows
  whatever the plan carries.
- **Print joins the Quick Access Toolbar** next to Save as an icon button
  (map-driven `quickAccess` placement, like Save/Undo/Redo in ticket 34).
- The printer-control answer stays documentation: the browser print dialog
  is the printer control; the preview hint carries the note.

## Acceptance

1. The preview lists the known paper sizes; choosing one re-plans the pages
   live, and the sheet's `pageSetup@paperSize` changes as one undoable edit.
2. The Quick Access Toolbar shows Print next to Save (icon, tooltip,
   map-authorised); it opens the preview.
3. Export PDF uses the chosen paper.
4. The existing print suites stay green.

## Resolution (2026-11-20)

The preview gained a Paper size list (18 named papers resolved from the
planner's own table; unknown imported codes show as Custom). Choosing one
runs the new `print.setup` command: one `printSetup.set` transaction that
writes `pageSetup@paperSize`, restores exactly on undo, and re-plans the
open preview live — including through undo. Export PDF follows the same
plan. Print joins the Quick Access Toolbar (Save 10, Print 20, Undo 30,
Redo 40) with an original printer glyph and tooltip. The printer-control
note names the browser print dialog explicitly. `ui-chrome.md` §1/§2
updated to the ruled toolbar list.
