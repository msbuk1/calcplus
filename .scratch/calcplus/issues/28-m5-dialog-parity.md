# 28 — M5 dialog and menu parity rulings

Status: resolved
Type: oracle parity
Blocked by: none

## Question

The 2026-11-18 Excel 365 oracle captures for the M5 menus and dialogs
(transcribed in [design-refs/m5-menus-dialogs.md](../../../docs/design-refs/m5-menus-dialogs.md))
disagree with the existing contracts and map in seven places (D1-D7 there).
Decide each: follow the capture or the contract.

## Answer

Ruled 2026-11-18 by the maintainer against the captures:

- **D1 (filter dropdown ends with Apply, not OK/Cancel): follow the capture.**
  `sort-autofilter` §3.8 amended: the dropdown ends with an Apply button;
  Enter applies and Esc cancels keep their meaning.
- **D2 (greyed out-of-scope entries: Sort By Colour, Filter By Colour, Sheet
  View): keep the contract.** `ui-chrome` §4.1 stands — out-of-matrix commands
  are hidden, not greyed — and §3.8 now records the deviation from Excel's
  greyed display explicitly.
- **D3 (Find dialog Direction dropdown: Up/Down): full parity.** The feature
  matrix's Find & Replace row gains "search direction" in its scope cell and
  its acceptance cell.
- **D4 (label casing: "Find next", "Go to", "Customised Sort"): full parity.**
  Chrome labels copy the oracle captures verbatim, sentence-case quirks
  included. The Sort dialog title is "Customised Sort" while the menu entry is
  "Custom Sort", exactly as the capture shows. This build is UK-English; the
  strings are chrome labels, not localised content, and the pinned reference is
  the maintainer's oracle machine.
- **D5 (Customised Sort entry in the filter dropdown): follow the capture.**
  §3.8's dropdown list now includes it; the entry dispatches the sort command.
- **D6 (Clear Filter from '<column>' wording): follow the capture.** The map
  label becomes "Clear Filter from '<column>'" with the column name in single
  quotes.
- **D7 (Clear/Reapply in the Sort & Filter menu as well as the Data tab):
  follow the capture.** §3.8 now places Reapply and Clear All in the Sort &
  Filter menu (Home → Editing) and in the Data tab.

Contract changes made under this ticket: `feature-matrix.md` (Find & Replace
row), `sort-autofilter.md` §3.8. The command map (`src/commands/command-inventory.json`)
and the ribbon follow in the M5 chrome pass.
