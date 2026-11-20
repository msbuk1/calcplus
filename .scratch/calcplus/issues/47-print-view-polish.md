# 47 — Print view polish: preview Page Setup button, Add to Print Area, centring

Status: resolved
Type: feature parity (oracle polish items from ticket 44)
Blocked by: none (the ticket 44 lane owns it)

## Question

The ticket 44 lane asked three questions whose answer is the same under the
PRD fidelity rule (match Excel — the oracle has all three):

1. **Page Setup button in the print preview** (Excel exposes it there), so
   page setup can be edited while the preview stays open. The preview
   already re-plans from the model on change.
2. **Add to Print Area** in the Print Area drop-down, for a second area
   (the model already supports several areas; the dialog's area field
   accepts comma-separated areas and preserves imported multi-area names).
3. **Centre on page** check boxes (horizontally / vertically) in the
   Margins tab — the model and planner already honour
   `horizontalCentered`/`verticalCentered`; only the authoring UI is
   missing.

## Ruled behaviour (2026-11-20, maintainer — via the fidelity rule)

All three are built. Each write is one undoable transaction like the rest
of the page-setup family. Add to Print Area appends to the existing system
name (the area splitter in `print-setup-xml.ts` is exported already).

## Acceptance

1. The preview shows a Page Setup button opening the dialog; edits re-plan
   the open preview.
2. The Print Area drop-down offers Set / Add to / Clear; Add keeps the
   existing areas and adds the selection.
3. The Margins tab has the two centring check boxes; they round-trip and
   the planner honours them.
4. Existing print suites stay green; `npm run inventory` green.

## Resolution (2026-11-20)

All three built: the print preview carries a Page Setup button that opens
the dialog over the preview (a suspended-dialog shell keeps the preview
live and re-planning behind it; Esc closes only the top); the Print Area
drop-down gains Add to Print Area (appends via the exported area splitter,
one name.set transaction, Excel's duplicate-append behaviour kept); and the
Margins tab gains the Centre on page check boxes (one printSetup.set, the
planner honours both axes, unchanged flags write nothing).
