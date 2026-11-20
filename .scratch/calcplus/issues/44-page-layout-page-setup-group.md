# 44 — Page Layout tab missing the Page Setup group (oracle parity)

Status: resolved
Type: feature parity (chrome against the Excel 365 oracle)
Blocked by: starts after the ticket 42/43 chrome lane lands (shared map and ribbon files)

## Question

Maintainer with a capture of Excel's **Page Layout** tab (2026-11-20):
"nothing like this in our page layout section". The oracle shows a **Page
Setup** group (Page Setup dialog launcher, Print Area drop-down, Breaks
drop-down) and **Sheet Options** (gridlines/headings with View and Print
check boxes; Sheet Right-to-Left). Calcplus currently shows only the two
view toggles (Gridlines, Headings) in Sheet Options — no Page Setup group
at all.

## Ruled shape (v1 subset, oracle layout; maintainer 2026-11-20)

The tab is laid out like the oracle, with the v1 items only. Out-of-scope
items are **hidden** (the map's hidden-not-greyed rule), not shown greyed.

**Page Setup group**
- **Page Setup** — a dialog (the Format Cells house style) carrying the v1
  settings we already model: Page (orientation, Scale-to-fit width/height,
  scale %), Margins, Sheet (print area, print gridlines/headings check
  boxes), plus the paper size list from ticket 42. Every change writes the
  sheet's page setup as one undoable transaction and re-plans any open
  preview.
- **Print Area** drop-down — Set Print Area / Clear Print Area (the
  contract marks print-area authoring as allowed: it is a system name, not
  defined-name authoring).
- **Breaks** — v2 per print-pdf.md §1: hidden.

**Sheet Options group**
- Gridlines and Headings rows with the oracle's two check boxes each:
  **View** (exists — `view.gridlines`/`view.headings`) and **Print**
  (`printOptions` — modelled in the M7 planner lane; authoring them is new).
- **Sheet Right-to-Left** — never in v1: hidden.

Themes, Background, Print Titles and Scale-to-Fit as a separate group:
v2/never per the contracts — hidden.

## Acceptance

1. The Page Layout tab shows Page Setup (Page Setup + Print Area) and
   Sheet Options (Gridlines View/Print, Headings View/Print).
2. The Page Setup dialog edits orientation, fit-to-width/height, scale,
   margins, paper and print area; each change is one undoable transaction
   and an open print preview re-plans.
3. Print Area drop-down sets and clears the area; the print preview and
   PDF honour it immediately.
4. The Print check boxes write `printOptions` (one transaction each) and
   affect the printed output only — the view toggles affect the screen only
   (the planner already reads them separately; keep that).
5. Nothing v2 appears anywhere (hidden, per the map rule).
6. Existing suites stay green; `npm run inventory` green.

## Resolution (2026-11-20)

The Page Layout tab carries the oracle's two groups with v1 items only
(everything else hidden): Page Setup (the dialog plus the Print Area
drop-down) and Sheet Options (Gridlines/Headings rows, View and Print check
boxes each). The dialog edits orientation, scale, fit-to pages, margins,
paper and print area; one OK/Apply is one transaction (at most one
printSetup.set plus one name.set) and shows a live plan line. Print area
set/clear writes the _xlnm.Print_Area system name with byte-exact save
round-trip. Print check boxes write printOptions (printed output only);
the View toggles stay screen-only. Breaks, Themes, Scale to Fit as a group,
Background, Print Titles and Sheet Right-to-Left have no placement.

Follow-up ticket 47 collects the three oracle polish items the lane raised
(Page Setup button in the preview, Add to Print Area, centring check boxes).
