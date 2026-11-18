# M5 design reference — menus and dialogs (Excel 365 oracle captures)

Captured from the maintainer's Excel 365 (UK build) on 2026-11-18 and transcribed
here as the visual/behavioural reference for the M5 chrome pass. Labels are
verbatim from the captures, including this build's sentence-case button labels
("Find next") and mixed "Custom Sort" / "Customised Sort" wording. Where the
captures and a contract disagree, the disagreement is listed in §9 and awaits a
ruling; until ruled, the contract wins.

## 1. Sort dialog (menu entry "Custom Sort"; dialog title "Customised Sort")

- Title bar: "Customised Sort" in accent green, close X top right.
- Toolbar row: `+ Add` · `Delete` (trash) · `Copy` · `↑` · `↓` · `Options ⌄` …
  right-aligned: checkbox **My data has headers** (checked in the capture).
- Key table: column headers **Column | Sort On | Order**; the first key row is
  labelled **Sort by**, subsequent rows **Then by**. Each row: three dropdowns —
  column picker (empty in capture), **Cell Values**, **Sort Ascending**.
- Buttons bottom right: green **OK**, white **Cancel**.
- Capture shows one key row; Add/Delete/Copy/↑↓ operate on the selected level.
- v1 notes: sort-on type is Cell Values only (colour/icon sorts are v2 — the
  Sort On dropdown ships with one option); `Options` holds left-to-right
  orientation and case sensitivity, both v2 — see §9 on whether to render it.
- Contract adds one thing the capture lacks: the target range as a read-only
  label (sort-autofilter §2.1). Keep it.

## 2. AutoFilter dropdown (open on column "City")

Top to bottom:

1. **Sort A to Z** (A↓Z icon) · **Sort Z to A** (Z↓A icon) · ~~Sort By Colour~~
2. **Customised Sort**
3. separator · ~~Sheet View ›~~
4. **Clear Filter from 'City'** (greyed when no filter is applied; includes the
   column name in single quotes)
5. ~~Filter By Colour~~
6. **Text Filters ›** (type-dependent: Text/Number/Date Filters)
7. Search box: magnifier icon, placeholder **Search**
8. Checkboxes: **Select All** (no parentheses), then the value checklist in one
   scrollable column (Birmingham, Boston, Bristol, Chicago, Cork, Dublin, …),
   all checked in the capture. Blanks entry sits in the list per contract §3.2.
9. Green **Apply** button (bottom left).

## 3. Sort & Filter menu (Home → Editing → Sort & Filter)

- Bold section heading: **Sort & Filter**
- **Sort Ascending** (A↓Z icon) · **Sort Descending** (Z↓A icon) · **Custom Sort**
- separator
- **Filter** (funnel) · **Clear** (funnel with red x) · **Reapply**
- Note: contract §3.8 places Reapply and Clear All on the Data tab; the capture
  shows them here too (see §9).

## 4. Find & Select menu (Home → Editing → Find & Select)

- Bold section heading: **Find & Select**
- **Find** (magnifier) `Ctrl+F` · **Replace** · **Go to** `Ctrl+G`
- (Go To Special and below are v2 and absent.)

## 5. Find and Replace dialog

- Title **Find and Replace** (green), close X. Tabs: **Find** | **Replace** with
  a green underline on the active tab.
- Find tab: label **Find**, input placeholder **Insert Text**; helper line
  "Wildcards can expand search. For example, "sm?th" finds "smith"." with a
  **Learn More** link; collapsed **> Search options** expander; buttons
  **Find next** (green) and **Find all**.
- Replace tab adds: label **Replace with**, input placeholder **Insert Text**;
  expanded **⌄ Search options** with **Within** dropdown (Sheet) and
  **Direction** dropdown (Down), checkboxes **Match case** and **Match entire
  cell contents**; buttons **Find next** (green), **Find all**, **Replace**,
  **Replace all**.
- Button labels are sentence case in this build.

## 6. Cross-cutting observations

- Primary buttons are solid accent green; secondary buttons are white with a
  thin border; dialog titles are accent green, not body colour.
- Menus show section headings in bold; disabled entries are shown greyed in the
  oracle (but see §9 on calcplus's hidden-not-greyed rule).
- Dropdown/menu icons are leading; shortcuts are trailing, grey.

## 7. Ribbon sub-menus (second capture, 2026-11-18)

Item scope mapped against the feature matrix; per ticket 28 D2, out-of-scope
entries are not rendered at all (the capture shows them greyed).

**Paste menu** (Clipboard → Paste dropdown):
- **Paste** `Ctrl+V` — IN.
- Bold heading **Paste Special**, then: **Values only** `Ctrl+Shift+V`,
  **Formulae only**, **Formatting only**, **Link to source**, **Keep Source
  Column Widths**, **Transpose Rows and Columns** — all Paste Special, out of
  v1 (matrix never-list) → hidden.

**Fill Colour menu** (Font → Fill Colour): **Theme Colours** grid (10 columns:
white/black then the 10 theme accents, each with lighter/darker tint rows;
hover tooltip reads "Green, Accent 6, Li…"), **Standard Colours** row of 10
swatches, then **No Fill**, ~~Highlight when…~~ (CF authoring, V2),
**More Colours…**, ~~Eyedropper~~ (out of scope). Theme swatch naming comes
from the workbook theme (`ThemeRecord`).

**Font Colour menu**: **Automatic** at top (black swatch), then the same Theme
Colours grid and Standard Colours row, **More Colours…**, ~~Eyedropper~~.

**Borders menu** (Font → Borders): a grid of border preset icons (bottom, top,
left, right, none, all, outside, thick box, bottom double, bottom thick…), then
**Draw Border**, **Draw Border Grid**, ~~Erase Border~~, **Border Colour ›**,
**Border Style ›**, **More Border Styles**. Draw/Erase pen modes are out of
matrix → hidden; presets, colour and style flyouts are IN ("Borders: all line
styles, weights and colours; outline and inside").

**Format menu** (Cells → Format):
- **Cell Size**: **Row Height…**, **Auto-fit Row Height**, **Column Width…**,
  **Auto-fit Column Width** — IN (`dimension.set`).
- **Visibility**: ~~Hide & Unhide ›~~ — matrix never-list → hidden.
- **Organise Sheets**: **Rename Sheet…** (IN), ~~Reorder Sheet…~~,
  ~~Copy Sheet~~, ~~Paste Sheet~~ — matrix never-list → hidden.
- **Format Cells** at the bottom — IN (`format.cellsDialog`).

**Insert menu** (Cells → Insert; the Delete menu mirrors it):
- ~~Insert Copied Cells…~~ — pending ruling, see below.
- **Insert Rows Above** `Ctrl+Shift+=`, **Insert Columns to the Left**
  `Ctrl+Shift+=`, **Insert Cells & Shift Right** `Ctrl+Shift+=`, **Insert
  Cells & Shift Down** `Ctrl+Shift+=`, **Insert Sheet** — IN (matrix row
  "insert/delete rows, columns and cells" + sheet insert; the M5 structural
  lane implements them).
- ~~Insert Table Rows Above~~, ~~Insert Table Columns to the Left~~ — tables
  are NEVER → hidden.

Open capture question: the Delete dropdown itself and the exact scope of
"Insert Copied Cells…" (clipboard insert vs the IN Shift-drag insert row).

## 8. Conditional Formatting menu (context for ticket 27)

The second capture shows the Styles group (Conditional Formatting, Format As
Table, Cell Styles) and the Conditional Formatting menu in full: bold heading
**Conditional Formatting**, then **Highlight Cell Rules ›**, **Top/Bottom
Rules ›**, **Data Bars ›**, **Colour Scales ›**, **Icon Sets ›**, **New
Rule…**, **Clear Rules ›**, **Manage Rules…**. Every entry is rule AUTHORING,
which is matrix V2. Ticket 27 ruled A: the button ships in v1 as
present-but-disabled (tooltip: authoring arrives in v2) and this menu does not
ship; the capture is stored as the v2 reference.


## 9. Deltas awaiting a ruling

All ruled 2026-11-18; see [ticket 28](../../.scratch/calcplus/issues/28-m5-dialog-parity.md).
D1/D5/D6/D7 follow the capture (§3.8 amended), D2 keeps the hidden-not-greyed
rule, D3/D4 take full parity (matrix row gains search direction; labels are
verbatim from the captures).

| # | Capture | Contract / map today | Options |
|---|---|---|---|
| D1 | Filter dropdown ends with **Apply** | sort-autofilter §3.8 says "OK and Cancel" | adopt Apply (amend §3.8) or keep OK/Cancel |
| D2 | ~~Sort By Colour~~, ~~Filter By Colour~~, ~~Sheet View~~ shown greyed | ui-chrome §4.1: out-of-matrix commands are hidden, not greyed | keep hidden (deviation from capture) or show greyed |
| D3 | Find dialog has **Direction** (Down/Up) | matrix row names sheet/workbook scope, match case and whole cell only | add Direction (scope note) or omit |
| D4 | Menu item "Custom Sort" vs dialog title "Customised Sort"; "Go to"; "Find next" | map labels are Title Case ("Go To", "Find Next") | match capture verbatim or normalise to one casing |
| D5 | Filter dropdown includes **Customised Sort** entry | §3.8's content list omits it | add to the dropdown (map: menu placement for the sort command) |
| D6 | **Clear Filter from 'City'** (column name) | map label "Clear filter" | adopt the column-name wording (map label change) |
| D7 | Sort & Filter menu contains **Clear** and **Reapply**; §3.8 puts Reapply/Clear All on the Data tab | | place in both (capture) or Data tab only (contract) |
