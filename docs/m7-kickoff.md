# M7 kickoff — output and hardening

Handover notes for the milestone **M7 — Output and hardening** (8-12 person
weeks). The milestone plan is
[`.scratch/calcplus/spec/milestones.md`](../.scratch/calcplus/spec/milestones.md)
(M7 section: the build list, the spike, the demo, the four exit criteria);
this file adds the entry points and the lane split the plan does not carry.
Added 2026-11-20.

## Read before building

1. `milestones.md` — the M7 section: build list, spike, demo, exit criteria.
2. `.scratch/calcplus/spec/print-pdf.md` — the governing contract for lanes
   1-3. Its §2 pipeline (one pure `PrintPlanner` → `PrintPlan`, three
   renderers: SVG preview, PDFKit vector PDF, 300 DPI canvas fallback) and
   §4 page-setup semantics (print area including multiple areas, orientation,
   fit-to-width, margins, gridlines and headings from the **print** settings,
   merged-cell clipping, renamed font subsets) are the deliverable list.
3. `.scratch/calcplus/spec/performance.md` — lane 4's contract (corpus,
   fan-out, memory, worker queue, degradation ladder, budgets and the
   reference machine).
4. `.scratch/calcplus/spec/persistence.md` — lane 5's contract (autosave
   recovery copy, recent list, external-edit conflict, close prompt, stable
   origin).
5. `.scratch/calcplus/spec/distribution.md` and
   `.scratch/calcplus/spec/licensing.md` — lane 6's contract (localhost
   launcher, portable artifacts, unsigned with the SmartScreen note, licence
   obligations for what ships).
6. `CONTRIBUTING.md` — process rules (signed commits, local checks only,
   dependency and asset reviews).

## What is already banked

- **Print settings are modelled** (M6): `Sheet.printSetup` (margins,
  orientation, fit-to-width, paper size) parsed and written by
  `src/ooxml/print-setup-xml.ts`; the print area is the sheet-scoped
  `_xlnm.Print_Area` name (read through `printAreaOf`). The `fitToPage`
  flag in `sheetPr/pageSetUpPr` is preserved but **not modelled** — lane 1
  models it.
- **The display records**: `src/view/projection.ts` (the `ProjectedCell`
  shape carries value, style, number format, merged rect, comments and
  hyperlinks) is the single source for what the grid paints. Print must
  render from the same records (print-pdf.md §4).
- **Conditional formatting paint** (`src/grid/cf-paint.ts`) and number
  formats carry into print unchanged.
- **Prompt patterns** for confirmations (`src/io/confirm.ts`: macro
  removal, signature invalidation) — the model for lane 5's close and
  conflict prompts.
- **The command map** is authoritative for chrome (`ui-chrome` §4.1);
  Print and Export PDF commands land in `command-inventory.json` before any
  UI appears.
- **Fixture groups** are single-file (`scripts/fixtures/groups/README.md`);
  `npm run fixtures` regenerates byte-identically.

## Lane split (added 2026-11-20)

| # | Lane | Scope | Key deliverables | Files | Effort |
|---|---|---|---|---|---|
| 1 | Print planner | The pure `PrintPlanner`/`PrintPlan`: print area (multiple areas as own page sets), orientation, fit-to-width (model `fitToPage`), margins, gridlines/headings from print settings, page boxes in points, row/column bands, uniform scale, the vector display list, merged-cell clipping (content on the anchor's page) | `src/print/planner.ts` + the `PrintPlan` type everything else renders | `src/print/*`, `src/ooxml/print-setup-xml.ts`, `src/model/types.ts` | 2 pw |
| 2 | SVG preview + print | Per-page SVG preview from the plan, page navigation, `window.print()`, map-driven File → Print, ARIA done properly (the demo wants an NVDA/VoiceOver pass) | preview window | `src/print/svg-renderer.ts`, chrome preview dialog, `src/styles.css` | 1.5-2 pw |
| 3 | PDF export + fallback | PDFKit vector PDF (subsetted fonts renamed per licensing.md), images, 300 DPI canvas fallback for fonts that cannot embed; the download path | `src/print/pdf-renderer.ts`, canvas fallback | `src/print/*`, build config | 2-2.5 pw |
| 4 | Performance hardening | Perf corpus, calculation fan-out, memory, worker queue, degradation ladder; fix the reference-machine budget problem (the rebuild budget flaking on slow machines) | `performance.md` made real | `src/engine/*`, perf fixtures/tests | 2 pw |
| 5 | Persistence | Autosave recovery copy, recent list, external-edit conflict, close prompt, stable origin | `persistence.md` flows | `src/io/*`, App hooks | 1.5-2 pw |
| 6 | Launcher + docs skeleton | Go static server on `127.0.0.1:47823` (localhost only), Windows portable `.exe`, Linux tarball, unsigned with the documented SmartScreen note; docs skeleton (PRD, technical design, ADR index, known deviations) | distribution artifacts | `launcher/`, build scripts, `docs/` | 1-1.5 pw |

### Before the waves

- **The plan's spike runs first** (it is part of the milestone): vector
  print/PDF fidelity and font embedding — SVG through `window.print()`,
  PDFKit in the browser (bundling, subsetting, renaming), and what triggers
  the 300 DPI fallback. Lane 3's design reads its findings
  (`.scratch/calcplus/research/26-print-pdf-fidelity.md`).
- **The `PrintPlan` type is lane 1's first deliverable** (half a day):
  lanes 2 and 3 only consume it, so freezing the shape early lets the
  print lanes run in parallel afterwards.

### Boundary rules

1. **`PrintPlan` is the contract** between lanes 1-3. Lane 1 shapes it;
   lanes 2 and 3 render exactly what it carries and report what is missing
   rather than extending it themselves.
2. **Wiring hooks are orchestrator-merged**: App.tsx, `src/styles.css` and
   the command map get additive last-step hook lines only (lane 2 owns
   preview styling, lane 5 owns persistence prompts; Print/Export PDF are
   map entries before UI).
3. **The display records are read-only** for print lanes: the planner takes
   plain data snapshots with a thin adapter. If the projection needs a new
   seam, it is additive (the `commentCells` precedent) and flagged.
4. **PDFKit is a dependency review** (MIT, chosen in the contract): the
   licence gate stays green and the commit carries the review note.
5. **Fixtures** are single-file group modules; regeneration stays
   byte-identical.

### Parallelism

All six lanes can run at once under the boundary rules; the only ordering is
the spike before lane 3 and the `PrintPlan` freeze before lanes 2 and 3
meaningfully progress. If staffing is limited: wave 1 is lanes 1, 4 and 6
(the interface lands and the independent hardening runs); wave 2 is lanes 2,
3 and 5. Three lanes at a time kept the shared dev server calmer during M6.

### Exit-criteria coverage

1 (print and PDF match the preview) from lanes 1-3; 2 (performance budgets
on the reference machine) from lane 4; 3 (recovery, conflict and close
prompts) from lane 5; 4 (launcher artifacts on Windows and Linux with the
SmartScreen note) from lane 6. The demo maps to lanes 2/3 (report to preview
and PDF), lane 2 (NVDA/VoiceOver), lane 4 (250k cells at 60 fps) and lane 6
(launcher offline).

### Loose ends that may ride along

Ticket 36's sweep items (small chrome/data-tools fixes; lane 4 takes any
that touch its files and reports the rest), the shared pattern-tile table
(tickets 39/40 follow-ups), and the dev-server dependency-reload flake that
sometimes wrecks the first e2e run (lane 4 territory).

## Status

- 2026-11-20: split recorded; the print/PDF spike launched; wave 1 (lanes 1,
  4, 6) launched.
- 2026-11-20: the demo/testing status is recorded in the milestone plan —
  the Linux launcher artifact is verified by hand; the print/PDF check
  against Excel, the 250k scroll, the recovery and conflict prompts and the
  Windows exe are **not tested** for this milestone (the automated suites
  cover their behaviour). The perf budgets ran advisory: no reference
  machine yet.
