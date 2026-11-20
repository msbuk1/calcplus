# 32 — Multi-part structural saves and the reference-update policy

Status: resolved
Type: architecture (closes the transform-or-block gap for cross-sheet references)
Blocked by: the M6 exit review (the maintainer ruled this starts after it)

## Question

When rows or columns are inserted or deleted, every reference in the file must
move with the data. Calcplus does this in one central transform
(`src/model/range-transform.ts`), which matches how the open-source
spreadsheets do it. The gap is the **save side**: one structural edit only
saves changes for the sheet being edited, so a hyperlink on another sheet that
points into the moved rows cannot be fixed inside the same transaction.

The current stop-gap (from the wave-2 cleanup, commit `e6a1c30`) is to **refuse
the edit** when it would break such a link. That keeps the "never a silently
stale anchor" rule, but it blocks edits the user expects to work.

Research note 25 (`.scratch/calcplus/research/25-structural-edit-reference-updates.md`,
2026-11-19) read the source of LibreOffice Calc, OnlyOffice, IronCalc, Univer
and Grist. Findings that bear on this ticket:

1. Every engine uses one central reference-update pass with registered
   reference holders, and one structural edit is one undo step.
2. No project refuses an edit to protect a hyperlink. Broken references become
   `#REF!` with a notice. Refusal is used only where data would be lost (an
   array formula cannot be split; a pivot would be damaged).
3. Univer re-bases hyperlink targets in the same edit and marks a gone target
   explicitly. OnlyOffice and IronCalc leave some targets stale; the note reads
   those as missing owners, not as policy.
4. Multi-part consistency comes from one model state and one save pass that
   writes every part the edit touched.

## Ruled direction (2026-11-19, maintainer)

Build the standard shape, on top of what exists:

1. **One edit, one transaction, one save.** A structural edit produces the
   renumbered bytes for every part it touches — the edited sheet, sibling
   sheets whose references move, charts, drawings, comments — as
   `StructuralPartPatch` entries (the type already exists) and the save writes
   them together. Undo restores every `before` byte exactly (the existing
   exact-inverse rule).
2. **Broken and marked, not refused.** A reference that cannot be fixed
   collapses to `#REF!` and the user gets one notice. This replaces the
   cross-sheet-link refusal when this work lands.
3. **Refuse only for data loss.** The block path stays, but only for the cases
   the research confirms: a split array formula, pivot damage, and constructs
   that would silently point at the wrong thing and cannot be marked.
4. **Retire the stop-gap.** The cross-sheet-link block and its message come
   out in the same change.

## Acceptance shape

- Insert/delete rows or columns moves hyperlink targets, names, formulas,
  chart series and drawing anchors on **every** sheet, in one undo step.
- A link or reference whose target is gone reads `#REF!` (or the agreed broken
  marker) and the report lists it once; the edit still applies.
- A save after such an edit writes all changed parts; reopening in Excel shows
  consistent references (this is checked in the maintainer's demo run).
- New fixtures: a cross-sheet link through insert/delete, and the broken-link
  file convention (research note 25 §9).

## Open questions (research note 25 §9)

1. Are chart and drawing parts in the registry scope for v1, or is the
   worksheet-plus-links scope enough?
2. When a row holding a comment is deleted, do comments move or delete? The
   research says every project moves them with the cells; our current answer
   should be checked against the Excel oracle.
3. The file convention for a broken internal link target (Univer writes an
   explicit error marker; plain `#REF!` in `location` may not round-trip) needs
   an Excel capture before it is fixed.

## Resolution (2026-11-19)

Implemented in `src/model/range-transform.ts` and `src/ooxml/`, plus the
structural command's one notice.

- **One edit, one transaction, one save.** The plan now carries
  `crossSheetHyperlinks` for every other sheet whose internal target moves. The
  command packs each sibling worksheet part and its hyperlink before/after map
  into the same `StructuralPartPatch` list, so the save writes the bytes and
  undo restores both. The save also writes any worksheet whose retained source
  bytes differ from the package, so a link-only sheet with no dirty cell is
  still saved. Target cells, formulas, defined names, chart series and drawing
  anchors already travelled in the same transaction.
- **Broken and marked, not refused.** The cross-sheet block
  (`foreignLocationChanges` / the `anchor-untransformable` message for a
  sibling link) is gone. A target the edit removes becomes `#REF!` in formulas,
  defined names, hyperlinks and chart series. The edit applies and the command
  returns one toast that names the count and the first broken reference.
  Refusal remains only for the existing data-loss cases: a split grouped
  formula, a pivot source or output that would be damaged, a merge or filter
  split, and a construct the module cannot mark.
- **Fixtures.** `fixtures/compat/cross-sheet-link.xlsx` (formula, name,
  hyperlink, chart series and a link-only sibling sheet through insert and
  delete) and `fixtures/compat/broken-internal-link.xlsx` (the provisional
  broken-link marker) are generated by `scripts/generate-fixtures.mjs` and
  covered by `tests/ticket32-multi-part-saves.test.ts`.

### Open questions

1. **Charts and drawings in v1 scope:** yes. Chart series ranges and drawing
   anchors already run through the one transform; the cross-sheet fixture
   proves a chart on another sheet follows an insert or delete on the source
   sheet. Pivot sources stay block-only.
2. **Comments on a deleted row:** they move with their cells. A comment whose
   anchor cell is deleted is dropped, because the model's comment map is keyed
   by cell and there is no cell left to attach it to. This matches the current
   contract and the research projects' row-delete behaviour (Grist deletes;
   the others move a comment while its cell survives). The Excel-oracle
   confirmation is deferred to the maintainer's demo run because Excel is not
   available in this environment.
3. **Broken internal link file convention:** **deferred.** The provisional
   marker is `location="#REF!"`, matching formulas and names, and
   `broken-internal-link.xlsx` pins it. An Excel capture is needed before the
   exact bytes Excel writes for a broken internal target can be confirmed; the
   capture cannot be made here. The in-app behaviour required by this ticket
   (target gone reads `#REF!`, edit applies) is implemented and tested.

## Placement

Planned to start after the M6 exit review (maintainer ruled 2026-11-19).
Candidate for M7 or a small follow-up milestone; the research note's design
options are the starting brief.
