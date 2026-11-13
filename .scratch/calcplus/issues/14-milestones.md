# 14 — Milestones and phasing

Status: resolved
Type: grilling
Blocked by: 01, 02, 03, 04, 05, 06, 07, 08, 09, 10, 11, 12, 13, 17, 18, 19, 20, 21, 22, 24, 25, 26

## Question

How is the v1 build sequenced?

Decompose the build into phased milestones with:

- Exit criteria tied to the feature matrix's acceptance criteria
- A demoable state at the end of each milestone
- Dependency ordering across model, engine, grid, chrome, and file I/O
- Risk callouts and spikes
- A definition of "v1 done", using the conformance bar from 13

This becomes the milestone section of the spec package.

## Answer

The plan is in [milestones.md](../spec/milestones.md).

**Shape:** hybrid vertical slices. Milestone 1 is a walking skeleton with the real OOXML layer, canonical model, minimal grid, one edit and undo. Every later milestone ends in a demoable user journey.

**Team and size:** three engineers in parallel (grid/render, model/engine/commands, file I/O/chrome), part-time design and QA. Estimates in person-weeks, no calendar dates: about 62-90 person-weeks plus a 20% contingency, roughly 75-108 person-weeks and 25-36 working weeks with three engineers.

**Milestones:**

1. **Walking skeleton** (4-6 pw) — open, render, edit, undo, save with untouched parts byte-identical.
2. **Grid foundation** (8-12 pw) — 1M-row canvas, panes, editing, keyboard, accessibility mirror. *Spike: AT mirror.*
3. **Formula engine and model depth** (8-12 pw) — 52 functions, shims, dates, names, undo, range-transform prototype. *Spikes: HyperFormula fan-out/memory; OOXML range transform.*
4. **Formatting and workbook** (8-12 pw) — full styles, number formats, merges, multi-sheet, find/Go To, chrome.
5. **Data tools** (8-12 pw) — sort, AutoFilter, conditional formatting, clipboard, fill, insert/delete. *Spike: Excel clipboard.*
6. **File fidelity** (10-14 pw) — full round-trip, preserved parts, structural edits, reports, producer matrix, hostile inputs.
7. **Output and hardening** (10-14 pw) — print/PDF, accessibility, performance, persistence, signed launcher. *Spike: vector print/PDF.*
8. **Conformance and release** (6-8 pw + reserve) — full golden suite, manual Excel repair check, AT sign-off, docs, release gate.

**Gates:** each milestone's exit criteria map to feature-matrix sections; a gate fails if any acceptance criterion fails, even when the demo looks good.

**v1 done:** the ticket 13 release gate passes in full, every feature-matrix IN item is present, launcher artifacts are signed/notarized, the spec package is complete, known deviations are documented and pinned, AT sign-off is recorded, and no P0/P1 or data-loss bug remains.
