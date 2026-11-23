# calcplus v2 roadmap

Status: planning draft, 2026-11-23. This directory is the v2 plan. It is kept
separate from the v1 spec package and the published docs. Nothing here is
committed build scope until the maintainer freezes a v2 release scope.

## Why this exists

calcplus 0.1.0 shipped the v1 scope. The [feature matrix](../spec/feature-matrix.md)
§9 defers a set of work to v2. This roadmap turns that deferred list into
tickets with prerequisites, blockers and acceptance criteria.

It does not replace the contracts. Where a ticket and a contract disagree, the
contract wins and the ticket is corrected.

## How to read it

- One ticket per file in [`tickets/`](tickets/).
- ID prefixes: **V2-S** structural, **V2-F** functional, **V2-V** visual.
- **Status** is `proposed`, `ready`, `in progress`, `done` or `cut`.
- **Prerequisites** must be done before the ticket starts. **Blocks** names
  the tickets that wait on it. Soft dependencies are named in the ticket but
  do not stop the start.
- Each ticket carries its own acceptance criteria. A ticket is done when its
  criteria pass, the local checks are green and the change is reviewed.

## Streams

- **Structural** — changes the model, engine, file layer or architecture.
  These gate functional work, so design them first.
- **Functional** — new commands and behaviour on the existing model.
- **Visual** — new rendering or appearance work. It is the most separable
  stream and can run alongside the other two.

## Ticket index

### Structural

| ID | Ticket | Effort | Prerequisites | Blocks |
|---|---|---|---|---|
| [V2-S-01](tickets/V2-S-01-tables-structured-references.md) | Tables and structured references | 4–6 pw | none | none |
| [V2-S-02](tickets/V2-S-02-dynamic-arrays-spill.md) | Dynamic arrays and spill | 4–7 pw | none | none (soft: V2-F-01, V2-F-17) |
| [V2-S-03](tickets/V2-S-03-reference-model.md) | Reference model extensions | 3–5 pw | none | none |
| [V2-S-04](tickets/V2-S-04-external-link-refresh.md) | External-link refresh | 2–3 pw | none | none |
| [V2-S-05](tickets/V2-S-05-iterative-manual-calculation.md) | Iterative and manual calculation | 2–3 pw | none | none |
| [V2-S-06](tickets/V2-S-06-protection-enforcement.md) | Protection enforcement | 2–3 pw | none | V2-F-02, V2-F-11, V2-F-12, V2-F-14 |
| [V2-S-07](tickets/V2-S-07-accessibility-mirror.md) | Accessibility mirror and WCAG claim | 4–7 pw | none | none |
| [V2-S-08](tickets/V2-S-08-localisation-foundation.md) | Localisation foundation | 4–6 pw | none | none (soft: V2-V-04) |

### Functional

| ID | Ticket | Effort | Prerequisites | Blocks |
|---|---|---|---|---|
| [V2-F-01](tickets/V2-F-01-function-tranche.md) | Tier 1 function tranche (37 functions) | 5–9 pw | none | none (soft: V2-F-16) |
| [V2-F-02](tickets/V2-F-02-data-validation.md) | Data validation authoring and enforcement | 3–5 pw | V2-S-06 | none |
| [V2-F-03](tickets/V2-F-03-advanced-filters.md) | Advanced, colour and Top 10 filters | 2–3 pw | none | none |
| [V2-F-04](tickets/V2-F-04-paste-special.md) | Paste Special | 1–2 pw | none | none |
| [V2-F-05](tickets/V2-F-05-fill-series.md) | Fill series | 2–3 pw | none | none |
| [V2-F-06](tickets/V2-F-06-entry-autocomplete.md) | Entry AutoComplete | 0.5–1 pw | none | none |
| [V2-F-07](tickets/V2-F-07-discontiguous-selection.md) | Discontiguous selection | 1–2 pw | none | none |
| [V2-F-08](tickets/V2-F-08-hide-unhide.md) | Hide/unhide rows and columns | 1 pw | none | V2-F-10, V2-F-19 |
| [V2-F-09](tickets/V2-F-09-sheet-management.md) | Sheet management and tab colours | 3–4 pw | none | none |
| [V2-F-10](tickets/V2-F-10-grouping-outline.md) | Grouping and outline | 1–2 pw | V2-F-08 | none |
| [V2-F-11](tickets/V2-F-11-comment-authoring.md) | Comment authoring | 1–2 pw | V2-S-06 (for protected sheets) | none |
| [V2-F-12](tickets/V2-F-12-hyperlink-authoring.md) | Hyperlink authoring | 1 pw | V2-S-06 (for protected sheets) | none |
| [V2-F-13](tickets/V2-F-13-name-manager.md) | Name Manager | 1–2 pw | none | none |
| [V2-F-14](tickets/V2-F-14-conditional-formatting-authoring.md) | Conditional-formatting authoring | 3–4 pw | V2-V-03 (for unsupported rule types) | none |
| [V2-F-15](tickets/V2-F-15-theme-style-authoring.md) | Theme and named-style authoring | 2–3 pw | none | none |
| [V2-F-16](tickets/V2-F-16-function-argument-wizard.md) | Insert Function argument wizard | 1–2 pw | none (soft: V2-F-01) | none |
| [V2-F-17](tickets/V2-F-17-formula-auditing.md) | Formula auditing | 2–4 pw | none (soft: V2-S-02) | none |
| [V2-F-18](tickets/V2-F-18-split-view.md) | Split view | 1–2 pw | none | V2-F-19 |
| [V2-F-19](tickets/V2-F-19-custom-views.md) | Custom views | 1–2 pw | V2-F-08, V2-F-18 | none |
| [V2-F-20](tickets/V2-F-20-page-break-preview.md) | Page-break preview | 1 pw | V2-F-21 | none |
| [V2-F-21](tickets/V2-F-21-full-page-setup.md) | Full page setup | 3–5 pw | none | V2-F-20 |

### Visual

| ID | Ticket | Effort | Prerequisites | Blocks |
|---|---|---|---|---|
| [V2-V-01](tickets/V2-V-01-chart-rendering.md) | Chart rendering | 6–12 pw | none | none |
| [V2-V-02](tickets/V2-V-02-images-shapes-sparklines.md) | Images, shapes and sparklines rendering | 4–7 pw | none | none |
| [V2-V-03](tickets/V2-V-03-conditional-format-rules.md) | Remaining conditional-format rule rendering | 3–5 pw | none | V2-F-14 (for those rule types) |
| [V2-V-04](tickets/V2-V-04-dark-mode.md) | Dark mode | 2–3 pw | none | none (soft: V2-S-08) |

## Dependency chains

Read `→` as "before".

```text
V2-S-01, V2-S-02, V2-S-03   independent structural builds
V2-S-06  →  V2-F-02, V2-F-11, V2-F-12, V2-F-14
V2-F-08  →  V2-F-10
V2-F-08  +  V2-F-18  →  V2-F-19
V2-F-21  →  V2-F-20
V2-F-01  →  V2-F-16            (soft: the wizard lists the full set)
V2-V-03  →  V2-F-14            (only for the remaining rule types)
V2-S-02  →  V2-F-01 array form, V2-F-17 spill auditing   (soft)
V2-S-08  +  V2-V-04            share tokens               (soft)
```

The soft edges do not stop a ticket from starting; they stop the dependent
part from meeting its acceptance.

## Phases

**Phase 0 — decisions.** Work through the structural tickets as design notes
first: tables, dynamic arrays, reference model, protection, accessibility,
localisation. No code until the model and engine shape is agreed. Freeze the
v2 release scope here; the register can be cut by phase without breaking the
dependency graph.

**Phase 1 — structural foundations.** Build the tickets that other work rests
on: V2-S-03, V2-S-01 and V2-S-02, then V2-S-06, V2-S-07, V2-S-08. V2-S-04 and
V2-S-05 can run in parallel once the engine work lands.

**Phase 2 — functional.** Start with V2-F-01 (the function tranche), then the
authoring and data tools. Respect the blockers: V2-F-08 before V2-F-10 and
V2-F-19; V2-F-18 before V2-F-19; V2-F-21 before V2-F-20; V2-S-06 before the
enforcement half of V2-F-02.

**Phase 3 — visual, in parallel.** V2-V-01 and V2-V-02 are independent of the
functional stream. V2-V-03 must land before V2-F-14 can author unsupported
rule types. V2-V-04 is independent.

**Phase 4 — conformance and release.** A v2 release gate, modelled on the v1
[release gate](../spec/testing-conformance.md) §13: every ticket's acceptance,
the regression gate, no P0/P1 or data-loss bug, and rebuilt artifacts with
checksums.

## Sizing

The register sums to roughly **75–125 person-weeks**, provisional. That is
wider than the earlier coarse estimate of 52–90 because every element now
carries its own ticket and acceptance work. The numbers are planning ranges,
not commitments.

## Definition of v2 done

- Every ticket in the frozen v2 scope is `done` and its acceptance passes.
- The v1 release gate still passes: no regression.
- No P0, P1 or data-loss bug remains.
- If V2-S-07 is in scope, the accessibility claim is backed by the manual
  assistive-technology matrix.
- The docs, known deviations and release notes are updated.
- Windows and Linux artifacts are rebuilt and published with checksums.

## Out of roadmap

- Real-time co-authoring, accounts, cloud sync and sharing.
- macOS artifacts.
- Store listings, auto-update and marketing.
- PivotTables and Power Query as features, Solver/Goal Seek, Data Analysis
  ToolPak, Consolidate, macro execution, native desktop and mobile apps,
  `.xls`/`.xlsb`/ODS.
- Chart, image and shape authoring. The visual stream renders preserved
  objects; authoring is a later decision and is not ticketed here.

## Sources

- [Feature matrix](../spec/feature-matrix.md) §9 and §10
- [Formula semantics](../spec/formula-semantics.md) §12 (function tranche)
- [Known deviations](../../../docs/known-deviations.md) (D-3, D-5, L-4)
- [Accessibility](../spec/accessibility.md)
- The 2026-11-23 v2 sizing discussion, recorded by this roadmap
