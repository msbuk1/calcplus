# 15 — Assemble the build-ready spec package

Status: resolved
Type: task
Blocked by: 14, 23, 25, 26

## Question

Assemble the final artifact from the resolved map.

Produce:

- **PRD**: goals, target user, v1 scope, non-goals, acceptance criteria
- **Technical design**: workbook model, architecture, component boundaries, ADR index
- **Milestones**: phased plan with exit criteria

Every decision recorded on the map must be reflected, and no two sections may contradict. Write it into `.scratch/calcplus/spec/` (or a single `spec.md` plus ADRs, whichever the team prefers).

## Answer

The build-ready package is assembled in [`.scratch/calcplus/spec/`](../spec/README.md):

- **[PRD](../spec/prd.md)** — product summary, target user, goals, fidelity promise, v1 scope, key journeys, non-goals, constraints, acceptance criteria, open questions, risks and success measures.
- **[Technical design](../spec/technical-design.md)** — architecture, eleven component boundaries, data flow for open/edit/save/undo, thirteen invariants, threading, performance, accessibility, security, failure handling and the ADR index.
- **[Milestones](../spec/milestones.md)** — eight phases with demos, exit criteria, estimates, dependency graph, risk register and the definition of v1 done.
- **[Package index](../spec/README.md)** — reading order, the contract table and the status of every open item.
- **Three missing contracts were written from their ticket answers:** [clipboard](../spec/clipboard.md), [number formats](../spec/number-formats.md) and [print/PDF](../spec/print-pdf.md), so every map decision now has a spec file.
- **ADRs:** 0001-0005 plus an [index](../../docs/adr/README.md), covering the canonical model, the custom grid, the preservation-first OOXML layer, HyperFormula and GPLv3.
- **Root `README.md`** now points at the package; `LICENSE`, `NOTICE`, `THIRD_PARTY_NOTICES.md`, `CONTRIBUTING.md` and `CONTRIBUTORS.md` are in place.

**Consistency pass:** every relative link in the spec, ADR and root documents resolves; the feature matrix remains the scope authority; the PRD, technical design, milestones, performance, accessibility, testing and licence contracts use the same numbers and the same v1/v2 cuts. The three deliberately open items are listed in the PRD and the package index: localization, formula-bar/name-box edge behaviour, and fill-handle/merged-selection/IME/selection-scrolling edge cases.

**No contradictions remain between:** 52 functions and HyperFormula; custom grid and the Univer rejection; preservation-first OOXML and the three file tiers; sort/filter/CF semantics and their fixtures; performance envelopes and milestone gates; WCAG 2.2 AA and the DOM mirror; Excel 365 oracle and the release gate; GPL-3.0-only and the dependency audit.
