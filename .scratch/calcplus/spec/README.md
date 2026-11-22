# calcplus v1 — spec package

This is the build-ready specification for **calcplus**, a browser-based, single-user clone of Excel desktop (Microsoft 365). It is assembled from the resolved decisions on the [map](../map.md) and is internally consistent: where a contract and a summary disagree, the contract wins.

## Reading order

1. **[PRD](prd.md)** — goals, target user, v1 scope, non-goals and product acceptance criteria.
2. **[Feature matrix](feature-matrix.md)** — the authoritative in/v2/never scope.
3. **[Technical design](technical-design.md)** — architecture, components, data flow, invariants and ADR index.
4. **[Milestones](milestones.md)** — the eight-phase build plan, exit criteria and definition of done.
5. The contracts below, by area.

## Contracts

| Area | Document |
|---|---|
| Scope | [Feature matrix](feature-matrix.md) |
| Workbook model | [Workbook model](workbook-model.md) |
| Commands, undo and keyboard | [Command system](command-system.md) |
| Formulas | [Formula semantics](formula-semantics.md) |
| Number formats | [Number formats](number-formats.md) |
| Grid and renderer | [Grid spike](grid-spike.md) |
| UI chrome | [UI chrome](ui-chrome.md) |
| `.xlsx` compatibility | [File compatibility](file-compat.md) |
| Sort and AutoFilter | [Sort and AutoFilter](sort-autofilter.md) |
| Conditional formatting | [Conditional formatting](conditional-formatting.md) |
| Clipboard | [Clipboard](clipboard.md) |
| Print and PDF | [Print and PDF](print-pdf.md) |
| Persistence | [Persistence](persistence.md) |
| Distribution | [Distribution](distribution.md) |
| Performance | [Performance](performance.md) |
| Accessibility (v2) | [Accessibility](accessibility.md) |
| Testing and conformance | [Testing and conformance](testing-conformance.md) |
| Licence | [Licensing](licensing.md) |

## Decisions and evidence

- [Map](../map.md) — the decision narrative.
- [Issue tracker](../issues/) — every ticket and its answer.
- [Research](../research/) — the evidence behind the engine, file, grid, clipboard, number-format, print, prior-art and launcher decisions.
- [ADRs](../../../docs/adr/) — the irreversible architecture choices.
- [Glossary](../../../CONTEXT.md) — the project vocabulary.

## Status

All decision and task tickets are resolved. The remaining open items are deliberately outside v1:

- Localization of the UI and function names.
- Formula-bar and name-box edge behaviours.
- Fill-handle options, merged-range selection and IME/bidi text. (Selection
  scrolling left this list on 2026-11-22: the range-autoscroll work in the
  chrome and data-tools sweep landed as a version one bug fix, ticket 36
  item 6.)

## How to use this package

- **Build:** start at the [milestones](milestones.md); each phase links to the contracts it implements.
- **Review:** the [PRD](prd.md) and [feature matrix](feature-matrix.md) define what is promised; the contracts define how it is proved.
- **Change:** a scope change goes through the feature matrix; an architecture change goes through a new ADR; a contract change updates the map and the ticket.
- **Release:** the [release gate](testing-conformance.md) is the definition of v1 done.

## Licence

GPL-3.0-only. See [LICENSE](../../../LICENSE), [NOTICE](../../../NOTICE) and [THIRD_PARTY_NOTICES.md](../../../THIRD_PARTY_NOTICES.md).
