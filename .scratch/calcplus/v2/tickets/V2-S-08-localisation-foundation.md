# V2-S-08 — Localisation foundation

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 4–6 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none (soft: V2-V-04 shared tokens)
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §9
  (localisation); [formula semantics](../../spec/formula-semantics.md) §3

## Scope

Extract UI strings into one catalogue, make number and date parsing and
formatting locale-aware, and support localised function names and list
separators. English ships first; a second locale proves the mechanism. Saved
files keep English function names and unchanged part XML.

## Acceptance

- [ ] UI strings come from the catalogue; no user-facing string is hard-coded.
- [ ] A second locale loads without code changes and keeps the layout intact.
- [ ] Number and date parsing and formatting follow the selected locale.
- [ ] Localised function names resolve; exported formulas stay in English.
- [ ] Existing files round-trip; Excel 365 opens the output without repair.

## Notes

Translation content is a follow-on task; this ticket delivers the mechanism,
the string extraction and one second locale as proof. Coordinate the token
layer with V2-V-04.
