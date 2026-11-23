# V2-V-04 — Dark mode

- **Stream:** Visual
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none (soft: V2-S-08 shared tokens)
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §9 (dark mode)

## Scope

Add a dark UI theme built on shared design tokens: chrome, ribbon, dialogs,
grid chrome and the canvas palette. Workbook colours and number formats do not
change, and saved files are identical in either theme. Honour Windows forced
colours.

## Acceptance

- [ ] Every surface has light and dark tokens; no hard-coded colour remains.
- [ ] Text and icon contrast meets the accessibility target.
- [ ] Forced-colours mode stays usable; the canvas palette adapts.
- [ ] Switching themes changes no document data or saved file bytes.
- [ ] Visual baselines cover both themes.

## Notes

This is UI theming, not workbook theme authoring (V2-F-15). The token layer
should be shared with V2-S-08 where the two overlap.
