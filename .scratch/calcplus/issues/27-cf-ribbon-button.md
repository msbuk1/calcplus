# 27 — Conditional Formatting button: present-but-disabled or hidden

Status: resolved
Type: contract conflict
Blocked by: none

## Question

Two contracts disagree about the Conditional Formatting ribbon button:

- [conditional-formatting](../spec/conditional-formatting.md) §8: "There is no CF
  authoring UI in v1. **The ribbon button is present but disabled**, with a
  tooltip that authoring arrives in v2."
- [ui-chrome](../spec/ui-chrome.md) §4.1: "Commands outside the matrix are
  **hidden** in v1, not greyed: a greyed control implies the command exists."

Conditional-formatting **rendering** is matrix IN; conditional-formatting
**authoring** (rule manager, custom rules, stop-if-true UI) is matrix V2. The
generated [command inventory](../../../docs/command-inventory.md) follows
ui-chrome: `format.conditionalRender` has no ribbon placement and the summary
states "OUT rows are absent from the chrome: no control is rendered and nothing
is greyed out to imply that the command exists."

The Excel 365 oracle shows a live **Conditional Formatting** button in a
**Styles** group (with Format As Table and Cell Styles, both out of v1 scope).
The 2026-11-18 ribbon-parity review against the oracle raised the gap.

Decide:

- **A. Follow conditional-formatting §8.** Render a Styles group with a disabled
  Conditional Formatting button and the v2 tooltip. Matches the oracle's
  silhouette; contradicts the ui-chrome hidden-not-greyed rule for this one
  control, which would need a stated carve-out.
- **B. Follow ui-chrome §4.1.** The button stays hidden; amend
  conditional-formatting §8 to drop the present-but-disabled sentence. The
  ribbon stays free of controls that do nothing.
- **C. Generalise the carve-out.** Matrix rows that are IN as render-only with
  V2 authoring may ship a disabled control whose tooltip names the v2
  limitation; all other out-of-matrix commands stay hidden. Amend ui-chrome
  §4.1 accordingly and record the deviation in the inventory.

Output: one amended contract sentence in the losing document, a matching
command-inventory entry, and the ribbon behaviour in the M5 chrome pass.

## Answer

Ruled 2026-11-18 by the maintainer: **A — present but disabled.** The Styles
group ships with a disabled Conditional Formatting button whose tooltip states
that authoring arrives in v2, exactly as conditional-formatting §8 says.
`ui-chrome` §4.1 gained the carve-out for render-only IN rows with V2
authoring; all other out-of-matrix commands remain hidden. The button has no
menu in v1 — the captured menu (design-refs §8) is entirely V2 authoring and is
stored as the v2 reference.
