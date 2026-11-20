# 33 — Ribbon icons missing; the ribbon is text-only

Status: resolved
Type: defect (visual completeness gap against the UI chrome contract)
Blocked by: none

## Question

The ribbon renders every control as text. There are no icons. The whole
surface reads as unfinished and does not look like the Excel 365 oracle the
[PRD](../spec/prd.md) names as the fidelity target ("When Excel and 'clean
design' disagree, match Excel").

Current-state evidence (2026-11-19, main at `a58d1ba`):

- The icon slot exists and renders. `Ribbon.tsx` lines 308, 468 and 490 emit a
  `<span className="ribbon-icon">` when `control.icon` is set.
- Nothing sets it. The whole `src/commands/command-inventory.json` has five
  icon entries: `"←"`, `"→"` and `"Σ"` used three times. Every other control
  is label only.
- There is no icon set anywhere in the repo. No SVG assets, no icon module. The
  comment in `Ribbon.tsx` around line 334 says "the icon set stays in one
  place", and that place is empty.

Three contracts already assume the icons exist:

1. [licensing](../spec/licensing.md) §"The icon sets are original SVG artwork
   drawn for calcplus. No Microsoft icons, logos, fonts". The licence path is
   decided. The artwork was never drawn.
2. [accessibility](../spec/accessibility.md) 1.1.1 requires every icon to have
   an accessible name, and 1.4.11 requires icon contrast of at least 3:1
   against the background. Both presume icons.
3. [ui-chrome](../spec/ui-chrome.md) lists `RibbonButton`, `RibbonToggle`,
   `RibbonSelect` and `RibbonSplitButton` as primitives. The narrow-width rule
   ("below 1100 px, hide ribbon group labels and wide controls; keep short
   labels and icons") only works if icons exist.

The milestone 4 exit criterion "The ribbon and compact bar meet the UI chrome
contract" passed on structure alone. Appearance was never covered. This ticket
records that gap rather than leaving it to look accepted.

## Ruled direction

Match the Excel 365 desktop ribbon visually: one consistent icon family across
every tab, the same icon size and baseline, colour reserved for state rather
than decoration. Original artwork only, no Microsoft assets. Every icon needs
an accessible name and must hold the 3:1 contrast rule.

## Acceptance shape

1. An icon set exists as original SVG artwork in the repo, covering every
   command in `command-inventory.json` that has a ribbon placement.
2. Every ribbon control with an icon renders it: `RibbonButton`,
   `RibbonToggle`, `RibbonSelect` and `RibbonSplitButton`, on the full ribbon
   and the compact bar.
3. Icons hold 3:1 contrast against `--chrome-surface` and the control states
   (default, hover, pressed, disabled, checked).
4. Every icon carries an accessible name; `aria-hidden` on the decorative span
   is kept where the adjacent label already names the command.
5. The narrow-width rule works: below 1100 px the icons stay visible while
   group labels and wide controls hide.
6. `command-inventory.json` and the generated `docs/command-inventory.md`
   carry the icon reference for each command, and the generator is the source
   of truth, matching ui-chrome §4.1.
7. No Microsoft icon, logo or font is used. The licensing audit stays clean.

## Placement

Owns its own lane inside milestone 7, where the accessibility and print work
already sits. It does not touch the model, engine or file layers. The Format
Cells border-preset defect is a separate ticket and is not in scope here.

## Answer

Pending.

## Resolution (2026-11-20)

Most of this ticket had already landed in the chrome sweep (`4051f8b`): the
original icon family in `src/chrome/ribbon-icons.ts`, the `RibbonIcon`
component, icon entries for every ribbon placement in
`src/commands/command-inventory.json`, the generated-doc icon column, and all
four control kinds drawing the icon (button, toggle, select, split) on the full
ribbon and the compact bar. The sweep also added most of
`tests/ribbon-icons.test.ts`.

The remainder was verification, not artwork: no further glyphs were needed.
Two test gaps were closed.

- `tests/ribbon-icons.test.ts` now resolves an icon for every Quick Access
  control too, and reads `docs/command-inventory.md` back to check the icon
  column still names each command, placement, split-menu item, compact and
  quick-access icon, and never names artwork outside the family.
- `tests/ribbon-model.test.ts` covers the Save split drop-down (ticket 34).

Acceptance: all seven items hold. The licence gate is green; every glyph is
hand-written SVG path data with no Microsoft assets, fonts or colours.
