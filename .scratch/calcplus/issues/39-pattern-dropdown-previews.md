# 39 — Pattern drop-down shows text only; it should show coloured previews

Status: resolved
Type: visual parity (Format Cells Fill tab vs the Excel 365 oracle)
Blocked by: none

## Question

Reported by the maintainer with a capture (2026-11-20): in **Format Cells →
Fill**, the Pattern drop-down is a plain text list ("Solid", …). In the
Excel 365 oracle the Pattern Style drop-down shows every pattern as a small
**drawn preview** — the pattern texture rendered in the chosen pattern
colour over the chosen background colour — so the user sees what each
pattern looks like before choosing. The closed control also shows the
current pattern's preview.

Today's implementation is `FILL_PATTERNS` in
`src/chrome/FormatCellsDialog.tsx` (a `{value, label}` text list rendered in
a native `<select>`). A native select cannot host previews; this needs the
house custom-drop-down pattern (the sweep's `src/chrome/ColorPicker.tsx` and
the `Menu` component show the shape).

## Ruled behaviour

- Each Pattern option renders as a drawn preview of that pattern, in colour
  (the dialog's current fill colour and pattern background), with the name
  beside or under it as the oracle shows it.
- The previews track the colour controls: changing Fill colour or Pattern
  background re-renders every preview in the new colours.
- The closed control shows the chosen pattern's preview.
- Keyboard and screen-reader behaviour matches the existing custom
  drop-downs (listbox semantics, arrows, Enter, Escape). Assistive-technology
  depth follows `accessibility.md` (V2); do not regress what exists.
- Original drawings only: the pattern textures are drawn by us (small canvas
  or CSS/SVG tiles), no external assets, licence gate green.

## Acceptance

1. The drop-down lists every pattern from `FILL_PATTERNS` as a coloured
   preview.
2. Changing the two colour controls updates the previews live.
3. The closed control previews the selection; choosing a pattern updates the
   cell preview state exactly as before (no command changes).
4. The existing Fill-tab behaviour and tests stay green
   (`tests/m4c-format-dialog.test.ts` and neighbours).

## Resolution (2026-11-20)

Every pattern renders as an 8x8 inline SVG tile (original art, no assets,
licence gate green) drawn in the dialog's fill colour over the pattern
background — washes by opacity, dots and lines at the same tile coordinates
as the grid's `applyPattern`, so the preview agrees with the painted cell.
The native select is replaced by the house custom listbox (combobox roles,
arrows/Home/End, Enter/Space, Escape closes and returns focus); the closed
control previews the choice and keeps the `format-fill-pattern` test id. No
command or model changes.

Follow-up recorded: share one tile table between `src/chrome/fill-patterns.ts`
and the canvas renderer so the two cannot drift. Open question left for the
maintainer: the rows show names beside the previews while Excel's list is
swatch-only; the PRD fidelity rule argues for swatch-only with accessible
names.

Tests: `tests/m4c-fill-pattern-previews.test.ts` plus the extended
`tests/e2e/format-dialog.spec.ts`.
