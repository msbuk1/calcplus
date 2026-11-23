# V2-F-15 — Theme and named-style authoring

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §3

## Scope

Author workbook themes (colour scheme and font scheme) and named cell styles:
create, edit, apply and delete. Imported themes and styles already resolve and
preserve; this ticket adds the editing surface.

## Acceptance

- [ ] Theme colours and fonts can be edited and apply across the workbook.
- [ ] Named styles can be created, modified, applied and deleted.
- [ ] Style changes update existing cells that use the style.
- [ ] Theme and style XML round-trips and opens correctly in Excel.
- [ ] No Microsoft theme file or asset is copied into the app.

## Notes

The v1 style table already handles imported styles; this ticket adds the
authoring commands and dialogs. Dark mode (V2-V-04) is UI theming and stays
separate.
