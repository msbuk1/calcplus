# V2-F-19 — Custom views

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** V2-F-08, V2-F-18
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §6

## Scope

Create, apply and delete named custom views. A view stores hidden rows and
columns, print settings, filter state and split state, as Excel's custom views
do. Show the Custom Views dialog and switch between saved views.

## Acceptance

- [ ] A view captures and restores the hidden, print, filter and split state.
- [ ] Deleting a view leaves the current state unchanged.
- [ ] Views round-trip and open correctly in Excel.
- [ ] Applying a view is one undo step.
- [ ] The dialog lists the views and their scope.

## Notes

Custom views depend on the hidden state (V2-F-08) and split state (V2-F-18).
Print settings come from V2-F-21, so coordinate the storage shape.
