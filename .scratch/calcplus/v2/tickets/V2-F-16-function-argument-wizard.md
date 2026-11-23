# V2-F-16 — Insert Function argument wizard

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** none (soft: V2-F-01)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §4

## Scope

Turn the v1 Insert Function dialog into a step-by-step wizard: one field per
argument, the signature and description, validation per argument, a live
preview and an insert action. Cover the full published set, including the
V2-F-01 tranche.

## Acceptance

- [ ] Every published function opens a wizard with its real arguments.
- [ ] Each argument field validates and explains its error.
- [ ] The preview updates as arguments are entered.
- [ ] Insert writes the formula with the caret on the next argument.
- [ ] Keyboard navigation and Esc behave as the v1 dialog does.

## Notes

The wizard reads the same registry and catalogue as the formula bar, so it
must not duplicate function descriptions.
