# V2-F-02 — Data validation authoring and enforcement

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 3–5 pw (provisional)
- **Prerequisites:** V2-S-06 (enforcement)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §2; known
  deviation D-1

## Scope

Author data validation: whole number, decimal, list, date, time, text length
and custom formula rules; input and error messages; stop, warning and
information styles; the in-cell dropdown. Enforce the rules on entry, paste
and fill, and respect sheet protection.

## Acceptance

- [ ] Each rule type can be created, edited and deleted from the dialog.
- [ ] A list rule shows a dropdown; an invalid entry follows its style.
- [ ] Enforcement works on typing, paste and fill; protected sheets behave as
      Excel does.
- [ ] Rules round-trip and survive structural edits; Excel opens without repair.
- [ ] The known deviation D-1 is removed and its pin updated.

## Notes

Authoring can start before V2-S-06, but the ticket cannot finish until
protection enforcement lands, because enforcement has to respect it.
