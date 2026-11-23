# V2-S-04 — External-link refresh

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §7

## Scope

Let the user refresh external-workbook links on demand. v1 preserves links and
shows their cached values but never refreshes them. v2 adds an explicit refresh
action: the user chooses the linked file through the file picker, calcplus reads
its cached values and updates the dependent cells. No silent network access and
no automatic fetch.

## Acceptance

- [ ] Refresh is an explicit action that names the links it will update.
- [ ] The user picks the linked file; missing files leave the cached value and
      report the link as stale.
- [ ] Refreshed values recalculate dependants and are written as cached values.
- [ ] Declining the refresh changes nothing.
- [ ] Link parts round-trip; Excel 365 opens the file without a repair prompt.

## Notes

This ticket does not add an external-link editor. Editing links stays out of
scope. Privacy rules hold: the workbook and the linked file stay on the
machine.
