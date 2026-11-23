# V2-S-06 — Protection enforcement

- **Stream:** Structural
- **Status:** proposed
- **Effort:** 2–3 pw (provisional)
- **Prerequisites:** none
- **Blocks:** V2-F-02, V2-F-11, V2-F-12, V2-F-14
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §5, §8; known
  deviation D-2

## Scope

Turn preserved protection into enforced protection. Add the Protect Sheet and
Protect Workbook dialogs and the allow-edit ranges list. Enforce locked and
unlocked cells, structure protection (insert, delete, rename, move) and the
allowed actions, with Excel's messages. v1 preserves the XML but lets every
edit through.

## Acceptance

- [ ] Locked cells refuse edits on a protected sheet; unlocked cells stay
      editable.
- [ ] Allow-edit ranges permit the listed actions.
- [ ] Structure protection blocks sheet insert, delete, rename and move.
- [ ] Unprotecting with the password restores editing; a wrong password is
      refused.
- [ ] Protection XML round-trips; Excel 365 opens the file without repair.
- [ ] The known deviation D-2 is removed and its pin updated.

## Notes

Worksheet protection is an editing guard, not a security boundary; match
Excel's behaviour and say so in the UI. Authoring dialogs are small; the work
is the command gate and its tests.
