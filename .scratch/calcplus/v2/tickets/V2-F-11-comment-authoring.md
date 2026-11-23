# V2-F-11 — Comment authoring

- **Stream:** Functional
- **Status:** proposed
- **Effort:** 1–2 pw (provisional)
- **Prerequisites:** V2-S-06 (for protected sheets)
- **Blocks:** none
- **Contract:** [feature matrix](../../spec/feature-matrix.md) §8

## Scope

Add, edit and delete legacy notes and threaded comments; reply to a thread;
show and hide the comment pane. v1 renders and preserves both kinds but cannot
author them.

## Acceptance

- [ ] New notes and threaded comments can be added, edited and deleted.
- [ ] Replies attach to the right thread; indicators and popups update.
- [ ] Comments respect sheet protection when V2-S-06 is on.
- [ ] Each change is one undo step.
- [ ] Both comment forms round-trip and open correctly in Excel 365.

## Notes

The VML note shape and threaded-comment parts already exist in the fixture
corpus; the authoring path must reuse them rather than invent a new format.
