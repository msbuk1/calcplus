# 30 — Insert Function dialog in v1, without the argument wizard

Status: resolved
Type: scope change
Blocked by: none

## Question

The maintainer's Excel 365 shows the modern **Insert function** dialog (search
box, Select category list, function list, signature line such as
`AVERAGE(number1, [number2], ...)`, description with Learn more, and the
argument wizard behind **Next** / **Insert with preview**). The feature matrix
fenced the whole dialog to V2 and ui-chrome §4.1 listed it among the dialogs
that do not ship in v1. Should any part of it come into v1?

## Answer

Ruled 2026-11-18 by the maintainer: **the dialog comes into v1 without the
argument wizard.**

- **In v1:** the Insert Function dialog with search box, category list,
  function list, and the signature plus description area — matching the
  captured layout. Choosing a function and confirming **inserts the function
  skeleton** into the formula bar with the cursor on the first argument. The
  `fx` button opens it.
- **Stays V2:** the argument wizard (the Next / Insert with preview
  step-through), and everything else in the old fence (Paste Special, Name
  Manager, the CF rule manager, Go To Special).

Contract changes made under this ticket: `feature-matrix.md` (the Insert
Function row is IN with the wizard carved out; the not-in list now says
"Insert Function argument wizard"), `ui-chrome.md` §4.1 (Insert Function joins
the v1 dialog inventory). The function names, categories, signatures and
descriptions come from the 52-function registry (M3B); the registry entry and
the command map land with the implementation lane, after the structural lane
releases `command-inventory.json` and `registry.ts`.
