# 09 — Command system, keyboard parity, and undo/redo

Status: resolved
Type: grilling
Blocked by: 07

## Question

How do user actions flow through calcplus, and what is undoable?

Decide:

- A single command registry consumed by ribbon, menus, keyboard, and context menus
- Command shape: execute/undo, enablement, checked state, labels
- Undo/redo model: inverse commands, snapshots, or an operation log
- Transaction boundaries (cell edit, paste, fill, format range, etc.)
- Memory bounds for history
- Keyboard parity: which Excel shortcuts v1 must honor, and the conflict policy
- Focus model across grid, formula bar, ribbon, and dialogs
- How the command log stays compatible with a future co-authoring operation log

Depends on the workbook data model (07); do not lock a model that contradicts it.

## Notes

- The model is in [workbook-model.md](../spec/workbook-model.md). Its command shape is the starting point for this ticket.

## Answer

The command system is in [command-system.md](../spec/command-system.md).

- **One registry:** every surface dispatches stable command IDs; commands return range-based change sets; no surface mutates the model directly.
- **Undo/redo:** one command log per document with inverse changes; one transaction per user action; typing coalesces; 200 transactions or 64 MB, whichever first.
- **Keyboard:** an Excel-compatible map for navigation, editing, formatting, data, file, find, sheets and formula. Browser-reserved shortcuts are documented with visible alternatives; full parity needs a desktop wrapper.
- **Focus:** explicit zones (grid, formula bar, name box, ribbon, dialog, panel); Esc returns to the grid; dialogs trap focus.
- **Co-authoring seam:** entries carry `seq`, `actorId`, `commandId`, `args`, `changes`, `inverseChanges`; view state and recalculated results are excluded; undo is an inverse transaction.
