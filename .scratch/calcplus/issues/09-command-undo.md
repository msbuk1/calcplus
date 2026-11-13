# 09 — Command system, keyboard parity, and undo/redo

Status: open
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
