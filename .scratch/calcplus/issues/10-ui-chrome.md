# 10 — UI chrome architecture and ribbon shell prototype

Status: claimed
Type: prototype
Blocked by: none

## Question

What does calcplus's shell look like, and how is its chrome built?

Build a cheap, clickable prototype of the app shell at functional, near-identical fidelity:

- Ribbon/toolbar with tabs and groups
- Name box and formula bar
- Grid viewport with headers, selection, and sheet tabs
- Status bar
- Context menu and dialog patterns

Then decide the UI architecture: reusable primitives, how ribbon controls dispatch commands, responsive behavior when the viewport narrows, and what chrome is deferred. Use a provisional command set pending the feature matrix (01); link the prototype artifact from this ticket.

## Notes

- All chrome controls dispatch command IDs from [command-system](spec/command-system.md); do not call the model directly.
