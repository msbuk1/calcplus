# 10 — UI chrome architecture and ribbon shell prototype

Status: resolved
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

## Answer

Verdict: **Variant A (ribbon) is the primary shell; Variant B (compact) is the responsive mode below roughly 900 px; Variant C (classic) is rejected.**

The architecture is in [ui-chrome.md](../spec/ui-chrome.md). The throwaway prototype lives on branch `prototype/ui-shell` under `prototype/ui-shell/` (screens in `screens/`).

- **Shell:** React 19 + TypeScript + Vite.
- **Primitives:** Ribbon family, CompactBar, NameBox, FormulaBar, SheetTabs, StatusBar, Menu/ContextMenu, Dialog.
- **Command-bound:** every control dispatches a command ID; no chrome business logic.
- **Layout:** header → ribbon/compact → formula row → grid → sheet tabs → status bar; only the grid scrolls vertically.
- **Deferred:** backstage File panel, command search, Alt key tips, ribbon customisation, dark mode, mobile/touch.
