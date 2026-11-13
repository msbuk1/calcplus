# calcplus v1 — UI chrome

Resolved by [10 — UI chrome architecture and ribbon shell prototype](../issues/10-ui-chrome.md) on 2026-11-13. The throwaway prototype lives on [branch `prototype/ui-shell`](https://github.com/msbuk1/calcplus/tree/prototype/ui-shell/prototype/ui-shell); command binding is in [command-system](command-system.md).

## 1. Verdict

- **Variant A — Ribbon** is the primary shell: Excel desktop fidelity.
- **Variant B — Compact** is the responsive mode below roughly 900 px, not a separate product.
- **Variant C — Classic** is rejected.

## 2. Layout

Top to bottom:

1. **App header** — document name, dirty indicator, quick access (save, undo, redo). The browser provides window controls.
2. **Ribbon** (primary) or **compact bar** (narrow).
3. **Formula row** — name box, `fx`, formula bar, cancel/enter.
4. **Grid viewport** — the only vertically scrollable region.
5. **Sheet tabs** — tabs, add, scroll.
6. **Status bar** — ready state, selection statistics, view toggles, zoom.

The grid is the flexible region; everything else has a fixed height.

## 3. Primitives

- `AppHeader`
- `Ribbon`, `RibbonTab`, `RibbonGroup`, `RibbonButton`, `RibbonToggle`, `RibbonSelect`, `RibbonSplitButton`
- `CompactBar` (responsive)
- `NameBox`, `FormulaBar`
- `SheetTabs`
- `StatusBar`
- `Menu`, `MenuItem`, `ContextMenu`
- `Dialog`, `DialogTabs`
- `Toolbar` (classic and overflow)

Every control takes a `commandId`. Enablement and checked state come from the command registry. No chrome component contains business logic.

## 4. Command binding

- Buttons, toggles, menus and context items dispatch command IDs.
- Toggle state comes from `isChecked(context)`; enablement from `isEnabled(context)`.
- Dialogs submit a command with a payload.
- No chrome code touches the workbook model directly.

## 5. Responsive behaviour

- **≥ 900 px:** full ribbon.
- **< 900 px:** compact bar, with a control to expand the full ribbon.
- **< 1100 px:** hide ribbon group labels and wide controls; keep short labels and icons.
- The ribbon scrolls horizontally rather than wrapping.
- No mobile/touch layout in v1.

## 6. Theming

- CSS custom properties for all chrome colours; light theme in v1, dark mode in v2.
- Excel-like palette: green accent, neutral chrome, white grid.
- Near-identical chrome, not pixel-perfect.

## 7. Accessibility

- Ribbon: `role="tablist"` / `tab` / `tabpanel`; menus: `role="menu"` / `menuitem`; dialogs: `role="dialog"` with a focus trap.
- Focus zones and Esc behaviour follow [command-system](command-system.md).
- Visible focus at all times; screen-reader grid mirror per tickets 05 and 22.

## 8. Deferred

- Full command inventory — derive from the feature matrix.
- Backstage File panel; v1 uses a simple File menu with New, Open, Save, Save As, Print and Recent.
- “Tell me” command search.
- Alt key tips.
- Ribbon customisation and quick-access toolbar.
- Dark mode and theme authoring.
- Mobile/touch layout.
- Multiple workbooks per tab.

## 9. Acceptance criteria

1. Variant A renders the full ribbon; Variant B takes over below 900 px; Variant C is not shipped.
2. Every control dispatches a command ID; no chrome code mutates the model.
3. Toggle and checked state reflect the command registry.
4. The grid is the only vertically scrollable region; chrome stays fixed.
5. The ribbon scrolls horizontally at narrow widths; group labels hide below 1100 px.
6. Dialogs trap focus and return it to the invoker; Esc returns to the grid.
7. Context menus are keyboard-navigable and dispatch command IDs.
8. Light theme only; every chrome colour uses a CSS custom property.
9. No mobile/touch layout in v1.
10. The prototype branch is the source for the verdict; main carries only this decision.
