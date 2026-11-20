# calcplus v1 — UI chrome

Resolved by [10 — UI chrome architecture and ribbon shell prototype](../issues/10-ui-chrome.md) on 2026-11-13. The throwaway prototype lives on [branch `prototype/ui-shell`](https://github.com/msbuk1/calcplus/tree/prototype/ui-shell/prototype/ui-shell); command binding is in [command-system](command-system.md).

## 1. Verdict

- **Variant A — Ribbon** is the primary shell: Excel desktop fidelity.
- **Variant B — Compact** is the responsive mode below roughly 900 px, not a separate product.
- **Variant C — Classic** is rejected.
- **Quick Access Toolbar.** It sits above the ribbon, anchored top left, and holds Save, Undo and Redo only. Each is a small 16 px monochrome icon button with a tooltip naming the command and its shortcut. Save As is not in it; it stays in the File menu and on the Save button's split arrow. Unsaved changes show on the Save button, not as a separate dot beside the filename. The document name stays in the app header.

## 2. Layout

Top to bottom:

1. **App header** — the Quick Access Toolbar at the top left, above the ribbon: Save, Undo and Redo as icon buttons with tooltips. The document name sits in the same strip. Unsaved changes show on the Save button. Save As stays in the File menu and on the Save button's split arrow.
2. **Ribbon** (primary) or **compact bar** (narrow).
3. **Formula row** — name box, `fx`, formula bar, cancel/enter.
4. **Grid viewport** — the only vertically scrollable region.
5. **Sheet tabs** — tabs, add, scroll.
6. **Status bar** — ready state, selection statistics, view toggles, zoom.

The grid is the flexible region; everything else has a fixed height.

## 3. Primitives

- `AppHeader`
- `QuickAccessToolbar`
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

### 4.1 Command inventory and ribbon map (v1 deliverable)

- Every feature-matrix **IN** row maps to one or more command IDs in the [command registry](command-system.md) — one per user-visible action, because a row such as "insert/delete rows, columns and cells" names several commands. The mapping is bidirectional: every command traces back to exactly one matrix row, and the generator fails if either direction is unmapped. The authoritative chrome map is a checked-in `command-inventory.md` generated from the matrix: command ID, matrix row, label, ribbon tab, ribbon group, menu placement, shortcut and enablement predicate. The generator lives in the repo; the output is reviewed like code.
- Ribbon tabs follow Excel 365 desktop's default tab set (Home, Insert, Page Layout, Formulas, Data, Review, View). A group renders only if it owns at least one IN command. Where calcplus is thinner than Excel, the fewer commands sit in the familiar group — never a regrouped ribbon.
- Commands outside the matrix are **hidden** in v1, not greyed: a greyed control implies the command exists.
- **Exception (ticket 27, ruled 2026-11-18):** a matrix row that is IN as render-only with V2 authoring ships its ribbon control present-but-disabled, with a tooltip naming the v2 limitation. The Conditional Formatting button (Home → Styles) is the v1 instance. Everything else outside the matrix stays hidden.
- Matrix rows qualified **IN (imported)**, **IN (preserve)** or **IN (reject)** ship as enabled commands whose limitation is stated in the tooltip and encoded in `isEnabled`. The matrix is the only scope vocabulary; there is no separate "partial" state.
- **AutoSum** (`Alt+=`, `formula.autosum`) is **IN** — it is a command surface over the `SUM` function contract — and appears in the inventory (Home → Editing).
- **v1 dialog inventory:** Format Cells, Insert Function, Sort, AutoFilter conditions, Go To, Find & Replace, Row Height / Column Width, Save As, Compatibility Report and About. Dialogs that belong to v2 features — Paste Special, Name Manager, the conditional-formatting rule manager, the Insert Function argument wizard and Go To Special — are not shipped in v1. Each dialog's behaviour is specified by its owning contract; this contract owns only its layout, focus trap, Esc and default-button semantics.

## 5. Responsive behaviour

- **≥ 900 px:** full ribbon.
- **< 900 px:** compact bar, with a control to expand the full ribbon.
- **< 1100 px:** hide ribbon group labels and wide controls; keep short labels and icons.
- The ribbon scrolls horizontally rather than wrapping.
- No mobile/touch layout in v1.

## 6. Theming

- CSS custom properties for all chrome colours; light theme in v1, dark mode in v2.
- Excel-like palette: green accent, neutral chrome, white grid. v1 tokens (confirm hexes against Excel 365 desktop's light theme before the task-10 freeze):

| Token | Value | Use |
|---|---|---|
| `--accent` | `#107C41` | Excel green: active tab underline, primary buttons |
| `--accent-hover` | `#0E6B37` | hover/pressed accent |
| `--chrome-surface` | `#F3F2F1` | ribbon, title strip |
| `--command-surface` | `#FFFFFF` | menus, dialogs, gallery popouts |
| `--grid-bg` | `#FFFFFF` | grid canvas |
| `--border` | `#D1D1D1` | chrome borders |
| `--text` | `#262626` | chrome text |
| `--disabled` | `#A19F9D` | disabled controls |
- Near-identical chrome, not pixel-perfect.

## 7. Keyboard and focus

- Every command is reachable by keyboard; focus zones and Esc behaviour follow [command-system](command-system.md).
- Visible focus at all times; dialogs trap focus and Esc closes them.
- Assistive-technology support (ARIA roles, screen-reader mirror) is v2; see [accessibility](accessibility.md).

## 8. Deferred

- (The full command inventory is a v1 deliverable — see §4.1; it is no longer deferred.)
- Backstage File panel; v1 uses a simple File menu with New, Open, Save, Save As, Print and Recent.
- “Tell me” command search.
- Alt key tips.
- Ribbon customisation and a user-editable quick-access toolbar. The fixed
  Save, Undo and Redo toolbar is v1.
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
