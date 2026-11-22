import rawInventory from './command-inventory.json';

/**
 * The checked-in chrome map. `scripts/generate-command-inventory.mjs` reads the
 * same JSON and the feature matrix, writes `docs/command-inventory.md`, and
 * fails when a matrix IN row has no command or a command names no IN row.
 *
 * A command appears once. `implemented` says whether the runtime registry has
 * it in this build; planned entries keep their target placement in the map so
 * later slices fill them in without moving the ribbon around. Every entry with
 * `implemented: false` also carries `unimplementedReason`, the reconciliation
 * classification the generator checks and renders (ticket 36 item 3).
 */
export type RibbonTabId =
  'Home' | 'Insert' | 'Page Layout' | 'Formulas' | 'Data' | 'Review' | 'View';

/**
 * Why an entry has no command. `not-a-command` means the capability is engine,
 * file or view behaviour and no command will ever carry the id (for example
 * `fidelity.*` and `name.resolve`). `v2-cut` means the command is a deliberate
 * v2 cut (`formula.recalculate`, `format.conditionalRender`).
 */
export type UnimplementedReason = 'not-a-command' | 'v2-cut';

export type RibbonControlKind = 'button' | 'toggle' | 'select' | 'split' | 'menu' | 'launcher';

export type ControlValueSource = 'fontName' | 'fontSize' | 'numberFormat';

export interface RibbonSelectOption {
  readonly label: string;
  /** The value key used by the host to report the current style. */
  readonly value?: string;
  /** The command arguments sent when the option is chosen. */
  readonly args?: Readonly<Record<string, unknown>>;
}

export interface RibbonMenuItem {
  readonly label: string;
  readonly commandId: string;
  readonly args?: Readonly<Record<string, unknown>>;
  /** The trailing shortcut text the capture shows, when it shows one. */
  readonly shortcut?: string;
  /** A name from `src/chrome/ribbon-icons.ts`; drawn ahead of the label. */
  readonly icon?: string;
  readonly dividerBefore?: boolean;
}

export interface RibbonPlacement {
  readonly tab: RibbonTabId;
  readonly group: string;
  readonly kind: RibbonControlKind;
  readonly order?: number;
  /** 1-based row inside the group band; the body packs up to two rows. */
  readonly row?: number;
  /** How many rows the control spans; a large control such as Paste spans three. */
  readonly rowSpan?: number;
  readonly label?: string;
  readonly title?: string;
  /**
   * The label shown while the command reports a checked state, as Excel's
   * Freeze Panes entry flips to Unfreeze Panes (ticket 43). Only a command
   * with an `isChecked` predicate may author it, and the generator rejects a
   * label that matches the default.
   */
  readonly activeLabel?: string;
  /** A name from `src/chrome/ribbon-icons.ts`; every rendered control has one. */
  readonly icon?: string;
  /** The icon alone carries the control's visible identity; the label stays in `title`. */
  readonly iconOnly?: boolean;
  /** Hidden below 1100 px; short labels and icons stay. */
  readonly wide?: boolean;
  /**
   * Present-but-disabled: the ticket-27 carve-out for a render-only IN row whose
   * authoring arrives in v2. No other planned command may use it.
   */
  readonly disabled?: boolean;
  readonly controlId?: string;
  readonly valueFrom?: ControlValueSource;
  readonly args?: Readonly<Record<string, unknown>>;
  readonly options?: readonly RibbonSelectOption[];
  readonly menu?: readonly RibbonMenuItem[];
  /**
   * A category menu the ribbon builds from the Insert Function catalogue at
   * render time. The map names the category, never a copy of its functions, so
   * the ribbon and the dialog cannot disagree. Must be a Select category other
   * than "All"; the generator checks it.
   */
  readonly functionCategory?: string;
  /** A bold section heading above the menu's items, as Excel shows it. */
  readonly menuHeading?: string;
  /**
   * A colour picker in place of a preset menu: `font` edits the text colour,
   * `fill` edits the fill colour. The picker itself is chrome-owned.
   */
  readonly picker?: 'font' | 'fill';
  /** Listed in the group's split/menu control but not rendered as its own control. */
  readonly menuOnly?: boolean;
}

/**
 * The Quick Access Toolbar above the ribbon. Save, Print, Undo and Redo are the
 * placements; Save As stays in the File menu and in the Save button's
 * drop-down, as the Excel 365 desktop oracle shows it. Print joined the
 * toolbar beside Save in ticket 42.
 */
export interface QuickAccessPlacement {
  readonly order?: number;
  /** A name from `src/chrome/ribbon-icons.ts`; defaults to the command icon. */
  readonly icon?: string;
  /**
   * A drop-down hanging off the control, opened by its split arrow. Only
   * `file.saveAs` may ride here: Save carries Save As without giving it a
   * quick-access button of its own.
   */
  readonly menu?: readonly RibbonMenuItem[];
}

export interface MenuPlacement {
  readonly menu: string;
  readonly label?: string;
  readonly order?: number;
  readonly dividerBefore?: boolean;
}

export interface CompactPlacement {
  readonly kind: 'button' | 'toggle';
  readonly order?: number;
  readonly label?: string;
  readonly title?: string;
  /** A name from `src/chrome/ribbon-icons.ts`; defaults to the command icon. */
  readonly icon?: string;
  /** The icon alone carries the compact control's visible identity. */
  readonly iconOnly?: boolean;
  readonly args?: Readonly<Record<string, unknown>>;
}

/** A one-level flyout of the grid context menu, as Excel's Insert/Delete show. */
export interface ContextMenuGroup {
  /** Stable id shared by every item inside the flyout. */
  readonly id: string;
  /** The flyout entry's label, verbatim from the oracle capture. */
  readonly label: string;
  /** The flyout entry's position among the flat context-menu items. */
  readonly order: number;
}

/** The grid region a placement belongs to; header regions show flat row/column items. */
export type ContextMenuRegion = 'cell' | 'rowHeader' | 'columnHeader';

export interface ContextMenuPlacement {
  /**
   * Which grid region shows this placement. A grouped placement that omits it
   * belongs to the cell menu's flyout; a flat placement that omits it is shared
   * chrome, shown on every region's menu.
   */
  readonly context?: ContextMenuRegion;
  /** Sort order: among the flat items, or inside the flyout when grouped. */
  readonly order?: number;
  readonly label?: string;
  /** The trailing shortcut text the capture shows, when it shows one. */
  readonly shortcut?: string;
  readonly args?: Readonly<Record<string, unknown>>;
  /** Present when the item sits in a one-level flyout. */
  readonly group?: ContextMenuGroup;
}

/**
 * One command may surface more than once: Insert Cells & Shift Right and Shift
 * Down are the same command with different arguments, as the ribbon menu
 * already authors them.
 */
export type AuthoredContextMenu = ContextMenuPlacement | readonly ContextMenuPlacement[];

export interface CommandInventoryEntry {
  readonly id: string;
  readonly matrixSection: string;
  readonly matrixFeature: string;
  readonly label: string;
  readonly category: string;
  readonly enablement: string;
  readonly implemented: boolean;
  /** Required on every unimplemented entry; the generator checks and renders it. */
  readonly unimplementedReason?: UnimplementedReason;
  readonly shortcut: readonly string[];
  /** The command's canonical icon; the generated docs carry this reference. */
  readonly icon?: string;
  readonly ribbon?: readonly RibbonPlacement[];
  readonly menu?: readonly MenuPlacement[];
  readonly contextMenu?: AuthoredContextMenu;
  readonly compact?: CompactPlacement;
  readonly quickAccess?: QuickAccessPlacement;
}

export const COMMAND_INVENTORY = rawInventory as readonly CommandInventoryEntry[];

const byId = new Map<string, CommandInventoryEntry>(
  COMMAND_INVENTORY.map((entry) => [entry.id, entry]),
);

export function inventoryCommand(id: string): CommandInventoryEntry | undefined {
  return byId.get(id);
}

/**
 * Normalise the map's authored context-menu shape: a single placement or a
 * list of them. Consumers always see a list.
 */
export function contextMenuPlacements(
  entry: CommandInventoryEntry,
): readonly ContextMenuPlacement[] {
  const authored = entry.contextMenu;
  if (authored === undefined) return [];
  return Array.isArray(authored)
    ? (authored as readonly ContextMenuPlacement[])
    : [authored as ContextMenuPlacement];
}

export function implementedInventoryCommands(): readonly CommandInventoryEntry[] {
  return COMMAND_INVENTORY.filter((entry) => entry.implemented);
}
