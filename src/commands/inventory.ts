import rawInventory from './command-inventory.json';

/**
 * The checked-in chrome map. `scripts/generate-command-inventory.mjs` reads the
 * same JSON and the feature matrix, writes `docs/command-inventory.md`, and
 * fails when a matrix IN row has no command or a command names no IN row.
 *
 * A command appears once. `implemented` says whether the runtime registry has
 * it in this build; planned entries keep their target placement in the map so
 * later slices fill them in without moving the ribbon around.
 */
export type RibbonTabId =
  'Home' | 'Insert' | 'Page Layout' | 'Formulas' | 'Data' | 'Review' | 'View';

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
  readonly icon?: string;
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
  /** A bold section heading above the menu's items, as Excel shows it. */
  readonly menuHeading?: string;
  /** Listed in the group's split/menu control but not rendered as its own control. */
  readonly menuOnly?: boolean;
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
  readonly args?: Readonly<Record<string, unknown>>;
}

export interface ContextMenuPlacement {
  readonly order?: number;
  readonly label?: string;
}

export interface CommandInventoryEntry {
  readonly id: string;
  readonly matrixSection: string;
  readonly matrixFeature: string;
  readonly label: string;
  readonly category: string;
  readonly enablement: string;
  readonly implemented: boolean;
  readonly shortcut: readonly string[];
  readonly ribbon?: readonly RibbonPlacement[];
  readonly menu?: readonly MenuPlacement[];
  readonly contextMenu?: ContextMenuPlacement;
  readonly compact?: CompactPlacement;
}

export const COMMAND_INVENTORY = rawInventory as readonly CommandInventoryEntry[];

const byId = new Map<string, CommandInventoryEntry>(
  COMMAND_INVENTORY.map((entry) => [entry.id, entry]),
);

export function inventoryCommand(id: string): CommandInventoryEntry | undefined {
  return byId.get(id);
}

export function implementedInventoryCommands(): readonly CommandInventoryEntry[] {
  return COMMAND_INVENTORY.filter((entry) => entry.implemented);
}
