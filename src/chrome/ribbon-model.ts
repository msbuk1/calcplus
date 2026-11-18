import { COMMAND_INVENTORY, inventoryCommand } from '@commands/inventory';
import type {
  CommandInventoryEntry,
  CompactPlacement,
  ControlValueSource,
  RibbonMenuItem,
  RibbonPlacement,
  RibbonSelectOption,
  RibbonTabId,
} from '@commands/inventory';
import type { MenuItemModel } from './Menu';

export interface CommandState {
  readonly enabled: boolean;
  readonly checked: boolean;
}

export type CommandStateResolver = (commandId: string) => CommandState;

export interface RibbonTabDefinition {
  readonly id: RibbonTabId;
  readonly label: string;
}

export interface RibbonLauncher {
  readonly commandId: string;
  readonly title: string;
  readonly args?: Readonly<Record<string, unknown>>;
}

export interface ResolvedRibbonControl {
  readonly key: string;
  readonly kind: RibbonPlacement['kind'];
  readonly commandId: string;
  readonly label: string;
  readonly title: string;
  readonly icon?: string;
  readonly wide: boolean;
  /** Present-but-disabled placement (ticket 27); the resolver cannot re-enable it. */
  readonly disabled: boolean;
  /** 1-based row inside the group band. */
  readonly row: number;
  /** Rows the control spans; Paste spans the whole Clipboard band. */
  readonly rowSpan: number;
  readonly controlId?: string;
  readonly valueFrom?: ControlValueSource;
  readonly args?: Readonly<Record<string, unknown>>;
  readonly options?: readonly RibbonSelectOption[];
  readonly menu?: readonly RibbonMenuItem[];
  readonly menuHeading?: string;
}

export interface ResolvedRibbonGroup {
  readonly id: string;
  readonly label: string;
  readonly controls: readonly ResolvedRibbonControl[];
  readonly rowCount: number;
  /** The Excel dialog launcher at the band's bottom-right corner. */
  readonly launcher?: RibbonLauncher;
}

export interface ResolvedCompactControl {
  readonly key: string;
  readonly kind: CompactPlacement['kind'];
  readonly commandId: string;
  readonly label: string;
  readonly title: string;
  readonly args?: Readonly<Record<string, unknown>>;
}

/** The Format Cells tab each group's dialog launcher opens. */
const FORMAT_TAB_BY_GROUP: Readonly<Record<string, string>> = {
  Font: 'font',
  Alignment: 'alignment',
  Number: 'number',
};

export const RIBBON_TABS: readonly RibbonTabDefinition[] = [
  { id: 'Home', label: 'Home' },
  { id: 'Insert', label: 'Insert' },
  { id: 'Page Layout', label: 'Page Layout' },
  { id: 'Formulas', label: 'Formulas' },
  { id: 'Data', label: 'Data' },
  { id: 'Review', label: 'Review' },
  { id: 'View', label: 'View' },
];

/**
 * Excel's familiar group order. A group with no implemented IN command is
 * dropped whole; it is never regrouped under another name.
 *
 * The Home order is the oracle's: Clipboard, Font, Alignment, Number, Styles,
 * Cells, Editing. The Styles group is thin in v1 — one present-but-disabled
 * Conditional Formatting button (ticket 27) — because Format As Table and Cell
 * Styles are out of matrix. No conditional-formatting authoring is built here.
 */
const GROUP_ORDER: Readonly<Record<RibbonTabId, readonly string[]>> = {
  Home: ['Clipboard', 'Font', 'Alignment', 'Number', 'Styles', 'Cells', 'Editing'],
  Insert: ['Tables', 'Illustrations', 'Links', 'Text', 'Sheets'],
  'Page Layout': ['Themes', 'Page Setup', 'Scale to Fit', 'Sheet Options'],
  Formulas: ['Function Library', 'Defined Names', 'Formula Auditing', 'Calculation'],
  Data: ['Sort & Filter', 'Data Tools', 'Outline'],
  Review: ['Proofing', 'Comments', 'Protect'],
  View: ['Show', 'Zoom', 'Window'],
};

/** The default grid: two rows of controls, the group label beneath them. */
const DEFAULT_ROW_COUNT = 2;

export function ribbonGroupsFor(tab: RibbonTabId): readonly ResolvedRibbonGroup[] {
  const groups = new Map<string, ResolvedRibbonControl[]>();
  const launchers = new Map<string, RibbonLauncher>();
  for (const entry of COMMAND_INVENTORY) {
    for (const placement of entry.ribbon ?? []) {
      if (placement.tab !== tab) continue;
      // An unimplemented command is hidden, not greyed (ui-chrome §4.1). The
      // single ruled exception is a present-but-disabled placement (ticket 27).
      if (!entry.implemented && placement.disabled !== true) continue;
      if (placement.menuOnly === true) continue;
      if (placement.kind === 'launcher') {
        if (!launchers.has(placement.group)) {
          launchers.set(placement.group, {
            commandId: entry.id,
            title: placement.title ?? entry.label,
            // Excel's group launcher opens the Format Cells dialog on its own
            // tab (ticket 28 captures); the tab follows the group name.
            args: { tab: FORMAT_TAB_BY_GROUP[placement.group] ?? 'font' },
          });
        }
        continue;
      }
      const controls = groups.get(placement.group) ?? [];
      controls.push({
        key: `${entry.id}:${placement.controlId ?? placement.label ?? placement.kind}:${placement.order ?? 100}`,
        kind: placement.kind,
        commandId: entry.id,
        label: placement.label ?? entry.label,
        title: placement.title ?? placement.label ?? entry.label,
        ...(placement.icon === undefined ? {} : { icon: placement.icon }),
        wide: placement.wide === true,
        disabled: placement.disabled === true,
        row: placement.row ?? 1,
        rowSpan: Math.max(1, placement.rowSpan ?? 1),
        ...(placement.controlId === undefined ? {} : { controlId: placement.controlId }),
        ...(placement.valueFrom === undefined ? {} : { valueFrom: placement.valueFrom }),
        ...(placement.args === undefined ? {} : { args: placement.args }),
        ...(placement.options === undefined ? {} : { options: placement.options }),
        ...(placement.menu === undefined ? {} : { menu: placement.menu }),
        ...(placement.menuHeading === undefined ? {} : { menuHeading: placement.menuHeading }),
      });
      groups.set(placement.group, controls);
    }
  }
  const order = GROUP_ORDER[tab] ?? [];
  return [...groups.entries()]
    .sort((left, right) => {
      const leftRank = order.indexOf(left[0]);
      const rightRank = order.indexOf(right[0]);
      return (
        (leftRank < 0 ? order.length : leftRank) - (rightRank < 0 ? order.length : rightRank) ||
        left[0].localeCompare(right[0])
      );
    })
    .map(([id, controls]) => {
      const sorted = [...controls].sort(
        (left, right) => (controlOrder(left) ?? 100) - (controlOrder(right) ?? 100),
      );
      const launcher = launchers.get(id);
      return {
        id,
        label: id,
        controls: sorted,
        rowCount: Math.max(
          DEFAULT_ROW_COUNT,
          ...sorted.map((control) => control.row + control.rowSpan - 1),
        ),
        ...(launcher === undefined ? {} : { launcher }),
      };
    });
}

/**
 * The authored ribbon menu, minus anything not implemented: a planned entry is
 * hidden rather than greyed (ui-chrome §4.1). Split menus and dropdown-only
 * menus share this so the Cells Insert shell and the Sort & Filter menu behave
 * the same way.
 */
export function ribbonMenuItems(
  menu: readonly RibbonMenuItem[],
  resolveState: CommandStateResolver,
): readonly MenuItemModel[] {
  const items: MenuItemModel[] = [];
  menu.forEach((item, index) => {
    if (inventoryCommand(item.commandId)?.implemented !== true) return;
    items.push({
      id: `${item.commandId}:${index}`,
      label: item.label,
      commandId: item.commandId,
      args: item.args,
      ...(item.shortcut === undefined ? {} : { shortcut: item.shortcut }),
      disabled: !resolveState(item.commandId).enabled,
      ...(item.dividerBefore === true ? { dividerBefore: true } : {}),
    });
  });
  return items;
}

export function compactControls(): readonly ResolvedCompactControl[] {
  return COMMAND_INVENTORY.filter(
    (entry): entry is CommandInventoryEntry & { compact: CompactPlacement } =>
      entry.implemented && entry.compact !== undefined,
  )
    .sort((left, right) => (left.compact.order ?? 100) - (right.compact.order ?? 100))
    .map((entry) => ({
      key: entry.id,
      kind: entry.compact.kind,
      commandId: entry.id,
      label: entry.compact.label ?? entry.label,
      title: entry.compact.title ?? entry.compact.label ?? entry.label,
      ...(entry.compact.args === undefined ? {} : { args: entry.compact.args }),
    }));
}

export function menuItemsFor(
  menu: string,
  resolveState: CommandStateResolver,
): readonly MenuItemModel[] {
  const items: MenuItemModel[] = [];
  for (const entry of COMMAND_INVENTORY) {
    if (!entry.implemented) continue;
    for (const placement of entry.menu ?? []) {
      if (placement.menu !== menu) continue;
      const state = resolveState(entry.id);
      const checked = isToggle(entry) ? state.checked : undefined;
      items.push({
        id: `${entry.id}:${placement.order ?? 100}`,
        label: placement.label ?? entry.label,
        commandId: entry.id,
        shortcut: entry.shortcut[0],
        disabled: !state.enabled,
        ...(checked === undefined ? {} : { checked }),
        ...(placement.dividerBefore === true ? { dividerBefore: true } : {}),
        ...(menuTestId(entry.id) === undefined ? {} : { testId: menuTestId(entry.id) }),
      });
    }
  }
  return items.sort((left, right) => menuRank(left.id) - menuRank(right.id));
}

export function contextMenuItems(resolveState: CommandStateResolver): readonly MenuItemModel[] {
  const items: MenuItemModel[] = [];
  for (const entry of COMMAND_INVENTORY) {
    if (!entry.implemented || entry.contextMenu === undefined) continue;
    const state = resolveState(entry.id);
    items.push({
      id: entry.id,
      label: entry.contextMenu.label ?? entry.label,
      commandId: entry.id,
      disabled: !state.enabled,
    });
  }
  return items.sort((left, right) => contextRank(left.id) - contextRank(right.id));
}

function controlOrder(control: ResolvedRibbonControl): number | null {
  // The key carries the authored order as its final segment.
  const suffix = control.key.slice(control.key.lastIndexOf(':') + 1);
  const value = Number(suffix);
  return Number.isFinite(value) ? value : null;
}

function menuRank(id: string): number {
  const value = Number(id.slice(id.lastIndexOf(':') + 1));
  return Number.isFinite(value) ? value : 100;
}

function contextRank(id: string): number {
  const entry = COMMAND_INVENTORY.find((candidate) => candidate.id === id);
  return entry?.contextMenu?.order ?? 100;
}

function menuTestId(commandId: string): string | undefined {
  if (commandId === 'file.new') return 'new-button';
  if (commandId === 'file.open') return 'open-button';
  if (commandId === 'file.saveAs') return 'save-as-button';
  return undefined;
}

function isToggle(entry: CommandInventoryEntry): boolean {
  return (
    entry.compact?.kind === 'toggle' ||
    (entry.ribbon ?? []).some((placement) => placement.kind === 'toggle')
  );
}
