import { COMMAND_INVENTORY, contextMenuPlacements, inventoryCommand } from '@commands/inventory';
import type {
  CommandInventoryEntry,
  CompactPlacement,
  ContextMenuGroup,
  ContextMenuPlacement,
  ContextMenuRegion,
  ControlValueSource,
  RibbonMenuItem,
  RibbonPlacement,
  RibbonSelectOption,
  RibbonTabId,
} from '@commands/inventory';
import { FUNCTION_CATEGORY_OPTIONS, functionsInCategory } from '@engine/function-catalog';
import type { FunctionCategoryOption } from '@engine/function-catalog';
import type { PublishedFunction } from '@engine/function-registry';
import type { MenuItemModel } from './Menu';
import { isRibbonIconName } from './ribbon-icons';
import type { RibbonIconName } from './ribbon-icons';

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
  readonly icon?: RibbonIconName;
  /** The icon alone is the control's visible content; `title` names it. */
  readonly iconOnly: boolean;
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
  /**
   * Authored visual line breaks for a tile label. The drop-down caret rides
   * the last line, so the break is fixed rather than left to the browser.
   */
  readonly labelLines?: readonly string[];
  /** The catalogue category this control's generated menu lists. */
  readonly functionCategory?: FunctionCategoryOption;
  readonly menuHeading?: string;
  /** A colour picker replaces the preset menu for the font and fill colours. */
  readonly picker?: 'font' | 'fill';
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
  readonly icon?: RibbonIconName;
  /** The icon alone is the compact control's visible content. */
  readonly iconOnly: boolean;
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

export interface ResolvedQuickAccessControl {
  readonly key: string;
  readonly commandId: string;
  readonly label: string;
  /** The tooltip and accessible name: the command and its first shortcut. */
  readonly title: string;
  readonly icon?: RibbonIconName;
  /** The split drop-down's authored rows; Save carries Save As here. */
  readonly menu?: readonly RibbonMenuItem[];
  readonly order: number;
}

/**
 * The Quick Access Toolbar above the ribbon: Save, Print, Undo and Redo, in map
 * order (ui-chrome §1, ticket 42). Save As is not a quick-access command; it
 * stays in the File menu. Enablement still comes from the registry at render
 * time.
 */
export function quickAccessControls(): readonly ResolvedQuickAccessControl[] {
  return COMMAND_INVENTORY.filter((entry) => entry.implemented && entry.quickAccess !== undefined)
    .sort((left, right) => (left.quickAccess?.order ?? 100) - (right.quickAccess?.order ?? 100))
    .map((entry) => {
      const iconName = entry.quickAccess?.icon ?? entry.icon;
      const icon = iconName !== undefined && isRibbonIconName(iconName) ? iconName : undefined;
      const shortcut = entry.shortcut[0];
      return {
        key: entry.id,
        commandId: entry.id,
        label: entry.label,
        title: shortcut === undefined ? entry.label : `${entry.label} (${shortcut})`,
        ...(icon === undefined ? {} : { icon }),
        ...(entry.quickAccess?.menu === undefined ? {} : { menu: entry.quickAccess.menu }),
        order: entry.quickAccess?.order ?? 100,
      };
    });
}

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

/**
 * The Function Library tiles break their labels where Excel 365 breaks them.
 * The drop-down caret joins the last line, so pinning the break keeps the
 * two-line shape at every band width and stops the caret wrapping alone.
 */
const FUNCTION_LIBRARY_LABEL_LINES: Readonly<Record<string, readonly string[]>> = {
  'Insert Function': ['Insert', 'Function'],
  AutoSum: ['AutoSum'],
  'Recently Used': ['Recently', 'Used'],
  Financial: ['Financial'],
  Logical: ['Logical'],
  Text: ['Text'],
  'Date & Time': ['Date &', 'Time'],
  'Lookup & Reference': ['Lookup &', 'Reference'],
  'Math & Trig': ['Math &', 'Trig'],
  'More Functions': ['More', 'Functions'],
};

/**
 * The Select-category values a generated ribbon menu may name. "All" belongs
 * to the dialog's searchable list, not to a Function Library menu.
 */
const FUNCTION_MENU_CATEGORIES: ReadonlySet<string> = new Set(
  FUNCTION_CATEGORY_OPTIONS.filter((option) => option !== 'All'),
);

function functionMenuCategory(value: string | undefined): FunctionCategoryOption | undefined {
  if (value === undefined || !FUNCTION_MENU_CATEGORIES.has(value)) return undefined;
  return value as FunctionCategoryOption;
}

export function ribbonGroupsFor(
  tab: RibbonTabId,
  resolveState?: CommandStateResolver,
  recentFunctions: readonly PublishedFunction[] = [],
): readonly ResolvedRibbonGroup[] {
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
      const placementIcon =
        placement.icon !== undefined && isRibbonIconName(placement.icon)
          ? placement.icon
          : undefined;
      // A checked-state label (ticket 43): Freeze Panes reads Unfreeze Panes
      // while the active sheet has a split. The authored default stays when the
      // command reports no checked state.
      const activeLabel =
        placement.activeLabel !== undefined && resolveState?.(entry.id).checked === true
          ? placement.activeLabel
          : undefined;
      // A Function Library category menu is generated from the same catalogue
      // the Insert Function dialog reads, so the two lists cannot drift.
      const functionCategory = functionMenuCategory(placement.functionCategory);
      const categoryMenu =
        functionCategory === undefined
          ? undefined
          : functionsInCategory(functionCategory, recentFunctions).map((functionEntry) => ({
              label: functionEntry.name,
              commandId: 'formula.insertFunction',
              args: { name: functionEntry.name },
            }));
      const menu = categoryMenu ?? placement.menu;
      const label = activeLabel ?? placement.label ?? entry.label;
      const labelLines =
        placement.group === 'Function Library' ? FUNCTION_LIBRARY_LABEL_LINES[label] : undefined;
      controls.push({
        key: `${entry.id}:${placement.controlId ?? placement.label ?? placement.kind}:${placement.order ?? 100}`,
        kind: placement.kind,
        commandId: entry.id,
        label,
        title: activeLabel ?? placement.title ?? placement.label ?? entry.label,
        ...(placementIcon === undefined ? {} : { icon: placementIcon }),
        iconOnly: placement.iconOnly === true,
        wide: placement.wide === true,
        disabled: placement.disabled === true,
        row: placement.row ?? 1,
        rowSpan: Math.max(1, placement.rowSpan ?? 1),
        ...(placement.controlId === undefined ? {} : { controlId: placement.controlId }),
        ...(placement.valueFrom === undefined ? {} : { valueFrom: placement.valueFrom }),
        ...(placement.args === undefined ? {} : { args: placement.args }),
        ...(placement.options === undefined ? {} : { options: placement.options }),
        ...(menu === undefined ? {} : { menu }),
        ...(labelLines === undefined ? {} : { labelLines }),
        ...(functionCategory === undefined ? {} : { functionCategory }),
        ...(placement.menuHeading === undefined ? {} : { menuHeading: placement.menuHeading }),
        ...(placement.picker === undefined ? {} : { picker: placement.picker }),
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
    const icon = item.icon !== undefined && isRibbonIconName(item.icon) ? item.icon : undefined;
    items.push({
      id: `${item.commandId}:${index}`,
      label: item.label,
      commandId: item.commandId,
      args: item.args,
      ...(item.shortcut === undefined ? {} : { shortcut: item.shortcut }),
      ...(icon === undefined ? {} : { icon }),
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
    .map((entry) => {
      const iconName = entry.compact.icon ?? entry.icon;
      const icon = iconName !== undefined && isRibbonIconName(iconName) ? iconName : undefined;
      return {
        key: entry.id,
        kind: entry.compact.kind,
        commandId: entry.id,
        label: entry.compact.label ?? entry.label,
        title: entry.compact.title ?? entry.compact.label ?? entry.label,
        ...(icon === undefined ? {} : { icon }),
        iconOnly: entry.compact.iconOnly === true,
        ...(entry.compact.args === undefined ? {} : { args: entry.compact.args }),
      };
    });
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

/**
 * The grid context menu: flat items plus the one-level Insert/Delete flyouts,
 * in map order (ui-chrome §4.1). A flyout is disabled only when every item in
 * it is disabled, so the interaction stays possible whenever a direction is.
 *
 * The region picks the placements: a cell right-click gets the grouped Insert
 * and Delete flyouts; a row or column header right-click gets the flat
 * Insert/Delete pair for that header, as Excel shows. Flat items without a
 * region are shared chrome and appear on every region's menu.
 */
export function contextMenuItems(
  resolveState: CommandStateResolver,
  region: ContextMenuRegion = 'cell',
): readonly MenuItemModel[] {
  const flat: { readonly order: number; readonly item: MenuItemModel }[] = [];
  const groups = new Map<
    string,
    {
      readonly group: ContextMenuGroup;
      readonly children: { readonly order: number; readonly item: MenuItemModel }[];
    }
  >();

  for (const entry of COMMAND_INVENTORY) {
    if (!entry.implemented) continue;
    for (const placement of contextMenuPlacements(entry)) {
      if (!placementAppliesTo(placement, region)) continue;
      const state = resolveState(entry.id);
      const group = placement.group;
      const item: MenuItemModel = {
        id: `${entry.id}:${group?.id ?? 'flat'}:${placement.order ?? 100}`,
        label: placement.label ?? entry.label,
        commandId: entry.id,
        ...(placement.args === undefined ? {} : { args: placement.args }),
        ...(placement.shortcut === undefined ? {} : { shortcut: placement.shortcut }),
        disabled: !state.enabled,
      };
      if (group === undefined) {
        flat.push({ order: placement.order ?? 100, item });
        continue;
      }
      const bucket = groups.get(group.id) ?? { group, children: [] };
      bucket.children.push({ order: placement.order ?? 100, item });
      groups.set(group.id, bucket);
    }
  }

  const entries = [...flat];
  for (const [id, bucket] of groups) {
    const children = [...bucket.children]
      .sort((left, right) => left.order - right.order)
      .map((child) => child.item);
    entries.push({
      order: bucket.group.order,
      item: {
        id: `context:${id}`,
        label: bucket.group.label,
        disabled: children.every((child) => child.disabled === true),
        items: children,
      },
    });
  }
  return entries.sort((left, right) => left.order - right.order).map((entry) => entry.item);
}

/**
 * A flat item with no region is shared chrome: it sits on every region's
 * menu. A flyout with no region is the cell menu's Insert/Delete shape, and a
 * placement that names a region belongs to that region alone.
 */
function placementAppliesTo(placement: ContextMenuPlacement, region: ContextMenuRegion): boolean {
  if (placement.context !== undefined) return placement.context === region;
  return region === 'cell' || placement.group === undefined;
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
