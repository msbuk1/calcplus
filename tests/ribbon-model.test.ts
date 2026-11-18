import { describe, expect, it } from 'vitest';
import { COMMAND_INVENTORY, inventoryCommand } from '@commands/inventory';
import { RIBBON_TABS, menuItemsFor, ribbonGroupsFor, ribbonMenuItems } from '@chrome/ribbon-model';
import type { ResolvedRibbonControl, ResolvedRibbonGroup } from '@chrome/ribbon-model';

/**
 * The ribbon is a view of the checked-in command map (ui-chrome §4.1). These
 * tests read the map through the model, so a label or placement change has to
 * be made in `command-inventory.json` first — the map is authoritative.
 */

const alwaysEnabled = (): { enabled: boolean; checked: boolean } => ({
  enabled: true,
  checked: false,
});

function groups(tab: string): readonly ResolvedRibbonGroup[] {
  return ribbonGroupsFor(tab as Parameters<typeof ribbonGroupsFor>[0]);
}

function group(tab: string, id: string): ResolvedRibbonGroup {
  const found = groups(tab).find((candidate) => candidate.id === id);
  if (found === undefined) throw new Error(`No ${id} group on ${tab}`);
  return found;
}

function control(groupId: ResolvedRibbonGroup, commandId: string): ResolvedRibbonControl {
  const found = groupId.controls.filter((candidate) => candidate.commandId === commandId);
  if (found.length === 0) throw new Error(`No ${commandId} control in ${groupId.id}`);
  return found[0]!;
}

function labels(groupId: ResolvedRibbonGroup): string[] {
  return groupId.controls.map((entry) => entry.label);
}

describe('ribbon tabs and Home group order', () => {
  it('follows Excel 365s default tab set, with File outside the tab list', () => {
    expect(RIBBON_TABS.map((tab) => tab.label)).toEqual([
      'Home',
      'Insert',
      'Page Layout',
      'Formulas',
      'Data',
      'Review',
      'View',
    ]);
  });

  it('packs Home in the oracle order, Styles between Number and Cells', () => {
    expect(groups('Home').map((entry) => entry.id)).toEqual([
      'Clipboard',
      'Font',
      'Alignment',
      'Number',
      'Styles',
      'Cells',
      'Editing',
    ]);
  });

  it('gives every group a two-row band, extended where a control spans rows', () => {
    for (const entry of groups('Home')) {
      expect(entry.rowCount).toBeGreaterThanOrEqual(2);
    }
    expect(group('Home', 'Clipboard').rowCount).toBe(3);
    expect(group('Home', 'Font').rowCount).toBe(2);
  });
});

describe('Clipboard group', () => {
  it('places a large Paste beside the stacked Cut, Copy and Format Painter', () => {
    const clipboard = group('Home', 'Clipboard');
    const paste = control(clipboard, 'edit.paste');
    expect(paste.label).toBe('Paste');
    expect(paste.row).toBe(1);
    expect(paste.rowSpan).toBe(3);

    expect(control(clipboard, 'edit.cut')).toMatchObject({ label: 'Cut', row: 1, rowSpan: 1 });
    expect(control(clipboard, 'edit.copy')).toMatchObject({ label: 'Copy', row: 2, rowSpan: 1 });
    // The oracle label is "Format Painter", not the old recorded "Painter".
    expect(control(clipboard, 'painter.pickup')).toMatchObject({
      label: 'Format Painter',
      row: 3,
    });
  });
});

describe('Font group', () => {
  it('puts family, size and quick font toggles in two rows', () => {
    const font = group('Home', 'Font');
    expect(
      font.controls.filter((entry) => entry.row === 1).map((entry) => entry.commandId),
    ).toEqual(['format.font', 'format.font']);
    // No prefix label on the two selects: the current value is the placeholder.
    for (const entry of font.controls.filter((candidate) => candidate.kind === 'select')) {
      expect(entry.label).toBe('');
      expect(entry.title).toBeTruthy();
    }
    expect(font.controls[0]).toMatchObject({ controlId: 'font-family', valueFrom: 'fontName' });
    expect(font.controls[1]).toMatchObject({ controlId: 'font-size', valueFrom: 'fontSize' });

    const rowTwo = font.controls.filter((entry) => entry.row === 2).map((entry) => entry.label);
    // Bold, Italic, Underline, Strikethrough, Borders, Fill Colour, Font Colour.
    expect(rowTwo).toEqual(['B', 'I', 'U', 'S', 'Borders', 'Fill', 'A']);
    expect(control(font, 'format.fill').title).toBe('Fill Colour');
    // The font colour control is the split "A", not one of the selects.
    expect(font.controls.find((entry) => entry.title === 'Font Colour')?.kind).toBe('split');
  });
});

describe('Alignment group', () => {
  it('renders the six implemented quick aligns plus Wrap Text, with no Merge control', () => {
    const alignment = group('Home', 'Alignment');
    expect(alignment.controls.map((entry) => entry.title)).toEqual([
      'Align Left',
      'Centre',
      'Align Right',
      'Align Top',
      'Align Middle',
      'Align Bottom',
      'Wrap Text',
    ]);
    // The merge family is unimplemented, so Merge & Centre is absent, not greyed.
    expect(alignment.controls.some((entry) => entry.commandId.startsWith('merge.'))).toBe(false);
    expect(control(alignment, 'format.wrapText').kind).toBe('toggle');
    expect(control(alignment, 'format.wrapText').row).toBe(1);
  });
});

describe('Number group', () => {
  it('has the format select in row one and the quick formats plus decimals in row two', () => {
    const number = group('Home', 'Number');
    const format = control(number, 'format.numberFormat');
    expect(format).toMatchObject({ kind: 'select', row: 1, label: '', title: 'Number Format' });

    const quick = number.controls
      .filter((entry) => entry.kind !== 'select')
      .map((entry) => ({
        label: entry.label,
        title: entry.title,
        icon: entry.icon,
        row: entry.row,
      }));
    expect(quick).toEqual([
      { label: '$', title: 'Currency', icon: undefined, row: 2 },
      { label: '%', title: 'Percent', icon: undefined, row: 2 },
      { label: ',', title: 'Comma', icon: undefined, row: 2 },
      { label: '.00', title: 'Increase Decimal', icon: '←', row: 2 },
      { label: '.00', title: 'Decrease Decimal', icon: '→', row: 2 },
    ]);
  });
});

describe('dialog launchers', () => {
  it('puts a Format Cells launcher on Font, Alignment and Number', () => {
    for (const id of ['Font', 'Alignment', 'Number']) {
      const launcher = group('Home', id).launcher;
      expect(launcher).toBeDefined();
      expect(launcher?.commandId).toBe('format.cellsDialog');
      expect(launcher?.title).toBe(`${id} Settings`);
    }
  });
});

describe('Styles group (ticket 27 carve-out)', () => {
  it('ships exactly one present-but-disabled Conditional Formatting button', () => {
    const styles = group('Home', 'Styles');
    expect(styles.controls).toHaveLength(1);
    const button = styles.controls[0]!;
    expect(button).toMatchObject({
      commandId: 'format.conditionalRender',
      label: 'Conditional Formatting',
      kind: 'button',
      disabled: true,
    });
    expect(button.title).toContain('authoring arrives in v2');
    // No menu in v1: the whole CF menu is v2 authoring.
    expect(button.menu).toBeUndefined();
  });

  it('is the only planned command that renders a control', () => {
    for (const tab of RIBBON_TABS) {
      for (const entry of groups(tab.id)) {
        for (const rendered of entry.controls) {
          if (inventoryCommand(rendered.commandId)?.implemented === true) continue;
          expect(rendered.commandId).toBe('format.conditionalRender');
        }
      }
    }
  });
});

describe('Cells and Editing groups', () => {
  it('offers the Insert menu shell, rendering only the implemented entries', () => {
    const cells = group('Home', 'Cells');
    expect(cells.controls.map((entry) => entry.label)).toEqual(['Insert']);
    expect(control(cells, 'sheet.insert').kind).toBe('menu');
    // The map reserves the capture's Insert menu; the structural commands are
    // still planned, so only Insert Sheet reaches the rendered menu.
    const menu = control(cells, 'sheet.insert').menu ?? [];
    expect(menu.map((item) => item.label)).toEqual([
      'Insert Rows Above',
      'Insert Columns to the Left',
      'Insert Cells & Shift Right',
      'Insert Cells & Shift Down',
      'Insert Sheet',
    ]);
    expect(ribbonMenuItems(menu, alwaysEnabled).map((item) => item.label)).toEqual([
      'Insert Sheet',
    ]);
  });

  it('places AutoSum, Clear, Sort & Filter and Find & Select in one row', () => {
    const editing = group('Home', 'Editing');
    expect(labels(editing)).toEqual(['AutoSum', 'Clear', 'Sort & Filter', 'Find & Select']);
    expect(editing.controls.every((entry) => entry.row === 1)).toBe(true);
  });

  it('labels the Sort & Filter menu verbatim from the oracle capture', () => {
    const sort = control(group('Home', 'Editing'), 'data.sort');
    expect(sort.kind).toBe('menu');
    expect(sort.menuHeading).toBe('Sort & Filter');
    expect(sort.menu?.map((item) => item.label)).toEqual([
      'Sort Ascending',
      'Sort Descending',
      'Custom Sort',
      'Filter',
      'Clear',
      'Reapply',
    ]);
    expect(sort.menu?.[3]).toMatchObject({ commandId: 'data.filter', dividerBefore: true });
    expect(sort.menu?.[0]?.args).toEqual({
      keys: [{ direction: 'ascending' }],
      hasHeaders: null,
    });
    expect(sort.menu?.[1]?.args).toEqual({
      keys: [{ direction: 'descending' }],
      hasHeaders: null,
    });
    // "Custom Sort" opens the dialog by asking the command for keys.
    expect(sort.menu?.[2]?.args).toBeUndefined();
  });

  it('labels the Find & Select menu with the capture casing', () => {
    const find = control(group('Home', 'Editing'), 'find.search');
    expect(find.label).toBe('Find & Select');
    expect(find.menuHeading).toBe('Find & Select');
    expect(find.menu?.map((item) => item.label)).toEqual(['Find', 'Replace', 'Go to']);
    expect(find.menu?.[2]?.commandId).toBe('goto.goto');
  });

  it('hides Clear All inside the Clear menu until the command exists', () => {
    const clear = control(group('Home', 'Editing'), 'edit.clearContents');
    expect(clear.kind).toBe('split');
    // The map keeps the planned entry so the menu does not move later.
    expect(clear.menu?.map((item) => item.label)).toContain('Clear All');
    expect(inventoryCommand('edit.clearAll')?.implemented).toBe(false);
    // The rendered menu drops it: hidden, not greyed.
    expect(ribbonMenuItems(clear.menu ?? [], alwaysEnabled).map((item) => item.label)).toEqual([
      'Clear Contents',
      'Clear Formats',
    ]);
  });
});

describe('Data tab', () => {
  it('keeps Sort, Filter, Clear and Reapply in Sort & Filter (D7)', () => {
    const sortAndFilter = group('Data', 'Sort & Filter');
    expect(sortAndFilter.controls.map((entry) => entry.label)).toEqual([
      'Sort',
      'Filter',
      'Clear',
      'Reapply',
    ]);
  });
});

describe('hidden-not-greyed outside the registered commands', () => {
  it('renders no control for a command that is only planned', () => {
    const rendered = new Set<string>();
    for (const tab of RIBBON_TABS) {
      for (const entry of groups(tab.id)) {
        for (const renderedControl of entry.controls) rendered.add(renderedControl.commandId);
        if (entry.launcher !== undefined) rendered.add(entry.launcher.commandId);
      }
    }
    for (const commandId of rendered) {
      const entry = inventoryCommand(commandId);
      if (entry === undefined) throw new Error(`${commandId} is not in the map`);
      if (!entry.implemented) expect(commandId).toBe('format.conditionalRender');
    }
    // Spot-check the neighbours that must stay out of the chrome.
    for (const hidden of ['merge.center', 'structure.deleteRows', 'edit.clearAll', 'view.freeze']) {
      expect(rendered.has(hidden)).toBe(false);
    }
  });

  it('offers no menu item for a planned command', () => {
    const items = menuItemsFor('File', alwaysEnabled);
    expect(items.length).toBeGreaterThan(0);
    expect(
      items.every((item) => inventoryCommand(item.commandId ?? '')?.implemented === true),
    ).toBe(true);
  });
});

describe('every rendered control maps to one command', () => {
  it('keeps all placements inside the authoritative map', () => {
    for (const entry of COMMAND_INVENTORY) {
      for (const placement of entry.ribbon ?? []) {
        if (placement.menuOnly === true) continue;
        expect(placement.group.length).toBeGreaterThan(0);
        for (const item of placement.menu ?? []) {
          expect(inventoryCommand(item.commandId)).toBeDefined();
        }
      }
    }
  });
});

describe('group dialog launchers', () => {
  it('open the Format Cells dialog on their own tab (ticket 28)', () => {
    expect(group('Home', 'Font').launcher?.args).toEqual({ tab: 'font' });
    expect(group('Home', 'Alignment').launcher?.args).toEqual({ tab: 'alignment' });
    expect(group('Home', 'Number').launcher?.args).toEqual({ tab: 'number' });
  });
});
