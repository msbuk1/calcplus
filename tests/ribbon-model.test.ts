import { describe, expect, it } from 'vitest';
import { COMMAND_INVENTORY, inventoryCommand } from '@commands/inventory';
import { functionsInCategory } from '@engine/function-catalog';
import type { PublishedFunction } from '@engine/function-registry';
import {
  RIBBON_TABS,
  menuItemsFor,
  quickAccessControls,
  ribbonGroupsFor,
  ribbonMenuItems,
} from '@chrome/ribbon-model';
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
    // Both colour split buttons open the shared picker, not a preset menu.
    expect(control(font, 'format.fill')).toMatchObject({ picker: 'fill' });
    expect(control(font, 'format.fill').menu).toBeUndefined();
    expect(font.controls.find((entry) => entry.title === 'Font Colour')).toMatchObject({
      picker: 'font',
    });
    expect(font.controls.find((entry) => entry.title === 'Font Colour')?.menu).toBeUndefined();
  });

  it('gives the Borders split button Excel 365s fifteen-item submenu', () => {
    const borders = control(group('Home', 'Font'), 'format.border');
    expect(borders.kind).toBe('split');
    // All Borders stays the split button's default action.
    expect(borders.args).toEqual({ preset: 'all', style: 'thin', color: '#000000' });
    expect(borders.menu?.map((item) => item.label)).toEqual([
      'Bottom Border',
      'Top Border',
      'Left Border',
      'Right Border',
      'No Border',
      'All Borders',
      'Outside Borders',
      'Thick Outside Borders',
      'Bottom Double Border',
      'Thick Bottom Border',
      'Top and Bottom Border',
      'Top and Thick Bottom Border',
      'Top and Double Bottom Border',
      'Inside Borders',
      'More Borders...',
    ]);
    expect(borders.menu?.map((item) => item.icon)).toEqual([
      'border-bottom',
      'border-top',
      'border-left',
      'border-right',
      'border-none',
      'border-all',
      'border-outside',
      'border-thick-outside',
      'border-bottom-double',
      'border-thick-bottom',
      'border-top-bottom',
      'border-top-thick-bottom',
      'border-top-double-bottom',
      'border-inside',
      'border-more',
    ]);
    // Dividers open the pattern group and the dialog entry.
    expect(
      (borders.menu ?? []).flatMap((item, index) => (item.dividerBefore === true ? [index] : [])),
    ).toEqual([5, 14]);
    // More Borders opens the Format Cells dialog on its Border tab.
    expect(borders.menu?.[14]?.commandId).toBe('format.cellsDialog');
    expect(borders.menu?.[14]?.args).toEqual({ tab: 'border' });
    // The single-edge and thick/double items carry the style the command accepts.
    expect(borders.menu?.[0]?.args).toEqual({
      edges: { bottom: { style: 'thin', color: '#000000', weight: 'thin' } },
    });
    expect(borders.menu?.[7]?.args).toEqual({
      preset: 'outline',
      style: 'thick',
      color: '#000000',
    });
    expect(borders.menu?.[8]?.args).toEqual({
      edges: { bottom: { style: 'double', color: '#000000', weight: 'thick' } },
    });
    // Every item is implemented, so the rendered menu keeps all fifteen rows
    // and copies each artwork name.
    const rendered = ribbonMenuItems(borders.menu ?? [], alwaysEnabled);
    expect(rendered).toHaveLength(15);
    expect(rendered.map((item) => item.icon)).toEqual(borders.menu?.map((item) => item.icon));
  });
});

describe('Alignment group', () => {
  it('renders the quick aligns, the Merge & Center split button and Wrap Text', () => {
    const alignment = group('Home', 'Alignment');
    expect(alignment.controls.map((entry) => entry.title)).toEqual([
      'Align Left',
      'Centre',
      'Align Right',
      'Align Top',
      'Align Middle',
      'Align Bottom',
      'Merge & Center',
      'Wrap Text',
    ]);
    // Merge sits in row 2 after the vertical aligns, as Excel places it.
    const merge = control(alignment, 'merge.center');
    expect(merge).toMatchObject({ kind: 'split', row: 2, icon: 'merge', wide: true });
    expect(merge.menu?.map((item) => item.label)).toEqual([
      'Merge & Center',
      'Merge Across',
      'Merge Cells',
      'Unmerge',
    ]);
    expect(merge.menu?.map((item) => item.commandId)).toEqual([
      'merge.center',
      'merge.across',
      'merge.cells',
      'merge.unmerge',
    ]);
    // Every item is live now, so the rendered menu keeps all four rows.
    expect(ribbonMenuItems(merge.menu ?? [], alwaysEnabled).map((item) => item.label)).toEqual([
      'Merge & Center',
      'Merge Across',
      'Merge Cells',
      'Unmerge',
    ]);
    expect(control(alignment, 'format.wrapText').kind).toBe('toggle');
    expect(control(alignment, 'format.wrapText').row).toBe(1);
    // The six quick aligns render icons, not the old letter labels.
    expect(alignment.controls.slice(0, 6).map((entry) => entry.icon)).toEqual([
      'align-left',
      'align-centre',
      'align-right',
      'align-top',
      'align-middle',
      'align-bottom',
    ]);
    expect(alignment.controls.slice(0, 6).every((entry) => entry.iconOnly)).toBe(true);
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
      { label: '$', title: 'Currency', icon: 'currency', row: 2 },
      { label: '%', title: 'Percent', icon: 'percent', row: 2 },
      { label: ',', title: 'Comma', icon: 'comma', row: 2 },
      { label: '.00', title: 'Increase Decimal', icon: 'increase-decimal', row: 2 },
      { label: '.00', title: 'Decrease Decimal', icon: 'decrease-decimal', row: 2 },
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
  it('renders the Insert and Delete menus now that the structural commands are implemented', () => {
    const cells = group('Home', 'Cells');
    expect(cells.controls.map((entry) => entry.label)).toEqual(['Insert', 'Delete']);
    expect(control(cells, 'sheet.insert').kind).toBe('menu');
    const menu = control(cells, 'sheet.insert').menu ?? [];
    expect(menu.map((item) => item.label)).toEqual([
      'Insert Rows Above',
      'Insert Columns to the Left',
      'Insert Cells & Shift Right',
      'Insert Cells & Shift Down',
      'Insert Sheet',
    ]);
    expect(ribbonMenuItems(menu, alwaysEnabled).map((item) => item.label)).toEqual([
      'Insert Rows Above',
      'Insert Columns to the Left',
      'Insert Cells & Shift Right',
      'Insert Cells & Shift Down',
      'Insert Sheet',
    ]);

    // The Delete menu mirrors the Insert menu (ticket 28's open capture
    // question): Delete Rows / Columns / Cells & Shift Up / Shift Left / Sheet.
    const deleteMenu = control(cells, 'structure.deleteCells').menu ?? [];
    expect(deleteMenu.map((item) => item.label)).toEqual([
      'Delete Rows',
      'Delete Columns',
      'Delete Cells & Shift Up',
      'Delete Cells & Shift Left',
      'Delete Sheet',
    ]);
    expect(ribbonMenuItems(deleteMenu, alwaysEnabled).map((item) => item.label)).toEqual([
      'Delete Rows',
      'Delete Columns',
      'Delete Cells & Shift Up',
      'Delete Cells & Shift Left',
      'Delete Sheet',
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

  it('lists Clear All, Clear Contents and Clear Formats in the Clear menu', () => {
    const clear = control(group('Home', 'Editing'), 'edit.clearContents');
    expect(clear.kind).toBe('split');
    expect(clear.menu?.map((item) => item.label)).toEqual([
      'Clear All',
      'Clear Contents',
      'Clear Formats',
    ]);
    expect(inventoryCommand('edit.clearAll')?.implemented).toBe(true);
    expect(ribbonMenuItems(clear.menu ?? [], alwaysEnabled).map((item) => item.label)).toEqual([
      'Clear All',
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

describe('Formulas Function Library category menus', () => {
  const recentFunctions: readonly PublishedFunction[] = ['AVERAGE', 'SUM'];

  function functionLibrary(): ResolvedRibbonGroup {
    const found = ribbonGroupsFor('Formulas', alwaysEnabled, recentFunctions).find(
      (candidate) => candidate.id === 'Function Library',
    );
    if (found === undefined) throw new Error('No Function Library group on Formulas');
    return found;
  }

  it('renders all ten controls in the oracle order', () => {
    expect(functionLibrary().controls.map((entry) => entry.label)).toEqual([
      'Insert Function',
      'AutoSum',
      'Recently Used',
      'Financial',
      'Logical',
      'Text',
      'Date & Time',
      'Lookup & Reference',
      'Math & Trig',
      'More Functions',
    ]);
  });

  it('keeps Insert Function tall beside the nine drop-downs in one row', () => {
    const controls = functionLibrary().controls;
    expect(controls[0]).toMatchObject({
      commandId: 'formula.insertFunction',
      row: 1,
      rowSpan: 3,
      icon: 'insert-function',
    });
    expect(controls[1]).toMatchObject({ commandId: 'formula.autosum', row: 1 });
    // The nine drop-downs carry no row assignment: how they lay out (one row
    // wide, three columns by three rows narrow) is presentation and lives in
    // the components and CSS.
    expect(controls.slice(1).every((entry) => entry.row === 1 && entry.rowSpan === 1)).toBe(true);
    expect(functionLibrary().rowCount).toBe(3);
  });

  it('breaks only the hero label, where the prototype breaks it', () => {
    const controls = functionLibrary().controls;
    expect(controls[0]?.labelLines).toEqual(['Insert', 'Function']);
    expect(controls[0]?.labelLines?.join(' ')).toBe(controls[0]?.label);
    // The nine drop-downs wrap naturally in the wide state and clip in the
    // narrow one; no authored breaks.
    expect(controls.slice(1).every((entry) => entry.labelLines === undefined)).toBe(true);
  });

  it('builds every category menu from the Insert Function catalogue', () => {
    const controls = functionLibrary().controls.filter(
      (entry) => entry.functionCategory !== undefined,
    );
    expect(controls.map((entry) => entry.functionCategory)).toEqual([
      'Recently used',
      'Financial',
      'Logical',
      'Text',
      'Date & Time',
      'Lookup & Reference',
      'Math & Trig',
    ]);
    // The ribbon and the dialog read one source, so their lists agree.
    for (const control of controls) {
      const category = control.functionCategory;
      if (category === undefined) continue;
      const expected = functionsInCategory(category, recentFunctions).map((entry) => entry.name);
      expect(
        control.menu?.map((item) => item.label),
        category,
      ).toEqual(expected);
      for (const item of control.menu ?? []) {
        expect(item.commandId).toBe('formula.insertFunction');
        expect(item.args).toEqual({ name: item.label });
      }
    }
    // Financial has no published function yet: the generated list is empty and
    // the renderer keeps the control closed rather than opening a blank menu.
    expect(controls[1]?.menu).toEqual([]);
    // Recently Used fills from the same session list the dialog reads.
    expect(controls[0]?.menu?.map((item) => item.label)).toEqual(['AVERAGE', 'SUM']);
  });

  it('maps More Functions to the Insert Function command', () => {
    expect(functionLibrary().controls.at(-1)).toMatchObject({
      commandId: 'formula.insertFunction',
      label: 'More Functions',
      kind: 'button',
      icon: 'function-more',
    });
  });

  it('keeps V2 formula and name controls off the Formulas tab', () => {
    const rendered = new Set<string>();
    for (const entry of groups('Formulas')) {
      expect(entry.id).toBe('Function Library');
      for (const control of entry.controls) rendered.add(control.commandId);
    }
    // Every planned row in §4 stays off the chrome, including the V2 auditing,
    // watch-window and manual-calculation commands.
    const planned = COMMAND_INVENTORY.filter(
      (entry) => entry.matrixSection === '4. Formulas and names' && entry.implemented === false,
    ).map((entry) => entry.id);
    expect(planned).toContain('formula.evaluate');
    expect(planned).toContain('formula.recalculate');
    expect(planned).toContain('name.resolve');
    for (const id of planned) expect(rendered.has(id), id).toBe(false);
    // The labels the capture rules out never appear, even if an id slips in.
    const labels = functionLibrary().controls.map((entry) => entry.label);
    for (const label of [
      'Trace Precedents',
      'Trace Dependents',
      'Show Formulas',
      'Error Checking',
      'Evaluate Formula',
      'Watch Window',
      'Calculation Options',
      'Calculate Now',
      'Calculate Sheet',
    ]) {
      expect(labels).not.toContain(label);
    }
  });
});

describe('the View tab freeze entry (ticket 43)', () => {
  const frozen = (): { enabled: boolean; checked: boolean } => ({ enabled: true, checked: true });

  it('reads Unfreeze Panes while the active sheet has a frozen split', () => {
    const windowGroup = ribbonGroupsFor('View', frozen).find(
      (candidate) => candidate.id === 'Window',
    );
    const freeze = windowGroup?.controls.find((candidate) => candidate.commandId === 'view.freeze');
    expect(freeze?.label).toBe('Unfreeze Panes');
    // The visible label and the accessible name move together.
    expect(freeze?.title).toBe('Unfreeze Panes');
    expect(freeze?.icon).toBe('freeze-panes');
  });

  it('reads Freeze Panes again once the split is gone', () => {
    const windowGroup = ribbonGroupsFor('View', alwaysEnabled).find(
      (candidate) => candidate.id === 'Window',
    );
    const freeze = windowGroup?.controls.find((candidate) => candidate.commandId === 'view.freeze');
    expect(freeze?.label).toBe('Freeze Panes');
    expect(freeze?.title).toBe('Freeze Panes');
  });

  it('keeps the authored label when no state resolver is supplied', () => {
    const windowGroup = group('View', 'Window');
    expect(control(windowGroup, 'view.freeze').label).toBe('Freeze Panes');
  });
});

describe('quick access toolbar (ticket 34)', () => {
  it('holds Save, Print, Undo and Redo in oracle order with shortcut tooltips', () => {
    const controls = quickAccessControls();
    expect(controls.map((entry) => entry.commandId)).toEqual([
      'file.save',
      'file.print',
      'edit.undo',
      'edit.redo',
    ]);
    expect(controls.map((entry) => entry.title)).toEqual([
      'Save (Ctrl+S)',
      'Print (Ctrl+P)',
      'Undo (Ctrl+Z)',
      'Redo (Ctrl+Y)',
    ]);
    for (const control of controls) {
      expect(control.icon).toBeDefined();
      expect(control.title).toContain(control.label);
    }
  });

  it('keeps Save As out of the toolbar and in the File menu', () => {
    expect(quickAccessControls().some((entry) => entry.commandId === 'file.saveAs')).toBe(false);
    expect(inventoryCommand('file.saveAs')?.quickAccess).toBeUndefined();
    expect(menuItemsFor('File', alwaysEnabled).map((item) => item.commandId)).toContain(
      'file.saveAs',
    );
  });

  it('hangs Save As off the Save split arrow, as the oracle shows it', () => {
    const [save, ...rest] = quickAccessControls();
    expect(save?.commandId).toBe('file.save');
    expect(save?.menu?.map((item) => item.commandId)).toEqual(['file.saveAs']);
    const rows = ribbonMenuItems(save?.menu ?? [], alwaysEnabled);
    expect(rows.map((item) => item.label)).toEqual(['Save As']);
    expect(rows[0]?.shortcut).toBe('Ctrl+Shift+S');
    // Only Save carries a drop-down; Undo and Redo stay single buttons.
    expect(rest.every((entry) => entry.menu === undefined)).toBe(true);
  });

  it('draws its icon from the shared original artwork', () => {
    for (const control of quickAccessControls()) {
      expect(inventoryCommand(control.commandId)?.icon).toBeDefined();
    }
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
    for (const hidden of ['structure.deleteRows', 'view.zoom']) {
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
