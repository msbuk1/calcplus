import { describe, expect, it } from 'vitest';
import { contextMenuItems } from '@chrome/ribbon-model';
import type { MenuItemModel } from '@chrome/Menu';

/**
 * The grid context menu is a view of the checked-in command map (ui-chrome
 * §4.1). Insert and Delete are one-level flyouts carrying the existing
 * structure commands: their labels and shortcuts come verbatim from the M5
 * capture (menu-dialogs §7), and their enablement is the registry's.
 */

function resolver(enabledByCommand: Readonly<Record<string, boolean>> = {}) {
  return (commandId: string): { enabled: boolean; checked: boolean } => ({
    enabled: enabledByCommand[commandId] ?? true,
    checked: false,
  });
}

function item(items: readonly MenuItemModel[], label: string): MenuItemModel {
  const found = items.find((candidate) => candidate.label === label);
  if (found === undefined) throw new Error(`No "${label}" item`);
  return found;
}

describe('the grid context menu', () => {
  it('puts Insert and Delete directly after the clipboard group', () => {
    const labels = contextMenuItems(resolver()).map((entry) => entry.label);
    expect(labels.slice(0, 6)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      'Format Painter',
      'Insert',
      'Delete',
    ]);
    // The old flat "Insert sheet" entry now lives inside the Insert flyout.
    expect(labels).not.toContain('Insert sheet');
    expect(labels[6]).toBe('Clear contents');
  });

  it('renders Insert as a four-item flyout in capture order', () => {
    const menu = contextMenuItems(resolver());
    const insert = item(menu, 'Insert');
    expect(insert.commandId).toBeUndefined();
    expect(insert.items?.map((child) => child.label)).toEqual([
      'Insert Rows Above',
      'Insert Columns to the Left',
      'Insert Cells & Shift Right',
      'Insert Cells & Shift Down',
    ]);
    expect(insert.items?.map((child) => child.shortcut ?? null)).toEqual([
      'Ctrl+Shift+=',
      'Ctrl+Shift+=',
      'Ctrl+Shift+=',
      'Ctrl+Shift+=',
    ]);
    expect(insert.items?.map((child) => child.commandId)).toEqual([
      'structure.insertRows',
      'structure.insertColumns',
      'structure.insertCells',
      'structure.insertCells',
    ]);
    // The two Insert Cells directions are one command with different arguments.
    expect(insert.items?.[2]?.args).toEqual({ shift: 'right' });
    expect(insert.items?.[3]?.args).toEqual({ shift: 'down' });
  });

  it('renders Delete as a four-item flyout in capture order', () => {
    const menu = contextMenuItems(resolver());
    const remove = item(menu, 'Delete');
    expect(remove.commandId).toBeUndefined();
    expect(remove.items?.map((child) => child.label)).toEqual([
      'Delete Rows',
      'Delete Columns',
      'Delete Cells & Shift Up',
      'Delete Cells & Shift Left',
    ]);
    // The structure deletes carry the map's Ctrl+- shortcut.
    expect(remove.items?.map((child) => child.shortcut ?? null)).toEqual([
      'Ctrl+-',
      'Ctrl+-',
      'Ctrl+-',
      'Ctrl+-',
    ]);
    expect(remove.items?.map((child) => child.commandId)).toEqual([
      'structure.deleteRows',
      'structure.deleteColumns',
      'structure.deleteCells',
      'structure.deleteCells',
    ]);
    expect(remove.items?.[2]?.args).toEqual({ shift: 'up' });
    expect(remove.items?.[3]?.args).toEqual({ shift: 'left' });
  });

  it('is one level deep: flyout children are leaves', () => {
    for (const parent of contextMenuItems(resolver())) {
      for (const child of parent.items ?? []) {
        expect(child.items).toBeUndefined();
      }
    }
  });

  it('follows the registry for each child', () => {
    const menu = contextMenuItems(resolver({ 'structure.deleteRows': false }));
    const insert = item(menu, 'Insert');
    const remove = item(menu, 'Delete');
    expect(insert.disabled).toBe(false);
    expect(remove.disabled).toBe(false);
    expect(item(insert.items ?? [], 'Insert Rows Above').disabled).toBe(false);
    expect(item(remove.items ?? [], 'Delete Rows').disabled).toBe(true);
    expect(item(remove.items ?? [], 'Delete Columns').disabled).toBe(false);
  });

  it('disables a flyout only when every child is disabled', () => {
    const menu = contextMenuItems(
      resolver({
        'structure.insertRows': false,
        'structure.insertColumns': false,
        'structure.insertCells': false,
      }),
    );
    const insert = item(menu, 'Insert');
    expect(insert.disabled).toBe(true);
    expect(insert.items?.every((child) => child.disabled === true)).toBe(true);
    // Delete stays available because its children are.
    expect(item(menu, 'Delete').disabled).toBe(false);
  });
});
