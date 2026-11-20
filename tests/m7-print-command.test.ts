import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { openBlankWorkbook } from '@ooxml/blank';
import type { GridRange } from '@model/types';
import inventory from '../src/commands/command-inventory.json';

/**
 * M7 lane 2 — File → Print is map-first: the command map names it before any
 * UI exists (ui-chrome §4.1), the registry registers it, and the File menu
 * reads the map's placement.
 */

interface InventoryEntry {
  readonly id: string;
  readonly label: string;
  readonly implemented: boolean;
  readonly shortcut: readonly string[];
  readonly menu?: readonly { readonly menu: string; readonly order?: number }[];
}

const entries = inventory as readonly InventoryEntry[];

function commandOf(id: string) {
  const command = createCommandRegistry().get(id);
  if (command === undefined) throw new Error(`No such command: ${id}`);
  return command;
}

const NO_SELECTION: GridRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

function context(sheets: number) {
  const workbook = openBlankWorkbook().workbook;
  return {
    workbook: sheets === 0 ? { ...workbook, sheets: [] } : workbook,
    activeSheetId: sheets === 0 ? null : (workbook.sheets[0]?.id ?? null),
    selection: NO_SELECTION,
    activeCell: { row: 0, col: 0 },
    view: {} as never,
    busy: false,
    canUndo: false,
    canRedo: false,
  };
}

describe('File → Print in the command map', () => {
  it('is implemented, on the File menu, and keeps Ctrl+P', () => {
    const entry = entries.find((candidate) => candidate.id === 'file.print');
    expect(entry).toBeDefined();
    expect(entry).toMatchObject({
      label: 'Print',
      implemented: true,
      shortcut: ['Ctrl+P'],
    });
    expect(entry?.menu).toEqual([{ menu: 'File', order: 50 }]);
  });

  it('is registered with the map’s label and shortcut', () => {
    const command = commandOf('file.print');
    expect(command.label).toBe('Print');
    expect(command.keybindings).toEqual(['Ctrl+P']);
  });

  it('is enabled only while a workbook is open and the app is idle', () => {
    const command = commandOf('file.print');
    expect(command.isEnabled(context(1) as never)).toBe(true);
    expect(command.isEnabled(context(0) as never)).toBe(false);
    expect(command.isEnabled({ ...context(1), busy: true } as never)).toBe(false);
  });
});
