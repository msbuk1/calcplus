import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { restoreWorkbook, snapshotWorkbook } from '@ooxml/snapshot';
import { createPartInventory } from '@ooxml/inventory';
import { saveWorkbook } from '@ooxml/save';
import { DocumentSession } from '@commands/session';
import { numberValue } from '@model/value';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

describe('workbook snapshot', () => {
  it('stores metadata as JSON and part payloads as binary records, never base64', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const snapshot = snapshotWorkbook(opened.workbook, opened.package);
    expect(snapshot.schemaVersion).toBe(1);
    expect(snapshot.parts.length).toBe(opened.package.entries.size);
    for (const part of snapshot.parts) {
      expect(part.bytes).toBeInstanceOf(Uint8Array);
    }
    expect('sourceParts' in snapshot.workbook).toBe(false);
    expect(snapshot.workbook.sheets[0]!.cells.length).toBe(10);
  });

  it('restores the model and all part bytes losslessly', async () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const snapshot = snapshotWorkbook(opened.workbook, opened.package);
    const restored = restoreWorkbook(snapshot);
    expect(restored.workbook.sheets[0]!.cells.size).toBe(10);
    expect(restored.workbook.sheets[0]!.cells.get('2:1')?.formula).toBe('B2*2');
    expect(restored.package.originalBytes).toBeNull();

    const before = await createPartInventory(opened.package.entries);
    const after = await createPartInventory(restored.package.entries);
    expect(after).toEqual(before);
    expect(
      Buffer.compare(
        Buffer.from(restored.package.entries.get('xl/charts/chart1.xml')!),
        Buffer.from(opened.package.entries.get('xl/charts/chart1.xml')!),
      ),
    ).toBe(0);
  });

  it('restores, edits and saves with Tier B parts intact', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const snapshot = snapshotWorkbook(opened.workbook, opened.package);
    const restored = restoreWorkbook(snapshot);
    const session = new DocumentSession(restored.workbook);
    const sheet = restored.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 1,
      col: 1,
      value: numberValue(99),
    });
    const result = saveWorkbook({
      workbook: restored.workbook,
      package: restored.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('1:1')?.value).toEqual({
      type: 'number',
      value: 99,
    });
    expect(reopened.workbook.opaqueParts.has('xl/charts/chart1.xml')).toBe(true);
    expect(
      Buffer.compare(
        Buffer.from(reopened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
        Buffer.from(opened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
      ),
    ).toBe(0);
  });
});
