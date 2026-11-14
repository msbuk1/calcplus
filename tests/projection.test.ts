import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { createSheetProjection, visibleRange } from '@view/projection';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

describe('sheet projection', () => {
  it('returns ready-to-paint records only for the requested window', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 7,
      defaultRowHeight: 22,
      defaultColumnWidth: 104,
    });
    expect(projection.version).toBe(7);
    expect(projection.cellAt(1, 1)?.displayText).toBe('42.5');
    expect(projection.cellAt(2, 1)?.formula).toBe('B2*2');
    expect(projection.cellAt(3, 1)?.displayText).toBe('#DIV/0!');
    expect(projection.window(0, 1, 0, 1).map((cell) => cell.displayText)).toEqual([
      'Name',
      'Value',
      'Alpha',
      '42.5',
    ]);
    expect(visibleRange(projection, 0, 0, 300, 200)).toEqual({
      startRow: 0,
      endRow: 9,
      startCol: 0,
      endCol: 2,
    });
  });

  it('answers geometry by index without iterating the grid', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      defaultRowHeight: 22,
      defaultColumnWidth: 104,
    });
    expect(projection.rowY(0)).toBe(0);
    expect(projection.rowY(10)).toBe(220);
    expect(projection.columnX(2)).toBe(208);
    expect(projection.rowAt(0)).toBe(0);
    expect(projection.rowAt(21)).toBe(0);
    expect(projection.rowAt(22)).toBe(1);
    expect(projection.columnAt(104)).toBe(1);
    expect(projection.columnLabel(27)).toBe('AB');
  });

  it('does not include selection or scroll in the document version', () => {
    const opened = openWorkbook(fixture('identity-minimal.xlsx'));
    const first = createSheetProjection(opened.workbook, opened.initialSheetId, { version: 3 });
    const second = createSheetProjection(opened.workbook, opened.initialSheetId, { version: 3 });
    expect(first.cellAt(0, 0)?.displayText).toBe('Hello');
    expect(second.cellAt(0, 0)?.displayText).toBe('Hello');
    expect(first).not.toBe(second);
  });
});
