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
      endRow: 8,
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
    expect(projection.columnX(2)).toBe(238);
    expect(projection.rowAt(0)).toBe(0);
    expect(projection.columnAt(135)).toBe(1);
    expect(projection.columnLabel(27)).toBe('AB');
  });

  it('honours imported row heights and column widths', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      defaultRowHeight: 22,
      defaultColumnWidth: 104,
    });
    // Row 1 carries ht="22" points, which is 29 pixels; the other rows keep
    // the default height.
    expect(projection.rowHeight(0)).toBe(29);
    expect(projection.rowHeight(1)).toBe(22);
    expect(projection.rowY(1)).toBe(29);
    expect(projection.rowY(10)).toBe(29 + 9 * 22);
    // Column A carries width="18.5" characters and column B width="14".
    expect(projection.columnWidth(0)).toBe(135);
    expect(projection.columnWidth(1)).toBe(103);
    expect(projection.columnX(1)).toBe(135);
    expect(projection.columnAt(134)).toBe(0);
    expect(projection.columnAt(135)).toBe(1);
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
