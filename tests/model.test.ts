import { describe, expect, it } from 'vitest';
import { appendStyle, createDefaultStyle, createEmptyWorkbook, createSheet } from '@model/workbook';
import {
  cellKey,
  formatCellRef,
  parseCellRef,
  setCellSnapshot,
  getCell,
  createCell,
} from '@model/types';
import { numberValue, parseUserValue, textValue } from '@model/value';

describe('sparse workbook model', () => {
  it('stores only touched cells in a huge grid', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);

    for (let index = 0; index < 100; index += 1) {
      setCellSnapshot(sheet, index * 10_000, index % 20, createCell(numberValue(index)));
    }

    expect(sheet.cells.size).toBe(100);
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(0));
    expect(getCell(sheet, 1, 1)).toBeNull();
  });

  it('removes a cell when clearing its last content', () => {
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    setCellSnapshot(sheet, 3, 4, createCell(textValue('gone')));
    setCellSnapshot(sheet, 3, 4, null);
    expect(sheet.cells.has(cellKey(3, 4))).toBe(false);
  });

  it('keeps style ids stable and rejects renumbering', () => {
    const workbook = createEmptyWorkbook();
    expect(workbook.styles[0]).toEqual(createDefaultStyle(0));
    const first = appendStyle(workbook, {
      numFmtId: 14,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      rawXml: '<xf numFmtId="14"/>',
    });
    const second = appendStyle(workbook, {
      numFmtId: 0,
      fontId: 1,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      rawXml: '<xf fontId="1"/>',
    });
    expect(first).toBe(1);
    expect(second).toBe(2);
    expect(workbook.styles.map((style) => style.id)).toEqual([0, 1, 2]);
    expect(() =>
      appendStyle(workbook, {
        id: 1,
        numFmtId: 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        xfId: 0,
        rawXml: '',
      }),
    ).toThrow(/append-only/);
  });

  it('addresses cells with A1 notation only at the edges', () => {
    expect(formatCellRef(0, 0)).toBe('A1');
    expect(formatCellRef(1_048_575, 16_383)).toBe('XFD1048576');
    expect(parseCellRef('XFD1048576')).toEqual({ row: 1_048_575, col: 16_383 });
    expect(parseCellRef('AA10')).toEqual({ row: 9, col: 26 });
  });
});

describe('user value parsing', () => {
  it('types numbers, booleans, errors and text', () => {
    expect(parseUserValue('42.5')).toEqual({ type: 'number', value: 42.5 });
    expect(parseUserValue('TRUE')).toEqual({ type: 'boolean', value: true });
    expect(parseUserValue('#DIV/0!')).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(parseUserValue('#CIRC!')).toEqual({ type: 'error', value: '#CIRC!' });
    expect(parseUserValue('hello')).toEqual({ type: 'text', value: 'hello' });
    expect(parseUserValue('')).toEqual({ type: 'blank' });
  });

  it('does not treat formulas as formula entry in M1', () => {
    expect(parseUserValue('=1+1')).toEqual({ type: 'text', value: '=1+1' });
  });
});
