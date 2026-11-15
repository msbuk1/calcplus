import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { createSheetProjection } from '@view/projection';
import { formatValue, numberFormatColor } from '@view/number-format';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

function m2Projection() {
  const opened = openWorkbook(fixture('m2-grid.xlsx'));
  return createSheetProjection(opened.workbook, opened.initialSheetId, { version: 1 });
}

describe('M2 grid projection', () => {
  it('reads the imported freeze panes', () => {
    const projection = m2Projection();
    expect(projection.freezeRows).toBe(1);
    expect(projection.freezeCols).toBe(1);
    expect(projection.rowCount).toBe(1_048_576);
    expect(projection.columnCount).toBe(16_384);
    expect(projection.showGridLines).toBe(true);
    expect(projection.showHeadings).toBe(true);
  });

  it('reports the merged range and the raw text of a cell', () => {
    const projection = m2Projection();
    const merge = projection.mergeAt(0, 0);
    expect(merge).toEqual({ anchorRow: 0, anchorCol: 0, rowSpan: 1, colSpan: 2 });
    expect(projection.mergeAt(0, 1)).toEqual(merge);
    expect(projection.mergeAt(5, 5)).toBeNull();
    expect(projection.rawTextAt(0, 0)).toBe('Quarterly report');
    expect(projection.rawTextAt(2, 1)).toBe('0.25');
  });

  it('resolves the imported font, fill, alignment and number format', () => {
    const projection = m2Projection();
    const header = projection.cellAt(1, 0);
    expect(header?.style.font.bold).toBe(true);
    expect(header?.style.fill.color).toBe('#FFF2CC');
    expect(header?.style.alignment.horizontal).toBe('center');

    const title = projection.cellAt(0, 0);
    expect(title?.style.font.size).toBe(14);
    expect(title?.style.font.bold).toBe(true);

    const note = projection.cellAt(2, 3);
    expect(note?.style.font.italic).toBe(true);
    expect(note?.style.font.color).toBe('#B00020');

    const amount = projection.cellAt(2, 1);
    expect(amount?.style.numberFormat).toBe('#,##0.00');
    expect(amount?.displayText).toBe('0.25');

    const share = projection.cellAt(2, 2);
    expect(share?.style.numberFormat).toBe('0.00%');
    expect(share?.displayText).toBe('0.00%');
    expect(projection.cellAt(3, 2)?.displayText).toBe('13.00%');
  });

  it('finds the used range and the data of a row and a column', () => {
    const projection = m2Projection();
    expect(projection.usedRange()).toEqual({
      startRow: 0,
      endRow: 301,
      startCol: 0,
      endCol: 3,
    });
    expect(projection.rowDataColumns(2)).toEqual([0, 1, 2, 3]);
    expect(projection.rowDataColumns(3)).toEqual([0, 1, 2]);
    expect(projection.columnDataRows(0)).toHaveLength(302);
  });
});

describe('number formats', () => {
  it('renders the built-in formats', () => {
    expect(formatValue({ type: 'number', value: 1234.5 }, '0')).toBe('1235');
    expect(formatValue({ type: 'number', value: 1234.5 }, '0.00')).toBe('1234.50');
    expect(formatValue({ type: 'number', value: 1234.5 }, '#,##0')).toBe('1,235');
    expect(formatValue({ type: 'number', value: 1234.5 }, '#,##0.00')).toBe('1,234.50');
    expect(formatValue({ type: 'number', value: 0.25 }, '0%')).toBe('25%');
    expect(formatValue({ type: 'number', value: 0.25 }, '0.00%')).toBe('25.00%');
    expect(formatValue({ type: 'number', value: 1234.5 }, null)).toBe('1234.5');
  });

  it('renders dates from the 1900 date system', () => {
    // Excel serial 45000 is 15 March 2023.
    expect(formatValue({ type: 'number', value: 45_000 }, 'm/d/yy')).toBe('3/15/23');
    expect(formatValue({ type: 'number', value: 45_000 }, 'yyyy-mm-dd')).toBe('2023-03-15');
  });

  it('reads the colour of a format code', () => {
    expect(numberFormatColor('[Red]0.00')).toBe('#ff0000');
    expect(numberFormatColor('0.00')).toBeNull();
  });

  it('does not add a second minus sign for a negative section', () => {
    // The negative section supplies the sign or the parentheses.
    expect(formatValue({ type: 'number', value: -42 }, '$#,##0_);($#,##0)')).toBe('($42)');
    expect(formatValue({ type: 'number', value: -42 }, '#,##0.00;(#,##0.00)')).toBe('(42.00)');
    expect(formatValue({ type: 'number', value: -42 }, '0.00;-0.00')).toBe('-42.00');
    // A single section keeps the minus sign.
    expect(formatValue({ type: 'number', value: -42 }, '0.00')).toBe('-42.00');
    expect(formatValue({ type: 'number', value: -0.25 }, '0%')).toBe('-25%');
    expect(formatValue({ type: 'number', value: 42 }, '$#,##0_);($#,##0)')).toBe('$42');
  });
});
