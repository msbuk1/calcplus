import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { numbersApproxEqual } from '@testing/compare';

const opened = openWorkbook(
  new Uint8Array(
    readFileSync(new URL('../fixtures/oracle/m1-values-errors.xlsx', import.meta.url)),
  ),
);
const sheet = opened.workbook.sheets[0]!;

function cellKey(ref: string): string {
  const match = /^([A-Z]+)(\d+)$/.exec(ref)!;
  let col = 0;
  for (const character of match[1]!) {
    col = col * 26 + (character.charCodeAt(0) - 64);
  }
  return `${Number(match[2]) - 1}:${col - 1}`;
}

const cell = (ref: string) => sheet.cells.get(cellKey(ref));

describe('M1 Excel oracle fixture', () => {
  it('opens and contains the typed values, errors and formula/constant cases the oracle captures', () => {
    expect(sheet.name).toBe('Sheet1');
    expect(sheet.cells.size).toBe(21);

    expect(cell('A1')?.value).toEqual({ type: 'text', value: 'Name' });
    expect(cell('B2')?.value).toEqual({ type: 'number', value: 42.5 });
    expect(cell('C2')?.value).toEqual({ type: 'text', value: 'inline text' });
    expect(cell('A4')?.value).toEqual({ type: 'boolean', value: true });
    expect(cell('B4')?.value).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(cell('E2')?.value).toEqual({ type: 'error', value: '#N/A' });
    expect(cell('E3')?.value).toEqual({ type: 'error', value: '#VALUE!' });
    expect(cell('A5')?.value).toEqual({ type: 'error', value: '#REF!' });
    expect(cell('B5')?.value).toEqual({ type: 'error', value: '#NAME?' });
    expect(cell('C5')?.value).toEqual({ type: 'error', value: '#NUM!' });
    expect(cell('D5')?.value).toEqual({ type: 'error', value: '#NULL!' });

    expect(cell('B3')?.formula).toBe('B2*2');
    expect(cell('B3')?.savedResult).toEqual({ type: 'number', value: 85 });
    const c3 = cell('C3');
    expect(c3?.value.type).toBe('number');
    if (c3?.value.type === 'number') {
      expect(numbersApproxEqual(c3.value.value, Math.PI)).toBe(true);
    }

    // A leading '=' in a constant is text. It must never be inferred as a formula.
    expect(cell('D2')?.value).toEqual({ type: 'text', value: '=not a formula' });
    expect(cell('D2')?.formula).toBeNull();
    expect(cell('D3')?.value).toEqual({ type: 'text', value: '=SUM(A1:A2)' });
    expect(cell('D3')?.formula).toBeNull();

    // Formula status is explicit, including when the result itself starts with '='.
    expect(cell('C4')?.formula).toBe('"=literal"');
    expect(cell('C4')?.savedResult).toEqual({ type: 'text', value: '=literal' });
    expect(cell('D4')?.formula).toBe('1/0');
    expect(cell('D4')?.savedResult).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(cell('E4')?.formula).toBe('SUM(B2:B3)');
    expect(cell('E4')?.savedResult).toEqual({ type: 'number', value: 127.5 });
  });

  it('puts the error cases in a narrow column so display text is not a usable identity', () => {
    expect(sheet.columnDimensions.get(4)?.width).toBe(4);
  });
});
