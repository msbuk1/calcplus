import { describe, expect, it } from 'vitest';
import {
  absolutizeFormulaReferences,
  findAreaReferences,
  parseWholeArea,
  resolveReferenceText,
} from '@model/references';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

function workbookWithTwoSheets() {
  const workbook = createEmptyWorkbook();
  const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  const second = createSheet('Sheet 2', 'xl/worksheets/sheet2.xml');
  workbook.sheets.push(first, second);
  workbook.names.push({
    name: 'TaxRange',
    formula: "'Sheet 2'!$A$1:$B$2",
    sheetId: null,
  });
  second.names.push({
    name: 'LocalAmount',
    formula: '$C$3',
    sheetId: second.id,
  });
  return { workbook, first, second };
}

describe('M3C A1 reference parsing', () => {
  it('records a sheet-qualified reference at the start of its prefix', () => {
    const references = findAreaReferences('Sheet2!A1+A1');
    expect(references).toHaveLength(2);
    expect(references[0]?.start).toBe(0);
    expect(references[0]?.prefix).toBe('Sheet2');
    expect(references[0]?.sheetName).toBe('Sheet2');
    expect(references[0]?.end).toBe('Sheet2!A1'.length);
    expect(references[1]?.start).toBe('Sheet2!A1+'.length);
    expect(references[1]?.prefix).toBeNull();
  });

  it('emits a sheet prefix once when absolutising', () => {
    expect(absolutizeFormulaReferences('Sheet2!A1+A1')).toBe('Sheet2!$A$1+$A$1');
    expect(absolutizeFormulaReferences("'Sheet 2'!A1")).toBe("'Sheet 2'!$A$1");
  });

  it('does not parse inside a defined name', () => {
    const references = findAreaReferences('Rate1+myName2+A1');
    expect(references).toHaveLength(1);
    expect(references[0]?.prefix).toBeNull();
    expect(references[0]?.start).toBe('Rate1+myName2+'.length);
    expect(absolutizeFormulaReferences('Rate1+myName2+A1')).toBe('Rate1+myName2+$A$1');
  });

  it('does not parse a name suffix after a sheet prefix', () => {
    expect(absolutizeFormulaReferences('Sheet1!Rate1')).toBe('Sheet1!Rate1');
  });

  it('parses whole columns and whole rows', () => {
    const columns = parseWholeArea('$A:$C');
    expect(columns?.kind).toBe('columns');
    expect(columns?.startCol).toBe(0);
    expect(columns?.endCol).toBe(2);
    expect(columns?.startColAbs).toBe(true);
    expect(columns?.endColAbs).toBe(true);

    const rows = parseWholeArea('1:3');
    expect(rows?.kind).toBe('rows');
    expect(rows?.startRow).toBe(0);
    expect(rows?.endRow).toBe(2);
  });

  it('parses quoted sheet references and defined names', () => {
    const { workbook, second } = workbookWithTwoSheets();
    const tax = resolveReferenceText(workbook, 'TaxRange', null);
    expect(tax.kind).toBe('area');
    if (tax.kind !== 'area') return;
    expect(tax.sheetId).toBe(second.id);
    expect(tax.range).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const local = resolveReferenceText(workbook, 'LocalAmount', second.id);
    expect(local.kind).toBe('area');
    if (local.kind !== 'area') return;
    expect(local.sheetId).toBe(second.id);
    expect(local.range).toEqual({ startRow: 2, startCol: 2, endRow: 2, endCol: 2 });

    const quoted = resolveReferenceText(workbook, "'Sheet 2'!B4", null);
    expect(quoted.kind).toBe('area');
    if (quoted.kind !== 'area') return;
    expect(quoted.range).toEqual({ startRow: 3, startCol: 1, endRow: 3, endCol: 1 });
  });
});
