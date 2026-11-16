import { describe, expect, it } from 'vitest';
import type { Cell, GridRange, Sheet } from '@model/types';
import { setCellSnapshot } from '@model/types';
import {
  planRangeTransform,
  transformFormulaReferences,
  transformMergeRange,
  type StructuralEdit,
} from '@model/range-transform';
import { numberValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { openWorkbook } from '@ooxml/open';
import { applyRangeTransformPrototype } from '@ooxml/range-transform';
import { buildXlsx } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function formulaCell(formula: string): Cell {
  return {
    value: numberValue(0),
    formula,
    formulaDerived: false,
    savedResult: numberValue(0),
    styleId: 0,
    richText: null,
  };
}

function setFormula(sheet: Sheet, row: number, col: number, formula: string): void {
  setCellSnapshot(sheet, row, col, formulaCell(formula));
}

function twoSheets() {
  const workbook = createEmptyWorkbook();
  const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
  workbook.sheets.push(first, second);
  return { workbook, first, second };
}

function insertRow(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'insert', sheetId, axis: 'row', index, count };
}

describe('M3C range transform formulas and names', () => {
  it('emits a sheet prefix once and transforms qualified references across sheets', () => {
    const { workbook, first, second } = twoSheets();
    setFormula(first, 1, 0, 'Sheet2!A2+A1');
    const edit = insertRow(second.id, 1);
    const plan = planRangeTransform(workbook, edit);
    expect(plan.formulaRewrites).toHaveLength(1);
    expect(plan.formulaRewrites[0]?.before).toBe('Sheet2!A2+A1');
    expect(plan.formulaRewrites[0]?.after).toBe('Sheet2!A3+A1');
  });

  it('keeps a quoted sheet prefix once', () => {
    const workbook = createEmptyWorkbook();
    const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    const quoted = createSheet('Sales Data', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(first, quoted);
    setFormula(first, 1, 0, "'Sales Data'!A2");
    const edit = insertRow(quoted.id, 1);
    expect(transformFormulaReferences("'Sales Data'!A2", edit, { workbook }, first.id)).toBe(
      "'Sales Data'!A3",
    );
  });

  it('shifts, shrinks and deletes ranges on delete', () => {
    const { workbook, first } = twoSheets();
    setFormula(first, 0, 0, 'A2');
    setFormula(first, 1, 1, 'A1');
    const plan = planRangeTransform(workbook, {
      kind: 'delete',
      sheetId: first.id,
      axis: 'row',
      index: 0,
      count: 1,
    });
    expect(plan.formulaRewrites.map((rewrite) => rewrite.after)).toEqual(['A1', '#REF!']);
  });

  it('rewrites whole columns and whole rows by the correct axis only', () => {
    const { workbook, first } = twoSheets();
    const edit = insertRow(first.id, 1);
    expect(transformFormulaReferences('SUM(A:A)', edit, { workbook }, first.id)).toBe('SUM(A:A)');
    expect(transformFormulaReferences('SUM(1:1)', edit, { workbook }, first.id)).toBe('SUM(1:1)');
    expect(
      transformFormulaReferences(
        'SUM(1:1)',
        { kind: 'insert', sheetId: first.id, axis: 'row', index: 0, count: 1 },
        { workbook },
        first.id,
      ),
    ).toBe('SUM(2:2)');
  });

  it('preserves grouping unions but rewrites function arguments', () => {
    const { workbook, first } = twoSheets();
    const edit = insertRow(first.id, 0);
    expect(transformFormulaReferences('(A1:A2,C1:C2)', edit, { workbook }, first.id)).toBe(
      '(A1:A2,C1:C2)',
    );
    expect(transformFormulaReferences('SUM(A1:A2,C1:C2)', edit, { workbook }, first.id)).toBe(
      'SUM(A2:A3,C2:C3)',
    );
  });

  it('does not parse a reference suffix inside a defined name', () => {
    const { workbook, first } = twoSheets();
    const edit = insertRow(first.id, 0);
    expect(transformFormulaReferences('Rate1+myName2+A1', edit, { workbook }, first.id)).toBe(
      'Rate1+myName2+A2',
    );
  });

  it('rewrites workbook and sheet defined names', () => {
    const { workbook, first, second } = twoSheets();
    workbook.names.push({
      name: 'TaxRange',
      formula: 'Sheet1!$A$1:$B$2',
      sheetId: null,
    });
    second.names.push({ name: 'LocalAmount', formula: '$C$3', sheetId: second.id });

    const firstPlan = planRangeTransform(workbook, insertRow(first.id, 1));
    expect(firstPlan.nameRewrites.map((rewrite) => rewrite.after)).toEqual(['Sheet1!$A$1:$B$3']);

    const secondPlan = planRangeTransform(workbook, insertRow(second.id, 1));
    expect(secondPlan.nameRewrites.map((rewrite) => rewrite.after)).toEqual(['$C$4']);
  });
});

describe('M3C range transform merges', () => {
  it('shifts, shrinks and removes merged ranges', () => {
    const range: GridRange = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    const insert = insertRow('sheet', 1);
    expect(transformMergeRange(range, insert)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 1,
    });
    expect(
      transformMergeRange(range, {
        kind: 'delete',
        sheetId: 'sheet',
        axis: 'row',
        index: 0,
        count: 1,
      }),
    ).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 });
    expect(
      transformMergeRange(
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
        {
          kind: 'delete',
          sheetId: 'sheet',
          axis: 'row',
          index: 0,
          count: 1,
        },
      ),
    ).toBeNull();
  });
});

describe('M3C OOXML range-transform prototype', () => {
  const worksheet = `${XML_HEADER}
<worksheet xmlns="${NS}">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="B2"><f>A1*2</f><v>2</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
</worksheet>`;

  function openedPrototypeWorkbook() {
    return openWorkbook(
      buildXlsx(worksheet, {
        workbookExtras:
          '<definedNames><definedName name="DataRange">Sheet1!$A$1:$B$2</definedName></definedNames>',
      }),
    );
  }

  it('rewrites formula, merge and defined-name text on insert', () => {
    const opened = openedPrototypeWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    const result = applyRangeTransformPrototype({
      workbook: opened.workbook,
      package: opened.package,
      edit: { kind: 'insert', sheetId: sheet.id, axis: 'row', index: 0, count: 1 },
    });
    const sheetXml = new TextDecoder().decode(result.entries.get(sheet.partPath)!);
    const workbookXml = new TextDecoder().decode(
      result.entries.get(opened.package.workbookPartPath)!,
    );
    expect(sheetXml).toContain('<f>A2*2</f>');
    expect(sheetXml).toContain('<mergeCell ref="A2:B2"/>');
    expect(sheetXml).toContain('<mergeCells count="1">');
    expect(workbookXml).toContain('<definedName name="DataRange">Sheet1!$A$2:$B$3</definedName>');
    expect(result.report.formulaRewrites).toBe(1);
    expect(result.report.mergesRewritten).toBe(1);
    expect(result.report.namesRewritten).toBe(1);
  });

  it('removes a deleted merge and leaves #REF! in the formula', () => {
    const opened = openedPrototypeWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    const result = applyRangeTransformPrototype({
      workbook: opened.workbook,
      package: opened.package,
      edit: { kind: 'delete', sheetId: sheet.id, axis: 'row', index: 0, count: 1 },
    });
    const sheetXml = new TextDecoder().decode(result.entries.get(sheet.partPath)!);
    const workbookXml = new TextDecoder().decode(
      result.entries.get(opened.package.workbookPartPath)!,
    );
    expect(sheetXml).toContain('<f>#REF!*2</f>');
    expect(sheetXml).not.toMatch(/<mergeCell\b[^>]*\bref=/);
    expect(sheetXml).toContain('<mergeCells count="0">');
    expect(workbookXml).toContain('<definedName name="DataRange">Sheet1!$A$1:$B$1</definedName>');
    expect(result.report.mergesRemoved).toBe(1);
  });
});
