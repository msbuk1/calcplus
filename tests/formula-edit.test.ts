import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { BLANK, numberValue } from '@model/value';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ws = (rows: string): string =>
  `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`;

/**
 * The formula engine belongs to the model-depth milestone. Until then the
 * grid may author formula text: the model keeps the authored text, the cached
 * result is dropped, and the saved file asks Excel to recalculate
 * (file-compat §6.1).
 */
describe('formula authoring without the engine', () => {
  const formulaSheet = ws(
    '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>',
  );

  it('stores the authored formula and drops the cached result', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: 'A1*3',
    });
    expect(result.changes).toHaveLength(1);
    const cell = sheet.cells.get('0:1')!;
    expect(cell.formula).toBe('A1*3');
    expect(cell.savedResult).toBeNull();
    expect(cell.value).toEqual({ type: 'number', value: 4 });
  });

  it('keeps the cached result when the formula is unchanged', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: 'A1*2',
    });
    expect(result.changes).toHaveLength(0);
    expect(sheet.cells.get('0:1')?.savedResult).toEqual({ type: 'number', value: 4 });
    expect(session.hasUnsavedChanges).toBe(false);
  });

  it('undoes a formula edit back to the authored text and the cached result', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: 'A1*3',
    });
    session.undo();
    const cell = sheet.cells.get('0:1')!;
    expect(cell.formula).toBe('A1*2');
    expect(cell.savedResult).toEqual({ type: 'number', value: 4 });
  });

  it('writes the formula without a cached value and asks Excel to recalculate', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: 'A1*3',
    });
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B1"><f>A1*3</f></c>');
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });

  it('replaces a formula with a value when the user types one', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: numberValue(7),
      formula: null,
    });
    const cell = sheet.cells.get('0:1')!;
    expect(cell.formula).toBeNull();
    expect(cell.value).toEqual({ type: 'number', value: 7 });
  });

  it('clears a formula cell like a content clear', () => {
    const opened = openWorkbook(buildXlsx(formulaSheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: null,
    });
    expect(sheet.cells.get('0:1')).toBeUndefined();
  });
});
