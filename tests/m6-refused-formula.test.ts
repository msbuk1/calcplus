import { describe, expect, it } from 'vitest';
import { Buffer } from 'node:buffer';
import { DocumentSession } from '@commands/session';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { buildCompatibilityReport, buildExportReport } from '@ooxml/report';
import { textValue } from '@model/value';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * A formula Excel refuses must never reach a saved file (file-compat §1, §7;
 * M6 exit criterion 3).
 *
 * The model keeps the formula visible so it can be corrected, and every save
 * omits it: `<f>` is dropped and the cached `<v>` is kept, which is exactly
 * what Excel's own repair does. Both reports name the cell — the open report
 * so the user learns before saving, the save report because dropping a formula
 * is a change to their data.
 *
 * `tests/formula-arity.test.ts` covers the same rule at the entry point.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ws = (rows: string): string =>
  `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`;

/** The `<f>` contents of the saved worksheet, in document order. */
function writtenFormulas(bytes: Uint8Array): string[] {
  const xml = unzipPart(bytes, 'xl/worksheets/sheet1.xml');
  return [...xml.matchAll(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/g)].map((match) => match[1] ?? '');
}

function worksheetText(bytes: Uint8Array): string {
  return unzipPart(bytes, 'xl/worksheets/sheet1.xml');
}

function save(opened: ReturnType<typeof openWorkbook>, dirtyCells = new Map()) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells,
  });
}

/** `SUM()` with its cached `#N/A`, and the two formulas that must survive. */
const CONTAMINATED = ws(
  '<row r="1">' +
    '<c r="A1" t="e"><f>SUM()</f><v>#N/A</v></c>' +
    '<c r="B1"><f>1+1</f><v>2</v></c>' +
    '</row>' +
    '<row r="2"><c r="A2"><f>_xlfn.FOO(1)</f><v>7</v></c></row>',
);

describe('a formula Excel refuses is kept in the model but never written', () => {
  it('keeps the formula visible after load', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const cell = opened.workbook.sheets[0]!.cells.get('0:0');
    expect(cell?.formula).toBe('SUM()');
    expect(cell?.savedResult).toEqual({ type: 'error', value: '#N/A' });
  });

  it('a no-op save drops it and keeps the cached value', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const result = save(opened);

    // The cell is force-marked dirty, so this save is no longer a no-op: the
    // whole point is that the original bytes must not survive untouched.
    expect(result.noOp).toBe(false);
    const xml = worksheetText(result.bytes);
    expect(writtenFormulas(result.bytes)).toEqual(['1+1', '_xlfn.FOO(1)']);
    // Excel's own repair keeps `<v>` and drops `<f>`; so do we.
    expect(xml).toContain('<c r="A1" t="e"><v>#N/A</v></c>');
  });

  it('an unrelated edit drops it too', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 10,
      col: 0,
      value: textValue('probe'),
      formula: null,
    });
    expect(result.changes).toHaveLength(1);

    const saved = save(opened, session.dirtyCellMap() as Map<string, Set<string>>);
    expect(writtenFormulas(saved.bytes)).toEqual(['1+1', '_xlfn.FOO(1)']);
  });

  it('writes a valid and an unsupported-function formula verbatim', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const xml = worksheetText(save(opened).bytes);
    expect(xml).toContain('<f>1+1</f>');
    expect(xml).toContain('<f>_xlfn.FOO(1)</f>');
  });

  it('leaves a workbook without one byte-identical on a no-op save', () => {
    const original = buildXlsx(ws('<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>'));
    const result = save(openWorkbook(original));
    expect(result.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(result.bytes), Buffer.from(original))).toBe(0);
  });
});

describe('both reports name the cell', () => {
  it('the open report lists it with sheet and anchor, and says it will not be written', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const report = buildCompatibilityReport({
      workbook: opened.workbook,
      compatibility: opened.compatibility,
    });
    const items = report.items.filter((item) => item.kind === 'refusedFormula');
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({
      sheetName: 'Sheet1',
      anchor: 'A1',
      title: 'Formula A1',
    });
    expect(items[0]!.detail).toContain('SUM()');
    expect(items[0]!.detail).toContain('does not write it to the file');
    expect(report.hasContent).toBe(true);
  });

  it('the export report warns, with the formula text, sheet and reference', () => {
    const opened = openWorkbook(buildXlsx(CONTAMINATED));
    const saved = save(opened);
    const report = buildExportReport({
      workbook: opened.workbook,
      before: opened.package,
      after: saved.package,
      report: saved.report,
      compatibility: opened.compatibility,
      noOp: saved.noOp,
    });
    const warnings = report.entries.filter((entry) => entry.kind === 'refusedFormula');
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({
      category: 'warning',
      sheetName: 'Sheet1',
      anchor: 'A1',
    });
    expect(warnings[0]!.detail).toContain('SUM()');
    expect(report.needsAttention).toBe(true);
    expect(report.text).toContain('Refused formula');
  });

  it('reports nothing when no formula is refused', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>')));
    const report = buildCompatibilityReport({
      workbook: opened.workbook,
      compatibility: opened.compatibility,
    });
    expect(report.items.filter((item) => item.kind === 'refusedFormula')).toEqual([]);
    expect(report.hasContent).toBe(false);
  });
});
