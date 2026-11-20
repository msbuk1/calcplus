import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import type { GridRange, Sheet } from '@model/types';
import { cellKey } from '@model/types';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { unzipPart } from './helpers/build-xlsx';

/**
 * Ticket 32 — multi-part structural saves and the reference-update policy.
 *
 * A structural edit is one transaction and one save. Every reference that can
 * follow the edit moves with it on every sheet; a reference whose target is
 * gone collapses to `#REF!` with one notice, and the save writes every changed
 * worksheet part, not only the edited sheet's.
 */

function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));
}

function range(row: number, col: number): GridRange {
  return { startRow: row, startCol: col, endRow: row, endCol: col };
}

function dispatch(session: DocumentSession, sheet: Sheet, id: string, selection: GridRange) {
  return session.dispatch(
    id,
    {},
    {
      activeSheetId: sheet.id,
      selection,
      activeCell: { row: selection.startRow, col: selection.startCol },
    },
  );
}

function save(opened: ReturnType<typeof openWorkbook>, session: DocumentSession) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
    dirtyDimensions: session.dirtyDimensionSnapshot(),
    dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
  });
}

function sheetByName(opened: ReturnType<typeof openWorkbook>, name: string): Sheet {
  const sheet = opened.workbook.sheets.find((candidate) => candidate.name === name);
  if (sheet === undefined) throw new Error(`Missing sheet ${name}`);
  return sheet;
}

describe('ticket 32 cross-sheet links survive insert and delete', () => {
  it('moves formulas, a name, hyperlinks and chart series in one transaction and one save', () => {
    const opened = openWorkbook(fixtureBytes('cross-sheet-link.xlsx'));
    const session = new DocumentSession(opened.workbook);
    const sheet1 = sheetByName(opened, 'Sheet1');
    const sheet2 = sheetByName(opened, 'Sheet2');
    const links = sheetByName(opened, 'Links');
    const beforeSheet2Source = session.workbook.sourceParts.get(sheet2.partPath);
    const beforeLinksSource = session.workbook.sourceParts.get(links.partPath);

    const result = dispatch(session, sheet1, 'structure.insertRows', range(2, 0));
    expect(result.structure?.state).toBe('applied');
    expect(result.ui).toBeUndefined();
    expect(session.transactionCount).toBe(1);

    expect(sheet2.cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B4');
    expect(sheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('Sheet1!A4');
    expect(links.hyperlinks.get(cellKey(1, 1))?.location).toBe('Sheet1!A4');
    expect(opened.workbook.names.find((name) => name.name === 'TargetCell')?.formula).toBe(
      'Sheet1!$A$4',
    );

    const saved = save(opened, session);
    expect(saved.report.patchedParts).toContain('xl/worksheets/sheet3.xml');
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet2.xml')).toContain(
      '<c r="A2"><f>Sheet1!B4</f>',
    );
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet2.xml')).toContain('location="Sheet1!A4"');
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet3.xml')).toContain('location="Sheet1!A4"');
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toContain('Sheet1!$A$4');
    const chart = unzipPart(saved.bytes, 'xl/charts/chart1.xml');
    expect(chart).toContain('Sheet1!$A$4');
    expect(chart).toContain('Sheet1!$B$4');
    expect(chart).not.toContain('Sheet1!$A$3');

    const reopened = openWorkbook(saved.bytes);
    const reopenedSheet2 = sheetByName(reopened, 'Sheet2');
    expect(reopenedSheet2.cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B4');
    expect(reopenedSheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('Sheet1!A4');
    expect(sheetByName(reopened, 'Links').hyperlinks.get(cellKey(1, 1))?.location).toBe(
      'Sheet1!A4',
    );
    expect(reopened.workbook.names.find((name) => name.name === 'TargetCell')?.formula).toBe(
      'Sheet1!$A$4',
    );

    // Undo restores every part byte exactly and redo applies it again.
    session.dispatch('edit.undo', {}, { activeSheetId: sheet1.id });
    expect(sheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('Sheet1!A3');
    expect(links.hyperlinks.get(cellKey(1, 1))?.location).toBe('Sheet1!A3');
    expect(session.workbook.sourceParts.get(sheet2.partPath)).toEqual(beforeSheet2Source);
    expect(session.workbook.sourceParts.get(links.partPath)).toEqual(beforeLinksSource);

    session.dispatch('edit.redo', {}, { activeSheetId: sheet1.id });
    expect(sheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('Sheet1!A4');
    expect(session.workbook.sourceParts.get(links.partPath)).not.toEqual(beforeLinksSource);
  });

  it('moves the same references when a delete removes a row above the target', () => {
    const opened = openWorkbook(fixtureBytes('cross-sheet-link.xlsx'));
    const session = new DocumentSession(opened.workbook);
    const sheet1 = sheetByName(opened, 'Sheet1');
    const sheet2 = sheetByName(opened, 'Sheet2');
    const links = sheetByName(opened, 'Links');

    const result = dispatch(session, sheet1, 'structure.deleteRows', range(0, 0));
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);

    expect(sheet2.cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B2');
    expect(sheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('Sheet1!A2');
    expect(links.hyperlinks.get(cellKey(1, 1))?.location).toBe('Sheet1!A2');
    expect(opened.workbook.names.find((name) => name.name === 'TargetCell')?.formula).toBe(
      'Sheet1!$A$2',
    );

    const saved = save(opened, session);
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet3.xml')).toContain('location="Sheet1!A2"');
    const chart = unzipPart(saved.bytes, 'xl/charts/chart1.xml');
    expect(chart).toContain('Sheet1!$A$2');
    expect(chart).toContain('Sheet1!$B$2');
    expect(chart).not.toContain('Sheet1!$A$3');

    const reopened = openWorkbook(saved.bytes);
    expect(sheetByName(reopened, 'Sheet2').cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B2');
    expect(sheetByName(reopened, 'Links').hyperlinks.get(cellKey(1, 1))?.location).toBe(
      'Sheet1!A2',
    );
  });

  it('breaks every link whose target the delete removes and reports one notice', () => {
    const opened = openWorkbook(fixtureBytes('cross-sheet-link.xlsx'));
    const session = new DocumentSession(opened.workbook);
    const sheet1 = sheetByName(opened, 'Sheet1');
    const sheet2 = sheetByName(opened, 'Sheet2');
    const links = sheetByName(opened, 'Links');

    const result = dispatch(session, sheet1, 'structure.deleteRows', range(2, 0));
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);
    expect(result.ui?.message).toContain('now read #REF!');

    expect(sheet2.cells.get(cellKey(1, 0))?.formula).toBe('#REF!');
    expect(sheet2.cells.get(cellKey(1, 0))?.stale).toBe(true);
    expect(sheet2.hyperlinks.get(cellKey(2, 1))?.location).toBe('#REF!');
    expect(links.hyperlinks.get(cellKey(1, 1))?.location).toBe('#REF!');
    expect(opened.workbook.names.find((name) => name.name === 'TargetCell')?.formula).toBe('#REF!');

    const saved = save(opened, session);
    const sheet2Xml = unzipPart(saved.bytes, 'xl/worksheets/sheet2.xml');
    expect(sheet2Xml).toContain('<f>#REF!</f>');
    expect(sheet2Xml).toContain('location="#REF!"');
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet3.xml')).toContain('location="#REF!"');
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toContain('>#REF!<');
    const chart = unzipPart(saved.bytes, 'xl/charts/chart1.xml');
    expect(chart).toContain('<c:f>#REF!</c:f>');
    expect(chart).not.toContain('Sheet1!$A$3');

    const reopened = openWorkbook(saved.bytes);
    expect(sheetByName(reopened, 'Sheet2').cells.get(cellKey(1, 0))?.formula).toBe('#REF!');
    expect(sheetByName(reopened, 'Sheet2').hyperlinks.get(cellKey(2, 1))?.location).toBe('#REF!');
    expect(sheetByName(reopened, 'Links').hyperlinks.get(cellKey(1, 1))?.location).toBe('#REF!');
    expect(reopened.workbook.names.find((name) => name.name === 'TargetCell')?.formula).toBe(
      '#REF!',
    );
  });
});

describe('ticket 32 broken internal link convention', () => {
  it('opens and saves an existing broken location byte-identically', () => {
    const original = fixtureBytes('broken-internal-link.xlsx');
    const opened = openWorkbook(original);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(result.noOp).toBe(true);
    expect(Buffer.from(result.bytes).equals(Buffer.from(original))).toBe(true);
  });

  it('keeps a broken location marked when a later structural edit applies', () => {
    const opened = openWorkbook(fixtureBytes('broken-internal-link.xlsx'));
    const session = new DocumentSession(opened.workbook);
    const sheet = sheetByName(opened, 'Sheet1');
    expect(sheet.hyperlinks.get(cellKey(2, 1))?.location).toBe('#REF!');

    const result = dispatch(session, sheet, 'structure.insertRows', range(0, 0));
    expect(result.structure?.state).toBe('applied');
    expect(sheet.hyperlinks.get(cellKey(3, 1))?.location).toBe('#REF!');

    const saved = save(opened, session);
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain('location="#REF!"');
    const reopened = openWorkbook(saved.bytes);
    const reopenedSheet = sheetByName(reopened, 'Sheet1');
    expect(reopenedSheet.hyperlinks.get(cellKey(3, 1))?.location).toBe('#REF!');
  });
});
