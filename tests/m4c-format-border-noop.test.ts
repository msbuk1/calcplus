import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { borderEdge } from '@model/style';
import type { CellAddress, GridRange, Sheet, StyleRecord } from '@model/types';
import { createCell, getCell, setCellSnapshot } from '@model/types';
import { BLANK, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * Ticket 35: a Format Cells border preset that changes no edge must leave the
 * cell completely untouched. A border change that does apply must change only
 * the border part, and one gesture stays one transaction.
 */

function blankSession() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet, session: new DocumentSession(workbook) };
}

function rangeOf(row: number, col: number, endRow = row, endCol = col): GridRange {
  return { startRow: row, startCol: col, endRow, endCol };
}

function contextFor(
  sheetId: string,
  range: GridRange,
  activeCell: CellAddress = { row: range.startRow, col: range.startCol },
) {
  return { activeSheetId: sheetId, selection: range, activeCell };
}

function styleAt(session: DocumentSession, sheet: Sheet, row: number, col: number): StyleRecord {
  const cell = getCell(sheet, row, col);
  if (cell === null) throw new Error(`expected a cell at ${row}:${col}`);
  const style = session.workbook.styles.find((candidate) => candidate.id === cell.styleId);
  if (style === undefined) throw new Error(`expected style ${cell.styleId} to exist`);
  return style;
}

function styleIdAt(sheet: Sheet, row: number, col: number): number | null {
  return getCell(sheet, row, col)?.styleId ?? null;
}

/** Give the whole range a font, fill, alignment and number format. */
function styleContent(
  session: DocumentSession,
  sheetId: string,
  range: GridRange,
  context: ReturnType<typeof contextFor>,
): void {
  session.dispatch(
    'format.font',
    { sheetId, range, name: 'Georgia', size: 14, bold: true, color: '#1a73e8' },
    context,
  );
  session.dispatch('format.fill', { sheetId, range, pattern: 'solid', color: '#fff2cc' }, context);
  session.dispatch(
    'format.alignment',
    { sheetId, range, horizontal: 'center', wrapText: true },
    context,
  );
  session.dispatch('format.numberFormat', { sheetId, range, code: '0.00' }, context);
}

describe('Ticket 35 border presets with no edge change', () => {
  it('Inside on a single cell writes nothing and keeps the whole style', () => {
    const { workbook, sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);
    styleContent(session, sheet.id, range, context);
    session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, edges: { top: borderEdge('medium', '#0000ff') } },
      context,
    );

    const styleIdBefore = styleIdAt(sheet, 0, 0);
    const styleBefore = styleAt(session, sheet, 0, 0);
    const styleSnapshot = JSON.parse(JSON.stringify(styleBefore)) as StyleRecord;
    const styleCountBefore = workbook.styles.length;
    const transactionsBefore = session.transactionCount;

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'inside', style: 'thick', color: '#00ff00' },
      context,
    );

    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(transactionsBefore);
    expect(workbook.styles).toHaveLength(styleCountBefore);
    expect(styleIdAt(sheet, 0, 0)).toBe(styleIdBefore);
    expect(styleAt(session, sheet, 0, 0)).toEqual(styleSnapshot);
  });

  it('Outline on the interior cells of a selection writes nothing for them', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0, 2, 2);
    const context = contextFor(sheet.id, range);
    styleContent(session, sheet.id, range, context);

    const interiorId = styleIdAt(sheet, 1, 1);
    const interiorStyle = JSON.parse(JSON.stringify(styleAt(session, sheet, 1, 1))) as StyleRecord;
    const idsBefore = [
      [styleIdAt(sheet, 0, 0), styleIdAt(sheet, 0, 1), styleIdAt(sheet, 0, 2)],
      [styleIdAt(sheet, 1, 0), styleIdAt(sheet, 1, 1), styleIdAt(sheet, 1, 2)],
      [styleIdAt(sheet, 2, 0), styleIdAt(sheet, 2, 1), styleIdAt(sheet, 2, 2)],
    ];
    const transactionsBefore = session.transactionCount;

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'outline', style: 'thick', color: '#00ff00' },
      context,
    );

    // One gesture is one transaction, and only the edge cells change.
    expect(result.changes).toHaveLength(1);
    expect(session.transactionCount).toBe(transactionsBefore + 1);
    expect(styleIdAt(sheet, 1, 1)).toBe(interiorId);
    expect(styleAt(session, sheet, 1, 1)).toEqual(interiorStyle);
    expect(styleAt(session, sheet, 0, 0).border.left.style).toBe('thick');
    expect(styleAt(session, sheet, 0, 0).font).toEqual(interiorStyle.font);
    expect(styleAt(session, sheet, 0, 0).fill).toEqual(interiorStyle.fill);

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(transactionsBefore);
    expect([
      [styleIdAt(sheet, 0, 0), styleIdAt(sheet, 0, 1), styleIdAt(sheet, 0, 2)],
      [styleIdAt(sheet, 1, 0), styleIdAt(sheet, 1, 1), styleIdAt(sheet, 1, 2)],
      [styleIdAt(sheet, 2, 0), styleIdAt(sheet, 2, 1), styleIdAt(sheet, 2, 2)],
    ]).toEqual(idsBefore);
  });

  it('a mixed selection writes only the cells whose edges change, in one undo', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0, 1, 1);
    const context = contextFor(sheet.id, range);
    styleContent(session, sheet.id, range, context);
    session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'all', style: 'thick', color: '#000000' },
      context,
    );

    // One cell loses its border, so the next outline only changes that cell.
    const lone = rangeOf(1, 1);
    session.dispatch(
      'format.border',
      { sheetId: sheet.id, range: lone, preset: 'none' },
      contextFor(sheet.id, lone),
    );

    const idsBefore = [
      [styleIdAt(sheet, 0, 0), styleIdAt(sheet, 0, 1)],
      [styleIdAt(sheet, 1, 0), styleIdAt(sheet, 1, 1)],
    ];
    const transactionsBefore = session.transactionCount;

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'outline', style: 'thick', color: '#000000' },
      context,
    );

    expect(result.changes).toHaveLength(1);
    expect(session.transactionCount).toBe(transactionsBefore + 1);
    expect(styleIdAt(sheet, 0, 0)).toBe(idsBefore[0]![0]);
    expect(styleIdAt(sheet, 0, 1)).toBe(idsBefore[0]![1]);
    expect(styleIdAt(sheet, 1, 0)).toBe(idsBefore[1]![0]);
    expect(styleIdAt(sheet, 1, 1)).not.toBe(idsBefore[1]![1]);
    expect(styleAt(session, sheet, 1, 1).border.bottom.style).toBe('thick');
    expect(styleAt(session, sheet, 1, 1).border.right.style).toBe('thick');

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(transactionsBefore);
    expect([
      [styleIdAt(sheet, 0, 0), styleIdAt(sheet, 0, 1)],
      [styleIdAt(sheet, 1, 0), styleIdAt(sheet, 1, 1)],
    ]).toEqual(idsBefore);
  });

  it('a border change that applies keeps font, fill, alignment and number format', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);
    styleContent(session, sheet.id, range, context);

    const before = styleAt(session, sheet, 0, 0);
    const borderBefore = JSON.parse(JSON.stringify(before.border)) as StyleRecord['border'];

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, edges: { top: borderEdge('medium', '#0000ff') } },
      context,
    );

    expect(result.changes).toHaveLength(1);
    const after = styleAt(session, sheet, 0, 0);
    expect(after.font).toEqual(before.font);
    expect(after.fill).toEqual(before.fill);
    expect(after.alignment).toEqual(before.alignment);
    expect(after.numFmtId).toBe(before.numFmtId);
    expect(after.numFmtCode).toBe(before.numFmtCode);
    expect(after.numFmtKind).toBe(before.numFmtKind);
    expect(after.border.top).toEqual({ style: 'medium', color: '#0000FF', weight: 'medium' });
    expect(after.border.left).toEqual(borderBefore.left);
    expect(after.border.right).toEqual(borderBefore.right);
    expect(after.border.bottom).toEqual(borderBefore.bottom);
    expect(after.border.diagonal).toEqual(borderBefore.diagonal);
  });

  it('a no-op preset keeps an imported style the table already duplicates', () => {
    // Excel files often carry several xf records that resolve to the same
    // appearance. Solving the patch must not swap the cell to an equal-looking
    // record from elsewhere in the table: that changes the cell's own style
    // object (and its source XML) without any visible edit.
    const { workbook, sheet, session } = blankSession();
    workbook.styles.push({
      ...workbook.styles[0]!,
      id: 1,
      rawXml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    });
    const range = rangeOf(0, 0);
    setCellSnapshot(sheet, 0, 0, createCell(textValue('keep me'), 1));
    const context = contextFor(sheet.id, range);
    const styleCountBefore = workbook.styles.length;

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'none' },
      context,
    );

    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
    expect(workbook.styles).toHaveLength(styleCountBefore);
    expect(getCell(sheet, 0, 0)?.styleId).toBe(1);
  });

  it('a no-op border write does not create a cell record for a blank cell', () => {
    const { sheet, session } = blankSession();
    setCellSnapshot(sheet, 0, 0, createCell(BLANK));
    const range = rangeOf(0, 0);

    const result = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, preset: 'none' },
      contextFor(sheet.id, range),
    );

    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
    expect(getCell(sheet, 0, 0)).toEqual(createCell(BLANK));
  });
});
