import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { borderEdge } from '@model/style';
import type { GridRange, StyleRecord } from '@model/types';
import { createCell, getCell, setCellSnapshot } from '@model/types';
import { textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import type { CellAddress } from '@model/types';

const SUBMISSION_IDS = [
  'format.font',
  'format.fill',
  'format.border',
  'format.alignment',
  'format.numberFormat',
  'painter.pickup',
  'painter.apply',
] as const;

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

function styleAt(
  session: DocumentSession,
  sheet: ReturnType<typeof createSheet>,
  row: number,
  col: number,
): StyleRecord {
  const cell = getCell(sheet, row, col);
  if (cell === null) throw new Error(`expected a cell at ${row}:${col}`);
  const style = session.workbook.styles.find((candidate) => candidate.id === cell.styleId);
  if (style === undefined) throw new Error(`expected style ${cell.styleId} to exist`);
  return style;
}

describe('M4C format dialog and painter command ids', () => {
  it('registers the submission ids the UI track calls', () => {
    const registry = createCommandRegistry();
    for (const id of SUBMISSION_IDS) expect(registry.require(id).id).toBe(id);
    // Esc cancels the painter by clearing the snapshot in view state; a cancel
    // command would create a document transaction and is not part of the seam.
    expect(registry.get('painter.cancel')).toBeUndefined();
  });
});

describe('M4C font dialog', () => {
  it('sets every supported font property as one transaction and reuses the style', () => {
    const { workbook, sheet, session } = blankSession();
    const range = rangeOf(0, 0, 1, 1);
    const args = {
      sheetId: sheet.id,
      range,
      name: 'Arial',
      size: 16,
      bold: true,
      italic: true,
      underlineStyle: 'double' as const,
      strike: true,
      color: '#c00000',
    };
    const result = session.dispatch('format.font', args, contextFor(sheet.id, range));

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('style.set');
    expect(session.transactionCount).toBe(1);
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]) {
      expect(styleAt(session, sheet, row!, col!).font).toEqual({
        name: 'Arial',
        size: 16,
        bold: true,
        italic: true,
        underline: true,
        underlineStyle: 'double',
        strike: true,
        color: '#C00000',
      });
    }

    const styleCount = workbook.styles.length;
    const again = session.dispatch('format.font', args, contextFor(sheet.id, range));
    expect(again.changes).toHaveLength(0);
    expect(workbook.styles).toHaveLength(styleCount);
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)).toBeNull();
    expect(getCell(sheet, 1, 1)).toBeNull();
  });

  it('maps the underline boolean and single/double variants', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);

    session.dispatch('format.font', { sheetId: sheet.id, range, underline: true }, context);
    expect(styleAt(session, sheet, 0, 0).font.underlineStyle).toBe('single');

    session.dispatch(
      'format.font',
      { sheetId: sheet.id, range, underlineStyle: 'double' },
      context,
    );
    expect(styleAt(session, sheet, 0, 0).font.underlineStyle).toBe('double');

    session.dispatch('format.font', { sheetId: sheet.id, range, underline: false }, context);
    expect(styleAt(session, sheet, 0, 0).font.underlineStyle).toBe('none');
    expect(styleAt(session, sheet, 0, 0).font.underline).toBe(false);
  });

  it('normalises a short hex colour and clears to automatic with null', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);

    session.dispatch('format.font', { sheetId: sheet.id, range, color: '#f00' }, context);
    expect(styleAt(session, sheet, 0, 0).font.color).toBe('#FF0000');

    session.dispatch('format.font', { sheetId: sheet.id, range, color: '#FF336699' }, context);
    expect(styleAt(session, sheet, 0, 0).font.color).toBe('#336699');

    session.dispatch('format.font', { sheetId: sheet.id, range, color: null }, context);
    expect(styleAt(session, sheet, 0, 0).font.color).toBeNull();
  });
});

describe('M4C fill dialog', () => {
  it('sets a solid or patterned fill with foreground and background colours', () => {
    const { sheet, session } = blankSession();
    const solid = rangeOf(0, 0);
    const patterned = rangeOf(1, 0);
    const context = contextFor(sheet.id, solid);

    session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range: solid, pattern: 'solid', color: '#fff2cc' },
      context,
    );
    expect(styleAt(session, sheet, 0, 0).fill).toEqual({
      pattern: 'solid',
      color: '#FFF2CC',
      backgroundColor: null,
    });

    session.dispatch(
      'format.fill',
      {
        sheetId: sheet.id,
        range: patterned,
        pattern: 'lightGrid',
        color: '#ffc000',
        backgroundColor: '#ffffff',
      },
      contextFor(sheet.id, patterned),
    );
    expect(styleAt(session, sheet, 1, 0).fill).toEqual({
      pattern: 'lightGrid',
      color: '#FFC000',
      backgroundColor: '#FFFFFF',
    });
    expect(session.transactionCount).toBe(2);
  });
});

describe('M4C border dialog', () => {
  it('sets only the named custom edges and leaves the other edges alone', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const result = session.dispatch(
      'format.border',
      {
        sheetId: sheet.id,
        range,
        edges: {
          top: borderEdge('medium', '#0000ff'),
          left: borderEdge('double', '#00ff00'),
        },
      },
      contextFor(sheet.id, range),
    );

    expect(result.changes).toHaveLength(1);
    const border = styleAt(session, sheet, 0, 0).border;
    expect(border.top).toEqual({ style: 'medium', color: '#0000FF', weight: 'medium' });
    expect(border.left).toEqual({ style: 'double', color: '#00FF00', weight: 'thick' });
    // An edge-only custom border must not inherit a thin default on the edges
    // the dialog did not name.
    expect(border.right).toEqual({ style: 'none', color: null, weight: null });
    expect(border.bottom).toEqual({ style: 'none', color: null, weight: null });
    expect(border.diagonal).toEqual({ style: 'none', color: null, weight: null });
  });

  it('keeps a style argument from bleeding into unnamed custom edges', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    session.dispatch(
      'format.border',
      {
        sheetId: sheet.id,
        range,
        style: 'thick',
        color: '#ff0000',
        edges: { top: borderEdge('hair', null) },
      },
      contextFor(sheet.id, range),
    );

    const border = styleAt(session, sheet, 0, 0).border;
    expect(border.top).toEqual({ style: 'hair', color: null, weight: 'hair' });
    expect(border.right.style).toBe('none');
    expect(border.bottom.style).toBe('none');
    expect(border.left.style).toBe('none');
  });

  it('applies outline, inside, all and none presets', () => {
    const outline = blankSession();
    const range = rangeOf(0, 0, 1, 1);
    outline.session.dispatch(
      'format.border',
      { sheetId: outline.sheet.id, range, preset: 'outline', style: 'medium', color: '#ff0000' },
      contextFor(outline.sheet.id, range),
    );
    const topLeft = styleAt(outline.session, outline.sheet, 0, 0).border;
    expect(topLeft.top.style).toBe('medium');
    expect(topLeft.left.style).toBe('medium');
    expect(topLeft.right.style).toBe('none');
    expect(topLeft.bottom.style).toBe('none');
    expect(styleAt(outline.session, outline.sheet, 1, 1).border.bottom.style).toBe('medium');

    const inside = blankSession();
    inside.session.dispatch(
      'format.border',
      { sheetId: inside.sheet.id, range, preset: 'inside' },
      contextFor(inside.sheet.id, range),
    );
    const insideTopLeft = styleAt(inside.session, inside.sheet, 0, 0).border;
    expect(insideTopLeft.top.style).toBe('none');
    expect(insideTopLeft.left.style).toBe('none');
    expect(insideTopLeft.right.style).toBe('thin');
    expect(insideTopLeft.bottom.style).toBe('thin');

    const all = blankSession();
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]) {
      setCellSnapshot(all.sheet, row!, col!, createCell(textValue('x')));
    }
    all.session.dispatch(
      'format.border',
      { sheetId: all.sheet.id, range, preset: 'all', style: 'dashed', color: '#000000' },
      contextFor(all.sheet.id, range),
    );
    for (const [row, col] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]) {
      const border = styleAt(all.session, all.sheet, row!, col!).border;
      expect(border.top).toEqual({ style: 'dashed', color: '#000000', weight: 'thin' });
      expect(border.right).toEqual({ style: 'dashed', color: '#000000', weight: 'thin' });
      expect(border.bottom).toEqual({ style: 'dashed', color: '#000000', weight: 'thin' });
      expect(border.left).toEqual({ style: 'dashed', color: '#000000', weight: 'thin' });
    }

    all.session.dispatch(
      'format.border',
      { sheetId: all.sheet.id, range, preset: 'none' },
      contextFor(all.sheet.id, range),
    );
    expect(styleAt(all.session, all.sheet, 0, 0).border).toEqual({
      left: { style: 'none', color: null, weight: null },
      right: { style: 'none', color: null, weight: null },
      top: { style: 'none', color: null, weight: null },
      bottom: { style: 'none', color: null, weight: null },
      diagonal: { style: 'none', color: null, weight: null },
      diagonalUp: false,
      diagonalDown: false,
    });
  });

  it('sets a style on every side when no custom edges are named', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, style: 'medium', color: '#107c41' },
      contextFor(sheet.id, range),
    );
    const border = styleAt(session, sheet, 0, 0).border;
    for (const edge of [border.left, border.right, border.top, border.bottom]) {
      expect(edge).toEqual({ style: 'medium', color: '#107C41', weight: 'medium' });
    }
  });
});

describe('M4C alignment dialog', () => {
  it('sets horizontal, vertical, indent, wrap, shrink and rotation in one transaction', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const result = session.dispatch(
      'format.alignment',
      {
        sheetId: sheet.id,
        range,
        horizontal: 'centerContinuous',
        vertical: 'distributed',
        wrapText: true,
        indent: 3,
        textRotation: 90,
        shrinkToFit: true,
      },
      contextFor(sheet.id, range),
    );

    expect(result.changes).toHaveLength(1);
    expect(styleAt(session, sheet, 0, 0).alignment).toEqual({
      horizontal: 'centerContinuous',
      vertical: 'distributed',
      wrapText: true,
      indent: 3,
      textRotation: 90,
      shrinkToFit: true,
    });
    expect(session.transactionCount).toBe(1);
  });
});

describe('M4C number-format dialog', () => {
  it('sets a built-in format id without materialising the code', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    session.dispatch(
      'format.numberFormat',
      { sheetId: sheet.id, range, code: '#,##0.00' },
      contextFor(sheet.id, range),
    );
    const style = styleAt(session, sheet, 0, 0);
    expect(style.numFmtId).toBe(4);
    expect(style.numFmtKind).toBe('builtin');
    expect(style.numFmtCode).toBe('#,##0.00');
    expect(session.transactionCount).toBe(1);
  });

  it('stores a custom format code and clears back to General', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);
    setCellSnapshot(sheet, 0, 0, createCell(textValue('42')));

    session.dispatch('format.numberFormat', { sheetId: sheet.id, range, code: '0.000' }, context);
    let style = styleAt(session, sheet, 0, 0);
    expect(style.numFmtId).toBe(0);
    expect(style.numFmtKind).toBe('custom');
    expect(style.numFmtCode).toBe('0.000');

    session.dispatch('format.numberFormat', { sheetId: sheet.id, range, code: null }, context);
    style = styleAt(session, sheet, 0, 0);
    expect(style.numFmtId).toBe(0);
    expect(style.numFmtKind).toBe('general');
    expect(style.numFmtCode).toBeNull();
  });

  it('refuses a malformed code with a toast and no transaction', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);
    const version = session.version;

    for (const code of ['0.00;0.00;0.00;@;garbage', '[NatNum1]']) {
      const result = session.dispatch(
        'format.numberFormat',
        { sheetId: sheet.id, range, code },
        context,
      );
      expect(result.changes).toHaveLength(0);
      expect(result.ui?.type).toBe('toast');
      expect(result.ui?.message).toMatch(/number format/i);
      expect(getCell(sheet, 0, 0)).toBeNull();
      expect(session.transactionCount).toBe(0);
      expect(session.version).toBe(version);
    }
  });
});

describe('M4C invalid colour input', () => {
  it('refuses a malformed colour in each formatter with a toast and no transaction', () => {
    const { sheet, session } = blankSession();
    const range = rangeOf(0, 0);
    const context = contextFor(sheet.id, range);
    const version = session.version;

    const fontResult = session.dispatch(
      'format.font',
      { sheetId: sheet.id, range, color: 'not-a-colour' },
      context,
    );
    const fillResult = session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range, pattern: 'solid', color: '#12345' },
      context,
    );
    const borderResult = session.dispatch(
      'format.border',
      { sheetId: sheet.id, range, style: 'thin', color: '#GG0000' },
      context,
    );
    const edgeResult = session.dispatch(
      'format.border',
      {
        sheetId: sheet.id,
        range,
        edges: { top: borderEdge('thin', 'blue') },
      },
      context,
    );

    for (const result of [fontResult, fillResult, borderResult, edgeResult]) {
      expect(result.changes).toHaveLength(0);
      expect(result.ui?.type).toBe('toast');
      expect(result.ui?.message).toMatch(/colour/i);
    }
    expect(getCell(sheet, 0, 0)).toBeNull();
    expect(session.transactionCount).toBe(0);
    expect(session.version).toBe(version);
  });
});

describe('M4C format painter', () => {
  function paintSource(session: DocumentSession, sheetId: string): number {
    const source = rangeOf(0, 0);
    const context = contextFor(sheetId, source);
    session.dispatch(
      'format.font',
      {
        sheetId,
        range: source,
        name: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        underlineStyle: 'double',
        strike: true,
        color: '#c00000',
      },
      context,
    );
    session.dispatch(
      'format.fill',
      { sheetId, range: source, pattern: 'solid', color: '#fff2cc' },
      context,
    );
    session.dispatch(
      'format.border',
      { sheetId, range: source, preset: 'all', style: 'medium', color: '#107c41' },
      context,
    );
    session.dispatch(
      'format.alignment',
      { sheetId, range: source, horizontal: 'center', wrapText: true, indent: 1 },
      context,
    );
    session.dispatch('format.numberFormat', { sheetId, range: source, code: '#,##0.00' }, context);
    const cell = getCell(session.workbook.sheets[0]!, 0, 0);
    if (cell === null) throw new Error('expected a painted source cell');
    return cell.styleId;
  }

  it('picks up the anchor style record without creating a transaction', () => {
    const { sheet, session } = blankSession();
    const sourceStyleId = paintSource(session, sheet.id);
    const source = rangeOf(0, 0, 1, 1);
    const context = contextFor(sheet.id, source, { row: 1, col: 1 });
    const before = session.transactionCount;

    const result = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: source },
      context,
    );

    expect(result.changes).toHaveLength(0);
    expect(result.painter?.styleId).toBe(sourceStyleId);
    expect(result.painter?.style).toEqual(
      session.workbook.styles.find((style) => style.id === sourceStyleId),
    );
    expect(result.painter?.sheetId).toBe(sheet.id);
    expect(result.painter?.range).toEqual(source);
    expect(session.transactionCount).toBe(before);
  });

  it('uses the source range anchor even when another cell is active', () => {
    const { sheet, session } = blankSession();
    const first = rangeOf(0, 0);
    const second = rangeOf(1, 1);
    session.dispatch(
      'format.font',
      { sheetId: sheet.id, range: first, color: '#ff0000' },
      contextFor(sheet.id, first),
    );
    session.dispatch(
      'format.font',
      { sheetId: sheet.id, range: second, color: '#0000ff' },
      contextFor(sheet.id, second),
    );
    const source = rangeOf(0, 0, 1, 1);
    const result = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: source },
      contextFor(sheet.id, source, { row: 1, col: 1 }),
    );
    expect(result.painter?.style.font.color).toBe('#FF0000');
  });

  it('picks up from the anchor of a whole-column selection beyond the format budget', () => {
    const { sheet, session } = blankSession();
    const sourceStyleId = paintSource(session, sheet.id);
    const wholeColumn = rangeOf(0, 0, 1_048_575, 0);
    const before = session.transactionCount;

    const result = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: wholeColumn },
      contextFor(sheet.id, wholeColumn),
    );

    expect(result.changes).toHaveLength(0);
    expect(result.painter?.styleId).toBe(sourceStyleId);
    expect(result.painter?.range).toEqual(wholeColumn);
    expect(session.transactionCount).toBe(before);
  });

  it('applies the whole style to a target range as one undo transaction', () => {
    const { workbook, sheet, session } = blankSession();
    const sourceStyleId = paintSource(session, sheet.id);
    const source = rangeOf(0, 0);
    const pickup = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: source },
      contextFor(sheet.id, source),
    );
    const snapshot = pickup.painter;
    if (snapshot === undefined) throw new Error('expected a painter snapshot');

    const target = rangeOf(3, 0, 4, 1);
    const before = session.transactionCount;
    const styleCount = workbook.styles.length;
    const result = session.dispatch(
      'painter.apply',
      { sheetId: sheet.id, range: target, snapshot },
      contextFor(sheet.id, target),
    );

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('style.set');
    expect(session.transactionCount).toBe(before + 1);
    expect(workbook.styles).toHaveLength(styleCount);
    for (const [row, col] of [
      [3, 0],
      [3, 1],
      [4, 0],
      [4, 1],
    ]) {
      expect(getCell(sheet, row!, col!)?.styleId).toBe(sourceStyleId);
    }

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(before);
    for (const [row, col] of [
      [3, 0],
      [3, 1],
      [4, 0],
      [4, 1],
    ]) {
      expect(getCell(sheet, row!, col!)).toBeNull();
    }
  });

  it('restores a formatted target cell exactly on undo', () => {
    const { sheet, session } = blankSession();
    const sourceStyleId = paintSource(session, sheet.id);
    const source = rangeOf(0, 0);
    const pickup = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: source },
      contextFor(sheet.id, source),
    );
    const snapshot = pickup.painter;
    if (snapshot === undefined) throw new Error('expected a painter snapshot');

    const target = rangeOf(5, 0);
    setCellSnapshot(sheet, 5, 0, createCell(textValue('keep me')));
    const before = getCell(sheet, 5, 0);
    session.dispatch(
      'painter.apply',
      { sheetId: sheet.id, range: target, snapshot },
      contextFor(sheet.id, target),
    );

    const painted = getCell(sheet, 5, 0);
    expect(painted?.styleId).toBe(sourceStyleId);
    expect(painted?.value).toEqual(textValue('keep me'));

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 5, 0)).toEqual(before);
  });

  it('clears target formatting when the picked-up source is the default style', () => {
    const { sheet, session } = blankSession();
    const source = rangeOf(0, 0);
    const target = rangeOf(1, 0);
    setCellSnapshot(sheet, 1, 0, createCell(textValue('target')));

    session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range: target, pattern: 'solid', color: '#fff2cc' },
      contextFor(sheet.id, target),
    );
    const pickup = session.dispatch(
      'painter.pickup',
      { sheetId: sheet.id, range: source },
      contextFor(sheet.id, source),
    );
    const snapshot = pickup.painter;
    if (snapshot === undefined) throw new Error('expected a painter snapshot');
    expect(snapshot.styleId).toBe(0);

    session.dispatch(
      'painter.apply',
      { sheetId: sheet.id, range: target, snapshot },
      contextFor(sheet.id, target),
    );
    expect(getCell(sheet, 1, 0)?.styleId).toBe(0);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 1, 0)?.styleId).not.toBe(0);
  });

  it('refuses to apply without a snapshot with a toast and no transaction', () => {
    const { sheet, session } = blankSession();
    const target = rangeOf(0, 0);
    const result = session.dispatch(
      'painter.apply',
      { sheetId: sheet.id, range: target },
      contextFor(sheet.id, target),
    );
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.type).toBe('toast');
    expect(result.ui?.message).toMatch(/pick up/i);
    expect(session.transactionCount).toBe(0);
  });
});

describe('M4C dialog Apply undo boundary', () => {
  it('records exactly one undo entry for one Apply and restores the prior style id', () => {
    const { sheet, session } = blankSession();
    setCellSnapshot(sheet, 0, 0, createCell(textValue('total')));
    const range = rangeOf(0, 0);
    const before = getCell(sheet, 0, 0);

    const result = session.dispatch(
      'format.font',
      {
        sheetId: sheet.id,
        range,
        name: 'Arial',
        size: 18,
        bold: true,
        color: '#1a73e8',
      },
      contextFor(sheet.id, range),
    );
    expect(result.changes).toHaveLength(1);
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(0);
    expect(getCell(sheet, 0, 0)).toEqual(before);
  });
});
