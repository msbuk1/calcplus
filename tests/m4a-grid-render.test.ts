import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { openBlankWorkbook } from '@ooxml/blank';
import { openWorkbook } from '@ooxml/open';
import { borderEdge } from '@model/style';
import { columnWidthCharsToPixels, pointsToPixels } from '@view/geometry';
import { createSheetProjection } from '@view/projection';
import type { MergeSpan, ProjectedCell, ProjectedStyle, SheetProjection } from '@view/types';
import { autofitColumnWidth, autofitRowHeight } from '@grid/autofit';
import { mergedBorderStyle } from '@grid/merged-borders';
import { resolveNavigation } from '@grid/navigation';
import { collapseSelection } from '@grid/selection';
import {
  fitFontSize,
  fontFor,
  fontSizePixels,
  horizontalFor,
  rotationFor,
  wrapTextLines,
} from '@grid/text-layout';
import { formattedSheetXlsx } from './helpers/formatted-sheet';

const measure = (text: string): number => text.length * 7;

function formattedProjection(): SheetProjection {
  const opened = openWorkbook(formattedSheetXlsx());
  return createSheetProjection(opened.workbook, opened.initialSheetId, { version: 1 });
}

function styleAt(projection: SheetProjection, row: number, col: number): ProjectedStyle {
  const cell = projection.cellAt(row, col);
  if (cell === null) throw new Error(`Missing cell ${row}:${col}`);
  return cell.style;
}

describe('M4A formatted grid projection and sizing', () => {
  it('collapses imported hidden rows and columns in the geometry', () => {
    const projection = formattedProjection();
    // Row 6 (index 5) and column F (index 5) are hidden in the fixture.
    expect(projection.isRowHidden(5)).toBe(true);
    expect(projection.rowHeight(5)).toBe(0);
    expect(projection.rowY(6)).toBe(projection.rowY(5));
    expect(projection.isColumnHidden(5)).toBe(true);
    expect(projection.columnWidth(5)).toBe(0);
    expect(projection.columnX(6)).toBe(projection.columnX(5));
  });

  it('skips imported hidden rows and columns when navigating', () => {
    const projection = formattedProjection();
    const down = resolveNavigation({
      key: 'ArrowDown',
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
      selection: collapseSelection({ row: 4, col: 0 }),
      projection,
      pageRows: 20,
      pageCols: 10,
    });
    expect(down.kind).toBe('move');
    if (down.kind !== 'move') throw new Error('expected a move');
    expect(down.selection.focus).toEqual({ row: 6, col: 0 });

    const right = resolveNavigation({
      key: 'ArrowRight',
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
      selection: collapseSelection({ row: 0, col: 4 }),
      projection,
      pageRows: 20,
      pageCols: 10,
    });
    expect(right.kind).toBe('move');
    if (right.kind !== 'move') throw new Error('expected a move');
    expect(right.selection.focus).toEqual({ row: 0, col: 6 });
  });

  it('never navigates to or clicks a hidden boundary row or column', () => {
    const opened = openBlankWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    sheet.rowDimensions.set(0, { hidden: true });
    sheet.rowDimensions.set(9, { hidden: true });
    sheet.columnDimensions.set(0, { hidden: true });
    sheet.columnDimensions.set(9, { hidden: true });
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      rowCount: 10,
      columnCount: 10,
    });
    const move = (
      key: string,
      address: { readonly row: number; readonly col: number },
    ): { readonly row: number; readonly col: number } => {
      const outcome = resolveNavigation({
        key,
        ctrlKey: false,
        shiftKey: false,
        altKey: false,
        selection: collapseSelection(address),
        projection,
        pageRows: 4,
        pageCols: 4,
      });
      if (outcome.kind !== 'move') throw new Error('expected a move');
      return outcome.selection.focus;
    };

    // The hidden boundary rows are collapsed; the arrow from the visible
    // edge stays on the visible edge instead of selecting the hidden row.
    expect(move('ArrowUp', { row: 1, col: 1 })).toEqual({ row: 1, col: 1 });
    expect(move('ArrowDown', { row: 8, col: 1 })).toEqual({ row: 8, col: 1 });
    expect(move('ArrowLeft', { row: 1, col: 1 })).toEqual({ row: 1, col: 1 });
    expect(move('ArrowRight', { row: 1, col: 8 })).toEqual({ row: 1, col: 8 });
    expect(move('Enter', { row: 8, col: 1 })).toEqual({ row: 8, col: 1 });
    expect(move('Tab', { row: 1, col: 8 })).toEqual({ row: 1, col: 8 });

    // A pointer at the collapsed boundary maps to the first/last visible row
    // and column, never the hidden one.
    expect(projection.rowAt(0)).toBe(1);
    expect(projection.columnAt(0)).toBe(1);
    expect(projection.rowAt(projection.totalHeight() - 1)).toBe(8);
    expect(projection.columnAt(projection.totalWidth() - 1)).toBe(8);
  });

  it('uses imported sheet defaults, row heights, column widths and freeze panes', () => {
    const projection = formattedProjection();
    expect(projection.freezeRows).toBe(2);
    expect(projection.freezeCols).toBe(1);
    expect(projection.rowHeight(0)).toBe(40);
    expect(projection.columnWidth(0)).toBe(173);
    // sheetFormatPr carries defaultRowHeight="15" and defaultColWidth="9".
    expect(projection.rowHeight(11)).toBe(20);
    expect(projection.columnWidth(6)).toBe(68);
  });

  it('projects every imported formatting property ready to paint', () => {
    const projection = formattedProjection();
    expect(styleAt(projection, 4, 0).alignment.shrinkToFit).toBe(true);
    expect(styleAt(projection, 4, 1).alignment.textRotation).toBe(45);
    expect(styleAt(projection, 4, 0).alignment.wrapText).toBe(false);
    expect(styleAt(projection, 3, 0).alignment.wrapText).toBe(true);
    expect(styleAt(projection, 3, 0).alignment.vertical).toBe('top');
    expect(styleAt(projection, 4, 2).font.underlineStyle).toBe('single');
    expect(styleAt(projection, 4, 3).font.underlineStyle).toBe('double');
    expect(styleAt(projection, 4, 3).font.strike).toBe(true);
    expect(styleAt(projection, 6, 0).fill.backgroundColor?.toLowerCase()).toBe('#ffffff');
    expect(styleAt(projection, 0, 0).border.left.weight).toBe('thin');
    expect(styleAt(projection, 8, 1).numberFormatColor?.toLowerCase()).toBe('#ff0000');
    expect(styleAt(projection, 8, 1).numberFormat).toBe('[Red]0.00');
    expect(projection.cellAt(10, 0)?.merge).toEqual({
      anchorRow: 10,
      anchorCol: 0,
      rowSpan: 1,
      colSpan: 4,
    });
  });

  it('takes every merged-range border edge from the cell that stores it', () => {
    const cells = new Map<string, ProjectedCell>();
    const put = (row: number, col: number, border: ProjectedStyle['border']): void => {
      cells.set(`${row}:${col}`, { style: { border } } as unknown as ProjectedCell);
    };
    const none = borderEdge('none', null);
    put(0, 0, {
      left: borderEdge('thin', null),
      right: none,
      top: borderEdge('double', null),
      bottom: none,
      diagonal: none,
      diagonalUp: false,
      diagonalDown: false,
    });
    // The right edge sits on the top-right cell, the bottom edge on the
    // bottom-left cell; the anchor keeps neither.
    put(0, 1, {
      left: none,
      right: borderEdge('medium', null),
      top: none,
      bottom: none,
      diagonal: none,
      diagonalUp: false,
      diagonalDown: false,
    });
    put(1, 0, {
      left: none,
      right: none,
      top: none,
      bottom: borderEdge('thick', null),
      diagonal: none,
      diagonalUp: false,
      diagonalDown: false,
    });
    put(1, 1, {
      left: none,
      right: none,
      top: none,
      bottom: none,
      diagonal: none,
      diagonalUp: false,
      diagonalDown: false,
    });
    const projection = {
      cellAt: (row: number, col: number) => cells.get(`${row}:${col}`) ?? null,
    } as unknown as SheetProjection;
    const merge: MergeSpan = { anchorRow: 0, anchorCol: 0, rowSpan: 2, colSpan: 2 };
    const border = mergedBorderStyle(projection, merge);
    expect(border.left.style).toBe('thin');
    expect(border.top.style).toBe('double');
    expect(border.right.style).toBe('medium');
    expect(border.bottom.style).toBe('thick');
  });

  it('measures autofit sizes for text and wrapped content', () => {
    const projection = formattedProjection();
    const column = autofitColumnWidth(projection, 0, measure);
    expect(column.pixels).toBeGreaterThan(projection.columnWidth(0));
    expect(column.modelValue).toBeGreaterThan(24);
    // D4 is indented and wrapped. The indent belongs to the left inset only;
    // the right side keeps the ordinary padding, so the fitted width is the
    // longest word plus one indent and two paddings, not two indents.
    expect(autofitColumnWidth(projection, 3, measure).pixels).toBe(100);
    // Row 4 (index 3) holds wrapped text and an imported custom height.
    const row = autofitRowHeight(projection, 3, measure);
    expect(row.pixels).toBeGreaterThan(20);
    expect(row.modelValue).toBeGreaterThan(15);
    expect(autofitRowHeight(projection, 2, measure).pixels).toBe(20);
  });

  it('stores an autofit result as an undoable dimension through the command', () => {
    const opened = openBlankWorkbook();
    const session = new DocumentSession(opened.workbook);
    const rowHeight = autofitRowHeight(formattedProjection(), 3, measure).modelValue;
    const rowChange = session.dispatch('dimension.set', {
      sheetId: opened.initialSheetId,
      axis: 'row',
      index: 3,
      size: rowHeight,
      custom: true,
    });
    expect(rowChange.changes).toHaveLength(1);
    const columnChange = session.dispatch('dimension.set', {
      sheetId: opened.initialSheetId,
      axis: 'column',
      index: 2,
      size: 20,
      custom: true,
      bestFit: true,
    });
    expect(columnChange.changes).toHaveLength(1);
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 2,
    });
    expect(projection.rowHeight(3)).toBe(pointsToPixels(rowHeight));
    expect(projection.rowDimension(3)?.customHeight).toBe(true);
    expect(projection.columnWidth(2)).toBe(columnWidthCharsToPixels(20));
    expect(projection.columnDimension(2)?.bestFit).toBe(true);
    expect(projection.columnDimension(2)?.customWidth).toBe(true);
    session.undo();
    session.undo();
    const restored = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 3,
    });
    expect(restored.rowHeight(3)).toBe(20);
    expect(restored.columnWidth(2)).toBe(64);
  });
});

describe('M4A text layout', () => {
  it('maps font weight, style and size to a canvas font', () => {
    const projection = formattedProjection();
    const header = styleAt(projection, 0, 0);
    const font = fontFor(header);
    expect(font).toContain('bold');
    expect(font).toContain('"Calibri"');
    expect(font).toContain(`${fontSizePixels(header)}px`);
  });

  it('maps Excel value alignment and rotation codes', () => {
    const projection = formattedProjection();
    const number = projection.cellAt(9, 4);
    const text = projection.cellAt(2, 0);
    expect(number).not.toBeNull();
    expect(text).not.toBeNull();
    expect(horizontalFor(styleAt(projection, 9, 4), number!.value)).toBe('right');
    expect(horizontalFor(styleAt(projection, 2, 0), text!.value)).toBe('left');
    expect(rotationFor(styleAt(projection, 4, 1))).toEqual({
      kind: 'angle',
      radians: -Math.PI / 4,
    });
    const base = styleAt(projection, 2, 0);
    const withRotation = (degrees: number): ProjectedStyle => ({
      ...base,
      alignment: { ...base.alignment, textRotation: degrees },
    });
    expect(rotationFor(withRotation(135))).toEqual({ kind: 'angle', radians: Math.PI / 4 });
    expect(rotationFor(withRotation(255))).toEqual({ kind: 'stacked' });
    expect(rotationFor(withRotation(0))).toEqual({ kind: 'none' });
  });

  it('wraps at word boundaries and shrinks to fit', () => {
    const projection = formattedProjection();
    const lines = wrapTextLines('alpha beta gamma delta', 70, measure);
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.every((line) => measure(line) <= 70)).toBe(true);
    expect(lines.join(' ')).toBe('alpha beta gamma delta');

    const style = styleAt(projection, 4, 0);
    const text = projection.cellAt(4, 0)?.displayText ?? '';
    const scaledMeasure = (value: string, font: string): number => {
      const size = Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? fontSizePixels(style));
      return value.length * size * 0.5;
    };
    const fitted = fitFontSize(style, text, 40, scaledMeasure);
    expect(fitted).toBeLessThan(fontSizePixels(style));
    expect(fitted).toBeGreaterThanOrEqual(6);
  });
});
