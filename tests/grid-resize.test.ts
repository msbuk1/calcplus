import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { openBlankWorkbook } from '@ooxml/blank';
import { createSheetProjection } from '@view/projection';
import { columnWidthCharsToPixels } from '@view/geometry';
import {
  computePaneLayout,
  headerResizeTarget,
  headerSelectionIndices,
  headerSelectionRange,
} from '@grid/panes';
import { resizeDimension } from '@grid/resize';
import { selectionForRange } from '@grid/selection';

function blankProjection() {
  const opened = openBlankWorkbook();
  return {
    opened,
    projection: createSheetProjection(opened.workbook, opened.initialSheetId, { version: 1 }),
  };
}

describe('grid header resize', () => {
  it('maps a pointer delta to the model unit at the live zoom', () => {
    expect(resizeDimension('column', 64, 70, 1)).toEqual({ pixels: 134, modelValue: 129 / 7 });
    expect(resizeDimension('column', 64, 70, 2)).toEqual({ pixels: 99, modelValue: 94 / 7 });
    expect(resizeDimension('row', 20, 40, 1)).toEqual({ pixels: 60, modelValue: 45 });
    expect(resizeDimension('row', 20, 40, 2)).toEqual({ pixels: 40, modelValue: 30 });
  });

  it('keeps a dragged row or column visible inside the Excel limits', () => {
    expect(resizeDimension('column', 64, -1_000_000, 1).pixels).toBe(8);
    expect(resizeDimension('column', 64, 1_000_000, 1).pixels).toBe(1790);
    expect(resizeDimension('row', 20, -1_000_000, 1).pixels).toBe(4);
    expect(resizeDimension('row', 20, 1_000_000, 1).pixels).toBe(545);
  });

  it('finds a header edge only inside the resize gutter', () => {
    const { projection } = blankProjection();
    const layout = computePaneLayout(projection, 800, 480);
    expect(headerResizeTarget(projection, layout, layout.headerWidth + 64, 10, 0, 0)).toEqual({
      axis: 'column',
      index: 0,
    });
    expect(headerResizeTarget(projection, layout, layout.headerWidth + 59, 10, 0, 0)).toBeNull();
    expect(headerResizeTarget(projection, layout, 10, layout.headerHeight + 20, 0, 0)).toEqual({
      axis: 'row',
      index: 0,
    });
    expect(headerResizeTarget(projection, layout, 10, layout.headerHeight + 15, 0, 0)).toBeNull();
  });

  it('previews a live dimension without mutating the model or unhiding a row', () => {
    const opened = openBlankWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    sheet.rowDimensions.set(2, { hidden: true });
    sheet.columnDimensions.set(2, { hidden: true });
    const columnPreview = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      dimensionPreview: { axis: 'column', indices: [0, 2, 4], size: 20 },
    });
    expect(columnPreview.columnWidth(0)).toBe(columnWidthCharsToPixels(20));
    expect(columnPreview.columnWidth(2)).toBe(0);
    expect(columnPreview.columnWidth(4)).toBe(columnWidthCharsToPixels(20));
    expect(columnPreview.columnDimension(0)).toBeNull();
    const rowPreview = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      dimensionPreview: { axis: 'row', indices: [2], size: 30 },
    });
    expect(rowPreview.rowHeight(2)).toBe(0);
  });

  it('reads a contiguous multi-header selection as rows or columns', () => {
    const { opened, projection } = blankProjection();
    const columns = selectionForRange({
      startRow: 0,
      startCol: 1,
      endRow: projection.rowCount - 1,
      endCol: 3,
    });
    expect(headerSelectionRange(projection, columns, 'column')).not.toBeNull();
    expect(headerSelectionIndices(projection, columns, 'column')).toEqual([1, 2, 3]);
    expect(headerSelectionIndices(projection, columns, 'row')).toBeNull();

    const rows = selectionForRange({
      startRow: 2,
      startCol: 0,
      endRow: 4,
      endCol: projection.columnCount - 1,
    });
    expect(headerSelectionIndices(projection, rows, 'row')).toEqual([2, 3, 4]);
    expect(headerSelectionIndices(projection, rows, 'column')).toBeNull();

    // Hidden members stay skipped; a cached projection sees the flag.
    opened.workbook.sheets[0]!.columnDimensions.set(2, { hidden: true });
    const withHidden = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 2,
    });
    expect(headerSelectionIndices(withHidden, columns, 'column')).toEqual([1, 3]);
  });

  it('changes selected dimensions in one undo transaction', () => {
    const opened = openBlankWorkbook();
    const session = new DocumentSession(opened.workbook);
    const outcome = session.dispatch('dimension.set', {
      sheetId: opened.initialSheetId,
      axis: 'column',
      entries: [
        { index: 0, size: 10 },
        { index: 1, size: 20 },
        { index: 2, size: 30 },
      ],
      custom: true,
      bestFit: false,
    });
    expect(outcome.changes).toHaveLength(3);
    expect(session.transactionCount).toBe(1);
    const next = createSheetProjection(opened.workbook, opened.initialSheetId, { version: 2 });
    expect(next.columnWidth(0)).toBe(columnWidthCharsToPixels(10));
    expect(next.columnWidth(1)).toBe(columnWidthCharsToPixels(20));
    expect(next.columnWidth(2)).toBe(columnWidthCharsToPixels(30));
    expect(next.columnDimension(2)?.bestFit).toBe(false);

    session.undo();
    expect(session.transactionCount).toBe(0);
    const restored = createSheetProjection(opened.workbook, opened.initialSheetId, { version: 3 });
    expect(restored.columnWidth(0)).toBe(64);
    expect(restored.columnWidth(1)).toBe(64);
    expect(restored.columnWidth(2)).toBe(64);
  });
});
