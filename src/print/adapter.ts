import { DEFAULT_FONT } from '@model/style';
import type { Workbook } from '@model/types';
import { printAreaOf } from '@ooxml/print-setup-xml';
import type { MergeSpan, ProjectedCell, SheetProjection } from '@view/types';
import type { PrintCell, PrintSheetSnapshot, PrintTextMeasure } from './types';

/** CSS pixels at 96 dpi to points. */
const PT_PER_PIXEL = 0.75;

/**
 * The display-record adapter: the one thin mapper between `SheetView` and the
 * pure print planner (print-pdf §4, M7 boundary rule 3).
 *
 * It adds no state and owns no layout. Every fact print needs is read from the
 * projection (value, display text, style, merge, hyperlink, width-aware number
 * format) or from the sheet's print settings, exactly as the grid reads them.
 * Sizes are taken at the projection's display scale and converted back to
 * sheet points, so a print built from a zoomed view is the same page.
 */
export function createPrintSnapshot(
  workbook: Workbook,
  projection: SheetProjection,
): PrintSheetSnapshot {
  const sheet = workbook.sheets.find((candidate) => candidate.id === projection.sheetId);
  if (sheet === undefined) throw new Error(`Unknown sheet: ${projection.sheetId}`);
  const pixel = PT_PER_PIXEL / projection.zoom;
  const headingFont = workbook.styles[0]?.font ?? DEFAULT_FONT;

  return {
    sheetId: projection.sheetId,
    sheetName: projection.sheetName,
    printAreas: printAreaOf(sheet),
    usedRange: projection.usedRange(),
    printSetup: sheet.printSetup ?? null,
    rowHeight: (row) => projection.rowHeight(row) * pixel,
    columnWidth: (col) => projection.columnWidth(col) * pixel,
    columnLabel: (col) => projection.columnLabel(col),
    rowLabel: (row) => projection.rowLabel(row),
    headingFont,
    cellAt(row, col) {
      const cell = projection.cellAt(row, col);
      return cell === null ? null : toPrintCell(cell);
    },
    cellsIn(startRow, endRow, startCol, endCol) {
      const cells: PrintCell[] = [];
      const firstRow = Math.max(0, Math.min(startRow, endRow));
      const lastRow = Math.max(startRow, endRow);
      for (let row = firstRow; row <= lastRow; row += 1) {
        for (const col of projection.rowDataColumns(row)) {
          if (col < startCol || col > endCol) continue;
          const cell = projection.cellAt(row, col);
          if (cell !== null) cells.push(toPrintCell(cell));
        }
      }
      return cells;
    },
    mergesIn(startRow, endRow, startCol, endCol) {
      const spans: MergeSpan[] = [];
      for (const merge of sheet.merges) {
        const range = merge.range;
        if (range.endRow < startRow || range.startRow > endRow) continue;
        if (range.endCol < startCol || range.startCol > endCol) continue;
        spans.push({
          anchorRow: range.startRow,
          anchorCol: range.startCol,
          rowSpan: range.endRow - range.startRow + 1,
          colSpan: range.endCol - range.startCol + 1,
        });
      }
      return spans;
    },
  };
}

/**
 * A `ProjectedCell` as print reads it. The width-aware layout is re-expressed
 * in sheet points: the projection measures in CSS pixels, and both sides of
 * every comparison scale by the same 4/3, so the result is the same `#`
 * overflow marker and the same `*` fill the grid paints.
 */
function toPrintCell(cell: ProjectedCell): PrintCell {
  return {
    row: cell.row,
    col: cell.col,
    value: cell.value,
    displayText: cell.displayText,
    style: cell.style,
    hyperlink: cell.hyperlink !== null,
    merge: cell.merge,
    layoutText(width: number, measure: PrintTextMeasure) {
      const font = {
        family: cell.style.font.name,
        size: cell.style.font.size,
        bold: cell.style.font.bold,
        italic: cell.style.font.italic,
        color: cell.style.font.color ?? '#202020',
      };
      return cell.layoutText(width / PT_PER_PIXEL, (text) => measure(text, font) / PT_PER_PIXEL);
    },
  };
}
