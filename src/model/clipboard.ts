import type { Cell, GridRange, Sheet, SheetId, StyleId } from './types';
import { getCell } from './types';
import { cloneCellValue, type CellValue } from './value';
import {
  MAX_COLUMN,
  MAX_ROW,
  findAreaReferences,
  formatAreaReference,
  type AreaReference,
} from './references';

/**
 * In-app clipboard snapshot (clipboard §6).
 *
 * The OS clipboard carries values and best-effort formulas; this snapshot is
 * the lossless in-app path. It stores only the authored cell state a paste
 * owns: value, formula and style id. Rich text, shared formulas and array
 * groups are not authored constructs here; a cell that uses them is copied as
 * its displayed value so the paste never invents a malformed formula group.
 */
export type ClipboardMode = 'copy' | 'cut';

export interface ClipboardCell {
  readonly value: CellValue;
  readonly formula: string | null;
  readonly styleId: StyleId;
}

export interface ClipboardSnapshot {
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly mode: ClipboardMode;
  /**
   * Source cells in row-major order. `null` is an empty cell. The rectangle is
   * complete, so pasting blank source cells clears the matching target cells.
   */
  readonly cells: readonly (readonly (ClipboardCell | null)[])[];
  /** Merges wholly inside the copied rectangle, in source coordinates. */
  readonly merges: readonly GridRange[];
}

/** The paste budget keeps one transaction inside the performance envelope. */
export const MAX_CLIPBOARD_CELLS = 250_000;

export function clipboardCellCount(range: GridRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

export function captureClipboard(
  sheet: Sheet,
  range: GridRange,
  mode: ClipboardMode,
): ClipboardSnapshot {
  const cells: (ClipboardCell | null)[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const cellsInRow: (ClipboardCell | null)[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      cellsInRow.push(captureCell(getCell(sheet, row, col)));
    }
    cells.push(cellsInRow);
  }
  const merges = sheet.merges
    .filter(
      (merge) =>
        merge.range.startRow >= range.startRow &&
        merge.range.startCol >= range.startCol &&
        merge.range.endRow <= range.endRow &&
        merge.range.endCol <= range.endCol,
    )
    .map((merge) => ({ ...merge.range }));
  return {
    sheetId: sheet.id,
    range: { ...range },
    mode,
    cells,
    merges,
  };
}

function captureCell(cell: Cell | null): ClipboardCell | null {
  if (cell === null) return null;
  return {
    value: cloneCellValue(cell.value),
    // Shared and array formulas are read-only in this milestone. Copying their
    // cached value is safer than copying a group whose master and followers are
    // not part of the pasted rectangle.
    formula: cell.formulaSource === undefined ? cell.formula : null,
    styleId: cell.styleId,
  };
}

/**
 * Shift the relative A1 references in a copied formula by the paste offset.
 * Absolute (`$`) components stay put, and a reference that leaves the grid
 * becomes `#REF!`, as Excel's copy does. A cut is a move, so it passes a zero
 * offset and keeps the authored formula unchanged.
 */
export function translateFormulaForPaste(
  formula: string,
  rowDelta: number,
  colDelta: number,
): string {
  if (formula.length === 0 || (rowDelta === 0 && colDelta === 0)) return formula;
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  let output = '';
  let cursor = 0;
  for (const reference of references) {
    output += formula.slice(cursor, reference.start);
    output += translateReference(reference, rowDelta, colDelta);
    cursor = reference.end;
  }
  return output + formula.slice(cursor);
}

function translateReference(reference: AreaReference, rowDelta: number, colDelta: number): string {
  if (reference.kind === 'columns') {
    if (colDelta === 0) return formatAreaReference(reference);
    const startCol = reference.startColAbs ? reference.startCol : reference.startCol + colDelta;
    const endCol = reference.endColAbs ? reference.endCol : reference.endCol + colDelta;
    if (startCol < 0 || endCol < 0 || startCol >= MAX_COLUMN || endCol >= MAX_COLUMN) {
      return '#REF!';
    }
    return formatAreaReference({ ...reference, startCol, endCol, single: false });
  }
  if (reference.kind === 'rows') {
    if (rowDelta === 0) return formatAreaReference(reference);
    const startRow = reference.startRowAbs ? reference.startRow : reference.startRow + rowDelta;
    const endRow = reference.endRowAbs ? reference.endRow : reference.endRow + rowDelta;
    if (startRow < 0 || endRow < 0 || startRow >= MAX_ROW || endRow >= MAX_ROW) return '#REF!';
    return formatAreaReference({ ...reference, startRow, endRow, single: false });
  }
  const startRow = reference.startRowAbs ? reference.startRow : reference.startRow + rowDelta;
  const endRow = reference.endRowAbs ? reference.endRow : reference.endRow + rowDelta;
  const startCol = reference.startColAbs ? reference.startCol : reference.startCol + colDelta;
  const endCol = reference.endColAbs ? reference.endCol : reference.endCol + colDelta;
  if (
    startRow < 0 ||
    endRow < 0 ||
    startCol < 0 ||
    endCol < 0 ||
    startRow >= MAX_ROW ||
    endRow >= MAX_ROW ||
    startCol >= MAX_COLUMN ||
    endCol >= MAX_COLUMN
  ) {
    return '#REF!';
  }
  return formatAreaReference({ ...reference, startRow, endRow, startCol, endCol });
}
