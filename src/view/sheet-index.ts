import type { Cell } from '@model/types';

/**
 * A lazy index of the cells that hold data.
 *
 * Navigation needs to answer "where is the next cell with data?" without
 * walking the sheet on every key press. The index is built once per document
 * version and answers by row and by column.
 */
export interface SheetIndex {
  /** Row index to the sorted columns that hold data in that row. */
  readonly rows: Map<number, number[]>;
  /** Column index to the sorted rows that hold data in that column. */
  readonly columns: Map<number, number[]>;
  readonly minRow: number;
  readonly maxRow: number;
  readonly minCol: number;
  readonly maxCol: number;
  readonly cellCount: number;
}

export function cellHoldsData(cell: Cell | undefined): boolean {
  if (cell === undefined) return false;
  if (cell.formula !== null) return true;
  return cell.value.type !== 'blank';
}

export function buildSheetIndex(cells: ReadonlyMap<string, Cell>): SheetIndex {
  const rows = new Map<number, number[]>();
  const columns = new Map<number, number[]>();
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = -1;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = -1;
  let cellCount = 0;
  for (const [key, cell] of cells) {
    if (!cellHoldsData(cell)) continue;
    const separator = key.indexOf(':');
    if (separator < 0) continue;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    cellCount += 1;
    if (row < minRow) minRow = row;
    if (row > maxRow) maxRow = row;
    if (col < minCol) minCol = col;
    if (col > maxCol) maxCol = col;
    const rowColumns = rows.get(row);
    if (rowColumns === undefined) rows.set(row, [col]);
    else rowColumns.push(col);
    const columnRows = columns.get(col);
    if (columnRows === undefined) columns.set(col, [row]);
    else columnRows.push(row);
  }
  for (const rowColumns of rows.values()) rowColumns.sort((left, right) => left - right);
  for (const columnRows of columns.values()) columnRows.sort((left, right) => left - right);
  return {
    rows,
    columns,
    minRow: Number.isFinite(minRow) ? minRow : 0,
    maxRow: maxRow < 0 ? 0 : maxRow,
    minCol: Number.isFinite(minCol) ? minCol : 0,
    maxCol: maxCol < 0 ? 0 : maxCol,
    cellCount,
  };
}
