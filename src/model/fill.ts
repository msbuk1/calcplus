import { translateFormulaForPaste } from './clipboard';
import type { Cell, CellAddress, GridRange, Sheet, SheetId } from './types';
import { cellStateEquals, cloneCell, getCell } from './types';
import { BLANK, cloneCellValue } from './value';

/**
 * Fill handle semantics (feature matrix §1: "Fill handle: copy,
 * relative-reference adjustment, double-click to adjacent data").
 *
 * v1 fills by **copying the source block**. There are no series and no custom
 * lists; those are v2. The source rectangle is tiled over the destination
 * rectangle, anchored at the source's own top-left corner, exactly like Excel's
 * plain drag-fill of a block:
 *
 * - a one-cell source fills the whole destination with that cell;
 * - a two-row source repeats two rows at a time, so the fifth filled row is the
 *   source's first row again;
 * - a fill upward repeats the block upward, so the row above a two-row source
 *   is the source's **bottom** row (the tile, not the destination, is the
 *   anchor);
 * - every copied formula shifts its relative row and column parts by the
 *   destination cell's distance from the source cell it came from. Absolute and
 *   mixed anchors stay put, and a reference that leaves the grid becomes
 *   `#REF!` — the copy/paste translator is reused verbatim so fill and paste
 *   cannot disagree.
 *
 * The plan is pure: it reads the sheet and returns the exact before/after
 * records of one transaction. Nothing is mutated. The command layer turns each
 * band into one `cell.set` patch; the whole gesture stays one undo entry.
 */

/** One fill is one bounded transaction, like a paste (command-system §5). */
export const MAX_FILL_CELLS = 250_000;

/** Excel's block message for a fill that would split or cover a merged range. */
export const MERGED_FILL_MESSAGE = 'You cannot change part of a merged cell.';

export type FillBlockCode = 'merged' | 'invalid-range' | 'too-large';

/** Why a fill produced no transaction. */
export type FillStopReason = 'no-destination' | 'nothing-changed';

/**
 * One destination cell the fill writes.
 *
 * `before` is the cell at `address` before the fill and `after` is the copied
 * cell. `source` names the cell it was copied from, so the log can replay the
 * block wrap without re-running the modulo.
 */
export interface FillCellRecord {
  readonly address: CellAddress;
  readonly source: CellAddress;
  readonly before: Cell | null;
  readonly after: Cell | null;
}

/** A destination row rewritten at `address`: the authored formula before and after. */
export interface FillFormulaRewrite {
  readonly address: CellAddress;
  readonly before: string | null;
  readonly after: string;
}

/**
 * One rectangular patch of the filled region. The destination minus the source
 * is at most four disjoint rectangles (above, below, left and right of the
 * source); each becomes one `cell.set` patch, so the source cells are never
 * touched and never marked dirty. `before`/`after` are row-major.
 */
export interface FillBand {
  readonly range: GridRange;
  readonly before: readonly (readonly (Cell | null)[])[];
  readonly after: readonly (readonly (Cell | null)[])[];
}

export interface FillPlan {
  readonly sheetId: SheetId;
  readonly source: GridRange;
  /** The whole target rectangle, source included. */
  readonly destination: GridRange;
  readonly bands: readonly FillBand[];
  readonly cellRecords: readonly FillCellRecord[];
  readonly formulaRewrites: readonly FillFormulaRewrite[];
  /** Destination cells written, not counting the source rectangle. */
  readonly cellCount: number;
  /** Written cells whose formula text changed. */
  readonly formulaCount: number;
}

export type FillResult =
  | { readonly kind: 'filled'; readonly plan: FillPlan }
  | {
      readonly kind: 'unchanged';
      readonly reason: FillStopReason;
      readonly source: GridRange;
      readonly destination: GridRange;
    }
  | {
      readonly kind: 'blocked';
      readonly code: FillBlockCode;
      readonly message: string;
      readonly source: GridRange;
      readonly destination: GridRange;
    };

export interface FillRequest {
  readonly sheet: Sheet;
  /** The block the gesture started from. */
  readonly source: GridRange;
  /** The full target rectangle; it must contain the source. */
  readonly destination: GridRange;
}

/**
 * Plan one fill. The destination is the whole rectangle the drag covers,
 * source included; every cell outside the source is written. A destination that
 * equals the source is a no-op, which is what releasing the handle without
 * moving produces.
 *
 * Merged ranges block: a merge that intersects the destination must sit wholly
 * inside the source, because v1 does not fill merges. Writing a merge's anchor
 * or covered cells, or splitting a merge that only partly overlaps the target,
 * is not a shape `.xlsx` or Excel would show, so the fill stops with Excel's
 * block message instead of writing something the file cannot represent.
 * A merge wholly inside the source is read for its anchor cell and is not
 * filled: the destination cells are plain cells.
 *
 * Hidden rows and columns in the destination are filled like any other cell
 * (Excel fills them).
 */
export function planFill(request: FillRequest): FillResult {
  const source = request.source;
  const destination = request.destination;
  if (!validRange(source) || !validRange(destination)) {
    return blocked('invalid-range', 'The fill range is outside the sheet.', source, destination);
  }
  if (!containsRange(destination, source)) {
    return blocked(
      'invalid-range',
      'The fill target must include the source range.',
      source,
      destination,
    );
  }

  const ranges = bandsFor(source, destination);
  if (ranges.length === 0)
    return { kind: 'unchanged', reason: 'no-destination', source, destination };

  if (cellCount(destination) > MAX_FILL_CELLS) {
    return blocked(
      'too-large',
      `A fill can cover at most ${MAX_FILL_CELLS.toLocaleString('en-US')} cells.`,
      source,
      destination,
    );
  }
  if (mergeBlocksFill(request.sheet, source, destination)) {
    return blocked('merged', MERGED_FILL_MESSAGE, source, destination);
  }

  const cellRecords: FillCellRecord[] = [];
  const formulaRewrites: FillFormulaRewrite[] = [];
  const bands: FillBand[] = [];
  for (const range of ranges) {
    const before: (Cell | null)[][] = [];
    const after: (Cell | null)[][] = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const beforeRow: (Cell | null)[] = [];
      const afterRow: (Cell | null)[] = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const origin = sourceCellFor(source, row, col);
        const existing = getCell(request.sheet, row, col);
        const copied = fillCell(getCell(request.sheet, origin.row, origin.col), {
          rowDelta: row - origin.row,
          colDelta: col - origin.col,
        });
        beforeRow.push(cloneCell(existing));
        afterRow.push(copied);
        cellRecords.push({
          address: { row, col },
          source: { ...origin },
          before: cloneCell(existing),
          after: cloneCell(copied),
        });
        if (copied !== null && copied.formula !== null) {
          const previous = existing?.formula ?? null;
          if (previous !== copied.formula) {
            formulaRewrites.push({
              address: { row, col },
              before: previous,
              after: copied.formula,
            });
          }
        }
      }
      before.push(beforeRow);
      after.push(afterRow);
    }
    if (rowChanges(before, after)) bands.push({ range, before, after });
  }

  if (bands.length === 0) {
    return { kind: 'unchanged', reason: 'nothing-changed', source, destination };
  }
  const written = cellRecords.filter((record) => !cellStateEquals(record.before, record.after));
  return {
    kind: 'filled',
    plan: {
      sheetId: request.sheet.id,
      source: { ...source },
      destination: { ...destination },
      bands,
      cellRecords: written,
      formulaRewrites,
      cellCount: written.length,
      formulaCount: formulaRewrites.length,
    },
  };
}

/**
 * The source cell a destination cell is copied from: the source block tiled
 * from its own top-left corner. The modulo is taken in both directions so a
 * fill above or left of the source repeats the block correctly.
 */
export function sourceCellFor(
  source: GridRange,
  row: number,
  col: number,
): { readonly row: number; readonly col: number } {
  const height = source.endRow - source.startRow + 1;
  const width = source.endCol - source.startCol + 1;
  return {
    row: source.startRow + modulo(row - source.startRow, height),
    col: source.startCol + modulo(col - source.startCol, width),
  };
}

/**
 * The cell a destination address receives. A formula is copied with its
 * relative references shifted; a shared or array group is copied as its
 * displayed value (the clipboard makes the same call, because a group whose
 * master is outside the copied block cannot be re-created safely); rich text is
 * copied as plain text, as an edit of it would be (workbook-model §9).
 */
export function fillCell(
  source: Cell | null,
  delta: { readonly rowDelta: number; readonly colDelta: number },
): Cell | null {
  if (source === null) return null;
  if (source.formulaSource !== undefined) {
    return staticCell(source.value, source.styleId);
  }
  if (source.formula !== null) {
    // The engine recalculates the new formula; the cached result is dropped so
    // a save before recalculation asks Excel to recalculate (file-compat §6.1).
    return {
      value: BLANK,
      formula: translateFormulaForPaste(source.formula, delta.rowDelta, delta.colDelta),
      formulaDerived: false,
      savedResult: null,
      styleId: source.styleId,
      richText: null,
      stale: true,
    };
  }
  return staticCell(source.value, source.styleId);
}

function staticCell(value: Cell['value'], styleId: number): Cell | null {
  if (value.type === 'blank' && styleId === 0) return null;
  return {
    value: cloneCellValue(value),
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId,
    richText: null,
    stale: false,
  };
}

/** Destination minus source, as at most four disjoint rectangles. */
function bandsFor(source: GridRange, destination: GridRange): GridRange[] {
  const bands: GridRange[] = [];
  if (destination.startRow < source.startRow) {
    bands.push({
      startRow: destination.startRow,
      endRow: source.startRow - 1,
      startCol: destination.startCol,
      endCol: destination.endCol,
    });
  }
  if (destination.endRow > source.endRow) {
    bands.push({
      startRow: source.endRow + 1,
      endRow: destination.endRow,
      startCol: destination.startCol,
      endCol: destination.endCol,
    });
  }
  if (destination.startCol < source.startCol) {
    bands.push({
      startRow: source.startRow,
      endRow: source.endRow,
      startCol: destination.startCol,
      endCol: source.startCol - 1,
    });
  }
  if (destination.endCol > source.endCol) {
    bands.push({
      startRow: source.startRow,
      endRow: source.endRow,
      startCol: source.endCol + 1,
      endCol: destination.endCol,
    });
  }
  return bands;
}

/**
 * A merge may sit wholly inside the source (it is read, not written). Every
 * other merge that touches the destination blocks the fill.
 */
function mergeBlocksFill(sheet: Sheet, source: GridRange, destination: GridRange): boolean {
  for (const merge of sheet.merges) {
    if (!rangesIntersect(merge.range, destination)) continue;
    if (!containsRange(source, merge.range)) return true;
  }
  return false;
}

function rowChanges(
  before: readonly (readonly (Cell | null)[])[],
  after: readonly (readonly (Cell | null)[])[],
): boolean {
  for (let row = 0; row < before.length; row += 1) {
    const beforeRow = before[row] ?? [];
    const afterRow = after[row] ?? [];
    for (let col = 0; col < beforeRow.length; col += 1) {
      if (!cellStateEquals(beforeRow[col] ?? null, afterRow[col] ?? null)) return true;
    }
  }
  return false;
}

function modulo(value: number, size: number): number {
  return ((value % size) + size) % size;
}

function rangesIntersect(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow <= right.endRow &&
    right.startRow <= left.endRow &&
    left.startCol <= right.endCol &&
    right.startCol <= left.endCol
  );
}

function containsRange(outer: GridRange, inner: GridRange): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startCol >= outer.startCol &&
    inner.endCol <= outer.endCol
  );
}

function validRange(range: GridRange): boolean {
  return (
    Number.isInteger(range.startRow) &&
    Number.isInteger(range.startCol) &&
    Number.isInteger(range.endRow) &&
    Number.isInteger(range.endCol) &&
    range.startRow >= 0 &&
    range.startCol >= 0 &&
    range.endRow >= range.startRow &&
    range.endCol >= range.startCol &&
    range.endRow < 1_048_576 &&
    range.endCol < 16_384
  );
}

export function cellCount(range: GridRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

function blocked(
  code: FillBlockCode,
  message: string,
  source: GridRange,
  destination: GridRange,
): FillResult {
  return { kind: 'blocked', code, message, source, destination };
}
