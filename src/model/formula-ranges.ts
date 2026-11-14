import type { Cell, Sheet } from './types';
import { parseRangeRef, rangeContains } from './types';

/**
 * Legacy array formulas cover their `ref` range even though only the master
 * cell carries the `<f t="array">` element. Any edit inside that range is
 * read-only in M1.
 */
export function arrayFormulaForCell(sheet: Sheet, row: number, col: number): Cell | null {
  for (const cell of sheet.cells.values()) {
    const source = cell.formulaSource;
    if (source?.kind !== 'array' || source.range === undefined) continue;
    try {
      if (rangeContains(parseRangeRef(source.range.replaceAll('$', '')), row, col)) return cell;
    } catch {
      // A malformed ref stays unmodelled; the worksheet patch preserves it.
    }
  }
  return null;
}
