import { describe, expect, it } from 'vitest';
import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import type { EngineBatchResult, EngineCellInput, EngineResultCell } from '@engine/port';
import { BLANK, numberValue } from '@model/value';

/**
 * Boundary behaviour of the batched `load` path.
 *
 * `load` screens each cell and then writes rectangular blocks instead of one
 * engine call per cell. The blocks must write exactly the listed cells: the
 * worker sends the snapshot in chunks that accumulate into one engine. These
 * cases pin the behaviours the batching must preserve.
 */
describe('Engine load batching', () => {
  function loadAll(cells: readonly EngineCellInput[]): EngineBatchResult {
    return loadAllInSequence([cells]);
  }

  function byPosition(cells: readonly EngineResultCell[]): EngineResultCell[] {
    return [...cells].sort((left, right) => left.row - right.row || left.col - right.col);
  }

  it('keeps formulas, cached literals and skipped state identical to the per-cell path', () => {
    const result = loadAll([
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(2) },
      { sheet: 0, row: 0, col: 1, formula: 'A1*3', value: BLANK },
      // Fails the engine parse after screening: keep the cached literal and
      // mark the cell skipped, never send the text to the engine.
      { sheet: 0, row: 0, col: 2, formula: '1e', value: numberValue(41) },
      // Stale input: displayed from its cached literal, never recalculated.
      { sheet: 0, row: 0, col: 3, formula: 'A1*5', value: numberValue(11), staleInput: true },
      // Depends on the stale input: computed from the literal but marked stale.
      { sheet: 0, row: 1, col: 3, formula: 'D1+1', value: BLANK },
    ]);

    expect(result.cells).toEqual([
      { sheet: 0, row: 0, col: 1, value: numberValue(6), status: 'value', stale: false },
      { sheet: 0, row: 1, col: 3, value: numberValue(12), status: 'value', stale: true },
    ]);
    expect(result.skipped).toEqual([
      { sheet: 0, row: 0, col: 2 },
      { sheet: 0, row: 0, col: 3 },
    ]);
    expect(result.circular).toEqual([]);
  });

  it('accumulates chunked loads without touching cells from earlier chunks', () => {
    const cells: EngineCellInput[] = [];
    for (let row = 0; row < 12; row += 1) {
      cells.push({ sheet: 0, row, col: 0, formula: null, value: numberValue(row + 1) });
      cells.push({ sheet: 0, row, col: 5, formula: `A${row + 1}*2`, value: BLANK });
    }
    // Two interleaved chunks: the second holds rows the first skips, so any
    // block that wrote holes would erase part of the first chunk.
    const first = cells.filter((cell) => cell.row % 3 === 0);
    const second = cells.filter((cell) => cell.row % 3 !== 0);

    const chunked = loadAllInSequence([first, second]);
    const single = loadAll(cells);
    expect(byPosition(chunked.cells)).toEqual(byPosition(single.cells));
    expect(chunked.skipped).toEqual(single.skipped);
    expect(chunked.circular).toEqual(single.circular);
    expect(chunked.cells).toHaveLength(12);
  });

  it('falls back to the per-cell path when one cell is out of bounds', () => {
    const lastRow = 1_048_576;
    const result = loadAll([
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(2) },
      { sheet: 0, row: 0, col: 1, formula: 'A1*3', value: BLANK },
      // Beyond the engine sheet size: the engine refuses every write attempt
      // for it, so it keeps its cached-literal state and nothing throws.
      { sheet: 0, row: lastRow, col: 0, formula: 'A1*3', value: numberValue(7) },
      { sheet: 0, row: lastRow, col: 1, formula: null, value: numberValue(8) },
    ]);

    expect(result.cells).toEqual([
      { sheet: 0, row: 0, col: 1, value: numberValue(6), status: 'value', stale: false },
    ]);
    expect(result.skipped).toEqual([{ sheet: 0, row: lastRow, col: 0 }]);
    expect(result.circular).toEqual([]);
  });

  it('writes the last listing of a repeated address', () => {
    const formulaLast = loadAll([
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(2) },
      { sheet: 0, row: 0, col: 0, formula: '1+2', value: BLANK },
    ]);
    expect(formulaLast.cells).toEqual([
      { sheet: 0, row: 0, col: 0, value: numberValue(3), status: 'value', stale: false },
    ]);

    const literalLast = loadAll([
      { sheet: 0, row: 0, col: 0, formula: '1+2', value: BLANK },
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(2) },
    ]);
    expect(literalLast.cells).toEqual([]);
    expect(literalLast.skipped).toEqual([]);
  });
});

function loadAllInSequence(chunks: readonly (readonly EngineCellInput[])[]): EngineBatchResult {
  const engine = new HyperFormulaEngine();
  engine.start({ date1904: false, sheetNames: ['Sheet1'] });
  for (const chunk of chunks) engine.load(chunk);
  const result = engine.finishLoad();
  engine.dispose();
  return result;
}
