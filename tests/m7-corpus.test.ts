import { describe, expect, it } from 'vitest';
import {
  CONTRACT_CORPUS,
  DEGRADED_1M,
  DENSE_FORMULA,
  FAN_OUT,
  MANY_SHEETS,
  PATHOLOGICAL,
  TINY,
  TYPICAL,
  WIDE,
  envelopeFor,
  formulaCells,
  generateShape,
  sheetCount,
  touchedCells,
} from '../bench/corpus/generate';
import { buildCorpusWorkbook } from '../bench/model/corpus';

/**
 * The committed corpus generator (performance §5 and acceptance criterion 11).
 *
 * The shapes are reproducible from the fixed seed, the counts match the
 * contract table, and the engine builder really produces those formulas. The
 * browser harness covers the paint side; this pins the shape definitions so a
 * benchmark cannot silently measure something else.
 */
describe('benchmark corpus', () => {
  it('lists the ten contract shapes in table order', () => {
    expect(CONTRACT_CORPUS.map((shape) => shape.name)).toEqual([
      'tiny',
      'typical',
      'dense-formula',
      'fan-out',
      'wide',
      'many-sheets',
      'styling-heavy',
      'cf-heavy',
      'degraded-1m',
      'pathological',
    ]);
    expect(new Set(CONTRACT_CORPUS.map((shape) => shape.seed)).size).toBeGreaterThan(1);
  });

  it('carries the touched-cell and formula counts the contract names', () => {
    expect([touchedCells(TINY), formulaCells(TINY)]).toEqual([1_000, 0]);
    expect([touchedCells(TYPICAL), formulaCells(TYPICAL)]).toEqual([50_100, 10_000]);
    expect(sheetCount(TYPICAL)).toBe(3);
    expect([touchedCells(DENSE_FORMULA), formulaCells(DENSE_FORMULA)]).toEqual([100_000, 50_000]);
    expect([touchedCells(FAN_OUT), formulaCells(FAN_OUT)]).toEqual([100_500, 100_000]);
    expect([touchedCells(WIDE), formulaCells(WIDE)]).toEqual([100_000, 0]);
    expect([touchedCells(MANY_SHEETS), formulaCells(MANY_SHEETS)]).toEqual([100_000, 5_000]);
    expect(sheetCount(MANY_SHEETS)).toBe(20);
    expect([touchedCells(DEGRADED_1M), formulaCells(DEGRADED_1M)]).toEqual([1_000_000, 200_000]);
    expect([touchedCells(PATHOLOGICAL), formulaCells(PATHOLOGICAL)]).toEqual([201_000, 200_000]);
  });

  it('places each shape in the documented envelope', () => {
    expect(envelopeFor(TINY)).toBe('guaranteed');
    expect(envelopeFor(TYPICAL)).toBe('guaranteed');
    expect(envelopeFor(DENSE_FORMULA)).toBe('guaranteed');
    expect(envelopeFor(WIDE)).toBe('guaranteed');
    expect(envelopeFor(FAN_OUT)).toBe('degraded');
    expect(envelopeFor(DEGRADED_1M)).toBe('degraded');
    expect(envelopeFor(PATHOLOGICAL)).toBe('degraded');
  });

  it('resolves a shape by name and rejects an unknown one', () => {
    expect(generateShape('fan-out')).toBe(FAN_OUT);
    expect(() => generateShape('nope')).toThrow(/Unknown corpus shape/);
  });

  it('builds the fan-out shape with every formula reading one cell', () => {
    const corpus = buildCorpusWorkbook(FAN_OUT);
    expect(corpus.cells).toHaveLength(touchedCells(FAN_OUT));
    expect(corpus.formulas).toBe(100_000);
    const formulas = corpus.cells.filter((cell) => cell.formula !== null);
    expect(formulas).toHaveLength(100_000);
    expect(formulas.every((cell) => cell.formula === '$A$1*2')).toBe(true);
    const anchor = corpus.cells.find(
      (cell) => cell.sheet === 0 && cell.row === 0 && cell.col === 0,
    );
    expect(anchor?.formula).toBeNull();
  });

  it('builds the pathological shape with whole-column references and no full-grid allocation', () => {
    const corpus = buildCorpusWorkbook(PATHOLOGICAL);
    expect(corpus.cells).toHaveLength(touchedCells(PATHOLOGICAL));
    expect(corpus.formulas).toBe(200_000);
    const formulas = corpus.cells.filter((cell) => cell.formula !== null);
    expect(formulas.every((cell) => cell.formula === 'SUM($A:$A)')).toBe(true);
    // The inputs are the touched cells only: 201,000, not the 1,048,576-row grid.
    expect(corpus.cells.length).toBeLessThan(202_000);
  });

  it('builds a sparse shape without materialising the million rows', () => {
    const corpus = buildCorpusWorkbook(generateShape('sparse-1m'));
    expect(corpus.cells).toHaveLength(100);
    expect(corpus.formulas).toBe(0);
  });

  it('builds the same workbook twice from the fixed seed', () => {
    const first = buildCorpusWorkbook(TYPICAL);
    const second = buildCorpusWorkbook(TYPICAL);
    expect(second.cells.length).toBe(first.cells.length);
    expect(second.cells[0]).toEqual(first.cells[0]);
    expect(second.cells[second.cells.length - 1]).toEqual(first.cells[first.cells.length - 1]);
    expect(second.formulas).toBe(first.formulas);
  });
});
