/**
 * The benchmark corpus (performance §5).
 *
 * The corpus is generated, never committed: a shape is a description that the
 * browser harness turns into a sheet. The seed is fixed so every run measures
 * the same workbook.
 *
 * M2 needs the guaranteed-envelope scroll shape and a sparse sheet that proves
 * the grid never allocates the sheet itself. The remaining corpus shapes of
 * performance §5 arrive with the performance-hardening milestone.
 */
export const CORPUS_SEED = 20_261_113;

export interface CorpusShape {
  readonly name: string;
  readonly rows: number;
  readonly cols: number;
  readonly seed: number;
  /** How many distinct styles the sheet cycles through. */
  readonly styleCount: number;
  /** When set, only every nth row holds a cell: a sparse sheet. */
  readonly touchedEvery?: number;
  /** What the shape is for. */
  readonly catches: string;
}

export const SCROLL_250K: CorpusShape = Object.freeze({
  name: 'scroll-250k',
  rows: 5_000,
  cols: 50,
  seed: CORPUS_SEED,
  styleCount: 6,
  catches: 'the scroll frame budget on a 250,000 touched-cell formatted sheet',
});

export const SPARSE_1M: CorpusShape = Object.freeze({
  name: 'sparse-1m',
  rows: 1_000_000,
  cols: 1,
  seed: CORPUS_SEED + 1,
  styleCount: 1,
  /** Only every tenthousandth row holds a cell, so the sheet stays sparse. */
  touchedEvery: 10_000,
  catches: 'scrolling a million-row sparse sheet without allocating the grid',
});

export const CORPUS: readonly CorpusShape[] = Object.freeze([SCROLL_250K, SPARSE_1M]);

export function generateShape(name: string): CorpusShape {
  const shape = CORPUS.find((candidate) => candidate.name === name);
  if (shape === undefined) throw new Error(`Unknown corpus shape: ${name}`);
  return shape;
}

/** The number of touched cells a shape produces. */
export function touchedCells(shape: CorpusShape): number {
  if (shape.touchedEvery !== undefined) {
    return Math.ceil(shape.rows / shape.touchedEvery) * shape.cols;
  }
  return shape.rows * shape.cols;
}
