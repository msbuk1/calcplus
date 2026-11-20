/**
 * The benchmark corpus (performance §5).
 *
 * The corpus is generated, never committed: a shape is a description that the
 * harness turns into a sheet. The seed is fixed so every run measures the same
 * workbook. The ten contract shapes cover startup, everyday workbooks, recalc
 * throughput, the fan-out cliff, horizontal scrolling, per-sheet overhead,
 * style lookup, conditional formatting, the degraded envelope and the
 * pathological ceiling. The M2 scroll shapes stay alongside them.
 */
export const CORPUS_SEED = 20_261_113;

export type CorpusFormulaPattern =
  'none' | 'row-local' | 'fan-out-single' | 'mixed' | 'whole-column';
export type CorpusEnvelope = 'guaranteed' | 'degraded';

export interface CorpusShape {
  readonly name: string;
  /** Rows per sheet. */
  readonly rows: number;
  readonly cols: number;
  readonly seed: number;
  /** How many distinct styles the sheet cycles through. */
  readonly styleCount: number;
  /** When set, only every nth row holds a cell: a sparse sheet. */
  readonly touchedEvery?: number;
  /** Sheet count; omitted means one sheet. */
  readonly sheets?: number;
  /** Formula cells in the whole workbook. */
  readonly formulas?: number;
  readonly formulaPattern?: CorpusFormulaPattern;
  /** Conditional-formatting rules in the whole workbook. */
  readonly cfRules?: number;
  readonly comments?: number;
  readonly hyperlinks?: number;
  readonly merges?: number;
  /** The performance §2 envelope this shape exercises. */
  readonly envelope?: CorpusEnvelope;
  /** What the shape is for. */
  readonly catches: string;
}

// ---------------------------------------------------------------------------
// The ten contract shapes (performance §5).
// ---------------------------------------------------------------------------

export const TINY: CorpusShape = Object.freeze({
  name: 'tiny',
  rows: 40,
  cols: 25,
  seed: CORPUS_SEED,
  styleCount: 1,
  formulas: 0,
  formulaPattern: 'none',
  envelope: 'guaranteed',
  catches: 'startup and smoke',
});

export const TYPICAL: CorpusShape = Object.freeze({
  name: 'typical',
  rows: 167,
  cols: 100,
  seed: CORPUS_SEED + 1,
  styleCount: 200,
  sheets: 3,
  formulas: 10_000,
  formulaPattern: 'mixed',
  comments: 1,
  hyperlinks: 1,
  merges: 1,
  envelope: 'guaranteed',
  catches: 'everyday workbooks',
});

export const DENSE_FORMULA: CorpusShape = Object.freeze({
  name: 'dense-formula',
  rows: 1_000,
  cols: 100,
  seed: CORPUS_SEED + 2,
  styleCount: 6,
  formulas: 50_000,
  formulaPattern: 'row-local',
  envelope: 'guaranteed',
  catches: 'recalc throughput',
});

export const FAN_OUT: CorpusShape = Object.freeze({
  name: 'fan-out',
  rows: 500,
  cols: 201,
  seed: CORPUS_SEED + 3,
  styleCount: 1,
  formulas: 100_000,
  formulaPattern: 'fan-out-single',
  envelope: 'degraded',
  catches: 'dependency-graph build cliff',
});

export const WIDE: CorpusShape = Object.freeze({
  name: 'wide',
  rows: 2_000,
  cols: 50,
  seed: CORPUS_SEED + 4,
  styleCount: 4,
  formulas: 0,
  formulaPattern: 'none',
  envelope: 'guaranteed',
  catches: 'horizontal scrolling and column widths',
});

export const MANY_SHEETS: CorpusShape = Object.freeze({
  name: 'many-sheets',
  rows: 250,
  cols: 20,
  seed: CORPUS_SEED + 5,
  styleCount: 3,
  sheets: 20,
  formulas: 5_000,
  formulaPattern: 'row-local',
  envelope: 'guaranteed',
  catches: 'sheet switching and per-sheet overhead',
});

export const STYLING_HEAVY: CorpusShape = Object.freeze({
  name: 'styling-heavy',
  rows: 1_000,
  cols: 100,
  seed: CORPUS_SEED + 6,
  styleCount: 500,
  formulas: 0,
  formulaPattern: 'none',
  envelope: 'guaranteed',
  catches: 'style-table lookup and paint',
});

export const CF_HEAVY: CorpusShape = Object.freeze({
  name: 'cf-heavy',
  rows: 1_000,
  cols: 100,
  seed: CORPUS_SEED + 7,
  styleCount: 3,
  formulas: 0,
  formulaPattern: 'none',
  cfRules: 50,
  envelope: 'guaranteed',
  catches: 'render-only rule evaluation',
});

export const DEGRADED_1M: CorpusShape = Object.freeze({
  name: 'degraded-1m',
  rows: 2_000,
  cols: 500,
  seed: CORPUS_SEED + 8,
  styleCount: 6,
  formulas: 200_000,
  formulaPattern: 'mixed',
  envelope: 'degraded',
  catches: 'the degraded envelope',
});

export const PATHOLOGICAL: CorpusShape = Object.freeze({
  name: 'pathological',
  rows: 1_000,
  cols: 201,
  seed: CORPUS_SEED + 9,
  styleCount: 2,
  formulas: 200_000,
  formulaPattern: 'whole-column',
  envelope: 'degraded',
  catches: '200k formulas reading one cell plus whole-column refs: must finish, must not hang',
});

/** The ten names in performance §5, in table order. */
export const CONTRACT_CORPUS: readonly CorpusShape[] = Object.freeze([
  TINY,
  TYPICAL,
  DENSE_FORMULA,
  FAN_OUT,
  WIDE,
  MANY_SHEETS,
  STYLING_HEAVY,
  CF_HEAVY,
  DEGRADED_1M,
  PATHOLOGICAL,
]);

// ---------------------------------------------------------------------------
// The M2 scroll shapes.
// ---------------------------------------------------------------------------

export const SCROLL_250K: CorpusShape = Object.freeze({
  name: 'scroll-250k',
  rows: 5_000,
  cols: 50,
  seed: CORPUS_SEED,
  styleCount: 6,
  formulas: 0,
  formulaPattern: 'none',
  envelope: 'guaranteed',
  catches: 'the scroll frame budget on a 250,000 touched-cell formatted sheet',
});

export const SPARSE_1M: CorpusShape = Object.freeze({
  name: 'sparse-1m',
  rows: 1_000_000,
  cols: 1,
  seed: CORPUS_SEED + 1,
  styleCount: 1,
  formulas: 0,
  formulaPattern: 'none',
  /** Only every tenthousandth row holds a cell, so the sheet stays sparse. */
  touchedEvery: 10_000,
  catches: 'scrolling a million-row sparse sheet without allocating the grid',
});

export const CORPUS: readonly CorpusShape[] = Object.freeze([
  ...CONTRACT_CORPUS,
  SCROLL_250K,
  SPARSE_1M,
]);

export function generateShape(name: string): CorpusShape {
  const shape = CORPUS.find((candidate) => candidate.name === name);
  if (shape === undefined) throw new Error(`Unknown corpus shape: ${name}`);
  return shape;
}

export function sheetCount(shape: CorpusShape): number {
  return shape.sheets ?? 1;
}

/** The number of touched cells a shape produces across every sheet. */
export function touchedCells(shape: CorpusShape): number {
  const perSheet =
    shape.touchedEvery !== undefined
      ? Math.ceil(shape.rows / shape.touchedEvery) * shape.cols
      : shape.rows * shape.cols;
  return perSheet * sheetCount(shape);
}

/** Formula cells in the whole workbook, capped by the touched cells. */
export function formulaCells(shape: CorpusShape): number {
  return Math.min(shape.formulas ?? 0, touchedCells(shape));
}

/** The performance §2 envelope the shape belongs to, from its counts. */
export function envelopeFor(shape: CorpusShape): CorpusEnvelope {
  return touchedCells(shape) > 250_000 || formulaCells(shape) > 50_000 ? 'degraded' : 'guaranteed';
}
