/**
 * The published function set (formula semantics §2).
 *
 * The engine registry is gated to exactly these 52 functions. `TRUE` and
 * `FALSE` are not counted: they are named expressions added to the engine
 * (technical design §3.5), not members of the published set.
 */

export const PUBLISHED_FUNCTIONS = [
  // Maths (10)
  'SUM',
  'PRODUCT',
  'ABS',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'INT',
  'MOD',
  'POWER',
  'SQRT',
  // Statistics (9)
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'COUNTIF',
  'SUMIF',
  'MEDIAN',
  // Logic (6)
  'IF',
  'AND',
  'OR',
  'NOT',
  'IFERROR',
  'IFNA',
  // Lookup (5)
  'VLOOKUP',
  'HLOOKUP',
  'INDEX',
  'MATCH',
  'CHOOSE',
  // Position (3)
  'ROW',
  'COLUMN',
  'ROWS',
  // Text (8)
  'CONCATENATE',
  'LEFT',
  'RIGHT',
  'MID',
  'LEN',
  'TRIM',
  'UPPER',
  'LOWER',
  // Dates (6)
  'TODAY',
  'NOW',
  'DATE',
  'DAY',
  'MONTH',
  'YEAR',
  // Checks (5)
  'ISBLANK',
  'ISNUMBER',
  'ISTEXT',
  'ISERROR',
  'ISNA',
] as const;

export type PublishedFunction = (typeof PUBLISHED_FUNCTIONS)[number];

export const PUBLISHED_FUNCTION_COUNT = PUBLISHED_FUNCTIONS.length;

/** How many arguments one function takes. */
export interface FunctionArity {
  readonly min: number;
  readonly max: number;
}

/**
 * The argument count Excel's parser accepts at entry for each published
 * function. Excel refuses a call outside this range when the user types it
 * ("There's a problem with this formula") and keeps the cell's previous value,
 * so the range doubles as the entry gate (formula semantics §2).
 *
 * Every entry was verified against Excel 365's parser, one call below the
 * minimum, at both boundaries and one above the maximum. The variable-argument
 * families cap at 255 like every Excel function. Empty arguments count as
 * arguments: `SUM(,)` is two arguments and `SUM( )` is none.
 *
 * `INDEX` is the one deliberate narrowing. Excel has two forms: the array
 * form takes 2-3 arguments and the reference form takes 3-4 (area_num). V1
 * implements the array form only (formula semantics §7, shim-plugin.ts), so
 * the range here is 2-3: a 4-argument call is refused at entry instead of
 * being accepted and evaluated as something Excel would not. Excel's parser
 * itself refuses `INDEX(array)` with one argument: at least one of row_num or
 * column_num must be present.
 *
 * `TRUE` and `FALSE` are named expressions, not calls, so they carry no entry
 * here.
 */
export const FUNCTION_ARITY: Readonly<Record<PublishedFunction, FunctionArity>> = {
  // Maths
  SUM: { min: 1, max: 255 },
  PRODUCT: { min: 1, max: 255 },
  ABS: { min: 1, max: 1 },
  ROUND: { min: 2, max: 2 },
  ROUNDUP: { min: 2, max: 2 },
  ROUNDDOWN: { min: 2, max: 2 },
  INT: { min: 1, max: 1 },
  MOD: { min: 2, max: 2 },
  POWER: { min: 2, max: 2 },
  SQRT: { min: 1, max: 1 },
  // Statistics
  AVERAGE: { min: 1, max: 255 },
  MIN: { min: 1, max: 255 },
  MAX: { min: 1, max: 255 },
  COUNT: { min: 1, max: 255 },
  COUNTA: { min: 1, max: 255 },
  COUNTBLANK: { min: 1, max: 1 },
  COUNTIF: { min: 2, max: 2 },
  SUMIF: { min: 2, max: 3 },
  MEDIAN: { min: 1, max: 255 },
  // Logic
  IF: { min: 2, max: 3 },
  AND: { min: 1, max: 255 },
  OR: { min: 1, max: 255 },
  NOT: { min: 1, max: 1 },
  IFERROR: { min: 2, max: 2 },
  IFNA: { min: 2, max: 2 },
  // Lookup
  VLOOKUP: { min: 3, max: 4 },
  HLOOKUP: { min: 3, max: 4 },
  // Array form only; the reference form (3-4 args) is v2 (see the note above).
  INDEX: { min: 2, max: 3 },
  MATCH: { min: 2, max: 3 },
  CHOOSE: { min: 2, max: 255 },
  // Position
  ROW: { min: 0, max: 1 },
  COLUMN: { min: 0, max: 1 },
  ROWS: { min: 1, max: 1 },
  // Text
  CONCATENATE: { min: 1, max: 255 },
  LEFT: { min: 1, max: 2 },
  RIGHT: { min: 1, max: 2 },
  MID: { min: 3, max: 3 },
  LEN: { min: 1, max: 1 },
  TRIM: { min: 1, max: 1 },
  UPPER: { min: 1, max: 1 },
  LOWER: { min: 1, max: 1 },
  // Dates
  TODAY: { min: 0, max: 0 },
  NOW: { min: 0, max: 0 },
  DATE: { min: 3, max: 3 },
  DAY: { min: 1, max: 1 },
  MONTH: { min: 1, max: 1 },
  YEAR: { min: 1, max: 1 },
  // Checks
  ISBLANK: { min: 1, max: 1 },
  ISNUMBER: { min: 1, max: 1 },
  ISTEXT: { min: 1, max: 1 },
  ISERROR: { min: 1, max: 1 },
  ISNA: { min: 1, max: 1 },
};

/**
 * The accepted argument count for a function name, or `null` when the name is
 * not a published function call (`TRUE`, `FALSE` and unknown names carry no
 * arity).
 */
export function functionArity(name: string): FunctionArity | null {
  const entry = (FUNCTION_ARITY as Readonly<Record<string, FunctionArity | undefined>>)[name];
  return entry ?? null;
}

/** Formula ids the engine may evaluate: the published set plus the boolean literals. */
const ALLOWED_FUNCTIONS: ReadonlySet<string> = new Set<string>([
  ...PUBLISHED_FUNCTIONS,
  'TRUE',
  'FALSE',
]);

/**
 * Engine-only shim ids. Authored text never contains them: the registry gate
 * runs first and unsupported names never reach the engine.
 */
export const SHIM_FUNCTION_IDS = {
  INT: 'CALCPLUS_INT',
  CHOOSE: 'CALCPLUS_CHOOSE',
  VLOOKUP: 'CALCPLUS_VLOOKUP',
  HLOOKUP: 'CALCPLUS_HLOOKUP',
  MATCH: 'CALCPLUS_MATCH',
  INDEX: 'CALCPLUS_INDEX',
  COUNTIF: 'CALCPLUS_COUNTIF',
  SUMIF: 'CALCPLUS_SUMIF',
  EMPTY_TEXT: 'CALCPLUS_EMPTY_TEXT',
} as const;

export function isPublishedFunction(name: string): name is PublishedFunction {
  return (PUBLISHED_FUNCTIONS as readonly string[]).includes(name);
}

export function isAllowedEngineFunction(name: string): boolean {
  return ALLOWED_FUNCTIONS.has(name);
}
