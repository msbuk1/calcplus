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
