import { PUBLISHED_FUNCTIONS } from './function-registry';
import type { PublishedFunction } from './function-registry';

/**
 * The Insert Function catalogue (feature matrix §4, ticket 30).
 *
 * Names and arities come from the M3B published set: the catalogue is typed
 * against `PublishedFunction`, so a missing or extra function fails the
 * compiler and the dialog can never offer a function the engine does not
 * evaluate. Categories use Excel 365's names; Financial stays in the picker
 * to match the Excel layout even though none of the 52 functions is financial.
 */

export type FunctionCategory =
  | 'Financial'
  | 'Date & Time'
  | 'Math & Trig'
  | 'Statistical'
  | 'Lookup & Reference'
  | 'Text'
  | 'Logical'
  | 'Information';

/** The Select category values, in the capture's order. */
export const FUNCTION_CATEGORY_OPTIONS = [
  'Recently used',
  'All',
  'Financial',
  'Date & Time',
  'Math & Trig',
  'Statistical',
  'Lookup & Reference',
  'Text',
  'Logical',
  'Information',
] as const;

export type FunctionCategoryOption = (typeof FUNCTION_CATEGORY_OPTIONS)[number];

export interface FunctionCatalogEntry {
  readonly name: PublishedFunction;
  readonly category: FunctionCategory;
  /** Excel's signature style: required arguments plain, optional ones bracketed. */
  readonly signature: string;
  /** One plain sentence, written from the function's contract semantics. */
  readonly description: string;
}

type FunctionDetails = Omit<FunctionCatalogEntry, 'name'>;

/**
 * The metadata for every published function. The record key type makes the
 * catalogue exhaustive: adding or removing a function in the registry without
 * updating this table is a type error.
 */
const FUNCTION_DETAILS: Readonly<Record<PublishedFunction, FunctionDetails>> = {
  // Maths (10)
  SUM: {
    category: 'Math & Trig',
    signature: 'SUM(number1, [number2], ...)',
    description: 'Adds the numbers in its arguments and returns the total.',
  },
  PRODUCT: {
    category: 'Math & Trig',
    signature: 'PRODUCT(number1, [number2], ...)',
    description: 'Multiplies its arguments and returns the product.',
  },
  ABS: {
    category: 'Math & Trig',
    signature: 'ABS(number)',
    description: 'Returns the absolute value of a number, without its sign.',
  },
  ROUND: {
    category: 'Math & Trig',
    signature: 'ROUND(number, num_digits)',
    description: 'Rounds a number to the given number of digits.',
  },
  ROUNDUP: {
    category: 'Math & Trig',
    signature: 'ROUNDUP(number, num_digits)',
    description: 'Rounds a number away from zero to the given number of digits.',
  },
  ROUNDDOWN: {
    category: 'Math & Trig',
    signature: 'ROUNDDOWN(number, num_digits)',
    description: 'Rounds a number towards zero to the given number of digits.',
  },
  INT: {
    category: 'Math & Trig',
    signature: 'INT(number)',
    description: 'Rounds a number down to the nearest integer.',
  },
  MOD: {
    category: 'Math & Trig',
    signature: 'MOD(number, divisor)',
    description: 'Returns the remainder after one number is divided by another.',
  },
  POWER: {
    category: 'Math & Trig',
    signature: 'POWER(number, power)',
    description: 'Returns a number raised to a power.',
  },
  SQRT: {
    category: 'Math & Trig',
    signature: 'SQRT(number)',
    description: 'Returns the positive square root of a number.',
  },
  // Statistics (9)
  AVERAGE: {
    category: 'Statistical',
    signature: 'AVERAGE(number1, [number2], ...)',
    description:
      'Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays or references that contain numbers.',
  },
  MIN: {
    category: 'Statistical',
    signature: 'MIN(number1, [number2], ...)',
    description: 'Returns the smallest number in a set of values.',
  },
  MAX: {
    category: 'Statistical',
    signature: 'MAX(number1, [number2], ...)',
    description: 'Returns the largest number in a set of values.',
  },
  COUNT: {
    category: 'Statistical',
    signature: 'COUNT(value1, [value2], ...)',
    description: 'Counts the cells in a range that contain numbers.',
  },
  COUNTA: {
    category: 'Statistical',
    signature: 'COUNTA(value1, [value2], ...)',
    description: 'Counts the cells in a range that are not empty.',
  },
  COUNTBLANK: {
    category: 'Statistical',
    signature: 'COUNTBLANK(range)',
    description: 'Counts the empty cells in a range.',
  },
  COUNTIF: {
    category: 'Statistical',
    signature: 'COUNTIF(range, criteria)',
    description: 'Counts the cells in a range that meet one condition.',
  },
  SUMIF: {
    category: 'Statistical',
    signature: 'SUMIF(range, criteria, [sum_range])',
    description: 'Adds the cells in a range that meet one condition.',
  },
  MEDIAN: {
    category: 'Statistical',
    signature: 'MEDIAN(number1, [number2], ...)',
    description: 'Returns the middle value of a sorted set of numbers.',
  },
  // Logic (6)
  IF: {
    category: 'Logical',
    signature: 'IF(logical_test, value_if_true, [value_if_false])',
    description: 'Returns one value when a condition is true and another when it is false.',
  },
  AND: {
    category: 'Logical',
    signature: 'AND(logical1, [logical2], ...)',
    description: 'Returns TRUE only when every condition is true.',
  },
  OR: {
    category: 'Logical',
    signature: 'OR(logical1, [logical2], ...)',
    description: 'Returns TRUE when any condition is true.',
  },
  NOT: {
    category: 'Logical',
    signature: 'NOT(logical)',
    description: 'Reverses the logical value of its argument.',
  },
  IFERROR: {
    category: 'Logical',
    signature: 'IFERROR(value, value_if_error)',
    description: 'Returns a fallback value when a result is an error, and the result otherwise.',
  },
  IFNA: {
    category: 'Logical',
    signature: 'IFNA(value, value_if_na)',
    description: 'Returns a fallback value when a result is #N/A, and the result otherwise.',
  },
  // Lookup (5)
  VLOOKUP: {
    category: 'Lookup & Reference',
    signature: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    description:
      'Looks for a value in the first column of a range and returns a value from the same row.',
  },
  HLOOKUP: {
    category: 'Lookup & Reference',
    signature: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    description:
      'Looks for a value in the first row of a range and returns a value from the same column.',
  },
  INDEX: {
    category: 'Lookup & Reference',
    signature: 'INDEX(array, row_num, [column_num])',
    description: 'Returns the value at a given row and column position in a range or array.',
  },
  MATCH: {
    category: 'Lookup & Reference',
    signature: 'MATCH(lookup_value, lookup_array, [match_type])',
    description: 'Returns the position of a value within a range or array.',
  },
  CHOOSE: {
    category: 'Lookup & Reference',
    signature: 'CHOOSE(index_num, value1, [value2], ...)',
    description: 'Returns a value from a list by its position.',
  },
  // Position (3). Excel files these under Lookup & Reference.
  ROW: {
    category: 'Lookup & Reference',
    signature: 'ROW([reference])',
    description: 'Returns the row number of a reference, or of the current cell.',
  },
  COLUMN: {
    category: 'Lookup & Reference',
    signature: 'COLUMN([reference])',
    description: 'Returns the column number of a reference, or of the current cell.',
  },
  ROWS: {
    category: 'Lookup & Reference',
    signature: 'ROWS(array)',
    description: 'Returns the number of rows in a range or array.',
  },
  // Text (8)
  CONCATENATE: {
    category: 'Text',
    signature: 'CONCATENATE(text1, [text2], ...)',
    description: 'Joins several text values into one string.',
  },
  LEFT: {
    category: 'Text',
    signature: 'LEFT(text, [num_chars])',
    description: 'Returns the first characters of a text value.',
  },
  RIGHT: {
    category: 'Text',
    signature: 'RIGHT(text, [num_chars])',
    description: 'Returns the last characters of a text value.',
  },
  MID: {
    category: 'Text',
    signature: 'MID(text, start_num, num_chars)',
    description:
      'Returns characters from the middle of a text value, starting at a given position.',
  },
  LEN: {
    category: 'Text',
    signature: 'LEN(text)',
    description: 'Returns the number of characters in a text value.',
  },
  TRIM: {
    category: 'Text',
    signature: 'TRIM(text)',
    description: 'Removes leading, trailing and repeated spaces from a text value.',
  },
  UPPER: {
    category: 'Text',
    signature: 'UPPER(text)',
    description: 'Converts text to upper case.',
  },
  LOWER: {
    category: 'Text',
    signature: 'LOWER(text)',
    description: 'Converts text to lower case.',
  },
  // Dates (6)
  TODAY: {
    category: 'Date & Time',
    signature: 'TODAY()',
    description: 'Returns the current date as a serial number.',
  },
  NOW: {
    category: 'Date & Time',
    signature: 'NOW()',
    description: 'Returns the current date and time as a serial number.',
  },
  DATE: {
    category: 'Date & Time',
    signature: 'DATE(year, month, day)',
    description: 'Builds a date serial number from a year, month and day.',
  },
  DAY: {
    category: 'Date & Time',
    signature: 'DAY(serial_number)',
    description: 'Returns the day of the month for a date serial number.',
  },
  MONTH: {
    category: 'Date & Time',
    signature: 'MONTH(serial_number)',
    description: 'Returns the month for a date serial number.',
  },
  YEAR: {
    category: 'Date & Time',
    signature: 'YEAR(serial_number)',
    description: 'Returns the year for a date serial number.',
  },
  // Checks (5)
  ISBLANK: {
    category: 'Information',
    signature: 'ISBLANK(value)',
    description: 'Returns TRUE when a value refers to an empty cell.',
  },
  ISNUMBER: {
    category: 'Information',
    signature: 'ISNUMBER(value)',
    description: 'Returns TRUE when a value is a number.',
  },
  ISTEXT: {
    category: 'Information',
    signature: 'ISTEXT(value)',
    description: 'Returns TRUE when a value is text.',
  },
  ISERROR: {
    category: 'Information',
    signature: 'ISERROR(value)',
    description: 'Returns TRUE when a value is any error.',
  },
  ISNA: {
    category: 'Information',
    signature: 'ISNA(value)',
    description: 'Returns TRUE when a value is the #N/A error.',
  },
};

/** The catalogue in published-set order; the picker sorts its lists by name. */
export const FUNCTION_CATALOG: readonly FunctionCatalogEntry[] = PUBLISHED_FUNCTIONS.map(
  (name) => ({ name, ...FUNCTION_DETAILS[name] }),
);

const BY_NAME = new Map<string, FunctionCatalogEntry>(
  FUNCTION_CATALOG.map((entry) => [entry.name, entry]),
);

export function functionEntry(name: string): FunctionCatalogEntry | undefined {
  return BY_NAME.get(name);
}

/** The skeleton the dialog inserts: `=NAME(` with the caret after the parenthesis. */
export function functionSkeleton(name: PublishedFunction): string {
  return `=${name}(`;
}

/**
 * Case-insensitive name search. Prefix matches rank above substring matches so
 * typing `AV` reaches AVERAGE before a name that merely contains those letters.
 */
export function searchFunctions(query: string): readonly FunctionCatalogEntry[] {
  const needle = query.trim().toLowerCase();
  const sorted = [...FUNCTION_CATALOG].sort((left, right) => left.name.localeCompare(right.name));
  if (needle === '') return sorted;
  const matches = sorted.filter((entry) => entry.name.toLowerCase().includes(needle));
  return [
    ...matches.filter((entry) => entry.name.toLowerCase().startsWith(needle)),
    ...matches.filter((entry) => !entry.name.toLowerCase().startsWith(needle)),
  ];
}

/**
 * The rows the list shows for a Select category value. "All" and every Excel
 * category are alphabetical; "Recently used" keeps its most-recent-first order
 * and is empty until a function has been inserted in this session.
 */
export function functionsInCategory(
  category: FunctionCategoryOption,
  recentlyUsed: readonly string[] = [],
): readonly FunctionCatalogEntry[] {
  if (category === 'Recently used') {
    const seen = new Set<string>();
    const entries: FunctionCatalogEntry[] = [];
    for (const name of recentlyUsed) {
      if (seen.has(name)) continue;
      seen.add(name);
      const entry = functionEntry(name);
      if (entry !== undefined) entries.push(entry);
    }
    return entries;
  }
  const matching =
    category === 'All'
      ? FUNCTION_CATALOG
      : FUNCTION_CATALOG.filter((entry) => entry.category === category);
  return [...matching].sort((left, right) => left.name.localeCompare(right.name));
}
