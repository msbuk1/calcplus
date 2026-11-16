import {
  CellError,
  EmptyValue,
  ErrorType,
  FunctionArgumentType,
  FunctionPlugin,
  HyperFormula,
  SimpleRangeValue,
} from 'hyperformula';
import { SHIM_FUNCTION_IDS } from './function-registry';
import { CALCPLUS_LANGUAGE } from './engine-language';

/** The only part of HyperFormula's procedure AST the shims need. */
type PluginAst = { readonly args: never };

/**
 * Excel-semantics shims for HyperFormula 3.5.0 (technical design §3.5).
 *
 * The functions are registered under engine-only ids. The registry gate in
 * `formula-text.ts` rewrites authored text to these ids for the engine only;
 * the model and the file never see either form.
 *
 * Covered here:
 * - `INT` floors toward negative infinity (HyperFormula truncates).
 * - `CHOOSE` returns `#VALUE!` when the index is out of range (HyperFormula
 *   returns `#NUM!`).
 * - `VLOOKUP`/`HLOOKUP` return the first wildcard match in exact mode and
 *   support the `~` escape; approximate mode follows Excel's last-smaller-or-
 *   equal rule (HyperFormula returns the last exact wildcard match).
 * - `MATCH` supports wildcards and the `~` escape for type 0, and types 1/-1
 *   are evaluated against the sorted data (HyperFormula ignores wildcards).
 * - `COUNTIF`/`SUMIF` support `*`, `?` and `~` criteria per formula semantics
 *   §3 (HyperFormula has no `~` escape).
 * - `INDEX` returns `#VALUE!` for a zero or omitted row/column (the whole-row
 *   and whole-column results are not scalar) and `#REF!` out of range
 *   (HyperFormula returns `#NUM!`).
 * - `""` used as an arithmetic operand becomes `#VALUE!` through
 *   `CALCPLUS_EMPTY_TEXT` (HyperFormula coerces it to 0).
 *
 * `INDEX` handles the array form only; the reference form is v2.
 */
export class CalcplusFunctionPlugin extends FunctionPlugin {
  intFunction(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.INT),
      (value: unknown) => Math.floor(asNumber(value)),
    );
  }

  chooseFunction(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.CHOOSE),
      (selector: unknown, ...values: unknown[]) =>
        values.length === 0
          ? error(ErrorType.VALUE)
          : (chooseValue(asNumber(selector), values) as never),
    );
  }

  vlookup(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.VLOOKUP),
      (key: unknown, range: unknown, index: unknown, sorted: unknown) =>
        lookupColumn(key, range, asNumber(index), sorted !== false) as never,
    );
  }

  hlookup(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.HLOOKUP),
      (key: unknown, range: unknown, index: unknown, sorted: unknown) =>
        lookupRow(key, range, asNumber(index), sorted !== false) as never,
    );
  }

  match(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.MATCH),
      (key: unknown, range: unknown, type: unknown) =>
        matchValue(key, range, asNumber(type)) as never,
    );
  }

  index(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.INDEX),
      (array: unknown, row: unknown, column: unknown) => indexValue(array, row, column) as never,
    );
  }

  countif(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.COUNTIF),
      (range: unknown, criteria: unknown) => countIf(range, criteria) as never,
    );
  }

  sumif(ast: PluginAst, state: never): unknown {
    return this.runFunction(
      ast.args,
      state,
      this.metadata(SHIM_FUNCTION_IDS.SUMIF),
      (range: unknown, criteria: unknown, sumRange: unknown) =>
        sumIf(range, criteria, sumRange) as never,
    );
  }

  emptyText(ast: PluginAst, state: never): unknown {
    return this.runFunction(ast.args, state, this.metadata(SHIM_FUNCTION_IDS.EMPTY_TEXT), () =>
      error(ErrorType.VALUE),
    );
  }
}

CalcplusFunctionPlugin.implementedFunctions = {
  [SHIM_FUNCTION_IDS.INT]: {
    method: 'intFunction',
    parameters: [{ argumentType: FunctionArgumentType.NUMBER }],
  },
  [SHIM_FUNCTION_IDS.CHOOSE]: {
    method: 'chooseFunction',
    parameters: [
      { argumentType: FunctionArgumentType.NUMBER },
      { argumentType: FunctionArgumentType.SCALAR, passSubtype: true },
    ],
    repeatLastArgs: 1,
  },
  [SHIM_FUNCTION_IDS.VLOOKUP]: {
    method: 'vlookup',
    parameters: [
      { argumentType: FunctionArgumentType.NOERROR },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.NUMBER },
      { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true },
    ],
  },
  [SHIM_FUNCTION_IDS.HLOOKUP]: {
    method: 'hlookup',
    parameters: [
      { argumentType: FunctionArgumentType.NOERROR },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.NUMBER },
      { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true },
    ],
  },
  [SHIM_FUNCTION_IDS.MATCH]: {
    method: 'match',
    parameters: [
      { argumentType: FunctionArgumentType.NOERROR },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.INTEGER, defaultValue: 1 },
    ],
  },
  [SHIM_FUNCTION_IDS.INDEX]: {
    method: 'index',
    parameters: [
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
      { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
    ],
  },
  [SHIM_FUNCTION_IDS.COUNTIF]: {
    method: 'countif',
    parameters: [
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.SCALAR },
    ],
  },
  [SHIM_FUNCTION_IDS.SUMIF]: {
    method: 'sumif',
    parameters: [
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.SCALAR },
      { argumentType: FunctionArgumentType.RANGE, optionalArg: true },
    ],
  },
  [SHIM_FUNCTION_IDS.EMPTY_TEXT]: {
    method: 'emptyText',
    parameters: [],
  },
};

let registered = false;

/**
 * Register the engine-only shims once. HyperFormula keeps function plugins in
 * a process-wide registry; the engine still never exposes these ids as
 * authored functions because the gate rejects them.
 */
export function registerCalcplusFunctions(): void {
  if (registered) return;
  const translations = Object.fromEntries(Object.values(SHIM_FUNCTION_IDS).map((id) => [id, id]));
  HyperFormula.registerFunctionPlugin(CalcplusFunctionPlugin, {
    enGB: translations,
    [CALCPLUS_LANGUAGE]: translations,
  });
  try {
    // A direct adapter user may register the plugin before the language exists;
    // extend the derived language once it is present.
    HyperFormula.getLanguage(CALCPLUS_LANGUAGE).extendFunctions(translations);
  } catch {
    // The engine falls back to enGB when the derived language is unavailable.
  }
  registered = true;
}

function error(type: ErrorType): CellError {
  return new CellError(type);
}

function asNumber(value: unknown): number {
  const numeric = numericValue(value);
  if (numeric !== null) return numeric;
  if (typeof value === 'boolean') return value ? 1 : 0;
  return Number.NaN;
}

function chooseValue(selector: number, values: readonly unknown[]): unknown {
  if (!Number.isFinite(selector)) return error(ErrorType.VALUE);
  const index = Math.trunc(selector);
  if (index < 1 || index > values.length) return error(ErrorType.VALUE);
  return values[index - 1];
}

function lookupColumn(key: unknown, range: unknown, index: number, sorted: boolean): unknown {
  if (!(range instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  if (!Number.isFinite(index)) return error(ErrorType.VALUE);
  const column = Math.trunc(index);
  if (column < 1) return error(ErrorType.VALUE);
  if (column > range.width()) return error(ErrorType.REF);
  const data = range.data;
  const row = sorted ? approximateColumnPosition(key, data) : exactColumnPosition(key, data);
  if (row < 0) return error(ErrorType.NA);
  return lookupResult(data[row]?.[column - 1]);
}

function lookupRow(key: unknown, range: unknown, index: number, sorted: boolean): unknown {
  if (!(range instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  if (!Number.isFinite(index)) return error(ErrorType.VALUE);
  const rowIndex = Math.trunc(index);
  if (rowIndex < 1) return error(ErrorType.VALUE);
  if (rowIndex > range.height()) return error(ErrorType.REF);
  const data = range.data;
  const column = sorted ? approximateRowPosition(key, data) : exactRowPosition(key, data);
  if (column < 0) return error(ErrorType.NA);
  const resultRow = data[rowIndex - 1];
  return lookupResult(resultRow?.[column]);
}

function matchValue(key: unknown, range: unknown, type: number): unknown {
  if (!(range instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  if (!Number.isFinite(type) || ![-1, 0, 1].includes(type)) return error(ErrorType.VALUE);
  const data = range.data;
  if (data.length > 1 && (data[0]?.length ?? 0) > 1) return error(ErrorType.NA);
  if (data.length === 1) {
    const row = data[0] ?? [];
    if (type === 0) {
      for (let column = 0; column < row.length; column += 1) {
        if (lookupKeyMatches(row[column], key)) return column + 1;
      }
      return error(ErrorType.NA);
    }
    const position = approximateRowPosition(key, data, type as 1 | -1);
    return position < 0 ? error(ErrorType.NA) : position + 1;
  }
  if (type === 0) {
    for (let row = 0; row < data.length; row += 1) {
      if (lookupKeyMatches(data[row]?.[0], key)) return row + 1;
    }
    return error(ErrorType.NA);
  }
  const position = approximateColumnPosition(key, data, type as 1 | -1);
  return position < 0 ? error(ErrorType.NA) : position + 1;
}

/**
 * INDEX array form only (formula semantics §7). Zero or omitted row/column
 * would return a whole row or column in Excel; v1 returns #VALUE! because a
 * formula yields one scalar. Out-of-range indices return #REF! like Excel
 * (HyperFormula returns #NUM!).
 */
function indexValue(array: unknown, row: unknown, column: unknown): unknown {
  if (!(array instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  const height = array.height();
  const width = array.width();
  const rowNumber = indexNumber(row);
  const columnNumber = indexNumber(column);

  if (rowNumber === null || columnNumber === null) return error(ErrorType.VALUE);
  if (rowNumber === 0 || columnNumber === 0) return error(ErrorType.VALUE);

  let selectedRow = rowNumber;
  let selectedColumn = columnNumber;
  if (selectedColumn === undefined) {
    if (width === 1) selectedColumn = 1;
    else if (height === 1) {
      selectedColumn = selectedRow;
      selectedRow = 1;
    } else {
      return error(ErrorType.VALUE);
    }
  }
  if (selectedRow === undefined || selectedColumn === undefined) {
    return error(ErrorType.VALUE);
  }
  const rowIndex = selectedRow;
  const columnIndex = selectedColumn;

  if (rowIndex < 1 || columnIndex < 1) return error(ErrorType.VALUE);
  if (rowIndex > height || columnIndex > width) return error(ErrorType.REF);
  return lookupResult(array.data[rowIndex - 1]?.[columnIndex - 1]);
}

function indexNumber(value: unknown): number | null | undefined {
  if (value === undefined) return undefined;
  const numeric = numericValue(value);
  if (numeric === null) {
    if (typeof value === 'boolean') return value ? 1 : 0;
    return null;
  }
  return Math.trunc(numeric);
}

function exactColumnPosition(key: unknown, data: readonly unknown[][]): number {
  for (let row = 0; row < data.length; row += 1) {
    if (lookupKeyMatches(data[row]?.[0], key)) return row;
  }
  return -1;
}

function exactRowPosition(key: unknown, data: readonly unknown[][]): number {
  const header = data[0];
  if (header === undefined) return -1;
  for (let column = 0; column < header.length; column += 1) {
    if (lookupKeyMatches(header[column], key)) return column;
  }
  return -1;
}

/**
 * Excel approximate lookup: the last value less than or equal to the key in a
 * sorted ascending column (MATCH type -1 uses the smallest value greater than
 * or equal in a descending column).
 */
function approximateColumnPosition(
  key: unknown,
  data: readonly unknown[][],
  mode: 1 | -1 = 1,
): number {
  const target = isEmpty(key) ? 0 : key;
  let found = -1;
  for (let row = 0; row < data.length; row += 1) {
    const cell = data[row]?.[0];
    if (isEmpty(cell) || cell instanceof CellError) continue;
    if (mode === 1) {
      if (compareApproximate(cell, target) <= 0) found = row;
    } else if (
      compareApproximate(cell, target) >= 0 &&
      (found < 0 || compareApproximate(cell, data[found]?.[0]) < 0)
    ) {
      found = row;
    }
  }
  return found;
}

function approximateRowPosition(
  key: unknown,
  data: readonly unknown[][],
  mode: 1 | -1 = 1,
): number {
  const header = data[0];
  if (header === undefined) return -1;
  const target = isEmpty(key) ? 0 : key;
  let found = -1;
  for (let column = 0; column < header.length; column += 1) {
    const cell = header[column];
    if (isEmpty(cell) || cell instanceof CellError) continue;
    if (mode === 1) {
      if (compareApproximate(cell, target) <= 0) found = column;
    } else if (
      compareApproximate(cell, target) >= 0 &&
      (found < 0 || compareApproximate(cell, header[found]) < 0)
    ) {
      found = column;
    }
  }
  return found;
}

function compareApproximate(left: unknown, right: unknown): number {
  const rankLeft = approximateRank(left);
  const rankRight = approximateRank(right);
  if (rankLeft !== rankRight) return rankLeft < rankRight ? -1 : 1;
  const leftNumber = numericValue(left);
  const rightNumber = numericValue(right);
  if (leftNumber !== null && rightNumber !== null) {
    return leftNumber === rightNumber ? 0 : leftNumber < rightNumber ? -1 : 1;
  }
  if (typeof left === 'string' && typeof right === 'string') {
    const leftText = left.toLowerCase();
    const rightText = right.toLowerCase();
    return leftText === rightText ? 0 : leftText < rightText ? -1 : 1;
  }
  if (typeof left === 'boolean' && typeof right === 'boolean') {
    return left === right ? 0 : left ? 1 : -1;
  }
  return 0;
}

function approximateRank(value: unknown): number {
  if (numericValue(value) !== null) return 0;
  if (typeof value === 'string') return 1;
  if (typeof value === 'boolean') return 2;
  return 3;
}

function countIf(range: unknown, criteria: unknown): unknown {
  if (!(range instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  const matcher = criteriaMatcher(criteria);
  if (matcher instanceof CellError) return matcher;
  let count = 0;
  for (const row of range.data) {
    for (const cell of row) {
      if (cell instanceof CellError) continue;
      if (matcher(cell)) count += 1;
    }
  }
  return count;
}

function sumIf(range: unknown, criteria: unknown, sumRange: unknown): unknown {
  if (!(range instanceof SimpleRangeValue)) return error(ErrorType.VALUE);
  const matcher = criteriaMatcher(criteria);
  if (matcher instanceof CellError) return matcher;
  const sumData = sumRange instanceof SimpleRangeValue ? sumRange.data : range.data;
  let total = 0;
  for (let row = 0; row < range.data.length; row += 1) {
    const sourceRow = range.data[row];
    if (sourceRow === undefined) continue;
    for (let column = 0; column < sourceRow.length; column += 1) {
      const cell = sourceRow[column];
      if (cell instanceof CellError) continue;
      if (!matcher(cell)) continue;
      const target = sumData[row]?.[column];
      if (target instanceof CellError) return target;
      const numeric = numericValue(target);
      if (numeric !== null) total += numeric;
    }
  }
  return total;
}

/** A cell value the criteria shims accept; ranges are never scalar criteria. */
type Scalar = unknown;

function lookupKeyMatches(cell: unknown, key: unknown): boolean {
  if (cell instanceof CellError) return false;
  if (isEmpty(cell)) return isEmpty(key);
  if (typeof key === 'string') {
    if (typeof cell !== 'string') return false;
    return hasWildcardSyntax(key)
      ? wildcardRegex(key).test(cell)
      : cell.toLowerCase() === key.toLowerCase();
  }
  const numericKey = numericValue(key);
  if (numericKey !== null) {
    const numericCell = numericValue(cell);
    return numericCell !== null && numericCell === numericKey;
  }
  if (typeof key === 'boolean') return typeof cell === 'boolean' && cell === key;
  return false;
}

/**
 * HyperFormula can wrap numbers in `RichNumber` (dates, percentages) even
 * outside array arithmetic. The shims compare their numeric value.
 */
function numericValue(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value === 'object' && value !== null && 'val' in value) {
    const wrapped = (value as { readonly val?: unknown }).val;
    if (typeof wrapped === 'number' && Number.isFinite(wrapped)) return wrapped;
  }
  return null;
}

function lookupResult(value: unknown): unknown {
  return isEmpty(value) ? 0 : value;
}

function isEmpty(value: unknown): boolean {
  return value === EmptyValue || value === undefined || value === null;
}

type CriteriaMatcher = (cell: Scalar) => boolean;

function criteriaMatcher(criteria: unknown): CriteriaMatcher | CellError {
  if (criteria instanceof CellError) return criteria;
  const numericCriteria = numericValue(criteria);
  if (numericCriteria !== null) {
    return (cell) => equalsValue(cell, numericCriteria);
  }
  if (typeof criteria === 'boolean') {
    return (cell) => equalsValue(cell, criteria);
  }
  const text = isEmpty(criteria) ? '' : String(criteria);
  let operator: string | null = null;
  let rest = text;
  for (const candidate of ['>=', '<=', '<>', '>', '<', '=']) {
    if (text.startsWith(candidate)) {
      operator = candidate;
      rest = text.slice(candidate.length);
      break;
    }
  }
  if (operator !== null) {
    const operand = parseCriteriaOperand(rest);
    return (cell) => matchesOperator(cell, operator, operand);
  }
  if (text === '') {
    return (cell) => isEmpty(cell) || cell === '';
  }
  if (hasWildcardSyntax(text)) {
    const regex = wildcardRegex(text);
    return (cell) => typeof cell === 'string' && regex.test(cell);
  }
  return (cell) => equalsValue(cell, text);
}

function matchesOperator(
  cell: Scalar,
  operator: string,
  operand: string | number | boolean,
): boolean {
  if (isEmpty(cell)) {
    // Excel counts blanks for "<>" and "=" with an empty operand, and ignores
    // them for ordering conditions.
    if (operator === '<>') return operand !== '';
    if (operator === '=') return operand === '';
    return false;
  }
  if (cell instanceof CellError) return false;
  if (operator === '=') return equalsValue(cell, operand);
  if (operator === '<>') return !equalsValue(cell, operand);
  // Ordering operators only compare like types; text and blanks are ignored
  // by numeric conditions such as ">10".
  if (!sameComparableKind(cell, operand)) return false;
  const comparison = compareValues(cell, operand);
  switch (operator) {
    case '>':
      return comparison > 0;
    case '<':
      return comparison < 0;
    case '>=':
      return comparison >= 0;
    case '<=':
      return comparison <= 0;
    default:
      return false;
  }
}

function parseCriteriaOperand(text: string): string | number | boolean {
  if (text === '') return '';
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return Number(text);
  if (/^(true|false)$/i.test(text)) return text.toLowerCase() === 'true';
  return text;
}

function equalsValue(cell: Scalar, criterion: string | number | boolean): boolean {
  if (isEmpty(cell)) return criterion === '';
  if (typeof criterion === 'string') {
    if (criterion === '') return typeof cell === 'string' && cell === '';
    if (hasWildcardSyntax(criterion)) {
      return typeof cell === 'string' && wildcardRegex(criterion).test(cell);
    }
    const criterionNumber = numericTextValue(criterion);
    if (criterionNumber !== null) {
      const cellNumber = numericValue(cell);
      if (cellNumber !== null && cellNumber === criterionNumber) return true;
      return typeof cell === 'string' && numericTextValue(cell) === criterionNumber;
    }
    return typeof cell === 'string' && cell.toLowerCase() === criterion.toLowerCase();
  }
  if (typeof criterion === 'number') {
    const numeric = numericValue(cell);
    if (numeric !== null && numeric === criterion) return true;
    // Excel criteria treat numeric-looking text as the same number for
    // equality and "<>10" conditions (formula semantics §3 and §7).
    return typeof cell === 'string' && numericTextValue(cell) === criterion;
  }
  return typeof cell === 'boolean' && cell === criterion;
}

const NUMERIC_TEXT_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function numericTextValue(value: string): number | null {
  const text = value.trim();
  if (!NUMERIC_TEXT_PATTERN.test(text)) return null;
  const numeric = Number(text);
  return Number.isFinite(numeric) ? numeric : null;
}

function sameComparableKind(cell: unknown, operand: string | number | boolean): boolean {
  if (typeof operand === 'number') return numericValue(cell) !== null;
  return typeof cell === typeof operand;
}

function compareValues(cell: unknown, operand: string | number | boolean): number {
  const cellNumber = numericValue(cell);
  if (cellNumber !== null && typeof operand === 'number') {
    return cellNumber === operand ? 0 : cellNumber < operand ? -1 : 1;
  }
  if (typeof cell === 'string' && typeof operand === 'string') {
    const left = cell.toLowerCase();
    const right = operand.toLowerCase();
    return left === right ? 0 : left < right ? -1 : 1;
  }
  if (typeof cell === 'boolean' && typeof operand === 'boolean') {
    return cell === operand ? 0 : cell ? 1 : -1;
  }
  return 0;
}

/** True when the pattern contains `*`, `?` or the `~` escape. */
function hasWildcardSyntax(pattern: string): boolean {
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index];
    if (character === '~' || character === '*' || character === '?') return true;
  }
  return false;
}

function wildcardRegex(pattern: string): RegExp {
  let source = '^';
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index];
    if (character === '~' && index + 1 < pattern.length) {
      index += 1;
      source += escapeRegExp(pattern[index] ?? '');
    } else if (character === '*') {
      source += '.*';
    } else if (character === '?') {
      source += '.';
    } else {
      source += escapeRegExp(character ?? '');
    }
  }
  return new RegExp(`${source}$`, 'i');
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
