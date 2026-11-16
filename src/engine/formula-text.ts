import { isAllowedEngineFunction, SHIM_FUNCTION_IDS } from './function-registry';

/**
 * Authored-formula screening and engine-only rewriting (formula semantics §9,
 * technical design §3.5).
 *
 * The model keeps the authored text exactly as written. This module decides
 * whether that text may be handed to the engine at all, and produces a
 * separate internal form for the engine. The internal form is never stored in
 * the model and never written to a file.
 */

export type FormulaSkipReason = 'unsupported-function' | 'unparseable';

export interface PreparedFormula {
  /** Engine form without the leading equals sign, or `null` when it must not be sent. */
  readonly text: string | null;
  readonly skip: FormulaSkipReason | null;
  /** Uppercase names outside the published set, in source order. */
  readonly unsupportedFunctions: readonly string[];
}

interface FormulaToken {
  readonly kind: 'string' | 'number' | 'identifier' | 'operator' | 'paren' | 'comma' | 'other';
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly emptyString: boolean;
}

interface FormulaCall {
  readonly name: string;
  readonly nameStart: number;
  readonly openParen: number;
}

interface FormulaScan {
  readonly tokens: readonly FormulaToken[];
  readonly calls: readonly FormulaCall[];
  readonly issues: readonly string[];
}

const UNSUPPORTED_PREFIXES = ['_xlfn.', '_xlws.'] as const;

const ERROR_LITERALS = [
  '#DIV/0!',
  '#VALUE!',
  '#N/A',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#NULL!',
] as const;

const ARITHMETIC_OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '^', '%']);

const PREPARED_CACHE_LIMIT = 4096;

/** Formulas repeat heavily in large sheets; the analysis is pure. */
const preparedCache = new Map<string, PreparedFormula>();

/**
 * Returns true when the authored text must stay out of the engine:
 * implicit intersection (`@`), future-function prefixes (`_xlfn.`/`_xlws.`),
 * structured/external references (`[...]`), a union comma outside a call, or
 * unbalanced quotes and parentheses.
 *
 * Quoted strings and quoted sheet names are scanned so `="a@b"` remains an
 * ordinary formula.
 */
export function formulaHasUnsupportedConstruct(formula: string): boolean {
  if (formula.length === 0 || formula.startsWith('=')) return true;
  return scanFormula(formula).issues.length > 0;
}

/** A formula the engine may be asked to parse. */
export function isEngineFormulaEligible(formula: string): boolean {
  return !formulaHasUnsupportedConstruct(formula);
}

/**
 * Screen the authored text against the 52-function registry and build the
 * engine form. Unsupported functions and unparseable constructs return
 * `text: null`; the caller keeps the authored text and cached result.
 */
export function prepareFormula(formula: string): PreparedFormula {
  const cached = preparedCache.get(formula);
  if (cached !== undefined) return cached;
  const prepared = prepareFormulaUncached(formula);
  if (preparedCache.size >= PREPARED_CACHE_LIMIT) preparedCache.clear();
  preparedCache.set(formula, prepared);
  return prepared;
}

function prepareFormulaUncached(formula: string): PreparedFormula {
  if (formula.length === 0 || formula.startsWith('=')) {
    return { text: null, skip: 'unparseable', unsupportedFunctions: [] };
  }
  const scan = scanFormula(formula);
  if (scan.issues.length > 0) {
    return { text: null, skip: 'unparseable', unsupportedFunctions: [] };
  }
  const unsupportedFunctions = uniqueFunctionNames(scan.calls).filter(
    (name) => !isAllowedEngineFunction(name),
  );
  if (unsupportedFunctions.length > 0) {
    return { text: null, skip: 'unsupported-function', unsupportedFunctions };
  }
  return { text: rewriteForEngine(formula, scan), skip: null, unsupportedFunctions: [] };
}

/**
 * The function names a formula calls, uppercased and de-duplicated in source
 * order. Strings and quoted sheet names are ignored.
 */
export function formulaFunctionNames(formula: string): readonly string[] {
  return uniqueFunctionNames(scanFormula(formula).calls);
}

function uniqueFunctionNames(calls: readonly FormulaCall[]): string[] {
  const names: string[] = [];
  const seen = new Set<string>();
  for (const call of calls) {
    const name = call.name.toUpperCase();
    if (seen.has(name)) continue;
    seen.add(name);
    names.push(name);
  }
  return names;
}

function rewriteForEngine(formula: string, scan: FormulaScan): string {
  const edits: { start: number; end: number; text: string }[] = [];

  for (const call of scan.calls) {
    const upper = call.name.toUpperCase();
    const engineName = shimFunctionFor(upper) ?? call.name;
    edits.push({
      start: call.nameStart,
      end: call.openParen + 1,
      text: `${engineName}(`,
    });
  }

  for (let index = 0; index < scan.tokens.length; index += 1) {
    const token = scan.tokens[index];
    if (!token.emptyString) continue;
    const previous = scan.tokens[index - 1];
    const next = scan.tokens[index + 1];
    if (!isArithmeticOperator(previous) && !isArithmeticOperator(next)) continue;
    edits.push({
      start: token.start,
      end: token.end,
      text: `${SHIM_FUNCTION_IDS.EMPTY_TEXT}()`,
    });
  }

  return applyEdits(formula, edits);
}

function shimFunctionFor(name: string): string | null {
  switch (name) {
    case 'INT':
      return SHIM_FUNCTION_IDS.INT;
    case 'CHOOSE':
      return SHIM_FUNCTION_IDS.CHOOSE;
    case 'VLOOKUP':
      return SHIM_FUNCTION_IDS.VLOOKUP;
    case 'HLOOKUP':
      return SHIM_FUNCTION_IDS.HLOOKUP;
    case 'MATCH':
      return SHIM_FUNCTION_IDS.MATCH;
    case 'INDEX':
      return SHIM_FUNCTION_IDS.INDEX;
    case 'COUNTIF':
      return SHIM_FUNCTION_IDS.COUNTIF;
    case 'SUMIF':
      return SHIM_FUNCTION_IDS.SUMIF;
    default:
      return null;
  }
}

function applyEdits(
  formula: string,
  edits: { start: number; end: number; text: string }[],
): string {
  const sorted = [...edits].sort((left, right) => left.start - right.start);
  let result = '';
  let cursor = 0;
  for (const edit of sorted) {
    if (edit.start < cursor) continue;
    result += formula.slice(cursor, edit.start);
    result += edit.text;
    cursor = edit.end;
  }
  return result + formula.slice(cursor);
}

function isArithmeticOperator(token: FormulaToken | undefined): boolean {
  return token !== undefined && token.kind === 'operator' && ARITHMETIC_OPERATORS.has(token.text);
}

function scanFormula(formula: string): FormulaScan {
  const tokens: FormulaToken[] = [];
  const calls: FormulaCall[] = [];
  const issues: string[] = [];
  if (hasThreeDimensionalRange(formula)) issues.push('unsupported-construct');
  const functionParens = new Set<number>();
  const parenStack: { readonly isFunction: boolean }[] = [];
  let index = 0;
  let depth = 0;
  let braceDepth = 0;

  while (index < formula.length) {
    const character = formula[index] ?? '';
    if (isWhitespace(character)) {
      index += 1;
      continue;
    }

    if (character === '"') {
      const start = index;
      const quoted = skipQuoted(formula, index, '"');
      index = quoted.end;
      if (!quoted.closed) issues.push('unbalanced');
      const text = formula.slice(start, index);
      tokens.push({
        kind: 'string',
        text,
        start,
        end: index,
        emptyString: text === '""',
      });
      continue;
    }

    if (character === "'") {
      const start = index;
      const quoted = skipQuoted(formula, index, "'");
      index = quoted.end;
      if (!quoted.closed) issues.push('unbalanced');
      tokens.push({
        kind: 'other',
        text: formula.slice(start, index),
        start,
        end: index,
        emptyString: false,
      });
      continue;
    }

    if (character === '#') {
      const literal = errorLiteralAt(formula, index);
      if (literal === null) {
        // A spill reference (A1#) or an unknown error literal. Excel's spill
        // operator is not a v1 construct, and unknown text is unparseable.
        issues.push('unsupported-construct');
        index += 1;
        continue;
      }
      tokens.push({
        kind: 'other',
        text: literal,
        start: index,
        end: index + literal.length,
        emptyString: false,
      });
      index += literal.length;
      continue;
    }

    if (character === '@' || character === '[' || character === ']') {
      issues.push('unsupported-construct');
      index += 1;
      continue;
    }

    if (character === '(') {
      tokens.push({
        kind: 'paren',
        text: character,
        start: index,
        end: index + 1,
        emptyString: false,
      });
      parenStack.push({ isFunction: functionParens.has(index) });
      depth += 1;
      index += 1;
      continue;
    }

    if (character === ')') {
      depth -= 1;
      if (depth < 0) issues.push('unbalanced');
      parenStack.pop();
      tokens.push({
        kind: 'paren',
        text: character,
        start: index,
        end: index + 1,
        emptyString: false,
      });
      index += 1;
      continue;
    }

    if (character === ',') {
      if (braceDepth === 0) {
        const context = parenStack[parenStack.length - 1];
        if (depth === 0 || context === undefined || !context.isFunction) {
          // A comma outside a function's argument list is Excel's union
          // operator, which is v2.
          issues.push('unsupported-construct');
        }
      }
      tokens.push({
        kind: 'comma',
        text: character,
        start: index,
        end: index + 1,
        emptyString: false,
      });
      index += 1;
      continue;
    }

    if (character === ';' || character === '{' || character === '}') {
      if (character === '{') braceDepth += 1;
      else if (character === '}') braceDepth = Math.max(0, braceDepth - 1);
      tokens.push({
        kind: 'other',
        text: character,
        start: index,
        end: index + 1,
        emptyString: false,
      });
      index += 1;
      continue;
    }

    if (isIdentifierStart(character)) {
      const start = index;
      index += 1;
      while (index < formula.length && isIdentifierPart(formula[index] ?? '')) index += 1;
      const name = formula.slice(start, index);
      if (hasUnsupportedPrefix(name)) issues.push('unsupported-construct');
      // R1C1 notation is v2; never hand it to the engine as a defined name.
      if (/^R\d+C\d+$/i.test(name)) issues.push('unsupported-construct');

      let openParen = index;
      while (openParen < formula.length && isWhitespace(formula[openParen] ?? '')) openParen += 1;
      if (formula[openParen] === '(') {
        if (!hasMatchingParen(formula, openParen)) issues.push('unbalanced');
        else {
          calls.push({ name, nameStart: start, openParen });
          functionParens.add(openParen);
        }
        continue;
      }

      tokens.push({
        kind: 'identifier',
        text: name,
        start,
        end: index,
        emptyString: false,
      });
      continue;
    }

    if (isDigit(character) || (character === '.' && isDigit(formula[index + 1] ?? ''))) {
      const start = index;
      index += 1;
      while (index < formula.length && /[0-9.eE+-]/.test(formula[index] ?? '')) {
        // Stop a numeric literal at a second sign that starts the next term.
        const current = formula[index] ?? '';
        const previous = formula[index - 1] ?? '';
        if ((current === '+' || current === '-') && previous !== 'e' && previous !== 'E') break;
        index += 1;
      }
      tokens.push({
        kind: 'number',
        text: formula.slice(start, index),
        start,
        end: index,
        emptyString: false,
      });
      continue;
    }

    if (isOperatorCharacter(character)) {
      const start = index;
      const next = formula[index + 1] ?? '';
      const twoCharacter = `${character}${next}`;
      if (character === '<' && (next === '=' || next === '>')) index += 2;
      else if (character === '>' && next === '=') index += 2;
      else if (character === '<' && next === '>') index += 2;
      else index += 1;
      tokens.push({
        kind: 'operator',
        text: twoCharacter.length === 2 && index - start === 2 ? twoCharacter : character,
        start,
        end: index,
        emptyString: false,
      });
      continue;
    }

    tokens.push({
      kind: 'other',
      text: character,
      start: index,
      end: index + 1,
      emptyString: false,
    });
    index += 1;
  }

  if (depth !== 0) issues.push('unbalanced');
  for (let index = 1; index < tokens.length; index += 1) {
    const previous = tokens[index - 1];
    const current = tokens[index];
    if (previous === undefined || current === undefined) continue;
    if (!isLeftOperand(previous) || !isRightOperand(current)) continue;
    if (/\s/.test(formula.slice(previous.end, current.start))) {
      // Two operands separated only by whitespace are Excel's intersection
      // operator (or malformed text). Both are v2/unsupported constructs.
      issues.push('unsupported-construct');
    }
  }
  checkOperatorStructure(tokens, issues);
  return { tokens, calls, issues };
}

/**
 * Reject malformed operator sequences before the text reaches HyperFormula:
 * trailing operators (`1+`, `SUM(1+)`), a missing operand (`1+*2`), a leading
 * binary operator (`*1`) and a postfix percent with no operand.
 */
function checkOperatorStructure(tokens: readonly FormulaToken[], issues: string[]): void {
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token === undefined || token.kind !== 'operator') continue;
    const previous = tokens[index - 1];
    const next = tokens[index + 1];

    if (token.text === '%') {
      const hasOperand =
        previous !== undefined &&
        (isLeftOperand(previous) || (previous.kind === 'operator' && previous.text === '%'));
      if (!hasOperand) issues.push('unsupported-construct');
      if (next !== undefined && isRightOperand(next)) issues.push('unsupported-construct');
      continue;
    }

    if (previous === undefined) {
      if ((token.text !== '+' && token.text !== '-') || next === undefined) {
        issues.push('unsupported-construct');
      }
      continue;
    }
    if (next === undefined) {
      issues.push('unsupported-construct');
      continue;
    }
    if (next.kind === 'comma' || (next.kind === 'paren' && next.text === ')')) {
      issues.push('unsupported-construct');
      continue;
    }
    if ((previous.kind === 'paren' && previous.text === '(') || previous.kind === 'comma') {
      if (token.text !== '+' && token.text !== '-') issues.push('unsupported-construct');
      continue;
    }
    if (next.kind === 'operator' && next.text !== '+' && next.text !== '-') {
      issues.push('unsupported-construct');
    }
  }
}

function errorLiteralAt(formula: string, start: number): string | null {
  const rest = formula.slice(start).toUpperCase();
  for (const literal of ERROR_LITERALS) {
    if (rest.startsWith(literal)) return formula.slice(start, start + literal.length);
  }
  return null;
}

/** A token that can end an operand; a closing parenthesis can. */
function isLeftOperand(token: FormulaToken): boolean {
  if (token.kind === 'identifier' || token.kind === 'number' || token.kind === 'string') {
    return true;
  }
  if (token.kind === 'paren') return token.text === ')';
  return (
    token.kind === 'other' &&
    (token.text === '$' || token.text.startsWith("'") || token.text.startsWith('#'))
  );
}

/** A token that can start an operand; an opening parenthesis can. */
function isRightOperand(token: FormulaToken): boolean {
  if (token.kind === 'identifier' || token.kind === 'number' || token.kind === 'string') {
    return true;
  }
  if (token.kind === 'paren') return token.text === '(';
  return (
    token.kind === 'other' &&
    (token.text === '$' || token.text.startsWith("'") || token.text.startsWith('#'))
  );
}

function hasMatchingParen(formula: string, openParen: number): boolean {
  let depth = 1;
  let index = openParen + 1;
  while (index < formula.length) {
    const character = formula[index] ?? '';
    if (character === '"' || character === "'") {
      index = skipQuoted(formula, index, character).end;
      continue;
    }
    if (character === '(') {
      depth += 1;
    } else if (character === ')') {
      depth -= 1;
      if (depth === 0) return true;
    }
    index += 1;
  }
  return false;
}

/** Skip a quoted string or sheet name; doubled quotes escape the delimiter. */
function skipQuoted(
  formula: string,
  start: number,
  quote: string,
): { end: number; closed: boolean } {
  let index = start + 1;
  while (index < formula.length) {
    if (formula[index] === quote) {
      if (formula[index + 1] === quote) {
        index += 2;
        continue;
      }
      return { end: index + 1, closed: true };
    }
    index += 1;
  }
  return { end: formula.length, closed: false };
}

function hasUnsupportedPrefix(name: string): boolean {
  const lower = name.toLowerCase();
  return UNSUPPORTED_PREFIXES.some((prefix) => lower.startsWith(prefix));
}

/** `Sheet1:Sheet3!A1` and quoted equivalents are 3D ranges, which are v2. */
function hasThreeDimensionalRange(formula: string): boolean {
  const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'/g, "''");
  return /(?:''|[A-Za-z_][\w.]*)\s*:\s*(?:''|[A-Za-z_][\w.]*)\s*!/.test(stripped);
}

function isWhitespace(character: string): boolean {
  return character === ' ' || character === '\t' || character === '\n' || character === '\r';
}

function isDigit(character: string): boolean {
  return character >= '0' && character <= '9';
}

function isIdentifierStart(character: string): boolean {
  return /[A-Za-z_\u00C0-\u02AF]/.test(character);
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.\u00C0-\u02AF]/.test(character);
}

function isOperatorCharacter(character: string): boolean {
  return '+ - * / ^ & % = < >'.includes(character);
}
