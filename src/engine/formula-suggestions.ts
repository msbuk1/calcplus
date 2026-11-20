import { functionEntry, searchFunctions } from './function-catalog';
import type { FunctionCatalogEntry } from './function-catalog';
import type { PublishedFunction } from './function-registry';

/**
 * Formula autocomplete and signature hints (feature matrix §4, command
 * `formula.autocomplete`).
 *
 * The module is pure: it reads the formula text and the caret offset and
 * returns either the function names that match the token being typed or the
 * signature of the call the caret sits inside. The cell editor and the formula
 * bar share it, so both surfaces suggest the same functions from the same
 * published catalogue and never offer a function the engine cannot evaluate.
 *
 * A completion is only offered for formula entry. Text inside a quoted string
 * literal never produces a suggestion.
 */

export interface NameSuggestions {
  readonly kind: 'names';
  /** The token starts here, in the edited text's offsets. */
  readonly start: number;
  /** The token ends here; the caret sits at this offset. */
  readonly end: number;
  readonly query: string;
  readonly matches: readonly FunctionCatalogEntry[];
}

export interface SignatureHint {
  readonly kind: 'signature';
  readonly name: PublishedFunction;
  readonly signature: string;
  /** 0-based argument the caret is in; `0` for an empty first argument. */
  readonly argumentIndex: number;
  readonly arguments: readonly string[];
}

export type FormulaCompletion = NameSuggestions | SignatureHint;

export interface FormulaSource {
  /** The edited text is formula entry (the editor may hide a leading `=`). */
  readonly formula: boolean;
}

export interface CompletionChoice {
  /** The whole edited text after the token was replaced. */
  readonly text: string;
  /** Where the caret belongs after the replacement. */
  readonly caret: number;
}

/** A character that may immediately precede a function name in a formula. */
const FUNCTION_BOUNDARIES = new Set(['=', '+', '-', '*', '/', '^', '&', '>', '<', ',', ':', '(']);

export function formulaCompletionAt(
  text: string,
  caret: number,
  source: FormulaSource,
): FormulaCompletion | null {
  if (!source.formula) return null;
  const offset = Math.min(Math.max(0, caret), text.length);
  if (insideStringLiteral(text, offset)) return null;

  const names = nameSuggestions(text, offset);
  if (names !== null) return names;
  return signatureHint(text, offset);
}

/**
 * Replace the typed token with `NAME(` and leave the caret on the first
 * argument, the same insertion the Insert Function dialog performs.
 */
export function applyFormulaCompletion(
  text: string,
  completion: NameSuggestions,
  name: PublishedFunction,
): CompletionChoice {
  const insertion = `${name}(`;
  return {
    text: `${text.slice(0, completion.start)}${insertion}${text.slice(completion.end)}`,
    caret: completion.start + insertion.length,
  };
}

export interface SignatureSegment {
  readonly text: string;
  readonly active: boolean;
}

/** The signature split into pieces, so the current argument can be emphasised. */
export function signatureSegments(hint: SignatureHint): readonly SignatureSegment[] {
  const open = hint.signature.indexOf('(');
  const close = hint.signature.lastIndexOf(')');
  if (open < 0 || close < open) return [{ text: hint.signature, active: false }];
  const segments: SignatureSegment[] = [{ text: hint.signature.slice(0, open + 1), active: false }];
  hint.arguments.forEach((argument, index) => {
    if (index > 0) segments.push({ text: ', ', active: false });
    segments.push({ text: argument, active: index === hint.argumentIndex });
  });
  segments.push({ text: ')', active: false });
  return segments;
}

// ---------------------------------------------------------------------------
// Name suggestions
// ---------------------------------------------------------------------------

function nameSuggestions(text: string, caret: number): NameSuggestions | null {
  const before = text.slice(0, caret);
  const token = /[A-Za-z][A-Za-z0-9.]*$/.exec(before);
  if (token === null) return null;
  const start = caret - token[0].length;
  if (start > 0) {
    const preceding = text[start - 1];
    // The start of the formula is a boundary too: the cell editor hides the
    // leading `=` and the formula bar keeps it.
    if (preceding !== undefined && !FUNCTION_BOUNDARIES.has(preceding)) return null;
  }
  const query = token[0];
  const matches = searchFunctions(query).filter((entry) =>
    entry.name.toLowerCase().startsWith(query.toLowerCase()),
  );
  if (matches.length === 0) return null;
  return { kind: 'names', start, end: caret, query, matches };
}

// ---------------------------------------------------------------------------
// Signature hints
// ---------------------------------------------------------------------------

function signatureHint(text: string, caret: number): SignatureHint | null {
  const masked = maskStringLiterals(text);
  const open = unmatchedOpenParenthesis(masked, caret);
  if (open < 0) return null;
  const nameMatch = /([A-Za-z][A-Za-z0-9.]*)\s*$/.exec(masked.slice(0, open));
  if (nameMatch === null) return null;
  const entry = functionEntry(nameMatch[1]!.toUpperCase());
  if (entry === undefined) return null;
  return {
    kind: 'signature',
    name: entry.name,
    signature: entry.signature,
    argumentIndex: argumentIndexAt(masked, open + 1, caret),
    arguments: signatureArguments(entry.signature),
  };
}

/** The innermost `(` the caret is still inside, or -1. */
function unmatchedOpenParenthesis(masked: string, caret: number): number {
  let depth = 0;
  for (let index = caret - 1; index >= 0; index -= 1) {
    const character = masked[index];
    if (character === ')') {
      depth += 1;
    } else if (character === '(') {
      if (depth === 0) return index;
      depth -= 1;
    }
  }
  return -1;
}

/** Count top-level commas between the call's `(` and the caret. */
function argumentIndexAt(masked: string, from: number, caret: number): number {
  let depth = 0;
  let index = 0;
  for (let position = from; position < caret; position += 1) {
    const character = masked[position];
    if (character === '(') depth += 1;
    else if (character === ')') depth = Math.max(0, depth - 1);
    else if (character === ',' && depth === 0) index += 1;
  }
  return index;
}

/** `SUM(number1, [number2], ...)` -> `['number1', '[number2]', '...']`. */
function signatureArguments(signature: string): readonly string[] {
  const open = signature.indexOf('(');
  const close = signature.lastIndexOf(')');
  if (open < 0 || close <= open) return [];
  const inner = signature.slice(open + 1, close).trim();
  if (inner === '') return [];
  return inner.split(',').map((argument) => argument.trim());
}

function insideStringLiteral(text: string, caret: number): boolean {
  let inside = false;
  for (let index = 0; index < caret; index += 1) {
    if (text[index] !== '"') continue;
    if (inside && text[index + 1] === '"') {
      index += 1;
      continue;
    }
    inside = !inside;
  }
  return inside;
}

/** String contents become spaces so parentheses and names inside text are inert. */
function maskStringLiterals(text: string): string {
  let result = '';
  let inside = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]!;
    if (character === '"') {
      if (inside && text[index + 1] === '"') {
        result += '  ';
        index += 1;
        continue;
      }
      inside = !inside;
      result += character;
      continue;
    }
    result += inside ? ' ' : character;
  }
  return result;
}
