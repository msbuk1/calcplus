import { invalidXmlCharacterMessage } from './text';
import type { Cell, SheetId, Workbook } from './types';
import { parseCellKey } from './types';
import { BLANK, formatCellValue, parseUserValue } from './value';

/**
 * Find & Replace over the workbook model (feature matrix §1).
 *
 * Search is literal, never a regular expression: a query such as `[` or `(`
 * is safe, and replacement text such as `$&` is inserted as typed. Values are
 * searched both as raw text (`1234.5`) and as display text (`1,234.50`) so a
 * formatted cell can be found either way. Formulas are searched as the
 * authored text the model stores, never as normalised engine text.
 */

export type FindScope = 'sheet' | 'workbook';

export interface FindOptions {
  readonly query: string;
  readonly scope: FindScope;
  /** The current sheet for `scope: 'sheet'`. */
  readonly sheetId?: SheetId | null;
  readonly matchCase?: boolean;
  readonly wholeCell?: boolean;
  readonly inFormulas?: boolean;
  readonly inValues?: boolean;
  /**
   * Display text after the cell's number format. The model does not import the
   * formatter; the command layer passes the projection's formatting seam.
   */
  readonly displayText?: (cell: Cell) => string;
  /**
   * Workbook-scope visibility filter. The model has no visibility field, so the
   * command layer passes the imported workbook state. Sheet scope keeps the
   * explicitly selected sheet; a workbook search must never land on a hidden or
   * veryHidden sheet, whose match the view cannot activate.
   */
  readonly includeSheet?: (sheetId: SheetId) => boolean;
}

/** One matching cell. Order is sheet order, then row, then column. */
export interface FindMatch {
  readonly sheetId: SheetId;
  readonly row: number;
  readonly col: number;
  /** The text that matched, shortened only when it is very long. */
  readonly preview: string;
}

export type FindValueSource = 'raw' | 'display';

/** The detail the replace path needs; the UI only consumes `FindMatch`. */
export interface CellFindMatch {
  readonly matchedIn: 'formula' | 'value';
  /** Which value text matched, or `null` for a formula match. */
  readonly source: FindValueSource | null;
  /** The exact candidate text that contains the query. */
  readonly text: string;
  readonly preview: string;
}

export interface ReplaceOptions extends FindOptions {
  readonly replacement: string;
}

/**
 * The outcome of replacing in one cell. `clear` removes an unstyled cell record
 * instead of storing an empty ghost; `set` carries the new snapshot; `unsafe`
 * refuses text that XML cannot store.
 */
export type ReplaceCellOutcome =
  | { readonly kind: 'no-change' }
  | { readonly kind: 'clear' }
  | { readonly kind: 'set'; readonly cell: Cell }
  | { readonly kind: 'unsafe'; readonly message: string };

const MAX_PREVIEW_LENGTH = 160;

interface LocatedMatch {
  readonly index: number;
  readonly length: number;
}

/** Every matching cell in the selected scope, in document order. */
export function searchWorkbook(workbook: Workbook, options: FindOptions): FindMatch[] {
  if (options.query === '' || (options.inFormulas === false && options.inValues === false)) {
    return [];
  }
  const sheets =
    options.scope === 'workbook'
      ? workbook.sheets.filter((sheet) => options.includeSheet?.(sheet.id) ?? true)
      : workbook.sheets.filter(
          (sheet) => sheet.id === (options.sheetId ?? workbook.sheets[0]?.id ?? null),
        );
  const matches: FindMatch[] = [];
  for (const sheet of sheets) {
    const entries: { row: number; col: number; cell: Cell }[] = [];
    for (const [key, cell] of sheet.cells) {
      try {
        const { row, col } = parseCellKey(key);
        entries.push({ row, col, cell });
      } catch {
        // A malformed key cannot be a real cell address; leave it out rather
        // than failing the whole search.
      }
    }
    entries.sort((left, right) => left.row - right.row || left.col - right.col);
    for (const entry of entries) {
      const match = searchCell(entry.cell, options);
      if (match === null) continue;
      matches.push({
        sheetId: sheet.id,
        row: entry.row,
        col: entry.col,
        preview: match.preview,
      });
    }
  }
  return matches;
}

/**
 * The first candidate in one cell that matches. Formulas win over values when
 * both match, so a cell is one match and Replace can act on it once.
 */
export function searchCell(cell: Cell, options: FindOptions): CellFindMatch | null {
  if (options.query === '') return null;

  if (options.inFormulas !== false && cell.formula !== null && !cell.formulaDerived) {
    const located = locateMatch(cell.formula, options);
    if (located !== null) {
      return {
        matchedIn: 'formula',
        source: null,
        text: cell.formula,
        preview: previewText(cell.formula, located),
      };
    }
  }

  if (options.inValues !== false) {
    const raw = formatCellValue(cell.value);
    if (raw !== '') {
      const located = locateMatch(raw, options);
      if (located !== null) {
        return {
          matchedIn: 'value',
          source: 'raw',
          text: raw,
          preview: previewText(raw, located),
        };
      }
    }
    const display = options.displayText?.(cell);
    if (display !== undefined && display !== '' && display !== raw) {
      const located = locateMatch(display, options);
      if (located !== null) {
        return {
          matchedIn: 'value',
          source: 'display',
          text: display,
          preview: previewText(display, located),
        };
      }
    }
  }

  return null;
}

/**
 * The next cell state after replacing every occurrence of the query in the
 * one matching cell. `no-change` covers a cell that does not match or is
 * read-only; `clear` removes an unstyled cell's record; `unsafe` refuses text
 * that XML cannot store, so the cell is skipped before a change is formed.
 *
 * Shared formulas, legacy arrays and rich text are read-only, matching the
 * cell-edit contract. An empty replacement clears the matched cell contents;
 * an unstyled cell is removed entirely, while a styled cell keeps its style as
 * a blank record. Clearing a formula also removes the authored text and its
 * cached result. A formula replacement that keeps a formula stores the new
 * authored text, drops the cached result and marks the cell stale so the engine
 * recalculates it; a value replacement stores a typed value parsed from the
 * replaced text.
 */
export function replaceCellMatches(cell: Cell, options: ReplaceOptions): ReplaceCellOutcome {
  if (cell.richText !== null || cell.formulaSource !== undefined) return { kind: 'no-change' };
  const match = searchCell(cell, options);
  if (match === null) return { kind: 'no-change' };

  if (match.matchedIn === 'formula') {
    if (cell.formula === null) return { kind: 'no-change' };
    const nextFormula = replacedText(cell.formula, options);
    if (nextFormula === null) return { kind: 'no-change' };
    const unsafe = unsafeTextMessage(options.replacement, nextFormula, 'Formula text');
    if (unsafe !== null) return { kind: 'unsafe', message: unsafe };
    if (nextFormula.trim() === '') {
      // Replacing the whole formula with nothing is a content clear: the
      // formula and its cached result are gone, so the cell itself has no
      // derived value left to mark stale. The engine recalculates dependants.
      return clearedCell(cell);
    }
    return {
      kind: 'set',
      cell: {
        value: cell.value,
        formula: nextFormula,
        formulaDerived: false,
        savedResult: null,
        styleId: cell.styleId,
        richText: null,
        stale: true,
      },
    };
  }

  // A value shown for a formula cell is derived. Changing it without changing
  // the formula would only last until the next recalculation.
  if (cell.formula !== null) return { kind: 'no-change' };
  const nextText = replacedText(match.text, options);
  if (nextText === null) return { kind: 'no-change' };
  const unsafe = unsafeTextMessage(options.replacement, nextText, 'Cell text');
  if (unsafe !== null) return { kind: 'unsafe', message: unsafe };
  const value = parseUserValue(nextText);
  return value.type === 'blank'
    ? clearedCell(cell)
    : { kind: 'set', cell: cellWithValue(cell, value) };
}

/**
 * The save layer rejects invalid XML characters, so a replacement that would
 * introduce one must be refused before a change exists. Both the replacement
 * itself and the resulting text are checked: a formula keeps its authored text,
 * while a value is stored after parsing.
 */
function unsafeTextMessage(
  replacement: string,
  resultText: string,
  resultContext: string,
): string | null {
  const replacementMessage = invalidXmlCharacterMessage(replacement, 'Replacement text');
  if (replacementMessage !== null) return `${replacementMessage} The cell was not changed.`;
  const resultMessage = invalidXmlCharacterMessage(resultText, resultContext);
  return resultMessage === null ? null : `${resultMessage} The cell was not changed.`;
}

/** An empty replacement removes an unstyled cell; a styled cell stays blank. */
function clearedCell(cell: Cell): ReplaceCellOutcome {
  if (cell.styleId === 0) return { kind: 'clear' };
  return { kind: 'set', cell: cellWithValue(cell, BLANK) };
}

function locateMatch(text: string, options: FindOptions): LocatedMatch | null {
  if (options.query === '') return null;
  if (options.wholeCell === true) {
    return equalsText(text, options.query, options.matchCase === true)
      ? { index: 0, length: text.length }
      : null;
  }
  if (options.matchCase === true) {
    const index = text.indexOf(options.query);
    return index < 0 ? null : { index, length: options.query.length };
  }
  const index = text.toLowerCase().indexOf(options.query.toLowerCase());
  return index < 0 ? null : { index, length: options.query.length };
}

function equalsText(left: string, right: string, matchCase: boolean): boolean {
  return matchCase ? left === right : left.toLowerCase() === right.toLowerCase();
}

function replacedText(text: string, options: ReplaceOptions): string | null {
  if (options.query === '') return null;
  let next: string;
  if (options.wholeCell === true) {
    next = options.replacement;
  } else {
    try {
      // The query is escaped, so a literal such as `[` cannot throw. A function
      // replacer keeps `$&` and other replacement patterns literal.
      const pattern = new RegExp(
        escapeRegExp(options.query),
        options.matchCase === true ? 'g' : 'gi',
      );
      next = text.replace(pattern, () => options.replacement);
    } catch {
      return null;
    }
  }
  return next === text ? null : next;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function previewText(text: string, located: LocatedMatch): string {
  if (text.length <= MAX_PREVIEW_LENGTH) return text;
  const start = Math.max(0, Math.min(located.index - 40, text.length - MAX_PREVIEW_LENGTH));
  const end = Math.min(
    text.length,
    Math.max(start + MAX_PREVIEW_LENGTH, located.index + located.length + 40),
  );
  return `${start > 0 ? '...' : ''}${text.slice(start, end)}${end < text.length ? '...' : ''}`;
}

function cellWithValue(cell: Cell, value: Cell['value']): Cell {
  return {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: cell.styleId,
    richText: null,
    stale: false,
  };
}
