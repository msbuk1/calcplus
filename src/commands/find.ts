import type { Cell, Sheet, Workbook } from '@model/types';
import { cellStateEquals, cloneCell, getCell, singleCellRange } from '@model/types';
import { arrayFormulaForCell } from '@model/formula-ranges';
import { getStyle } from '@model/workbook';
import {
  replaceCellMatches,
  searchWorkbook,
  type FindMatch,
  type FindOptions,
  type FindScope,
  type ReplaceOptions,
} from '@model/find';
import { formatValue } from '@view/number-format';
import { sheetVisibilityById } from '@view/sheet-visibility';
import type { CommandChange, CommandContext, CommandDefinition, CommandResult } from './types';

/**
 * Find & Replace command surface (feature matrix §1).
 *
 * `find.search` returns matches and changes nothing. `find.replaceOne` and
 * `find.replaceAll` each dispatch one command, so each is one undoable
 * transaction whatever the number of cells. The dialog and its open commands
 * belong to the chrome layer; these ids are the submission seam it calls.
 */

export type FindDirection = 'down' | 'up';

export interface FindArgs {
  readonly query: string;
  readonly scope?: FindScope;
  readonly matchCase?: boolean;
  readonly wholeCell?: boolean;
  readonly inFormulas?: boolean;
  readonly inValues?: boolean;
  /**
   * Direction constrains the scan from the active cell (matrix row 1, ticket
   * 28 D3). Results arrive in scan order: document order downwards, reverse
   * document order upwards. 'down' is the default and the prior behaviour.
   */
  readonly direction?: FindDirection;
}

export interface ReplaceArgs extends FindArgs {
  readonly replacement: string;
}

const hasSheet = (context: CommandContext): boolean => context.workbook.sheets.length > 0;

const ARRAY_READ_ONLY =
  'Legacy array formulas are read-only in this milestone; the cell was not changed.';
const SHARED_READ_ONLY =
  'Shared formulas are read-only in this milestone; the cell was not changed.';
const RICH_TEXT_READ_ONLY =
  'This cell contains rich text. Rich text is read-only in this milestone; editing it would flatten its formatting, so the cell was not changed.';

type MatchChange =
  | { readonly kind: 'change'; readonly change: CommandChange }
  | { readonly kind: 'read-only'; readonly message: string }
  | { readonly kind: 'unsafe'; readonly message: string };

export const findSearchCommand: CommandDefinition<FindArgs> = {
  id: 'find.search',
  label: 'Find',
  category: 'edit',
  keybindings: ['Ctrl+F'],
  isEnabled: hasSheet,
  run(context, args): CommandResult {
    const options = findOptions(context, args);
    if (options.query === '') {
      return { changes: [], matches: [], find: { matches: [], total: 0 } };
    }
    const matches = scanOrder(searchWorkbook(context.workbook, options), args?.direction);
    return { changes: [], matches, find: { matches, total: matches.length } };
  },
};

/** Results in scan order: document order downwards, reversed upwards. */
function scanOrder(matches: FindMatch[], direction: FindDirection | undefined): FindMatch[] {
  return direction === 'up' ? [...matches].reverse() : matches;
}

export const findReplaceOneCommand: CommandDefinition<ReplaceArgs> = {
  id: 'find.replaceOne',
  label: 'Replace',
  category: 'edit',
  keybindings: ['Ctrl+H'],
  isEnabled: hasSheet,
  run(context, args): CommandResult {
    const options = replaceOptions(context, args);
    const matches = scanOrder(searchWorkbook(context.workbook, options), args?.direction);
    if (matches.length === 0) return toast('No matches found.');
    for (const match of orderedMatches(context, matches, args?.direction)) {
      const outcome = changeForMatch(context.workbook, match, options);
      if (outcome === null) continue;
      if (outcome.kind !== 'change') return toast(outcome.message);
      return {
        changes: [outcome.change],
        selectionAfter: singleCellRange(match.row, match.col),
        activeSheetId: match.sheetId,
        replaced: 1,
        find: { matches: [], total: 0, replaced: 1 },
      };
    }
    return toast('No replaceable matches found.');
  },
};

export const findReplaceAllCommand: CommandDefinition<ReplaceArgs> = {
  id: 'find.replaceAll',
  label: 'Replace all',
  category: 'edit',
  keybindings: [],
  isEnabled: hasSheet,
  run(context, args): CommandResult {
    const options = replaceOptions(context, args);
    const matches = searchWorkbook(context.workbook, options);
    const changes: CommandChange[] = [];
    let skippedMessage: string | null = null;
    for (const match of matches) {
      const outcome = changeForMatch(context.workbook, match, options);
      if (outcome === null) continue;
      if (outcome.kind !== 'change') {
        skippedMessage ??= outcome.message;
        continue;
      }
      changes.push(outcome.change);
    }
    if (changes.length === 0) {
      return toast(
        skippedMessage ??
          (matches.length === 0 ? 'No matches found.' : 'No replaceable matches found.'),
      );
    }
    return {
      changes,
      replaced: changes.length,
      find: { matches: [], total: 0, replaced: changes.length },
    };
  },
};

function findOptions(context: CommandContext, args: FindArgs): FindOptions {
  const input = args ?? { query: '' };
  const visibility = sheetVisibilityById(context.workbook);
  return {
    query: typeof input.query === 'string' ? input.query : '',
    scope: input.scope ?? 'sheet',
    sheetId: context.activeSheetId,
    matchCase: input.matchCase === true,
    wholeCell: input.wholeCell === true,
    inFormulas: input.inFormulas !== false,
    inValues: input.inValues !== false,
    displayText: (cell) => displayTextFor(context.workbook, cell),
    // A workbook search never lands on an imported hidden sheet: the view
    // refuses to activate it, so its match would trap Find Next.
    includeSheet: (sheetId) => (visibility.get(sheetId) ?? 'visible') === 'visible',
  };
}

function replaceOptions(context: CommandContext, args: ReplaceArgs): ReplaceOptions {
  const input = args ?? { query: '', replacement: '' };
  return {
    ...findOptions(context, input),
    replacement: typeof input.replacement === 'string' ? input.replacement : '',
  };
}

/** Display text after the cell's number format, the view's formatting seam. */
function displayTextFor(workbook: Workbook, cell: Cell): string {
  const style = getStyle(workbook, cell.styleId);
  return formatValue(cell.value, style.numFmtCode, workbook.date1904);
}

/**
 * Excel repeats Replace from the current selection: replace the selected match
 * if it is one, otherwise start at the next match in scan order and wrap. The
 * matches arrive in scan order, so the direction only flips the "next" test.
 */
function orderedMatches(
  context: CommandContext,
  matches: readonly FindMatch[],
  direction: FindDirection | undefined,
): FindMatch[] {
  const active =
    context.activeCell ??
    (context.selection === null
      ? null
      : { row: context.selection.startRow, col: context.selection.startCol });
  if (active === null) return [...matches];

  const exact = matches.findIndex(
    (match) =>
      match.sheetId === context.activeSheetId &&
      match.row === active.row &&
      match.col === active.col,
  );
  if (exact >= 0) return [...matches.slice(exact), ...matches.slice(0, exact)];

  const activeSheetIndex = sheetIndexOf(context, context.activeSheetId);
  const next = matches.findIndex((match) => {
    const sheetIndex = sheetIndexOf(context, match.sheetId);
    const position =
      sheetIndex !== activeSheetIndex
        ? Math.sign(sheetIndex - activeSheetIndex)
        : match.row !== active.row
          ? Math.sign(match.row - active.row)
          : Math.sign(match.col - active.col);
    return direction === 'up' ? position < 0 : position > 0;
  });
  if (next < 0) return [...matches];
  return [...matches.slice(next), ...matches.slice(0, next)];
}

function sheetIndexOf(context: CommandContext, sheetId: string | null): number {
  return context.workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
}

function changeForMatch(
  workbook: Workbook,
  match: FindMatch,
  options: ReplaceOptions,
): MatchChange | null {
  const sheet = workbook.sheets.find((candidate) => candidate.id === match.sheetId);
  if (sheet === undefined) return null;
  const cell = getCell(sheet, match.row, match.col);
  if (cell === null) return null;
  const readOnly = readOnlyMessage(sheet, cell, match);
  if (readOnly !== null) return { kind: 'read-only', message: readOnly };
  const before = cloneCell(cell);
  const outcome = replaceCellMatches(cell, options);
  if (outcome.kind === 'no-change') return null;
  if (outcome.kind === 'unsafe') return { kind: 'unsafe', message: outcome.message };
  const after = outcome.kind === 'clear' ? null : outcome.cell;
  if (cellStateEquals(before, after)) return null;
  const range = singleCellRange(match.row, match.col);
  return {
    kind: 'change',
    change: {
      type: 'cell.set',
      sheetId: sheet.id,
      range,
      before: [[before]],
      after: [[after]],
    },
  };
}

/**
 * The same read-only rules as `cell.set`: a legacy array's whole `ref` range,
 * including follower cells that carry no marker of their own, plus shared
 * formulas and rich text. Refusing here keeps a replace from dirtying a cell
 * the save path would reject.
 */
function readOnlyMessage(sheet: Sheet, cell: Cell, match: FindMatch): string | null {
  if (cell.formulaSource?.kind === 'array' || arrayFormulaForCell(sheet, match.row, match.col)) {
    return ARRAY_READ_ONLY;
  }
  if (
    cell.formulaSource?.kind === 'shared-master' ||
    cell.formulaSource?.kind === 'shared-follower'
  ) {
    return SHARED_READ_ONLY;
  }
  if (cell.richText !== null) return RICH_TEXT_READ_ONLY;
  return null;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
