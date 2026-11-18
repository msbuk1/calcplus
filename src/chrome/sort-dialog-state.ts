import { MAX_SORT_KEYS } from '@model/sort';
import type { SortDirection, SortKey } from '@model/sort';

/**
 * The Sort dialog's draft state (sort-autofilter §2, design-refs
 * m5-menus-dialogs §1). The dialog owns the form; the pure helpers here own the
 * level list so the add/delete/copy/move rules and the 64-key limit are
 * testable without a DOM.
 */

/** One selectable column of the sort target range. */
export interface SortDialogColumn {
  /** Absolute model column index. */
  readonly column: number;
  /** Column letters, used when headers are off or the header cell is not text. */
  readonly letter: string;
  /** The header text, or `''` when the first-row cell is not text. */
  readonly header: string;
}

/** One sort level: a column inside the range and a direction. */
export interface SortLevel {
  readonly column: number;
  readonly direction: SortDirection;
}

export interface SortDraft {
  readonly levels: readonly SortLevel[];
  readonly hasHeaders: boolean;
}

/**
 * The capture's Order dropdown options. The Sort On dropdown ships with one
 * option: colour and icon sorts are v2 (§1, §6).
 */
export const SORT_DIRECTION_OPTIONS: readonly {
  readonly value: SortDirection;
  readonly label: string;
}[] = [
  { value: 'ascending', label: 'Sort Ascending' },
  { value: 'descending', label: 'Sort Descending' },
];

/** The single §2.3 sort-on type. */
export const SORT_ON_LABEL = 'Cell Values';

/** The first key row is "Sort by"; every later one is "Then by" (§1). */
export function sortLevelLabel(index: number): string {
  return index === 0 ? 'Sort by' : 'Then by';
}

/**
 * A draft with the given keys, or one ascending key on the first column. Keys
 * outside the range are dropped; the command rejects them anyway (§2.3).
 */
export function createSortDraft(
  columns: readonly SortDialogColumn[],
  hasHeaders: boolean,
  keys: readonly SortLevel[] = [],
): SortDraft {
  const known = new Set(columns.map((column) => column.column));
  const levels = keys
    .filter((key) => known.has(key.column))
    .slice(0, MAX_SORT_KEYS)
    .map((key) => ({ column: key.column, direction: key.direction }));
  if (levels.length === 0) {
    const first = columns[0];
    if (first !== undefined) levels.push({ column: first.column, direction: 'ascending' });
  }
  return { levels, hasHeaders };
}

/** The label a Column dropdown shows for one level (§2.2: header text or letters). */
export function sortColumnLabel(
  columns: readonly SortDialogColumn[],
  column: number,
  hasHeaders: boolean,
): string {
  const match = columns.find((candidate) => candidate.column === column);
  if (match === undefined) return '';
  if (hasHeaders && match.header !== '') return match.header;
  return match.letter;
}

/** Append a level, inheriting the last column and starting ascending. */
export function addSortLevel(draft: SortDraft, columns: readonly SortDialogColumn[]): SortDraft {
  if (draft.levels.length >= MAX_SORT_KEYS) return draft;
  const last = draft.levels[draft.levels.length - 1];
  const column = last?.column ?? columns[0]?.column;
  if (column === undefined) return draft;
  return { ...draft, levels: [...draft.levels, { column, direction: 'ascending' }] };
}

/** Remove one level; the table always keeps at least one. */
export function deleteSortLevel(draft: SortDraft, index: number): SortDraft {
  if (draft.levels.length <= 1 || index < 0 || index >= draft.levels.length) return draft;
  return { ...draft, levels: draft.levels.filter((_, position) => position !== index) };
}

/** Duplicate one level directly below itself, as the capture's Copy does. */
export function copySortLevel(draft: SortDraft, index: number): SortDraft {
  const level = draft.levels[index];
  if (level === undefined || draft.levels.length >= MAX_SORT_KEYS) return draft;
  const levels = [...draft.levels];
  levels.splice(index + 1, 0, { ...level });
  return { ...draft, levels };
}

/** Move one level up (`offset` -1) or down (1), clamped at the ends. */
export function moveSortLevel(draft: SortDraft, index: number, offset: 1 | -1): SortDraft {
  const target = index + offset;
  if (index < 0 || index >= draft.levels.length) return draft;
  if (target < 0 || target >= draft.levels.length) return draft;
  const levels = [...draft.levels];
  const [level] = levels.splice(index, 1);
  if (level === undefined) return draft;
  levels.splice(target, 0, level);
  return { ...draft, levels };
}

/** Change one level's column or direction. */
export function updateSortLevel(
  draft: SortDraft,
  index: number,
  patch: Partial<SortLevel>,
): SortDraft {
  if (index < 0 || index >= draft.levels.length) return draft;
  const levels = draft.levels.map((level, position) =>
    position === index ? { ...level, ...patch } : level,
  );
  return { ...draft, levels };
}

export function setSortHeaders(draft: SortDraft, hasHeaders: boolean): SortDraft {
  return { ...draft, hasHeaders };
}

/** The keys in level order, as `data.sort` takes them (§2.3). */
export function sortKeys(draft: SortDraft): readonly SortKey[] {
  return draft.levels.map((level) => ({ column: level.column, direction: level.direction }));
}

/**
 * Why the draft cannot be submitted, or `null`. The command re-checks every
 * rule; the dialog validates first so the user sees the problem in place.
 */
export function sortDraftProblem(
  draft: SortDraft,
  columns: readonly SortDialogColumn[],
): string | null {
  if (columns.length === 0) return 'The sort range has no columns to sort by.';
  if (draft.levels.length === 0) return 'Add at least one sort level.';
  if (draft.levels.length > MAX_SORT_KEYS) {
    return `A sort can use at most ${MAX_SORT_KEYS} keys.`;
  }
  const known = new Set(columns.map((column) => column.column));
  if (draft.levels.some((level) => !known.has(level.column))) {
    return 'Every sort key must be a column inside the sorted range.';
  }
  return null;
}
