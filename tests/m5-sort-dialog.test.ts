import { describe, expect, it } from 'vitest';
import {
  addSortLevel,
  copySortLevel,
  createSortDraft,
  deleteSortLevel,
  moveSortLevel,
  setSortHeaders,
  sortColumnLabel,
  sortDraftProblem,
  sortKeys,
  sortLevelLabel,
  updateSortLevel,
} from '@chrome/sort-dialog-state';
import type { SortDialogColumn } from '@chrome/sort-dialog-state';
import { MAX_SORT_KEYS } from '@model/sort';

/**
 * The Sort dialog's level-list rules (sort-autofilter §2.3, design-refs
 * m5-menus-dialogs §1). Pure state only: the dialog renders this and the
 * command owns the transaction, so add/delete/copy/move and the 64-key limit
 * are proven here rather than through the DOM.
 */

const COLUMNS: readonly SortDialogColumn[] = [
  { column: 0, letter: 'A', header: 'Dept' },
  { column: 1, letter: 'B', header: 'City' },
  { column: 2, letter: 'C', header: '' },
];

describe('Sort dialog draft', () => {
  it('starts with one ascending key on the first column of the range', () => {
    const draft = createSortDraft(COLUMNS, true);
    expect(draft.levels).toEqual([{ column: 0, direction: 'ascending' }]);
    expect(draft.hasHeaders).toBe(true);
  });

  it('labels the first row Sort by and every later row Then by', () => {
    expect(sortLevelLabel(0)).toBe('Sort by');
    expect(sortLevelLabel(1)).toBe('Then by');
    expect(sortLevelLabel(4)).toBe('Then by');
  });

  it('uses the header text with headers on and the column letters with them off', () => {
    expect(sortColumnLabel(COLUMNS, 0, true)).toBe('Dept');
    expect(sortColumnLabel(COLUMNS, 0, false)).toBe('A');
    // A non-text header cell falls back to its letters, as `keyLabelsFor` does.
    expect(sortColumnLabel(COLUMNS, 2, true)).toBe('C');
  });

  it('appends a level inheriting the last column and drops keys outside the range', () => {
    let draft = createSortDraft(COLUMNS, true, [
      { column: 1, direction: 'descending' },
      { column: 9, direction: 'ascending' },
    ]);
    expect(draft.levels).toEqual([{ column: 1, direction: 'descending' }]);

    draft = addSortLevel(draft, COLUMNS);
    expect(draft.levels).toEqual([
      { column: 1, direction: 'descending' },
      { column: 1, direction: 'ascending' },
    ]);
  });

  it('deletes, copies and moves the selected level without losing the last one', () => {
    let draft = createSortDraft(COLUMNS, true, [
      { column: 0, direction: 'ascending' },
      { column: 1, direction: 'descending' },
      { column: 2, direction: 'ascending' },
    ]);

    draft = copySortLevel(draft, 1);
    expect(draft.levels.map((level) => level.column)).toEqual([0, 1, 1, 2]);

    draft = moveSortLevel(draft, 2, -1);
    expect(draft.levels.map((level) => level.column)).toEqual([0, 1, 1, 2]);
    draft = moveSortLevel(draft, 0, -1);
    expect(draft.levels.map((level) => level.column)).toEqual([0, 1, 1, 2]);

    draft = deleteSortLevel(draft, 2);
    expect(draft.levels.map((level) => level.column)).toEqual([0, 1, 2]);
    draft = deleteSortLevel(draft, 2);
    draft = deleteSortLevel(draft, 1);
    // The table always keeps one key, like Excel's Delete button.
    expect(draft.levels).toHaveLength(1);
    expect(deleteSortLevel(draft, 0)).toBe(draft);
  });

  it('stops at the 64-key limit', () => {
    let draft = createSortDraft(COLUMNS, true);
    while (draft.levels.length < MAX_SORT_KEYS) draft = addSortLevel(draft, COLUMNS);
    expect(draft.levels).toHaveLength(MAX_SORT_KEYS);
    expect(addSortLevel(draft, COLUMNS)).toBe(draft);
    expect(copySortLevel(draft, 0)).toBe(draft);
  });

  it('changes a level column or direction and the header override', () => {
    let draft = createSortDraft(COLUMNS, true);
    draft = updateSortLevel(draft, 0, { column: 2, direction: 'descending' });
    expect(draft.levels[0]).toEqual({ column: 2, direction: 'descending' });
    draft = setSortHeaders(draft, false);
    expect(draft.hasHeaders).toBe(false);
    // An out-of-bounds update is ignored rather than throwing.
    expect(updateSortLevel(draft, 5, { column: 1 })).toBe(draft);
  });

  it('submits the keys in level order', () => {
    const draft = createSortDraft(COLUMNS, false, [
      { column: 1, direction: 'descending' },
      { column: 0, direction: 'ascending' },
    ]);
    expect(sortKeys(draft)).toEqual([
      { column: 1, direction: 'descending' },
      { column: 0, direction: 'ascending' },
    ]);
  });

  it('reports the problems the command would reject', () => {
    expect(sortDraftProblem(createSortDraft([], true), [])).toBe(
      'The sort range has no columns to sort by.',
    );
    expect(sortDraftProblem({ levels: [], hasHeaders: true }, COLUMNS)).toBe(
      'Add at least one sort level.',
    );
    expect(
      sortDraftProblem(
        { levels: [{ column: 7, direction: 'ascending' }], hasHeaders: true },
        COLUMNS,
      ),
    ).toBe('Every sort key must be a column inside the sorted range.');
    expect(sortDraftProblem(createSortDraft(COLUMNS, true), COLUMNS)).toBeNull();
  });
});
