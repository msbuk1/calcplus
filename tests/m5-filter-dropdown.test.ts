import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import {
  buildConditions,
  checklistProblem,
  conditionsDraftFrom,
  criterionFromChecklist,
  defaultConditionsDraft,
  detectFilterColumnType,
  filterMenuLabel,
  filterRecordCount,
  initialChecklistState,
  isDateFormatCode,
  operatorAcceptsRange,
  selectAllState,
  toggleChecklistBlanks,
  toggleChecklistEntry,
  toggleSelectAll,
} from '@chrome/filter-dropdown-state';
import type { ConditionsDraft } from '@chrome/filter-dropdown-state';
import { collectValueChecklist, createAutoFilter } from '@model/autofilter';
import type { ChecklistEntry, ColumnFilter, ValueChecklist } from '@model/autofilter';
import type { GridRange, Sheet } from '@model/types';
import { createCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The AutoFilter dropdown's state (sort-autofilter §3.2, §3.8): checklist
 * selection, Select All, the criteria Apply builds, the type-dependent
 * condition windows and the status bar's `X of Y records`. The checklist
 * itself comes from the model, so these tests use the real
 * `collectValueChecklist` rather than a hand-built entry list.
 */

function workbookWithSheet() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet };
}

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

function put(sheet: Sheet, row: number, col: number, value: CellValue): void {
  setCellSnapshot(sheet, row, col, createCell(value));
}

/** A1:B6: a header and five city rows, with one blank score. */
function citySheet() {
  const { workbook, sheet } = workbookWithSheet();
  put(sheet, 0, 0, textValue('City'));
  put(sheet, 0, 1, textValue('Sales'));
  put(sheet, 1, 0, textValue('Bristol'));
  put(sheet, 1, 1, numberValue(120));
  put(sheet, 2, 0, textValue('Bath'));
  put(sheet, 2, 1, numberValue(80));
  put(sheet, 3, 0, textValue('Leeds'));
  put(sheet, 3, 1, numberValue(200));
  put(sheet, 4, 0, textValue('York'));
  put(sheet, 5, 0, textValue('Cork'));
  put(sheet, 5, 1, numberValue(60));
  return { workbook, sheet };
}

function checklistFor(sheet: Sheet, columnOffset: number, endRow = 5): ValueChecklist {
  const filter = createAutoFilter(sheet.id, range(0, 0, endRow, 1));
  return collectValueChecklist(sheet, filter, columnOffset);
}

function valuesFilter(...values: CellValue[]): ColumnFilter {
  return { kind: 'values', selected: values, includeBlanks: false };
}

describe('value checklist state', () => {
  it('ticks every entry for a column with no criterion', () => {
    const { sheet } = citySheet();
    const checklist = checklistFor(sheet, 0);
    const state = initialChecklistState(checklist.entries, checklist.hasBlanks, null);
    expect(checklist.entries.map((entry) => entry.display)).toEqual([
      'Bath',
      'Bristol',
      'Cork',
      'Leeds',
      'York',
    ]);
    expect([...state.checked]).toEqual([0, 1, 2, 3, 4]);
    expect(selectAllState(state, checklist.entries)).toEqual({ all: true, some: false });
    // Every value is visible, so Apply means "no column filter".
    expect(criterionFromChecklist(state, checklist.entries)).toBeNull();
  });

  it('ticks the entries an existing value criterion selects', () => {
    const { sheet } = citySheet();
    const checklist = checklistFor(sheet, 0);
    const state = initialChecklistState(
      checklist.entries,
      checklist.hasBlanks,
      valuesFilter(textValue('Bath'), textValue('York')),
    );
    expect([...state.checked]).toEqual([0, 4]);
    expect(selectAllState(state, checklist.entries)).toEqual({ all: false, some: true });
    const criterion = criterionFromChecklist(state, checklist.entries);
    expect(criterion).toEqual({
      kind: 'values',
      selected: [textValue('Bath'), textValue('York')],
      includeBlanks: false,
    });
  });

  it('leaves a conditions criterion fully ticked and builds a value criterion from toggles', () => {
    const { sheet } = citySheet();
    const checklist = checklistFor(sheet, 0);
    const conditions: ColumnFilter = {
      kind: 'conditions',
      join: 'and',
      conditions: [{ kind: 'text', operator: 'contains', value: 'o' }],
    };
    let state = initialChecklistState(checklist.entries, checklist.hasBlanks, conditions);
    expect(state.checked.size).toBe(checklist.entries.length);

    state = toggleSelectAll(state, checklist.entries);
    expect(state.checked.size).toBe(0);
    state = toggleChecklistEntry(state, 2);
    expect(criterionFromChecklist(state, checklist.entries)).toEqual({
      kind: 'values',
      selected: [textValue('Cork')],
      includeBlanks: false,
    });
    expect(checklistProblem(state, checklist.entries)).toBeNull();

    state = toggleChecklistEntry(state, 2);
    expect(checklistProblem(state, checklist.entries)).toBe(
      'Select at least one value or (Blanks) before applying the filter.',
    );
  });

  it('offers (Blanks) only when a data row is blank, and applies it with the checklist', () => {
    const { sheet } = citySheet();
    const checklist = checklistFor(sheet, 1);
    expect(checklist.hasBlanks).toBe(true);

    let state = initialChecklistState(checklist.entries, checklist.hasBlanks, null);
    state = toggleSelectAll(state, checklist.entries);
    expect(state.blanksChecked).toBe(false);
    state = toggleChecklistBlanks(state);
    expect(criterionFromChecklist(state, checklist.entries)).toEqual({
      kind: 'values',
      selected: [],
      includeBlanks: true,
    });

    // Reselecting everything restores the "no filter" state, blanks included.
    const all = { ...state, checked: new Set(checklist.entries.map((_, index) => index)) };
    expect(selectAllState(all, checklist.entries).all).toBe(true);
    expect(criterionFromChecklist(all, checklist.entries)).toBeNull();
  });

  it('groups values that format alike into one entry', () => {
    const grouped: ChecklistEntry = {
      display: '10.00',
      values: [numberValue(10.001), numberValue(10.002)],
    };
    const other: ChecklistEntry = { display: '20.00', values: [numberValue(20)] };
    const state = { checked: new Set([0]), blanksChecked: false, hasBlanks: false };
    expect(criterionFromChecklist(state, [grouped, other])).toEqual({
      kind: 'values',
      selected: [numberValue(10.001), numberValue(10.002)],
      includeBlanks: false,
    });
  });
});

describe('filter type and operators', () => {
  it('picks Text, Number or Date from the column values', () => {
    const { sheet } = citySheet();
    const cities = checklistFor(sheet, 0);
    const sales = checklistFor(sheet, 1);
    expect(detectFilterColumnType(cities, false)).toBe('text');
    expect(detectFilterColumnType(sales, false)).toBe('number');
    expect(detectFilterColumnType(sales, true)).toBe('date');
    expect(filterMenuLabel('text')).toBe('Text Filters');
    expect(filterMenuLabel('number')).toBe('Number Filters');
    expect(filterMenuLabel('date')).toBe('Date Filters');
  });

  it('recognises date formats without mistaking quoted text for one', () => {
    expect(isDateFormatCode('yyyy-mm-dd')).toBe(true);
    expect(isDateFormatCode('d/m/yyyy')).toBe(true);
    expect(isDateFormatCode('h:mm:ss AM/PM')).toBe(true);
    expect(isDateFormatCode('#,##0.00')).toBe(false);
    expect(isDateFormatCode('0.00 "days"')).toBe(false);
    expect(isDateFormatCode('[Red]0.00')).toBe(false);
    expect(isDateFormatCode('@')).toBe(false);
    expect(isDateFormatCode(null)).toBe(false);
    expect(operatorAcceptsRange('between')).toBe(true);
    expect(operatorAcceptsRange('equals')).toBe(false);
  });
});

describe('Custom AutoFilter conditions', () => {
  it('starts with the first operator of the type and an empty second row', () => {
    expect(defaultConditionsDraft('text').first.operator).toBe('equals');
    expect(defaultConditionsDraft('number').first.operator).toBe('=');
    expect(defaultConditionsDraft('date').first.operator).toBe('equals');
  });

  it('builds a text condition and ignores a blank second row', () => {
    const draft: ConditionsDraft = {
      join: 'and',
      first: { operator: 'contains', value: 'is', to: '' },
      second: { operator: 'equals', value: '', to: '' },
    };
    const built = buildConditions(draft, 'text', false);
    expect(built).toEqual({
      ok: true,
      join: 'and',
      conditions: [{ kind: 'text', operator: 'contains', value: 'is' }],
    });
  });

  it('builds two number conditions joined by OR, including a range', () => {
    const draft: ConditionsDraft = {
      join: 'or',
      first: { operator: 'between', value: '50', to: '150' },
      second: { operator: '>', value: '200', to: '' },
    };
    expect(buildConditions(draft, 'number', false)).toEqual({
      ok: true,
      join: 'or',
      conditions: [
        { kind: 'number', operator: 'between', from: 50, to: 150 },
        { kind: 'number', operator: '>', value: 200 },
      ],
    });
  });

  it('rejects an unparseable number or an empty text value', () => {
    const draft = defaultConditionsDraft('number');
    const bad: ConditionsDraft = { ...draft, first: { operator: '>', value: 'ten', to: '' } };
    expect(buildConditions(bad, 'number', false)).toEqual({
      ok: false,
      message: 'Enter a number for the filter condition.',
    });
    const text = defaultConditionsDraft('text');
    expect(buildConditions(text, 'text', false)).toEqual({
      ok: false,
      message: 'Enter a value for the filter condition.',
    });
  });

  it('converts a date input in the workbook date system and round-trips it', () => {
    const draft: ConditionsDraft = {
      join: 'and',
      first: { operator: 'after', value: '2026-11-18', to: '' },
      second: { operator: 'equals', value: '', to: '' },
    };
    const built = buildConditions(draft, 'date', false);
    expect(built.ok).toBe(true);
    if (built.ok === false) return;
    const condition = built.conditions[0];
    expect(condition?.kind).toBe('date');
    if (condition?.kind !== 'date' || !('value' in condition)) return;
    // 2026-11-18 is serial 46344 in the 1900 system.
    expect(condition.value).toBe(46_344);
    const criterion: ColumnFilter = {
      kind: 'conditions',
      join: built.join,
      conditions: built.conditions,
    };
    const restored = conditionsDraftFrom(criterion, 'date', false);
    expect(restored?.first).toEqual({ operator: 'after', value: '2026-11-18', to: '' });
  });

  it('prefills the window from an existing conditions criterion', () => {
    const criterion: ColumnFilter = {
      kind: 'conditions',
      join: 'or',
      conditions: [
        { kind: 'number', operator: 'between', from: 10, to: 20 },
        { kind: 'number', operator: '<>', value: 15 },
      ],
    };
    expect(conditionsDraftFrom(criterion, 'number', false)).toEqual({
      join: 'or',
      first: { operator: 'between', value: '10', to: '20' },
      second: { operator: '<>', value: '15', to: '' },
    });
    expect(conditionsDraftFrom(valuesFilter(textValue('Bath')), 'text', false)).toBeNull();
    expect(conditionsDraftFrom(null, 'text', false)).toBeNull();
  });
});

describe('status bar records', () => {
  it('matches the hidden-row count the filter payload reports', () => {
    const { workbook, sheet } = citySheet();
    const session = new DocumentSession(workbook);
    session.dispatch('data.filter', { range: range(0, 0, 5, 1) }, { activeSheetId: sheet.id });
    const result = session.dispatch(
      'data.filterColumn',
      { column: 0, criterion: valuesFilter(textValue('Bath'), textValue('Cork')) },
      { activeSheetId: sheet.id },
    );
    const filter = sheet.autoFilter;
    expect(filter).not.toBeNull();
    if (filter === null) return;
    const records = filterRecordCount(sheet, filter);
    expect(records).toEqual({ visible: 2, total: 5 });
    expect(records.visible).toBe(result.filter?.rowsVisible);
    expect(records.total).toBe(result.filter?.rowsTotal);

    // With no criteria every data row is visible; the header never counts.
    session.dispatch('data.filterClear', undefined, { activeSheetId: sheet.id });
    const cleared = sheet.autoFilter;
    expect(cleared).not.toBeNull();
    if (cleared === null) return;
    expect(filterRecordCount(sheet, cleared)).toEqual({ visible: 5, total: 5 });
  });
});
