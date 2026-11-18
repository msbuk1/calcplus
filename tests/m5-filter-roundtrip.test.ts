import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import type { AutoFilter } from '@model/types';
import { createAutoFilter, restoreHiddenLayers } from '@model/autofilter';
import { isUnsupportedFilterSource } from '@ooxml/autofilter-xml';
import { numberValue, textValue } from '@model/value';
import { DocumentSession } from '@commands/session';
import { buildXlsx, elementRanges, unzipPart, type BuildXlsxOptions } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The exact source text of the first `<autoFilter>` element in a worksheet part. */
function autoFilterSlice(worksheetXml: string): string | null {
  const ranges = elementRanges(worksheetXml, 'autoFilter');
  const first = ranges[0];
  return first === undefined ? null : worksheetXml.slice(first.start, first.end);
}

function worksheetPart(bytes: Uint8Array): string {
  return unzipPart(bytes, 'xl/worksheets/sheet1.xml');
}

function workbookPart(bytes: Uint8Array): string {
  return unzipPart(bytes, 'xl/workbook.xml');
}

function openFixture(worksheetXml: string, options: BuildXlsxOptions = {}) {
  return openWorkbook(buildXlsx(worksheetXml, options));
}

function filterDatabaseName(xml: string): string | null {
  const match =
    /<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*>([\s\S]*?)<\/definedName>/.exec(xml);
  return match === null ? null : match[1]!;
}

/** A plain, comparable view of the filter's document state. */
function filterState(filter: AutoFilter | null): unknown {
  if (filter === null) return null;
  return {
    range: { ...filter.range },
    columns: [...filter.columns.entries()]
      .sort((left, right) => left[0] - right[0])
      .map(([offset, column]) => {
        if (column.kind === 'values') {
          return [
            offset,
            { kind: column.kind, selected: column.selected, blanks: column.includeBlanks },
          ];
        }
        if (column.kind === 'conditions') {
          return [offset, { kind: column.kind, join: column.join, conditions: column.conditions }];
        }
        const source = isUnsupportedFilterSource(column.raw) ? column.raw : null;
        return [
          offset,
          { kind: column.kind, reason: source?.reason ?? null, inner: source?.inner ?? null },
        ];
      }),
  };
}

// ---------------------------------------------------------------------------
// F13 filter-roundtrip
// ---------------------------------------------------------------------------

const FILTER_DATABASE_EXTRAS = `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$B$6</definedName></definedNames>`;

/**
 * Rows 2-7. The filter selects A in {10, 30} and B > 5, so rows 2, 3, 5 and 6
 * fail it. Row 3 is deliberately not flagged hidden (a producer inconsistency
 * the open-side rebuild must absorb); row 4 passes but is manually hidden;
 * row 7 sits outside the range and is manually hidden.
 */
const ROUNDTRIP_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:B7"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2" hidden="1"><c r="A2"><v>10</v></c><c r="B2"><v>4</v></c></row>
    <row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>6</v></c></row>
    <row r="4" hidden="1"><c r="A4"><v>30</v></c><c r="B4"><v>8</v></c></row>
    <row r="5" hidden="1"><c r="A5"><v>40</v></c><c r="B5"><v>2</v></c></row>
    <row r="6" hidden="1"><c r="A6"><v>50</v></c><c r="B6"><v>9</v></c></row>
    <row r="7" hidden="1"><c r="A7"><v>60</v></c><c r="B7"><v>1</v></c></row>
  </sheetData>
  <autoFilter ref="A1:B6" foo="bar"><filterColumn colId="0"><filters><filter val="10"/><filter val="30"/></filters></filterColumn><filterColumn colId="1" hiddenButton="1"><customFilters and="0"><customFilter operator="greaterThan" val="5"/></customFilters></filterColumn><sortState ref="A2:B6"><sortCondition ref="B2:B6" descending="1"/></sortState></autoFilter>
</worksheet>`;

function openRoundTrip() {
  return openFixture(ROUNDTRIP_WORKSHEET, {
    sharedStrings: ['Name', 'Value'],
    workbookExtras: FILTER_DATABASE_EXTRAS,
  });
}

describe('F13 filter-roundtrip', () => {
  it('reads the autoFilter into the model and rebuilds the visibility layers', () => {
    const opened = openRoundTrip();
    const sheet = opened.workbook.sheets[0]!;
    const filter = sheet.autoFilter;
    expect(filter).not.toBeNull();
    expect(filter!.range).toEqual({ startRow: 0, startCol: 0, endRow: 5, endCol: 1 });

    const values = filter!.columns.get(0);
    expect(values).toEqual({
      kind: 'values',
      selected: [numberValue(10), numberValue(30)],
      includeBlanks: false,
    });
    expect(filter!.columns.get(1)).toEqual({
      kind: 'conditions',
      join: 'or',
      conditions: [{ kind: 'number', operator: '>', value: 5 }],
    });

    // §3.3: the saved flags split into the filter layer and the manual layer.
    const savedHidden = [1, 3, 4, 5, 6];
    const layers = restoreHiddenLayers(sheet, filter!, savedHidden);
    expect([...layers.filterHidden].sort((left, right) => left - right)).toEqual([1, 2, 4, 5]);
    expect([...layers.manualHidden].sort((left, right) => left - right)).toEqual([3, 6]);

    // Every row the criteria hide is hidden in the model, even the row the
    // file forgot to flag.
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);
    expect(sheet.rowDimensions.get(6)?.hidden).toBe(true);
  });

  it('leaves an untouched autoFilter byte-identical on a no-op save', () => {
    const original = buildXlsx(ROUNDTRIP_WORKSHEET, {
      sharedStrings: ['Name', 'Value'],
      workbookExtras: FILTER_DATABASE_EXTRAS,
    });
    const opened = openWorkbook(original);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(result.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(result.bytes), Buffer.from(original))).toBe(0);
  });

  it('keeps the autoFilter, hidden rows and _FilterDatabase across an unrelated edit', () => {
    const opened = openRoundTrip();
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const before = worksheetPart(opened.package.originalBytes!);

    // C9 is outside the filter range and touches no filtered column.
    session.dispatch('cell.set', { sheetId: sheet.id, row: 8, col: 2, value: numberValue(7) });

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const after = worksheetPart(result.bytes);

    expect(autoFilterSlice(after)).toBe(autoFilterSlice(before));
    expect(autoFilterSlice(after)).toContain('foo="bar"');
    expect(after).toContain('<row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>6</v></c></row>');
    expect(after).toContain('<row r="4" hidden="1">');
    expect(after).toContain('<c r="C9"><v>7</v></c>');
    expect(filterDatabaseName(workbookPart(result.bytes))).toBe('Sheet1!$A$1:$B$6');
  });

  it('writes a changed filter in the §3.6 shape and retargets _FilterDatabase', () => {
    const opened = openRoundTrip();
    const sheet = opened.workbook.sheets[0]!;
    const filter = sheet.autoFilter!;

    // Keep the raw source, change one criterion (with blanks), the second
    // column to a between pair, and widen the range by one row.
    filter.columns.set(0, {
      kind: 'values',
      selected: [numberValue(10), numberValue(20)],
      includeBlanks: true,
    });
    filter.columns.set(1, {
      kind: 'conditions',
      join: 'and',
      conditions: [{ kind: 'number', operator: 'between', from: 3, to: 9 }],
    });
    sheet.autoFilter = { ...filter, range: { startRow: 0, startCol: 0, endRow: 6, endCol: 1 } };

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    const part = worksheetPart(result.bytes);

    expect(autoFilterSlice(part)).toBe(
      '<autoFilter ref="A1:B7" foo="bar">' +
        '<filterColumn colId="0"><filters blank="1"><filter val="10"/><filter val="20"/></filters></filterColumn>' +
        '<filterColumn colId="1" hiddenButton="1"><customFilters and="1">' +
        '<customFilter operator="greaterThanOrEqual" val="3"/>' +
        '<customFilter operator="lessThanOrEqual" val="9"/></customFilters></filterColumn>' +
        '<sortState ref="A2:B6"><sortCondition ref="B2:B6" descending="1"/></sortState>' +
        '</autoFilter>',
    );
    // Row 2 now passes, so its stale hidden flag is removed; row 4 stays hidden.
    expect(part).toContain('<row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>4</v></c></row>');
    expect(part).toContain('<row r="4" hidden="1">');
    expect(filterDatabaseName(workbookPart(result.bytes))).toBe('Sheet1!$A$1:$B$7');

    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(filterState(reopenedSheet.autoFilter)).toEqual({
      range: { startRow: 0, startCol: 0, endRow: 6, endCol: 1 },
      columns: [
        [
          0,
          {
            kind: 'values',
            selected: [numberValue(10), numberValue(20)],
            blanks: true,
          },
        ],
        [
          1,
          {
            kind: 'conditions',
            join: 'and',
            conditions: [{ kind: 'number', operator: 'between', from: 3, to: 9 }],
          },
        ],
      ],
    });
    expect(reopenedSheet.names.map((name) => [name.name, name.formula])).toContainEqual([
      '_xlnm._FilterDatabase',
      'Sheet1!$A$1:$B$7',
    ]);
    // The reopened layers follow the saved criteria.
    const hidden = [...reopenedSheet.rowDimensions.entries()]
      .filter(([, dimension]) => dimension.hidden === true)
      .map(([row]) => row)
      .sort((left, right) => left - right);
    expect(hidden).toEqual([3, 4, 5, 6]);
  });

  it('creates the element and the _FilterDatabase name when the filter is switched on', () => {
    const opened = openFixture(
      `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>4</v></c></row>
    <row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>6</v></c></row>
    <row r="4"><c r="A4"><v>30</v></c><c r="B4"><v>8</v></c></row>
  </sheetData>
</worksheet>`,
      { sharedStrings: ['Name', 'Value'] },
    );
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.autoFilter).toBeNull();
    sheet.autoFilter = createAutoFilter(
      sheet.id,
      { startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
      [[0, { kind: 'values', selected: [numberValue(10)], includeBlanks: false }]],
    );

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    const part = worksheetPart(result.bytes);
    expect(autoFilterSlice(part)).toBe(
      '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="10"/></filters></filterColumn></autoFilter>',
    );
    expect(part.indexOf('<autoFilter')).toBeGreaterThan(part.indexOf('</sheetData>'));
    // §3.3: the effective hidden flag lands in the file for the rows the new
    // criteria hide, while the row that passes stays untouched.
    expect(part).toContain('<row r="2"><c r="A2"><v>10</v></c>');
    expect(part).toContain('<row r="3" hidden="1"><c r="A3"><v>20</v></c>');
    expect(part).toContain('<row r="4" hidden="1"><c r="A4"><v>30</v></c>');
    expect(filterDatabaseName(workbookPart(result.bytes))).toBe("'Sheet1'!$A$1:$B$4");

    const reopened = openWorkbook(result.bytes).workbook.sheets[0]!;
    expect(reopened.autoFilter?.columns.get(0)).toEqual({
      kind: 'values',
      selected: [numberValue(10)],
      includeBlanks: false,
    });
    expect(reopened.names.map((name) => name.name)).toContain('_xlnm._FilterDatabase');
  });

  it('removes the element and the name when the filter is cleared', () => {
    const opened = openRoundTrip();
    const sheet = opened.workbook.sheets[0]!;
    sheet.autoFilter = null;

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(autoFilterSlice(worksheetPart(result.bytes))).toBeNull();
    expect(filterDatabaseName(workbookPart(result.bytes))).toBeNull();
    // Rows only the old criteria hid become visible; manual rows stay hidden.
    const part = worksheetPart(result.bytes);
    expect(part).toContain('<row r="2"><c r="A2"><v>10</v></c>');
    expect(part).toContain('<row r="3"><c r="A3"><v>20</v></c>');
    expect(part).toContain('<row r="4" hidden="1">');
    expect(part).toContain('<row r="7" hidden="1">');

    const reopened = openWorkbook(result.bytes).workbook.sheets[0]!;
    expect(reopened.autoFilter).toBeNull();
    expect(reopened.names.map((name) => name.name)).not.toContain('_xlnm._FilterDatabase');
  });

  it('reads a prefixed worksheet part and writes a changed filter back', () => {
    const prefixed = `${HEADER}
<x:worksheet xmlns:x="${NS}" xmlns:r="${REL_NS}">
  <x:sheetData>
    <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c></x:row>
    <x:row r="2"><x:c r="A2"><x:v>10</x:v></x:c></x:row>
    <x:row r="3"><x:c r="A3"><x:v>20</x:v></x:c></x:row>
  </x:sheetData>
  <x:autoFilter ref="A1:A3"><x:filterColumn colId="0"><x:filters><x:filter val="10"/></x:filters></x:filterColumn></x:autoFilter>
</x:worksheet>`;
    const opened = openWorkbook(buildXlsx(prefixed));
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.autoFilter?.columns.get(0)).toEqual({
      kind: 'values',
      selected: [numberValue(10)],
      includeBlanks: false,
    });

    sheet.autoFilter = {
      ...sheet.autoFilter!,
      columns: new Map([
        [0, { kind: 'values' as const, selected: [numberValue(20)], includeBlanks: false }],
      ]),
    };
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    const part = worksheetPart(result.bytes);
    expect(part).toContain(
      '<x:autoFilter ref="A1:A3"><x:filterColumn colId="0"><x:filters><x:filter val="20"/></x:filters></x:filterColumn></x:autoFilter>',
    );
    expect(openWorkbook(result.bytes).workbook.sheets[0]!.autoFilter?.columns.get(0)).toEqual({
      kind: 'values',
      selected: [numberValue(20)],
      includeBlanks: false,
    });
  });

  it('writes a changed filter through the tree fallback too', () => {
    const opened = openRoundTrip();
    const sheet = opened.workbook.sheets[0]!;
    const filter = sheet.autoFilter!;
    filter.columns.set(0, { kind: 'values', selected: [numberValue(10)], includeBlanks: false });

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
      options: { worksheet: { forceTreeFallback: true } },
    });
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('tree');
    const slice = autoFilterSlice(worksheetPart(result.bytes));
    expect(slice).toContain('<autoFilter ref="A1:B6" foo="bar">');
    expect(slice).toContain('<filterColumn colId="0"><filters><filter val="10"/></filters>');
    expect(slice).toContain('<sortCondition ref="B2:B6" descending="1"/>');
    expect(worksheetPart(result.bytes)).toContain('<row r="2" hidden="1">');
  });

  it('writes a filter on a sheet inserted in the same save', () => {
    const opened = openFixture(
      `${HEADER}
<worksheet xmlns="${NS}">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
    );
    const session = new DocumentSession(opened.workbook);
    const inserted = session.dispatch('sheet.insert', { name: 'Filtered' });
    expect(inserted.changes.length).toBeGreaterThan(0);
    const sheet = opened.workbook.sheets.find((candidate) => candidate.name === 'Filtered')!;
    for (const [row, value] of [
      [0, 'Name'],
      [1, 10],
    ] as const) {
      session.dispatch('cell.set', {
        sheetId: sheet.id,
        row,
        col: 0,
        value: typeof value === 'number' ? numberValue(value) : textValue(value),
      });
    }
    sheet.autoFilter = createAutoFilter(
      sheet.id,
      { startRow: 0, startCol: 0, endRow: 2, endCol: 0 },
      [[0, { kind: 'values', selected: [numberValue(10)], includeBlanks: false }]],
    );

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const newPart = unzipPart(result.bytes, 'xl/worksheets/sheet2.xml');
    expect(autoFilterSlice(newPart)).toContain('<filterColumn colId="0">');
    expect(filterDatabaseName(workbookPart(result.bytes))).toBe("'Filtered'!$A$1:$A$3");
    const reopenedWorkbook = openWorkbook(result.bytes).workbook;
    const reopened = reopenedWorkbook.sheets[1]!;
    expect(reopened.autoFilter?.range).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 0,
    });
    expect(reopened.names.map((name) => [name.name, name.formula])).toContainEqual([
      '_xlnm._FilterDatabase',
      "'Filtered'!$A$1:$A$3",
    ]);
    // The name is sheet-scoped to the inserted sheet, not to the first one.
    expect(reopenedWorkbook.sheets[0]!.names.map((name) => name.name)).not.toContain(
      '_xlnm._FilterDatabase',
    );
  });
});

// ---------------------------------------------------------------------------
// F14 filter-unsupported
// ---------------------------------------------------------------------------

const UNSUPPORTED_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:C4"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
    <row r="2" hidden="1"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>3</v></c><c r="C2"><v>7</v></c></row>
    <row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3"><v>1</v></c><c r="C3"><v>9</v></c></row>
    <row r="4" hidden="1"><c r="A4" t="s"><v>5</v></c><c r="B4"><v>2</v></c><c r="C4"><v>4</v></c></row>
  </sheetData>
  <autoFilter ref="A1:C4"><filterColumn colId="0"><colorFilter dxfId="0" cellColor="1"/></filterColumn><filterColumn colId="1"><dynamicFilter type="thisMonth"/></filterColumn><filterColumn colId="2"><top10 val="10" top="1"/></filterColumn></autoFilter>
</worksheet>`;

const UNSUPPORTED_OPTIONS: BuildXlsxOptions = {
  sharedStrings: ['Name', 'Value', 'Score', 'a', 'b', 'c'],
  workbookExtras: `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$C$4</definedName></definedNames>`,
};

describe('F14 filter-unsupported', () => {
  it('reads colour, dynamic and Top 10 criteria as unsupported and keeps them verbatim', () => {
    const opened = openFixture(UNSUPPORTED_WORKSHEET, UNSUPPORTED_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    const filter = sheet.autoFilter!;
    const reasons = [0, 1, 2].map((offset) => {
      const column = filter.columns.get(offset);
      expect(column?.kind).toBe('unsupported');
      if (column?.kind !== 'unsupported') return null;
      return isUnsupportedFilterSource(column.raw) ? column.raw.reason : null;
    });
    expect(reasons).toEqual(['colour', 'dynamic', 'top10']);

    // Their hidden rows still render hidden: the filter layer cannot evaluate
    // them, so the imported flags stay the manual layer (§3.7).
    const layers = restoreHiddenLayers(sheet, filter, [1, 3]);
    expect(layers.filterHidden.size).toBe(0);
    expect([...layers.effectiveHidden].sort((left, right) => left - right)).toEqual([1, 3]);
  });

  it('preserves the criteria byte-identically on a no-op save and after an edit', () => {
    const original = buildXlsx(UNSUPPORTED_WORKSHEET, UNSUPPORTED_OPTIONS);
    const opened = openWorkbook(original);
    const noOp = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(Buffer.compare(Buffer.from(noOp.bytes), Buffer.from(original))).toBe(0);

    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 2, col: 2, value: numberValue(11) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(autoFilterSlice(worksheetPart(edited.bytes))).toBe(
      autoFilterSlice(worksheetPart(original)),
    );
  });

  it('keeps an unsupported column verbatim when another column is filtered', () => {
    const opened = openFixture(UNSUPPORTED_WORKSHEET, UNSUPPORTED_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    const filter = sheet.autoFilter!;
    filter.columns.set(0, {
      kind: 'values',
      selected: [textValue('b')],
      includeBlanks: false,
    });

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(autoFilterSlice(worksheetPart(result.bytes))).toBe(
      '<autoFilter ref="A1:C4">' +
        '<filterColumn colId="0"><filters><filter val="b"/></filters></filterColumn>' +
        '<filterColumn colId="1"><dynamicFilter type="thisMonth"/></filterColumn>' +
        '<filterColumn colId="2"><top10 val="10" top="1"/></filterColumn>' +
        '</autoFilter>',
    );

    const reopened = openWorkbook(result.bytes).workbook.sheets[0]!;
    const reopenedFilter = reopened.autoFilter!;
    expect(reopenedFilter.columns.get(0)).toEqual({
      kind: 'values',
      selected: [textValue('b')],
      includeBlanks: false,
    });
    for (const [offset, detail] of [
      [1, 'dynamic'],
      [2, 'top10'],
    ] as const) {
      const column = reopenedFilter.columns.get(offset);
      expect(column?.kind).toBe('unsupported');
      if (column?.kind !== 'unsupported') throw new Error(`column ${offset} is not unsupported`);
      const source = isUnsupportedFilterSource(column.raw) ? column.raw : null;
      expect(source?.reason).toBe(detail);
      if (offset === 1) expect(source?.inner).toBe('<dynamicFilter type="thisMonth"/>');
    }
  });
});

// ---------------------------------------------------------------------------
// F15 filter-table
// ---------------------------------------------------------------------------

const TABLE_PART = `${HEADER}
<table xmlns="${NS}" id="1" name="Table1" displayName="Table1" ref="A1:B4" totalsRowShown="0">
  <autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="10"/></filters></filterColumn></autoFilter>
  <tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Value"/></tableColumns>
</table>`;

const TABLE_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>4</v></c></row>
    <row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>6</v></c></row>
    <row r="4"><c r="A4"><v>30</v></c><c r="B4"><v>8</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`;

const TABLE_OPTIONS: BuildXlsxOptions = {
  sharedStrings: ['Name', 'Value'],
  extraParts: {
    'xl/tables/table1.xml': TABLE_PART,
    'xl/worksheets/_rels/sheet1.xml.rels': `${HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_NS}/table" Target="../tables/table1.xml"/>
</Relationships>`,
  },
  contentTypeOverrides: [
    [
      'xl/tables/table1.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    ],
  ],
};

describe('F15 filter-table', () => {
  it('does not read a table filter into the sheet and preserves the table part', () => {
    const original = buildXlsx(TABLE_WORKSHEET, TABLE_OPTIONS);
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    // A table's own filter lives in the table part and is not editable (§3.6).
    expect(sheet.autoFilter).toBeNull();

    const noOp = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(Buffer.compare(Buffer.from(noOp.bytes), Buffer.from(original))).toBe(0);

    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 2, col: 1, value: numberValue(7) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(edited.bytes, 'xl/tables/table1.xml')).toBe(TABLE_PART);
    expect(worksheetPart(edited.bytes)).toContain('<tableParts count="1">');
    expect(autoFilterSlice(worksheetPart(edited.bytes))).toBeNull();
  });

  it('leaves the table part alone when the sheet has its own filter', () => {
    const original = buildXlsx(
      ROUNDTRIP_WORKSHEET.replace(
        '  <autoFilter',
        '  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>\n  <autoFilter',
      ),
      {
        ...TABLE_OPTIONS,
        sharedStrings: ['Name', 'Value'],
        workbookExtras: FILTER_DATABASE_EXTRAS,
      },
    );
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.autoFilter).not.toBeNull();

    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 8, col: 2, value: numberValue(7) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(edited.bytes, 'xl/tables/table1.xml')).toBe(TABLE_PART);
    expect(autoFilterSlice(worksheetPart(edited.bytes))).toBe(
      autoFilterSlice(worksheetPart(original)),
    );
    // Only one `autoFilter` per sheet: the worksheet carries it, not the table.
    expect(elementRanges(worksheetPart(edited.bytes), 'autoFilter')).toHaveLength(1);
  });
});
