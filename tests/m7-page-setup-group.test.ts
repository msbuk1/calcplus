import { describe, expect, it } from 'vitest';
import {
  DEFAULT_PRINT_MARGINS,
  buildPageSetup,
  formatPrintAreaRanges,
  normalisePrintAreaText,
  pageSetupValuesOf,
  printAreaTextForRange,
} from '@commands/print';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { COMMAND_INVENTORY, inventoryCommand } from '@commands/inventory';
import { ribbonGroupsFor } from '@chrome/ribbon-model';
import {
  pageSetupDraftFrom,
  pageSetupDraftPayload,
  pageSetupPlanText,
  pageSetupValuesFromDraft,
} from '@chrome/page-setup-state';
import { MAX_ROW } from '@model/references';
import { openBlankWorkbook } from '@ooxml/blank';
import type { OpenedWorkbook } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { printAreaOf } from '@ooxml/print-setup-xml';
import { createSheetProjection } from '@view/projection';
import { createPrintPlan } from '../src/print/planner';
import { createPrintSnapshot } from '../src/print/adapter';
import type { PrintPlan } from '../src/print/types';
import { fixtureBytes, openFixture, partText } from './helpers/roundtrip';

/**
 * Ticket 44 — the Page Layout tab's Page Setup and Sheet Options groups.
 *
 * The map test proves the v2 items are hidden (they have no placement). The
 * command tests prove one Apply is one transaction, the print area round-trips
 * through the workbook part and undo restores the exact prior state. The
 * planner assertions prove printed gridlines and headings come from
 * `printOptions`, never from the screen toggles.
 */

const WORKBOOK_PART = 'xl/workbook.xml';

function planFor(opened: OpenedWorkbook): PrintPlan {
  const sheet = opened.workbook.sheets[0]!;
  const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });
  return createPrintPlan(createPrintSnapshot(opened.workbook, projection));
}

function save(opened: OpenedWorkbook) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: new Map(),
  });
}

function sheetOf(opened: OpenedWorkbook) {
  return opened.workbook.sheets[0]!;
}

describe('Page Layout map', () => {
  it('shows only the Page Setup and Sheet Options groups', () => {
    // Breaks, Themes, Background, Print Titles, Scale to Fit and Sheet
    // Right-to-Left are v2 or never: they have no placement, so they are
    // hidden rather than greyed (ui-chrome §4.1).
    expect(ribbonGroupsFor('Page Layout').map((group) => group.id)).toEqual([
      'Page Setup',
      'Sheet Options',
    ]);
  });

  it('places the dialog, the Print Area drop-down and the four Sheet Options toggles', () => {
    const pageSetup = ribbonGroupsFor('Page Layout').find((group) => group.id === 'Page Setup')!;
    expect(pageSetup.controls.map((control) => [control.commandId, control.label])).toEqual([
      ['print.pageSetup', 'Page Setup'],
      ['print.setArea', 'Print Area'],
    ]);

    const sheetOptions = ribbonGroupsFor('Page Layout').find(
      (group) => group.id === 'Sheet Options',
    )!;
    expect(
      sheetOptions.controls.map((control) => [control.commandId, control.label, control.title]),
    ).toEqual([
      ['view.gridlines', 'Gridlines', 'Gridlines: View'],
      ['print.gridlines', 'Print', 'Gridlines: Print'],
      ['view.headings', 'Headings', 'Headings: View'],
      ['print.headings', 'Print', 'Headings: Print'],
    ]);
  });

  it('lists the Print Area commands as the drop-down items', () => {
    const placement = inventoryCommand('print.setArea')?.ribbon?.[0];
    expect(placement?.kind).toBe('menu');
    expect(placement?.menu?.map((item) => [item.label, item.commandId])).toEqual([
      ['Set Print Area', 'print.setArea'],
      ['Add to Print Area', 'print.addArea'],
      ['Clear Print Area', 'print.clearArea'],
    ]);
  });

  it('has no v2 Page Layout command in the map', () => {
    const pageLayout = COMMAND_INVENTORY.flatMap((entry) =>
      (entry.ribbon ?? []).filter((placement) => placement.tab === 'Page Layout'),
    );
    expect(pageLayout.every((placement) => placement.group !== 'Breaks')).toBe(true);
    expect(pageLayout.every((placement) => placement.group !== 'Themes')).toBe(true);
    expect(pageLayout.every((placement) => placement.group !== 'Scale to Fit')).toBe(true);
    for (const id of ['print.breaks', 'print.titles', 'print.background', 'view.rightToLeft']) {
      expect(inventoryCommand(id)).toBeUndefined();
    }
  });
});

describe('print.applySetup', () => {
  it('writes every changed field as one transaction, and undo restores the record', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = structuredClone(sheet.printSetup);

    const result = session.dispatch('print.applySetup', {
      sheetId: sheet.id,
      setup: {
        orientation: 'portrait',
        fitToPage: false,
        scale: 80,
        paperSize: 6,
        margins: { ...DEFAULT_PRINT_MARGINS, top: 1, bottom: 1 },
      },
    });

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('printSetup.set');
    expect(session.canUndo).toBe(true);
    expect(sheet.printSetup).toMatchObject({
      orientation: 'portrait',
      fitToPage: false,
      scale: 80,
      paperSize: 6,
      margins: { ...DEFAULT_PRINT_MARGINS, top: 1, bottom: 1 },
      // A field the dialog did not touch keeps its authored value.
      fitToWidth: 1,
      fitToHeight: 0,
    });

    session.undo();
    expect(sheet.printSetup).toEqual(before);
    expect(session.canUndo).toBe(false);
  });

  it('selects fit-to-page when the dialog chooses Fit to', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = structuredClone(sheet.printSetup);

    session.dispatch('print.applySetup', {
      sheetId: sheet.id,
      setup: { fitToPage: true, fitToWidth: 2, fitToHeight: 2 },
    });

    expect(sheet.printSetup).toMatchObject({
      fitToPage: true,
      fitToWidth: 2,
      fitToHeight: 2,
    });
    const plan = planFor(opened);
    expect(plan.fitToPage).toBe(true);
    expect(plan.scale).toBeLessThanOrEqual(1);

    session.undo();
    expect(sheet.printSetup).toEqual(before);
  });

  it('rejects an invalid value without writing anything', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = sheet.printSetup;

    for (const setup of [
      { scale: 5 },
      { scale: 401 },
      { fitToWidth: -1 },
      { margins: { ...DEFAULT_PRINT_MARGINS, left: -1 } },
    ]) {
      const result = session.dispatch('print.applySetup', { sheetId: sheet.id, setup });
      expect(result.changes).toHaveLength(0);
    }
    expect(sheet.printSetup).toBe(before);
    expect(session.canUndo).toBe(false);
  });

  it('writes each single setting as one transaction, and undo restores it exactly', () => {
    const cases: { readonly setup?: object; readonly printArea?: string }[] = [
      { setup: { orientation: 'portrait' } },
      { setup: { fitToPage: false, scale: 70 } },
      { setup: { fitToPage: true, fitToWidth: 2, fitToHeight: 1 } },
      { setup: { paperSize: 6 } },
      { setup: { margins: { ...DEFAULT_PRINT_MARGINS, left: 1.5 } } },
      { setup: { printGridlines: true } },
      { setup: { printHeadings: true } },
      { printArea: 'A1:B2' },
    ];
    for (const args of cases) {
      const opened = openFixture('print-setup.xlsx');
      const sheet = sheetOf(opened);
      const session = new DocumentSession(opened.workbook);
      const beforeSetup = structuredClone(sheet.printSetup);
      const beforeArea = printAreaOf(sheet);
      const beforeVersion = session.version;

      const result = session.dispatch('print.applySetup', { sheetId: sheet.id, ...args });
      expect(result.changes, JSON.stringify(args)).toHaveLength(1);
      expect(session.version).toBe(beforeVersion + 1);
      session.undo();
      expect(sheet.printSetup, JSON.stringify(args)).toEqual(beforeSetup);
      expect(printAreaOf(sheet), JSON.stringify(args)).toEqual(beforeArea);
      expect(session.version).toBe(beforeVersion + 2);
      expect(session.canUndo).toBe(false);
    }
  });

  it('does nothing when the draft matches the sheet', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch('print.applySetup', { sheetId: sheet.id, setup: {} });
    expect(result.changes).toHaveLength(0);
    expect(session.canUndo).toBe(false);
  });
});

describe('Page Setup values and the dialog payload', () => {
  it('reads the sheet’s values, with Excel’s default margins when none is declared', () => {
    const opened = openBlankWorkbook();
    const sheet = sheetOf(opened);
    sheet.printSetup = null;
    const values = pageSetupValuesOf(sheet);
    expect(values).toMatchObject({
      orientation: 'portrait',
      fitToPage: false,
      fitToWidth: 1,
      fitToHeight: 0,
      scale: 100,
      paperSize: null,
      printAreaText: '',
      printGridlines: false,
      printHeadings: false,
    });
    expect(values.margins).toEqual(DEFAULT_PRINT_MARGINS);
  });

  it('sends only the changed fields, so a one-field edit is a one-field patch', () => {
    const opened = openFixture('print-setup.xlsx');
    const values = pageSetupValuesOf(sheetOf(opened));
    const draft = pageSetupDraftFrom(values);
    expect(pageSetupDraftPayload(draft, values, 'Report').payload).toEqual({});

    const changed = pageSetupDraftPayload({ ...draft, orientation: 'portrait' }, values, 'Report');
    expect(changed.errors).toEqual({});
    expect(changed.payload).toEqual({ setup: { orientation: 'portrait' } });
  });

  it('validates the scale, fit pages, margins and print area', () => {
    const opened = openFixture('print-setup.xlsx');
    const values = pageSetupValuesOf(sheetOf(opened));
    const draft = pageSetupDraftFrom(values);

    expect(
      pageSetupDraftPayload({ ...draft, scaling: 'scale', scale: '5' }, values, 'Report').errors,
    ).toHaveProperty('scale');
    expect(
      pageSetupDraftPayload({ ...draft, scaling: 'fit', fitToWidth: '-1' }, values, 'Report')
        .errors,
    ).toHaveProperty('fitToWidth');
    expect(
      pageSetupDraftPayload(
        { ...draft, margins: { ...draft.margins, top: '50' } },
        values,
        'Report',
      ).errors,
    ).toHaveProperty('margins');
    expect(
      pageSetupDraftPayload({ ...draft, printArea: 'not a range' }, values, 'Report').errors,
    ).toHaveProperty('printArea');
    expect(
      pageSetupDraftPayload({ ...draft, printArea: "'Other'!A1:B2" }, values, 'Report').errors,
    ).toHaveProperty('printArea');
  });

  it('keeps an imported multi-area print area when the field is untouched', () => {
    const opened = openFixture('print-areas.xlsx');
    const values = pageSetupValuesOf(sheetOf(opened));
    expect(values.printAreaText).toBe('$A$1:$C$30, $E$1:$F$6');
    const built = pageSetupDraftPayload(pageSetupDraftFrom(values), values, 'Sales');
    expect(built.errors).toEqual({});
    expect(built.payload).toEqual({});
  });

  it('carries a changed print area in the same payload and nothing when unchanged', () => {
    const opened = openFixture('print-setup.xlsx');
    const values = pageSetupValuesOf(sheetOf(opened));
    const draft = pageSetupDraftFrom(values);

    const same = pageSetupDraftPayload({ ...draft, printArea: 'A1:D20' }, values, 'Report');
    expect(same.payload.printArea).toBeUndefined();

    const moved = pageSetupDraftPayload({ ...draft, printArea: 'A1:C5' }, values, 'Report');
    expect(moved.errors).toEqual({});
    expect(moved.payload).toEqual({ printArea: 'A1:C5' });

    const cleared = pageSetupDraftPayload({ ...draft, printArea: '  ' }, values, 'Report');
    expect(cleared.payload).toEqual({ printArea: null });
  });

  it('reports the candidate plan in one line', () => {
    expect(
      pageSetupPlanText({ ok: true, pages: 1, paperName: 'A4', orientation: 'landscape' }),
    ).toBe('Plan: 1 page on A4 landscape.');
    expect(
      pageSetupPlanText({ ok: true, pages: 3, paperName: 'Letter', orientation: 'portrait' }),
    ).toBe('Plan: 3 pages on Letter portrait.');
    expect(
      pageSetupPlanText({ ok: true, pages: 0, paperName: 'A4', orientation: 'portrait' }),
    ).toBe('The plan has no pages to print.');
    expect(pageSetupPlanText({ ok: false, message: 'Too many pages.' })).toBe('Too many pages.');
  });

  it('builds the committed values from a valid draft', () => {
    const opened = openFixture('print-setup.xlsx');
    const values = pageSetupValuesOf(sheetOf(opened));
    const draft = { ...pageSetupDraftFrom(values), scale: '90', orientation: 'portrait' as const };
    expect(pageSetupValuesFromDraft(draft, 'Report')).toMatchObject({
      scale: 90,
      orientation: 'portrait',
    });
  });
});

describe('print area', () => {
  it('sets the area from a range as one name transaction, and undo clears it', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = printAreaOf(sheet);
    expect(before).toHaveLength(2);

    const result = session.dispatch('print.setArea', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 9, endCol: 2 },
    });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('name.set');
    expect(sheet.names.map((name) => name.formula)).toEqual(["'Sales'!$A$1:$C$10"]);
    // The preview follows immediately: one page set instead of two.
    expect(planFor(opened).pageSets.map((set) => set.range)).toEqual([
      { startRow: 0, startCol: 0, endRow: 9, endCol: 2 },
    ]);

    session.undo();
    expect(printAreaOf(sheet)).toEqual(before);
    expect(planFor(opened).pageSets).toHaveLength(2);
  });

  it('creates the name on a sheet that imported none, and round-trips the file', () => {
    const opened = openBlankWorkbook();
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    expect(sheet.names).toEqual([]);

    session.dispatch('print.setArea', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
    });
    const saved = save(opened);
    expect(saved.noOp).toBe(false);
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0" hidden="1">'Sheet1'!$A$1:$C$5</definedName>`,
    );
    expect(printAreaOf(openWorkbook(saved.bytes).workbook.sheets[0]!)).toEqual([
      { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
    ]);
  });

  it('clears the name and removes it from the saved workbook part', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);

    const result = session.dispatch('print.clearArea', { sheetId: sheet.id });
    expect(result.changes).toHaveLength(1);
    expect(printAreaOf(sheet)).toEqual([]);
    expect(planFor(opened).pageSets).toHaveLength(1);

    const saved = save(opened);
    expect(partText(saved.bytes, WORKBOOK_PART)).not.toContain('_xlnm.Print_Area');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.names).toEqual([]);

    session.undo();
    expect(printAreaOf(sheet)).toEqual([{ startRow: 0, startCol: 0, endRow: 19, endCol: 3 }]);
  });

  it('moves the name with the sheet list when a sheet is deleted', () => {
    const opened = openFixture('print-setup.xlsx');
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', { index: 1 });
    const added = opened.workbook.sheets[1]!;
    session.dispatch('print.setArea', {
      sheetId: added.id,
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
    });
    session.dispatch('sheet.delete', { sheetId: sheetOf(opened).id, confirm: true });

    const saved = save(opened);
    const workbookXml = partText(saved.bytes, WORKBOOK_PART);
    expect(workbookXml).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0" hidden="1">'${added.name}'!$A$1:$B$3</definedName>`,
    );
    expect(workbookXml).not.toContain('_xlnm.Print_Area" localSheetId="1"');
    const reopened = openWorkbook(saved.bytes);
    expect(printAreaOf(reopened.workbook.sheets[0]!)).toEqual([
      { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
    ]);
  });

  it('follows a sheet rename in the print area formula', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.rename', { sheetId: sheet.id, name: 'Report 2027' });

    expect(sheet.names.map((name) => name.formula)).toEqual(["'Report 2027'!$A$1:$D$20"]);
    const saved = save(opened);
    const workbookXml = partText(saved.bytes, WORKBOOK_PART);
    expect(workbookXml).toContain(`'Report 2027'!$A$1:$D$20`);
    expect(workbookXml).not.toContain(`'Report'!$A$1:$D$20`);
  });

  it('keeps an untouched print area byte-identical on save', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const saved = save(opened);
    expect(saved.noOp).toBe(true);
    expect(partText(saved.bytes, WORKBOOK_PART)).toBe(partText(original, WORKBOOK_PART));
  });

  it('writes whole columns and whole rows in Excel’s form', () => {
    expect(
      printAreaTextForRange({ startRow: 0, startCol: 0, endRow: MAX_ROW - 1, endCol: 2 }),
    ).toBe('$A:$C');
    expect(printAreaTextForRange({ startRow: 0, startCol: 0, endRow: 4, endCol: 16_383 })).toBe(
      '$1:$5',
    );
    expect(printAreaTextForRange({ startRow: 1, startCol: 1, endRow: 3, endCol: 4 })).toBe(
      '$B$2:$E$4',
    );
    expect(normalisePrintAreaText('A1:C5', 'My Sheet')).toEqual({
      ok: true,
      formula: "'My Sheet'!$A$1:$C$5",
      ranges: [{ startRow: 0, startCol: 0, endRow: 4, endCol: 2 }],
    });
    // Several areas are the shape Excel writes for a multi-area name.
    expect(normalisePrintAreaText('$A$1:$C$5, $E$1:$F$6', 'My Sheet')).toEqual({
      ok: true,
      formula: "'My Sheet'!$A$1:$C$5,'My Sheet'!$E$1:$F$6",
      ranges: [
        { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
        { startRow: 0, startCol: 4, endRow: 5, endCol: 5 },
      ],
    });
    expect(
      formatPrintAreaRanges([
        { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
        { startRow: 0, startCol: 4, endRow: 5, endCol: 5 },
      ]),
    ).toBe('$A$1:$C$5, $E$1:$F$6');
  });

  it('writes the print area and page setup in one transaction from the dialog', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = structuredClone(sheet.printSetup);
    const beforeArea = printAreaOf(sheet);

    const result = session.dispatch('print.applySetup', {
      sheetId: sheet.id,
      setup: { orientation: 'portrait' },
      printArea: 'B2:C6',
    });
    expect(result.changes.map((change) => change.type)).toEqual(['printSetup.set', 'name.set']);
    expect(session.canUndo).toBe(true);
    expect(printAreaOf(sheet)).toEqual([{ startRow: 1, startCol: 1, endRow: 5, endCol: 2 }]);

    session.undo();
    expect(sheet.printSetup).toEqual(before);
    expect(printAreaOf(sheet)).toEqual(beforeArea);
    expect(sheet.names.map((name) => name.formula)).toEqual(["'Report'!$A$1:$D$20"]);
  });
});

describe('print gridlines and headings', () => {
  it('writes printOptions as one transaction and changes the printed output only', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const viewBefore = structuredClone(sheet.view);
    const planBefore = planFor(opened);
    expect(planBefore.showGridLines).toBe(false);
    expect(planBefore.showHeadings).toBe(false);

    const gridlines = session.dispatch('print.gridlines', { sheetId: sheet.id });
    expect(gridlines.changes).toHaveLength(1);
    expect(gridlines.changes[0]?.type).toBe('printSetup.set');
    expect(sheet.printSetup?.options).toMatchObject({ gridLines: true, gridLinesSet: true });
    expect(planFor(opened).showGridLines).toBe(true);

    session.dispatch('print.headings', { sheetId: sheet.id });
    expect(planFor(opened).showHeadings).toBe(true);

    // The view flags are chrome state; the command never writes them.
    expect(sheet.view).toEqual(viewBefore);

    session.undo();
    session.undo();
    expect(planFor(opened).showGridLines).toBe(false);
    expect(planFor(opened).showHeadings).toBe(false);
  });

  it('round-trips the printed flags through the worksheet part', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('print.gridlines', { sheetId: sheet.id });

    const saved = save(opened);
    const sheetXml = partText(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('gridLines="1"');
    expect(sheetXml).toContain('gridLinesSet="1"');
    const reopened = openWorkbook(saved.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(
      createPrintPlan(
        createPrintSnapshot(
          reopened.workbook,
          createSheetProjection(reopened.workbook, reopenedSheet.id, { version: 1 }),
        ),
      ).showGridLines,
    ).toBe(true);
  });

  it('reports its checked state from the sheet, not the view', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const context = {
      workbook: opened.workbook,
      activeSheetId: sheet.id,
      selection: null,
      activeCell: null,
    };
    const registry = createCommandRegistry();
    // print-areas.xlsx prints gridlines and headings.
    expect(registry.get('print.gridlines')?.isChecked?.(context)).toBe(true);
    expect(registry.get('print.headings')?.isChecked?.(context)).toBe(true);

    // The screen toggles are different commands; they never move this state.
    session.dispatch('view.gridlines', undefined);
    session.dispatch('view.headings', undefined);
    expect(registry.get('print.gridlines')?.isChecked?.(context)).toBe(true);
    expect(registry.get('print.headings')?.isChecked?.(context)).toBe(true);
  });
});

describe('buildPageSetup', () => {
  it('keeps the imported source so a save rewrites only the changed element', () => {
    const opened = openFixture('print-setup.xlsx');
    const before = sheetOf(opened).printSetup!;
    const built = buildPageSetup(before, { orientation: 'portrait' });
    expect(built.errors).toEqual({});
    expect(built.setup.raw).toBe(before.raw);
    expect(built.setup.margins).toBe(before.margins);
    expect(built.setup.options).toBe(before.options);
    expect(built.setup.orientation).toBe('portrait');
  });
});

describe('ticket 47 — print view polish', () => {
  it('adds the selection to the existing areas as one name transaction', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = sheet.names[0]!.formula;

    const result = session.dispatch('print.addArea', {
      sheetId: sheet.id,
      range: { startRow: 9, startCol: 4, endRow: 11, endCol: 5 },
    });

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('name.set');
    expect(sheet.names[0]!.formula).toBe(
      "'Sales'!$A$1:$C$30,'Sales'!$E$1:$F$6,'Sales'!$E$10:$F$12",
    );
    expect(printAreaOf(sheet)).toHaveLength(3);
    expect(planFor(opened).pageSets).toHaveLength(3);

    // Undo restores the previous name exactly, not a rewritten equivalent.
    session.undo();
    expect(sheet.names[0]!.formula).toBe(before);
    expect(planFor(opened).pageSets).toHaveLength(2);
  });

  it('adds from the selection when no range is given', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const selection = { startRow: 0, startCol: 4, endRow: 2, endCol: 5 };

    const result = session.dispatch(
      'print.addArea',
      { sheetId: sheet.id },
      { activeSheetId: sheet.id, selection, activeCell: { row: 0, col: 4 } },
    );

    expect(result.changes).toHaveLength(1);
    expect(printAreaOf(sheet)).toEqual([
      { startRow: 0, startCol: 0, endRow: 29, endCol: 2 },
      { startRow: 0, startCol: 4, endRow: 5, endCol: 5 },
      { startRow: 0, startCol: 4, endRow: 2, endCol: 5 },
    ]);
  });

  it('sets the area when the sheet has none, and round-trips through the file', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('print.addArea', {
      sheetId: sheet.id,
      range: { startRow: 2, startCol: 4, endRow: 4, endCol: 5 },
    });
    expect(sheet.names.map((name) => name.formula)).toEqual([
      "'Report'!$A$1:$D$20,'Report'!$E$3:$F$5",
    ]);

    const saved = save(opened);
    const workbookXml = partText(saved.bytes, WORKBOOK_PART);
    expect(workbookXml).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0">'Report'!$A$1:$D$20,'Report'!$E$3:$F$5</definedName>`,
    );
    expect(printAreaOf(openWorkbook(saved.bytes).workbook.sheets[0]!)).toHaveLength(2);
  });

  it('writes both centring flags as one transaction and the planner honours them', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const before = structuredClone(sheet.printSetup);
    const planBefore = planFor(opened);
    const firstColumnX = planBefore.pages[0]!.columns[0]!.x;
    const firstRowY = planBefore.pages[0]!.rows[0]!.y;

    const result = session.dispatch('print.applySetup', {
      sheetId: sheet.id,
      setup: { horizontalCentered: true, verticalCentered: true },
    });

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('printSetup.set');
    // The other print options keep their authored values.
    expect(sheet.printSetup?.options).toMatchObject({
      horizontalCentered: true,
      verticalCentered: true,
      gridLines: true,
      headings: true,
    });

    const plan = planFor(opened);
    expect(plan.horizontalCentered).toBe(true);
    expect(plan.verticalCentered).toBe(true);
    expect(plan.pages[0]!.columns[0]!.x).toBeGreaterThan(firstColumnX);
    expect(plan.pages[0]!.rows[0]!.y).toBeGreaterThan(firstRowY);

    session.undo();
    expect(sheet.printSetup).toEqual(before);
    expect(planFor(opened).horizontalCentered).toBe(false);
    expect(planFor(opened).pages[0]!.columns[0]!.x).toBe(firstColumnX);
  });

  it('round-trips the centring flags through the worksheet part', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('print.applySetup', {
      sheetId: sheet.id,
      setup: { horizontalCentered: true, verticalCentered: true },
    });

    const saved = save(opened);
    const sheetXml = partText(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('horizontalCentered="1"');
    expect(sheetXml).toContain('verticalCentered="1"');
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.printSetup?.options).toMatchObject({
      horizontalCentered: true,
      verticalCentered: true,
    });
  });

  it('sends one centring check box as one patch field', () => {
    const values = pageSetupValuesOf(sheetOf(openFixture('print-areas.xlsx')));
    const draft = pageSetupDraftFrom(values);
    const built = pageSetupDraftPayload({ ...draft, horizontalCentered: true }, values, 'Sales');
    expect(built.errors).toEqual({});
    expect(built.payload).toEqual({ setup: { horizontalCentered: true } });
  });

  it('reads the imported centring flags', () => {
    const values = pageSetupValuesOf(sheetOf(openFixture('print-setup.xlsx')));
    expect(values.horizontalCentered).toBe(true);
    expect(values.verticalCentered).toBe(false);
  });
});
