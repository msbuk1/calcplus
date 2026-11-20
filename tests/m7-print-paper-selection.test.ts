import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { EMPTY_PRINT_SETUP } from '@commands/print';
import { openBlankWorkbook } from '@ooxml/blank';
import type { OpenedWorkbook } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { createSheetProjection } from '@view/projection';
import { createPrintPlan } from '../src/print/planner';
import { createPrintSnapshot } from '../src/print/adapter';
import type { PrintPlan } from '../src/print/types';
import { printPaperOptions, printSelectedPaperCode } from '../src/chrome/print-preview-state';
import { openFixture, partText } from './helpers/roundtrip';

/**
 * Ticket 42 — paper selection in the print preview.
 *
 * The media list is the planner's own table (probed through its public API, so
 * there is no second copy), a choice writes the sheet's page setup as one
 * transaction, and undo puts the previous paper back so the plan follows.
 */

const SHEET1 = 'xl/worksheets/sheet1.xml';

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

describe('the preview media list', () => {
  it('names every paper the planner knows, in code order', () => {
    expect(printPaperOptions()).toEqual([
      { code: 1, name: 'Letter' },
      { code: 3, name: 'Tabloid' },
      { code: 4, name: 'Ledger' },
      { code: 5, name: 'Legal' },
      { code: 6, name: 'Statement' },
      { code: 7, name: 'Executive' },
      { code: 8, name: 'A3' },
      { code: 9, name: 'A4' },
      { code: 10, name: 'A5' },
      { code: 11, name: 'B4' },
      { code: 12, name: 'B5' },
      { code: 13, name: 'Folio' },
      { code: 14, name: 'Quarto' },
      { code: 15, name: 'Standard 10x14' },
      { code: 16, name: 'Standard 11x17' },
      { code: 17, name: 'Note' },
      { code: 18, name: 'Envelope 9' },
      { code: 19, name: 'Envelope 10' },
    ]);
  });

  it('resolves every listed code through the planner itself', () => {
    const opened = openBlankWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    const snapshot = createPrintSnapshot(
      opened.workbook,
      createSheetProjection(opened.workbook, sheet.id, { version: 1 }),
    );
    for (const option of printPaperOptions()) {
      const plan = createPrintPlan({
        ...snapshot,
        printAreas: [],
        usedRange: null,
        printSetup: { ...EMPTY_PRINT_SETUP, paperSize: option.code },
      });
      expect({ code: plan.paper.code, name: plan.paper.name }).toEqual({
        code: option.code,
        name: option.name,
      });
    }
    // A code the table does not name resolves to Custom and is not a choice.
    expect(printPaperOptions().some((option) => option.code === 2)).toBe(false);
    expect(printPaperOptions().some((option) => option.name === 'Custom')).toBe(false);
  });

  it('selects the plan’s own paper, the default, or no option for a custom code', () => {
    const options = printPaperOptions();
    const opened = openFixture('print-areas.xlsx');
    const a4 = planFor(opened);
    expect(printSelectedPaperCode(a4, options)).toBe(a4.paper.code);

    // A sheet that names no paper resolves to Letter even though its code is null.
    const letter = {
      ...a4,
      paper: { ...a4.paper, name: 'Letter', code: null },
    };
    expect(printSelectedPaperCode(letter, options)).toBe(1);

    const custom = {
      ...a4,
      paper: { ...a4.paper, name: 'Custom', code: 256 },
    };
    expect(printSelectedPaperCode(custom, options)).toBeNull();
  });
});

describe('print.setup', () => {
  it('writes the paper as one transaction and re-plans the pages', () => {
    const opened = openFixture('print-areas.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const before = planFor(opened);
    expect(before.paper).toMatchObject({ name: 'A4', code: 9, orientation: 'landscape' });
    expect(session.canUndo).toBe(false);

    const result = session.dispatch(
      'print.setup',
      { sheetId: sheet.id, paperSize: 6 },
      { activeSheetId: sheet.id },
    );

    // One change, one history entry.
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('printSetup.set');
    expect(session.canUndo).toBe(true);
    expect(sheet.printSetup?.paperSize).toBe(6);

    const after = planFor(opened);
    expect(after.paper).toMatchObject({ name: 'Statement', code: 6 });
    expect(after.paper.width).not.toBe(before.paper.width);
    expect(after.pages.length).not.toBe(before.pages.length);

    // Undo restores the previous paper and the plan follows it exactly.
    session.undo();
    expect(sheet.printSetup?.paperSize).toBe(9);
    expect(session.canUndo).toBe(false);
    expect(planFor(opened).paper).toEqual(before.paper);
    expect(planFor(opened).pages.length).toBe(before.pages.length);
  });

  it('keeps every other page-setup value and round-trips through the file', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const setupBefore = { ...sheet.printSetup! };
    const session = new DocumentSession(opened.workbook);
    session.dispatch('print.setup', { sheetId: sheet.id, paperSize: 6 });

    expect(sheet.printSetup).toMatchObject({
      paperSize: 6,
      orientation: setupBefore.orientation,
      fitToWidth: setupBefore.fitToWidth,
      fitToHeight: setupBefore.fitToHeight,
      scale: setupBefore.scale,
      margins: setupBefore.margins,
    });

    const saved = save(opened);
    expect(saved.noOp).toBe(false);
    expect(partText(saved.bytes, SHEET1)).toContain(
      '<pageSetup paperSize="6" orientation="landscape" fitToWidth="1" fitToHeight="0" scale="85" firstPageNumber="1" useFirstPageNumber="1" cellComments="none"/>',
    );
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.printSetup?.paperSize).toBe(6);
    // The margins element was not rewritten by a paper-only change.
    expect(partText(saved.bytes, SHEET1)).toContain(
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    );
  });

  it('adds a pageSetup element to a sheet that declared none', () => {
    const opened = openBlankWorkbook();
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    session.dispatch('print.setup', { sheetId: sheet.id, paperSize: 3 });

    const saved = save(opened);
    expect(partText(saved.bytes, SHEET1)).toContain('<pageSetup paperSize="3"/>');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.printSetup?.paperSize).toBe(3);
  });

  it('does nothing when the sheet already uses that paper', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch('print.setup', { sheetId: sheet.id, paperSize: 9 });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toBe('The sheet already uses that paper size.');
    expect(session.canUndo).toBe(false);
  });

  it('rejects an invalid code without changing the sheet', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const before = sheet.printSetup;
    for (const paperSize of [0, -1, 4.5, Number.NaN]) {
      const result = session.dispatch('print.setup', { sheetId: sheet.id, paperSize });
      expect(result.changes).toHaveLength(0);
    }
    expect(sheet.printSetup).toBe(before);
    expect(session.canUndo).toBe(false);
  });
});
