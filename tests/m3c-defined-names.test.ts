import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { CalculationController } from '@engine/calculation-controller';
import { cellKey, createCell, setCellSnapshot, type Sheet, type Workbook } from '@model/types';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { BLANK, numberValue } from '@model/value';
import { InlineEngineWorker } from './helpers/inline-engine-worker';

/**
 * Defined names inside formulas (feature matrix §4, ticket 36 item 3).
 *
 * The contract puts defined names IN to "evaluate, preserve, use in formulas
 * and Go To", with no editor in v1. Go To and the name box were already tested
 * (`tests/m4c-find.test.ts`), and preservation round-trips
 * (`tests/m6-sheet-features.test.ts`). Nothing before this file ever put a
 * defined name inside a formula, so the `name.resolve` inventory flag had no
 * evidence behind it either way.
 *
 * What is under test is the production path end to end:
 * `CalculationController.namedExpressionsFor` maps the model's names to engine
 * named expressions, and `HyperFormulaEngine` registers them. The controller
 * owns a Web Worker, so the inline worker from `tests/helpers/inline-engine-worker.ts`
 * stands in and runs the real engine adapter.
 */

interface NameFixture {
  readonly session: DocumentSession;
  readonly workbook: Workbook;
  readonly sheet: Sheet;
  /** The cached value of a cell, as a plain number or NaN. */
  readonly number: (row: number, col: number) => number;
}

/**
 * A sheet where A1 is 5, A2 is 3, and a named range `Rates` covers A1:A2.
 *
 * `TaxRate` is workbook-scoped, `LocalAmount` is sheet-scoped, and `Rates`
 * names the two-cell range so a name can be tested inside an aggregate.
 */
function nameFixture(): NameFixture {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  setCellSnapshot(sheet, 0, 0, createCell(numberValue(5)));
  setCellSnapshot(sheet, 1, 0, createCell(numberValue(3)));

  workbook.names.push({ name: 'TaxRate', formula: 'Sheet1!$A$1', sheetId: null });
  workbook.names.push({ name: 'Rates', formula: 'Sheet1!$A$1:$A$2', sheetId: null });
  sheet.names.push({ name: 'LocalAmount', formula: 'Sheet1!$A$2', sheetId: sheet.id });

  return {
    session: new DocumentSession(workbook),
    workbook,
    sheet,
    number: (row, col) => {
      const value = sheet.cells.get(cellKey(row, col))?.value;
      return value !== undefined && value.type === 'number' ? value.value : Number.NaN;
    },
  };
}

function formula(sheet: Sheet, row: number, col: number, text: string): void {
  setCellSnapshot(sheet, row, col, { ...createCell(BLANK), formula: text });
}

async function waitFor(predicate: () => boolean, label: string, timeoutMs = 20_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error(`Timed out waiting for ${label}`);
    await new Promise((resolve) => {
      setTimeout(resolve, 1);
    });
  }
}

/** Run the controller until the engine is ready and every formula has a value. */
async function calculate(fixture: NameFixture, done: () => boolean): Promise<string | null> {
  let error: string | null = null;
  const controller = new CalculationController(
    fixture.session,
    {
      onModelChanged: () => undefined,
      onCircular: () => undefined,
      onError: (message) => {
        error = message;
      },
    },
    {
      workerFactory: () => new InlineEngineWorker(true) as unknown as Worker,
    },
  );
  try {
    controller.start();
    await waitFor(() => controller.state().ready, 'engine ready');
    await waitFor(done, 'named formulas to evaluate');
  } finally {
    controller.dispose();
  }
  return error;
}

describe('defined names inside formulas', () => {
  it('evaluates a workbook-scoped name in a formula', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'TaxRate*2');

    const error = await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));

    expect(error).toBeNull();
    expect(fixture.number(0, 1)).toBe(10);
  });

  it('evaluates a sheet-scoped name in a formula', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'LocalAmount+1');

    const error = await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));

    expect(error).toBeNull();
    expect(fixture.number(0, 1)).toBe(4);
  });

  it('resolves a name passed to a function', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'SUM(TaxRate)');

    const error = await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));

    expect(error).toBeNull();
    expect(fixture.number(0, 1)).toBe(5);
  });

  it('resolves a name that covers a range', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'SUM(Rates)');
    formula(fixture.sheet, 0, 2, 'COUNT(Rates)');

    const error = await calculate(
      fixture,
      () => Number.isFinite(fixture.number(0, 1)) && Number.isFinite(fixture.number(0, 2)),
    );

    expect(error).toBeNull();
    // A1:A2 holds 5 and 3.
    expect(fixture.number(0, 1)).toBe(8);
    expect(fixture.number(0, 2)).toBe(2);
  });

  it('uses the sheet-scoped name over a workbook name of the same spelling', async () => {
    const fixture = nameFixture();
    // `Shadow` exists at both scopes with different values; the sheet one wins
    // on Sheet1 (model/references.ts `findDefinedName` looks local first).
    fixture.workbook.names.push({ name: 'Shadow', formula: 'Sheet1!$A$1', sheetId: null });
    fixture.sheet.names.push({ name: 'Shadow', formula: 'Sheet1!$A$2', sheetId: fixture.sheet.id });
    formula(fixture.sheet, 0, 1, 'Shadow*10');

    const error = await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));

    expect(error).toBeNull();
    // Sheet-scoped names resolve first, so this is A2 (3), not A1 (5).
    expect(fixture.number(0, 1)).toBe(30);
  });

  it('mixes two names in one formula', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'TaxRate+LocalAmount');

    const error = await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));

    expect(error).toBeNull();
    expect(fixture.number(0, 1)).toBe(8);
  });

  it('recomputes a formula when the named target changes', async () => {
    const fixture = nameFixture();
    formula(fixture.sheet, 0, 1, 'TaxRate*2');
    await calculate(fixture, () => Number.isFinite(fixture.number(0, 1)));
    expect(fixture.number(0, 1)).toBe(10);

    // The name still points at A1; changing A1 must move the result.
    setCellSnapshot(fixture.sheet, 0, 0, createCell(numberValue(7)));

    let error: string | null = null;
    const controller = new CalculationController(
      fixture.session,
      {
        onModelChanged: () => undefined,
        onCircular: () => undefined,
        onError: (message) => {
          error = message;
        },
      },
      { workerFactory: () => new InlineEngineWorker(true) as unknown as Worker },
    );
    try {
      controller.start();
      await waitFor(() => controller.state().ready, 'engine ready');
      await waitFor(() => fixture.number(0, 1) === 14, 'the named formula to follow A1');
      expect(error).toBeNull();
    } finally {
      controller.dispose();
    }
  });
});
