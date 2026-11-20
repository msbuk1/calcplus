import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { CalculationController } from '@engine/calculation-controller';
import { FormulaEngineClient } from '@engine/formula-engine-client';
import type { EngineEdit } from '@engine/port';
import type { ToEngineWorkerMessage } from '@engine/protocol';
import { cellKey, createCell, setCellSnapshot, type Sheet } from '@model/types';
import { BLANK, numberValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { InlineEngineWorker } from './helpers/inline-engine-worker';

/**
 * M7 queue, backpressure and fan-out correctness (performance §3.2).
 *
 * The real controller owns a Web Worker, so these tests inject an inline worker
 * that speaks the production protocol and runs the real HyperFormula adapter.
 * What is under test is the production code: the client's chunked load with
 * acks, its per-cell coalescing and single in-flight apply, and the
 * controller's sliced application of a large fan-out.
 */
describe('formula worker queue', () => {
  it('loads one chunk at a time and waits for each ack', async () => {
    const worker = new InlineEngineWorker(false);
    let clientError: string | null = null;
    const client = new FormulaEngineClient(
      {
        onResult: () => undefined,
        onError: (message) => {
          clientError = message;
        },
      },
      { workerFactory: () => worker as unknown as Worker, chunkSize: 10 },
    );
    const cells = Array.from({ length: 25 }, (_, index) => ({
      sheet: 0,
      row: index,
      col: 0,
      formula: null,
      value: numberValue(index),
    }));

    client.initialize({ date1904: false, sheetNames: ['Sheet1'] }, cells, 0);
    // Start plus exactly one chunk: the rest waits for the worker's ack.
    expect(worker.inbox.map((message) => message.type)).toEqual(['start', 'load']);
    expect(worker.inbox[1]).toMatchObject({ type: 'load', final: false });
    expect(worker.inbox[1]?.type === 'load' ? worker.inbox[1].cells.length : 0).toBe(10);

    await worker.step();
    await worker.step();
    // The ack requested chunk two synchronously; chunk three waits again.
    expect(worker.inbox).toHaveLength(1);
    expect(worker.inbox[0]).toMatchObject({ type: 'load', final: false });

    await worker.settle();
    expect(worker.loadChunks).toEqual([
      { cells: 10, final: false },
      { cells: 10, final: false },
      { cells: 5, final: true },
    ]);
    expect(client.isReady).toBe(true);
    expect(client.hasFailed).toBe(false);
    expect(clientError).toBeNull();
  });

  it('coalesces edits queued while an apply is in flight, newest per cell wins', async () => {
    const worker = new InlineEngineWorker(false);
    let clientError: string | null = null;
    const client = new FormulaEngineClient(
      {
        onResult: () => undefined,
        onError: (message) => {
          clientError = message;
        },
      },
      { workerFactory: () => worker as unknown as Worker, chunkSize: 10 },
    );
    client.initialize({ date1904: false, sheetNames: ['Sheet1'] }, [], 0);
    await worker.settle();
    expect(client.isReady).toBe(true);

    client.applyEdits([edit(0, 0, 1)], 1);
    // The first request is in flight; these two cells supersede the queued
    // write to A1 and add one more cell.
    client.applyEdits([edit(0, 0, 2)], 2);
    client.applyEdits([edit(0, 1, 3)], 2);
    expect(client.queuedEditCount).toBe(2);
    expect(client.isBusy).toBe(true);

    await worker.settle();

    const applies = worker.received.filter(
      (message): message is Extract<ToEngineWorkerMessage, { type: 'apply' }> =>
        message.type === 'apply',
    );
    expect(applies).toHaveLength(2);
    expect(applies[0]?.edits).toEqual([edit(0, 0, 1)]);
    expect(applies[1]?.edits).toEqual([edit(0, 0, 2), edit(0, 1, 3)]);
    expect(client.queuedEditCount).toBe(0);
    expect(client.hasFailed).toBe(false);
    expect(clientError).toBeNull();
  });

  it('drops queued edits a rebuild snapshot already contains', async () => {
    const worker = new InlineEngineWorker(false);
    const client = new FormulaEngineClient(
      { onResult: () => undefined, onError: () => undefined },
      { workerFactory: () => worker as unknown as Worker, chunkSize: 10 },
    );
    client.initialize({ date1904: false, sheetNames: ['Sheet1'] }, [], 0);
    await worker.settle();

    client.applyEdits([edit(0, 0, 1)], 1);
    client.applyEdits([edit(0, 1, 2)], 2);
    client.applyEdits([edit(0, 2, 3)], 3);
    expect(client.queuedEditCount).toBe(2);
    client.discardPendingBefore(2);
    expect(client.queuedEditCount).toBe(1);
    await worker.settle();

    const applies = worker.received.filter(
      (message): message is Extract<ToEngineWorkerMessage, { type: 'apply' }> =>
        message.type === 'apply',
    );
    expect(applies[1]?.edits).toEqual([edit(0, 2, 3)]);
  });

  it('holds the next request until a result has finished applying', async () => {
    const worker = new InlineEngineWorker(false);
    let releaseResult!: () => void;
    const resultApplied = new Promise<void>((resolve) => {
      releaseResult = resolve;
    });
    let results = 0;
    let clientError: string | null = null;
    const client = new FormulaEngineClient(
      {
        onResult: async () => {
          results += 1;
          if (results === 1) await resultApplied;
        },
        onError: (message) => {
          clientError = message;
        },
      },
      { workerFactory: () => worker as unknown as Worker, chunkSize: 10 },
    );

    client.initialize({ date1904: false, sheetNames: ['Sheet1'] }, [], 0);
    await worker.settle();
    expect(client.isReady).toBe(true);
    expect(results).toBe(1);

    // The first result callback is still applying: the edit must wait rather
    // than clone a second request into the worker.
    client.applyEdits([edit(0, 0, 1)], 1);
    expect(worker.inbox).toHaveLength(0);
    expect(client.queuedEditCount).toBe(1);

    releaseResult();
    await worker.settle();
    expect(worker.received.filter((message) => message.type === 'apply')).toHaveLength(1);
    expect(client.queuedEditCount).toBe(0);
    expect(client.hasFailed).toBe(false);
    expect(clientError).toBeNull();
  });
});

describe('calculation fan-out', () => {
  it('builds the engine snapshot in bounded slices', async () => {
    const fixture = fanOutFixture(400);
    const yields: number[] = [];
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
        snapshotChunkSize: 50,
        snapshotSliceBudgetMs: 1_000,
        resultChunkSize: 10_000,
        resultSliceBudgetMs: 1_000,
        yieldToUi: async () => {
          yields.push(performance.now());
          await Promise.resolve();
        },
      },
    );

    try {
      controller.start();
      await waitFor(() => controller.state().ready, 'engine ready');
      await waitFor(() => fixture.numbers(399).at(1) === 2, 'fan-out');
      expect(error).toBeNull();
      // 401 cells in 50-cell snapshot slices: at least seven yields before the
      // engine request is posted.
      expect(yields.length).toBeGreaterThanOrEqual(7);
    } finally {
      controller.dispose();
    }
  });

  it('applies a 500-cell fan-out in bounded slices with no long main-thread block', async () => {
    const fixture = fanOutFixture(500);
    let modelChanges = 0;
    let error: string | null = null;
    const controller = new CalculationController(
      fixture.session,
      {
        onModelChanged: () => {
          modelChanges += 1;
        },
        onCircular: () => undefined,
        onError: (message) => {
          error = message;
        },
      },
      {
        workerFactory: () => new InlineEngineWorker(true) as unknown as Worker,
        resultSliceBudgetMs: 1_000,
        resultChunkSize: 100,
        yieldToUi: async () => {
          await Promise.resolve();
        },
      },
    );

    try {
      controller.start();
      await waitFor(() => controller.state().ready, 'engine ready');
      await waitFor(() => fixture.numbers(499).at(1) === 2, 'every formula in the load result');
      expect(error).toBeNull();
      // 500 updates in 100-cell slices: five model changes, so no slice could
      // have swallowed the whole fan-out.
      expect(modelChanges).toBe(5);

      modelChanges = 0;
      fixture.session.dispatch('cell.set', {
        sheetId: fixture.sheet.id,
        row: 0,
        col: 0,
        value: numberValue(3),
      });
      await waitFor(() => fixture.numbers(499).at(1) === 6, 'recalculated fan-out');
      await waitFor(() => !controller.state().applying, 'result application finished');

      for (let row = 0; row < 500; row += 1) {
        expect(fixture.numbers(row).at(1), `B${row + 1}`).toBe(6);
      }
      expect(modelChanges).toBe(5);
      expect(controller.state().staleCount).toBe(0);
      expect(controller.state().unsupportedCount).toBe(0);
    } finally {
      controller.dispose();
    }
  });

  it('lets an edit made during result slicing win over the older derived value', async () => {
    const fixture = fanOutFixture(400);
    let armed = false;
    let injected = false;
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
        resultSliceBudgetMs: 0,
        resultChunkSize: 50,
        yieldToUi: async () => {
          if (armed && !injected) {
            injected = true;
            fixture.session.dispatch('cell.set', {
              sheetId: fixture.sheet.id,
              row: 0,
              col: 1,
              value: numberValue(99),
            });
          }
          await Promise.resolve();
        },
      },
    );

    try {
      controller.start();
      await waitFor(() => controller.state().ready, 'engine ready');
      await waitFor(() => fixture.numbers(399).at(1) === 2, 'initial fan-out');
      armed = true;

      fixture.session.dispatch('cell.set', {
        sheetId: fixture.sheet.id,
        row: 0,
        col: 0,
        value: numberValue(5),
      });
      await waitFor(() => injected, 'injected edit');
      await waitFor(() => fixture.numbers(399).at(1) === 10, 'new fan-out applied');
      await waitFor(() => !controller.state().applying, 'result application finished');

      // The injected edit won: B1 is the authored literal, not the 10 from the
      // older batch. Every other dependent is the new result.
      expect(fixture.numbers(0).at(1)).toBe(99);
      expect(fixture.numbers(1).at(1)).toBe(10);
      expect(fixture.numbers(399).at(1)).toBe(10);
      expect(error).toBeNull();
    } finally {
      controller.dispose();
    }
  });

  it('keeps the stale and unsupported counts correct without rescanning the workbook', async () => {
    const fixture = fanOutFixture(20);
    const sheet = fixture.sheet;
    sheet.cells.set(cellKey(20, 0), {
      ...createCell(BLANK),
      formula: '_xlfn.XLOOKUP(A1,A1:A1,A1:A1)',
    });
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
        yieldToUi: async () => {
          await Promise.resolve();
        },
      },
    );

    try {
      controller.start();
      await waitFor(() => controller.state().ready, 'engine ready');
      await waitFor(() => fixture.numbers(19).at(1) === 2, 'fan-out');
      await waitFor(() => controller.state().staleCount === 1, 'stale count');
      expect(controller.state().unsupportedCount).toBe(1);

      fixture.session.dispatch('cell.set', {
        sheetId: sheet.id,
        row: 20,
        col: 0,
        value: numberValue(7),
      });
      await waitFor(() => controller.state().unsupportedCount === 0, 'unsupported count');
      await waitFor(() => controller.state().staleCount === 0, 'stale count after edit');
      expect(error).toBeNull();
    } finally {
      controller.dispose();
    }
  });
});

function edit(sheet: number, row: number, col: number, value = 1): EngineEdit {
  return { sheet, row, col, formula: null, value: numberValue(value) };
}

function fanOutFixture(formulas: number): {
  readonly session: DocumentSession;
  readonly sheet: Sheet;
  readonly numbers: (row: number) => readonly number[];
} {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  setCellSnapshot(sheet, 0, 0, createCell(numberValue(1)));
  for (let row = 0; row < formulas; row += 1) {
    setCellSnapshot(sheet, row, 1, { ...createCell(BLANK), formula: 'A1*2' });
  }
  return {
    session: new DocumentSession(workbook),
    sheet,
    numbers: (row) => {
      const values: number[] = [];
      for (let col = 0; col < 2; col += 1) {
        const cell = sheet.cells.get(cellKey(row, col));
        values.push(cell?.value.type === 'number' ? cell.value.value : Number.NaN);
      }
      return values;
    },
  };
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
