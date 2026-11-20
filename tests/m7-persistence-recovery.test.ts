import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { numberValue } from '@model/value';
import type { OpenedWorkbook } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { snapshotWorkbook, type WorkbookSnapshot } from '@ooxml/snapshot';
import {
  MAX_PENDING_RECOVERY_DOCS,
  RECOVERY_BATCH_BYTES,
  RECOVERY_MAX_AGE_MS,
  RecoveryQuotaError,
  RecoveryRepository,
  recoveryIsNewerThanSavedFile,
  type RecoverySaveInput,
} from '@io/recovery';
import {
  MemoryOriginStorage,
  originPrefix,
  type OriginStorage,
  type StorageEntry,
} from '@io/storage';

/**
 * M7 lane 5 — the autosave recovery copy (persistence §2).
 *
 * The contract's storage shape is exercised end to end: JSON metadata and
 * binary part records, incremental keyed writes, a manifest written last, the
 * pending list a restarted app reads, discard and pruning.
 */

const ORIGIN = 'http://127.0.0.1:47823';

function fixture(): Uint8Array {
  return new Uint8Array(
    readFileSync(new URL('../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url)),
  );
}

function opened(): OpenedWorkbook {
  return openWorkbook(fixture());
}

function edited(openedWorkbook: OpenedWorkbook, value = 2): DocumentSession {
  const session = new DocumentSession(openedWorkbook.workbook);
  const sheet = openedWorkbook.workbook.sheets[0]!;
  session.dispatch('cell.set', { sheetId: sheet.id, row: 1, col: 0, value: numberValue(value) });
  return session;
}

function input(
  docId: string,
  openedWorkbook: OpenedWorkbook,
  session: DocumentSession,
  overrides: Partial<RecoverySaveInput> = {},
): RecoverySaveInput {
  return {
    docId,
    displayName: 'book.xlsx',
    snapshot: snapshotWorkbook(session.workbook, openedWorkbook.package),
    compatibility: openedWorkbook.compatibility,
    savedFileModifiedAt: 1_000,
    savedFileSize: 4_000,
    activeSheetId: session.workbook.sheets[0]!.id,
    now: 2_000,
    ...overrides,
  };
}

/** Fails the manifest transaction once, to model a crash mid-write. */
class CrashOnManifestStorage implements OriginStorage {
  failNextManifest = false;
  quotaNextRecord = false;
  constructor(private readonly inner: MemoryOriginStorage) {}
  get origin(): string {
    return this.inner.origin;
  }
  get prefix(): string {
    return this.inner.prefix;
  }
  get(key: string): Promise<unknown> {
    return this.inner.get(key);
  }
  getAll(prefix: string): Promise<readonly (readonly [string, unknown])[]> {
    return this.inner.getAll(prefix);
  }
  keys(prefix: string): Promise<readonly string[]> {
    return this.inner.keys(prefix);
  }
  async setMany(entries: readonly StorageEntry[]): Promise<void> {
    if (
      this.quotaNextRecord &&
      entries.some((entry) => entry.key.includes('/sheet/') || entry.key.includes('/part/'))
    ) {
      this.quotaNextRecord = false;
      const error = new Error('quota');
      error.name = 'QuotaExceededError';
      throw error;
    }
    if (this.failNextManifest && entries.some((entry) => entry.key.endsWith('/manifest'))) {
      this.failNextManifest = false;
      throw new Error('simulated crash before the manifest write');
    }
    await this.inner.setMany(entries);
  }
  deleteMany(keys: readonly string[]): Promise<void> {
    return this.inner.deleteMany(keys);
  }
}

describe('recovery copy lifecycle', () => {
  it('stores JSON metadata and binary part records, never base64 part data', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    const session = edited(workbook);

    await repository.save(input('doc-1', workbook, session));

    const pending = await repository.listPending();
    expect(pending).toHaveLength(1);
    expect(pending[0]!.docId).toBe('doc-1');
    expect(pending[0]!.displayName).toBe('book.xlsx');
    expect(pending[0]!.sizeBytes).toBeGreaterThan(0);

    const manifestKey = `${originPrefix(ORIGIN)}doc/doc-1/manifest`;
    const storedManifest = await storage.get(manifestKey);
    expect(typeof storedManifest).toBe('string');
    const manifest = JSON.parse(storedManifest as string) as {
      workbook: { sheets: unknown[] };
      parts: { path: string; recordKey: string }[];
    };
    // Sheet records are separate; the manifest keeps only the refs.
    expect(manifest.workbook.sheets).toEqual([]);
    const chartRef = manifest.parts.find((part) => part.path === 'xl/charts/chart1.xml');
    expect(chartRef).toBeDefined();
    const chartBytes = await storage.get(chartRef!.recordKey);
    expect(chartBytes).toBeInstanceOf(Uint8Array);
    expect((chartBytes as Uint8Array).byteLength).toBeGreaterThan(0);

    // A restarted app creates a new repository against the same origin store.
    const restarted = new RecoveryRepository(storage, () => 9_000);
    const afterRestart = await restarted.listPending();
    expect(afterRestart).toHaveLength(1);
    expect(afterRestart[0]!.docId).toBe('doc-1');
    expect(await restarted.restore('doc-1')).not.toBeNull();
  });

  it('restores the unsaved edit and later saves Tier B parts byte-identical', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    const session = edited(workbook, 43);
    await repository.save(input('doc-1', workbook, session));

    const restored = await repository.restore('doc-1');
    expect(restored).not.toBeNull();
    const sheet = restored!.workbook.sheets[0]!;
    expect(sheet.cells.get('1:0')?.value).toEqual({ type: 'number', value: 43 });
    expect(restored!.activeSheetId).toBe(sheet.id);

    const restoredSession = new DocumentSession(restored!.workbook);
    // A restored copy has no command history, so the session marks every
    // record dirty: the next Save must write the recovered model.
    restoredSession.markRestored();
    expect(restoredSession.hasUnsavedChanges).toBe(true);
    const saved = saveWorkbook({
      workbook: restored!.workbook,
      package: restored!.package,
      dirtyCells: restoredSession.dirtyCellMap(),
      dirtyDimensions: restoredSession.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(restoredSession.dirtyMergeSheets()),
    });
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 43,
    });
    const chart = reopened.workbook.opaqueParts.get('xl/charts/chart1.xml');
    const originalChart = workbook.workbook.opaqueParts.get('xl/charts/chart1.xml');
    expect(Buffer.compare(Buffer.from(chart!.bytes), Buffer.from(originalChart!.bytes))).toBe(0);
  });

  it('offers a copy only when it is newer than the saved file', () => {
    expect(recoveryIsNewerThanSavedFile({ updatedAt: 2_000, savedFileModifiedAt: 1_000 })).toBe(
      true,
    );
    expect(recoveryIsNewerThanSavedFile({ updatedAt: 1_000, savedFileModifiedAt: 1_000 })).toBe(
      false,
    );
    // A new workbook has no file stamp, so there is unsaved work by definition.
    expect(recoveryIsNewerThanSavedFile({ updatedAt: 1_000, savedFileModifiedAt: null })).toBe(
      true,
    );
  });

  it('discards the records and the pending entry together', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    await repository.save(input('doc-1', workbook, edited(workbook)));

    await repository.discard('doc-1');

    expect(await repository.listPending()).toEqual([]);
    expect(await repository.restore('doc-1')).toBeNull();
    expect(await storage.keys(`${originPrefix(ORIGIN)}doc/doc-1/`)).toEqual([]);
  });

  it('leaves the previous copy restorable when a write crashes before the manifest', async () => {
    const inner = new MemoryOriginStorage(ORIGIN);
    const storage = new CrashOnManifestStorage(inner);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();

    await repository.save(input('doc-1', workbook, edited(workbook, 1), { now: 1_500 }));
    storage.failNextManifest = true;
    await expect(
      repository.save(input('doc-1', workbook, edited(workbook, 2), { now: 2_000 })),
    ).rejects.toThrow(/simulated crash/);

    const restored = await repository.restore('doc-1');
    expect(restored!.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 1,
    });
  });

  it('keeps the last good copy when local storage reports a quota error', async () => {
    const inner = new MemoryOriginStorage(ORIGIN);
    const storage = new CrashOnManifestStorage(inner);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();

    await repository.save(input('doc-1', workbook, edited(workbook, 1), { now: 1_500 }));
    storage.quotaNextRecord = true;
    await expect(
      repository.save(input('doc-1', workbook, edited(workbook, 2), { now: 2_000 })),
    ).rejects.toBeInstanceOf(RecoveryQuotaError);

    const restored = await repository.restore('doc-1');
    expect(restored!.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 1,
    });
  });

  it('rewrites only the changed sheet and never rewrites unchanged parts', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    const session = edited(workbook, 1);
    await repository.save(input('doc-1', workbook, session, { now: 1_500 }));
    const manifestBefore = await repository.readManifest('doc-1');
    const partKeys = new Set(manifestBefore!.parts.map((part) => part.recordKey));

    session.dispatch('cell.set', {
      sheetId: session.workbook.sheets[0]!.id,
      row: 1,
      col: 1,
      value: numberValue(7),
    });
    const savedAfter = await repository.save(
      input('doc-1', workbook, session, {
        now: 2_000,
        changedSheets: new Set([session.workbook.sheets[0]!.id]),
        partsChanged: false,
      }),
    );

    const newBatches = savedAfter.batches;
    expect(newBatches).toHaveLength(3);
    const writtenKeys = newBatches.flat().filter((key) => !key.startsWith('-'));
    expect(writtenKeys.some((key) => partKeys.has(key))).toBe(false);
    expect(writtenKeys.filter((key) => key.includes('/sheet/'))).toHaveLength(1);
    // The manifest follows the records; the old sheet record is then removed.
    expect(newBatches[1]![0]).toContain('/manifest');
    expect(newBatches[2]!.every((key) => key.startsWith('-'))).toBe(true);
    // The unchanged part refs carry forward, so a restore still has every part.
    const manifestAfter = await repository.readManifest('doc-1');
    expect(new Set(manifestAfter!.parts.map((part) => part.recordKey))).toEqual(partKeys);
  });

  it('writes only the manifest when nothing changed', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    const session = edited(workbook, 1);
    await repository.save(input('doc-1', workbook, session, { now: 1_500 }));

    const savedAfter = await repository.save(
      input('doc-1', workbook, session, {
        now: 2_000,
        changedSheets: new Set([session.workbook.sheets[0]!.id]),
        partsChanged: false,
      }),
    );

    const newBatches = savedAfter.batches;
    expect(newBatches).toHaveLength(1);
    expect(newBatches[0]).toHaveLength(1);
    expect(newBatches[0]![0]).toContain('/manifest');
  });

  it('splits a write above 8 MiB into batches and writes the manifest last', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    const session = edited(workbook, 1);
    const snapshot: WorkbookSnapshot = snapshotWorkbook(session.workbook, workbook.package);
    const large: WorkbookSnapshot = {
      ...snapshot,
      parts: [
        ...snapshot.parts,
        { path: 'xl/large.bin', kind: 'opaque', bytes: new Uint8Array(9 * 1024 * 1024) },
      ],
    };

    const saved = await repository.save(
      input('doc-1', workbook, session, { snapshot: large, now: 2_000 }),
    );

    expect(saved.batches.length).toBeGreaterThan(2);
    const last = saved.batches[saved.batches.length - 1]!;
    expect(last).toHaveLength(1);
    expect(last[0]).toContain('/manifest');
    // The write was chunked instead of one 9 MB transaction.
    const recordBatches = saved.batches.slice(0, -1);
    expect(recordBatches.length).toBeGreaterThan(1);
    for (const batch of recordBatches) {
      expect(batch.length).toBeGreaterThan(0);
    }
    expect(RECOVERY_BATCH_BYTES).toBeLessThan(9 * 1024 * 1024);
  });

  it('prunes expired copies and keeps the newest pending documents', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 100_000);
    const workbook = opened();
    const session = edited(workbook);

    for (let index = 0; index < MAX_PENDING_RECOVERY_DOCS + 2; index += 1) {
      await repository.save(
        input(`doc-${index}`, workbook, session, {
          now: 1_000 + index,
          snapshot: snapshotWorkbook(session.workbook, workbook.package),
        }),
      );
    }
    const pending = await repository.listPending();
    expect(pending).toHaveLength(MAX_PENDING_RECOVERY_DOCS);
    expect(pending[0]!.docId).toBe(`doc-${MAX_PENDING_RECOVERY_DOCS + 1}`);
    // The two evicted copies left no records behind.
    expect(await storage.keys(`${originPrefix(ORIGIN)}doc/doc-0/`)).toEqual([]);

    const removed = await repository.prune(1_000 + MAX_PENDING_RECOVERY_DOCS + 2);
    expect(removed).toEqual([]);

    const expired = await repository.save(
      input('doc-old', workbook, session, {
        now: 1_000 + MAX_PENDING_RECOVERY_DOCS + 3,
      }),
    );
    expect(expired.manifest.docId).toBe('doc-old');
    const removedAfterAge = await repository.prune(
      1_000 + MAX_PENDING_RECOVERY_DOCS + 3 + RECOVERY_MAX_AGE_MS + 1,
    );
    expect(removedAfterAge).toContain('doc-old');
    expect(await repository.restore('doc-old')).toBeNull();
  });

  it('lists only copies whose manifest still exists', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const repository = new RecoveryRepository(storage, () => 2_000);
    const workbook = opened();
    await repository.save(input('doc-1', workbook, edited(workbook)));

    // A broken pending row (for example a crash between writes) is dropped.
    await storage.setMany([
      {
        key: `${originPrefix(ORIGIN)}pending`,
        value: [{ docId: 'ghost', displayName: 'ghost.xlsx', origin: ORIGIN, updatedAt: 9_000 }],
      },
    ]);
    expect(await repository.listPending()).toEqual([]);
  });
});

describe('origin-scoped persistence', () => {
  it('keeps two origins from seeing each other’s recovery copies', async () => {
    const first = new MemoryOriginStorage('http://127.0.0.1:47823');
    const second = new MemoryOriginStorage('http://127.0.0.1:50000');
    const firstRepository = new RecoveryRepository(first, () => 2_000);
    const secondRepository = new RecoveryRepository(second, () => 2_000);
    const workbook = opened();
    await firstRepository.save(input('doc-1', workbook, edited(workbook)));

    expect(await secondRepository.listPending()).toEqual([]);
    expect(await secondRepository.restore('doc-1')).toBeNull();
    expect((await first.keys('')).every((key) => key.startsWith(originPrefix(first.origin)))).toBe(
      true,
    );
    expect(originPrefix(first.origin)).not.toBe(originPrefix(second.origin));
  });

  it('keeps the recent list per origin', async () => {
    const { RecentRepository } = await import('@io/recent');
    const first = new RecentRepository(new MemoryOriginStorage('http://127.0.0.1:47823'));
    const second = new RecentRepository(new MemoryOriginStorage('http://127.0.0.1:50000'));
    await first.remember({
      id: 'recent:a',
      name: 'a.xlsx',
      fileSize: 1,
      lastModified: 1,
      handle: null,
    });
    expect(await first.list()).toHaveLength(1);
    expect(await second.list()).toEqual([]);
  });
});
