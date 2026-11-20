import { describe, expect, it } from 'vitest';
import {
  RECENT_ENTRY_LIMIT,
  RecentRepository,
  recentDocumentId,
  type RecentEntry,
} from '@io/recent';
import { MemoryOriginStorage } from '@io/storage';
import type { FileHandleLike } from '@io/file-system';

/**
 * M7 lane 5 — the local recent list (persistence §3).
 *
 * Remembering a file, ordering by last opened, pruning past the cap, removing
 * an entry whose file is gone, and the origin-scoped separation the launcher's
 * stable port provides.
 */

const ORIGIN = 'http://127.0.0.1:47823';

function handle(name: string): FileHandleLike {
  return {
    name,
    async createWritable() {
      return { write: async () => undefined, close: async () => undefined };
    },
  };
}

function repository(now: () => number = () => 1_000): RecentRepository {
  return new RecentRepository(new MemoryOriginStorage(ORIGIN), now);
}

describe('recent list', () => {
  it('remembers a display name, origin, stamp and last-opened time', async () => {
    const store = repository(() => 5_000);
    const entries = await store.remember({
      id: recentDocumentId(ORIGIN, 'budget.xlsx'),
      name: 'budget.xlsx',
      fileSize: 4_096,
      lastModified: 1_234,
      handle: handle('budget.xlsx'),
    });

    expect(entries).toHaveLength(1);
    const entry = entries[0]!;
    expect(entry.name).toBe('budget.xlsx');
    expect(entry.origin).toBe(ORIGIN);
    expect(entry.lastOpenedAt).toBe(5_000);
    expect(entry.fileSize).toBe(4_096);
    expect(entry.lastModified).toBe(1_234);
    expect(entry.handle?.name).toBe('budget.xlsx');
    // No absolute path is kept or shown.
    expect(JSON.stringify({ ...entry, handle: undefined })).not.toContain('/home/');
  });

  it('puts the newest entry first and refreshes an existing one', async () => {
    let now = 1_000;
    const store = repository(() => now);
    await store.remember({
      id: 'a',
      name: 'a.xlsx',
      fileSize: 1,
      lastModified: 1,
      handle: null,
      now: 1_000,
    });
    await store.remember({
      id: 'b',
      name: 'b.xlsx',
      fileSize: 1,
      lastModified: 1,
      handle: null,
      now: 2_000,
    });
    expect((await store.list()).map((entry) => entry.name)).toEqual(['b.xlsx', 'a.xlsx']);

    now = 3_000;
    const refreshed = await store.remember({
      id: 'a',
      name: 'a.xlsx',
      fileSize: 2,
      lastModified: 5,
      handle: null,
      now: 3_000,
    });
    expect(refreshed.map((entry) => entry.name)).toEqual(['a.xlsx', 'b.xlsx']);
    expect(refreshed[0]!.fileSize).toBe(2);
    expect(refreshed).toHaveLength(2);
  });

  it(`prunes past the ${RECENT_ENTRY_LIMIT}-entry cap, oldest first`, async () => {
    const store = repository();
    for (let index = 0; index < RECENT_ENTRY_LIMIT + 3; index += 1) {
      await store.remember({
        id: `doc-${index}`,
        name: `doc-${index}.xlsx`,
        fileSize: index,
        lastModified: index,
        handle: null,
        now: 1_000 + index,
      });
    }
    const entries = await store.list();
    expect(entries).toHaveLength(RECENT_ENTRY_LIMIT);
    expect(entries[0]!.name).toBe(`doc-${RECENT_ENTRY_LIMIT + 2}.xlsx`);
    expect(entries.some((entry) => entry.name === 'doc-0.xlsx')).toBe(false);
    expect(entries.some((entry) => entry.name === 'doc-1.xlsx')).toBe(false);
  });

  it('removes an entry whose permission was denied or whose file is gone', async () => {
    const store = repository();
    await store.remember({
      id: 'gone',
      name: 'gone.xlsx',
      fileSize: 1,
      lastModified: 1,
      handle: handle('gone.xlsx'),
    });
    const remaining = await store.remove('gone');
    expect(remaining).toEqual([]);
    expect(await store.list()).toEqual([]);
  });

  it('keeps a name-only entry on a browser without handle persistence', async () => {
    const store = repository();
    const [entry] = await store.remember({
      id: 'name-only',
      name: 'name-only.xlsx',
      fileSize: null,
      lastModified: null,
      handle: null,
    });
    expect(entry).toBeDefined();
    expect((entry as RecentEntry).handle).toBeNull();
  });

  it('survives a restart: a new repository reads the same storage', async () => {
    const storage = new MemoryOriginStorage(ORIGIN);
    const first = new RecentRepository(storage);
    await first.remember({
      id: 'kept',
      name: 'kept.xlsx',
      fileSize: 12,
      lastModified: 34,
      handle: handle('kept.xlsx'),
    });
    const restarted = new RecentRepository(storage);
    const [entry] = await restarted.list();
    expect(entry?.name).toBe('kept.xlsx');
    expect(entry?.origin).toBe(ORIGIN);
    expect(entry?.handle?.name).toBe('kept.xlsx');
  });

  it('keeps a stable id per origin and file name, with no path', () => {
    expect(recentDocumentId(ORIGIN, 'book.xlsx')).toBe(`recent:${ORIGIN}:book.xlsx`);
    expect(recentDocumentId('http://127.0.0.1:50000', 'book.xlsx')).not.toBe(
      recentDocumentId(ORIGIN, 'book.xlsx'),
    );
    expect(recentDocumentId(ORIGIN, 'book.xlsx')).not.toContain('/home/');
    expect(recentDocumentId(ORIGIN, 'book.xlsx')).not.toContain('\\');
  });

  it('scopes the list to one origin', async () => {
    const first = new RecentRepository(new MemoryOriginStorage('http://127.0.0.1:47823'));
    const second = new RecentRepository(new MemoryOriginStorage('http://127.0.0.1:50000'));
    await first.remember({
      id: 'a',
      name: 'a.xlsx',
      fileSize: 1,
      lastModified: 1,
      handle: null,
    });
    expect((await first.list())[0]!.origin).toBe('http://127.0.0.1:47823');
    expect(await second.list()).toEqual([]);
  });
});
