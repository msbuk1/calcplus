/**
 * The local recent list (persistence §3).
 *
 * It remembers a display name, the last-opened time, a stable document id and,
 * on a browser with the File System Access API, the file handle. It never
 * stores or shows an absolute path. The list is bounded; the oldest entries
 * are pruned as new files are opened.
 */
import type { FileHandleLike } from './file-system';
import type { OriginStorage } from './storage';

export const RECENT_ENTRY_LIMIT = 10;

export interface RecentEntry {
  readonly id: string;
  /** The file's display name only. No directory is ever kept. */
  readonly name: string;
  readonly origin: string;
  readonly lastOpenedAt: number;
  readonly fileSize: number | null;
  readonly lastModified: number | null;
  /** `null` on a browser without handle persistence (Firefox). */
  readonly handle: FileHandleLike | null;
}

export interface RememberRecentInput {
  readonly id: string;
  readonly name: string;
  readonly fileSize: number | null;
  readonly lastModified: number | null;
  readonly handle: FileHandleLike | null;
  readonly now?: number;
}

interface StoredRecentEntry {
  readonly id: string;
  readonly name: string;
  readonly lastOpenedAt: number;
  readonly fileSize: number | null;
  readonly lastModified: number | null;
  /** A structured-cloneable FileSystemFileHandle, or absent. */
  readonly handle?: unknown;
}

/**
 * The stable id a recent entry uses. Two files with the same name on one
 * origin share an entry because no absolute path is stored: the contract
 * forbids showing or keeping one (persistence §3).
 */
export function recentDocumentId(origin: string, name: string): string {
  return `recent:${origin}:${name}`;
}

export class RecentRepository {
  constructor(
    private readonly storage: OriginStorage,
    private readonly clock: () => number = () => Date.now(),
  ) {}

  async list(): Promise<readonly RecentEntry[]> {
    return (await this.read()).map((entry) => toRecentEntry(entry, this.storage.origin));
  }

  /** Adds or refreshes an entry and prunes the oldest past the cap. */
  async remember(input: RememberRecentInput): Promise<readonly RecentEntry[]> {
    const now = input.now ?? this.clock();
    const entries = (await this.read()).filter((entry) => entry.id !== input.id);
    entries.unshift({
      id: input.id,
      name: input.name,
      lastOpenedAt: now,
      fileSize: input.fileSize,
      lastModified: input.lastModified,
      ...(input.handle === null ? {} : { handle: input.handle }),
    });
    const pruned = entries
      .sort((left, right) => right.lastOpenedAt - left.lastOpenedAt)
      .slice(0, RECENT_ENTRY_LIMIT);
    await this.write(pruned);
    return pruned.map((entry) => toRecentEntry(entry, this.storage.origin));
  }

  /** Drops an entry whose handle was denied or whose file is gone. */
  async remove(id: string): Promise<readonly RecentEntry[]> {
    const entries = (await this.read()).filter((entry) => entry.id !== id);
    await this.write(entries);
    return entries.map((entry) => toRecentEntry(entry, this.storage.origin));
  }

  private async read(): Promise<StoredRecentEntry[]> {
    const value = await this.storage.get(this.key());
    if (!Array.isArray(value)) return [];
    return value.filter(
      (entry): entry is StoredRecentEntry =>
        typeof entry === 'object' &&
        entry !== null &&
        typeof (entry as StoredRecentEntry).id === 'string' &&
        typeof (entry as StoredRecentEntry).name === 'string' &&
        typeof (entry as StoredRecentEntry).lastOpenedAt === 'number',
    );
  }

  private async write(entries: readonly StoredRecentEntry[]): Promise<void> {
    if (entries.length === 0) {
      await this.storage.deleteMany([this.key()]);
      return;
    }
    await this.storage.setMany([{ key: this.key(), value: entries }]);
  }

  private key(): string {
    return `${this.storage.prefix}recent`;
  }
}

function toRecentEntry(entry: StoredRecentEntry, origin: string): RecentEntry {
  return {
    id: entry.id,
    name: entry.name,
    origin,
    lastOpenedAt: entry.lastOpenedAt,
    fileSize: entry.fileSize ?? null,
    lastModified: entry.lastModified ?? null,
    handle: asFileHandle(entry.handle),
  };
}

function asFileHandle(value: unknown): FileHandleLike | null {
  if (typeof value !== 'object' || value === null) return null;
  const candidate = value as Partial<FileHandleLike>;
  if (typeof candidate.name !== 'string' || typeof candidate.createWritable !== 'function') {
    return null;
  }
  return candidate as FileHandleLike;
}
