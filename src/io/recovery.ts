/**
 * The autosave recovery copy (persistence §2).
 *
 * The copy is stored under the open document's stable id. The workbook
 * metadata, sheet records and package parts are written as separate records:
 *
 * - metadata is JSON, never base64 part bytes;
 * - every part is a binary record;
 * - records are content-addressed, so an unchanged sheet or part is not
 *   rewritten, and an edit only adds the records it changed;
 * - records are written before the manifest, which names the exact records a
 *   restore needs. A crash mid-write therefore leaves the previous manifest,
 *   and its records, intact.
 *
 * A snapshot above 8 MiB is written in part-sized batches, so no one IndexedDB
 * transaction carries the whole workbook. The manifest is always its own last
 * transaction.
 */
import type { CompatibilityInfo } from '@ooxml/open';
import {
  SNAPSHOT_SCHEMA_VERSION,
  restoreWorkbook,
  type SnapshotPackageMetadata,
  type SnapshotPart,
  type SnapshotSheet,
  type SnapshotWorkbookMetadata,
  type WorkbookSnapshot,
} from '@ooxml/snapshot';
import type { SheetId } from '@model/types';
import type { OriginStorage, StorageValue } from './storage';

export const RECOVERY_SCHEMA_VERSION = 1;
/** Persistence §2: debounce after the last document edit. */
export const RECOVERY_DEBOUNCE_MS = 3_000;
/** Above this, the write is split into batches with the manifest written last. */
export const RECOVERY_SINGLE_TRANSACTION_BYTES = 8 * 1024 * 1024;
/** The batch size used once a write is above the single-transaction limit. */
export const RECOVERY_BATCH_BYTES = 1 * 1024 * 1024;
/** A small, bounded pending list; the oldest copies are pruned. */
export const MAX_PENDING_RECOVERY_DOCS = 10;
/** Copies older than this are pruned on startup. */
export const RECOVERY_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

export interface RecoverySheetRef {
  readonly id: SheetId;
  readonly name: string;
  readonly recordKey: string;
  readonly fingerprint: string;
  readonly size: number;
}

export interface RecoveryPartRef {
  readonly path: string;
  readonly kind: SnapshotPart['kind'];
  readonly recordKey: string;
  readonly fingerprint: string;
  readonly size: number;
}

export interface RecoveryManifest {
  readonly kind: 'calcplus-recovery';
  readonly schemaVersion: number;
  readonly snapshotSchemaVersion: number;
  readonly docId: string;
  readonly displayName: string;
  readonly origin: string;
  readonly createdAt: number;
  readonly updatedAt: number;
  readonly savedFileModifiedAt: number | null;
  readonly savedFileSize: number | null;
  /** The active sheet at write time, so a restore opens where the crash left it. */
  readonly activeSheetId: SheetId | null;
  readonly workbook: SnapshotWorkbookMetadata;
  readonly package: SnapshotPackageMetadata;
  readonly sheets: readonly RecoverySheetRef[];
  readonly parts: readonly RecoveryPartRef[];
  readonly compatibility: CompatibilityInfo | null;
}

export interface PendingRecoveryDocument {
  readonly docId: string;
  readonly displayName: string;
  readonly origin: string;
  readonly updatedAt: number;
  readonly savedFileModifiedAt: number | null;
  readonly savedFileSize: number | null;
  /** A File System Access handle, when the browser gave one (structured-cloneable). */
  readonly handle?: unknown;
}

export interface PendingRecovery extends PendingRecoveryDocument {
  /** The bytes a restore would read, for the prompt and diagnostics. */
  readonly sizeBytes: number;
}

export interface RecoverySaveInput {
  readonly docId: string;
  readonly displayName: string;
  readonly snapshot: WorkbookSnapshot;
  readonly compatibility: CompatibilityInfo;
  readonly savedFileModifiedAt: number | null;
  readonly savedFileSize: number | null;
  readonly activeSheetId?: SheetId | null;
  readonly handle?: unknown;
  /**
   * The sheets whose model records may have changed since the previous write.
   * `all` (the default) reads every sheet. The hint only avoids serialising
   * unchanged sheets; the record key still proves whether the bytes changed.
   */
  readonly changedSheets?: ReadonlySet<SheetId> | 'all';
  /**
   * `false` reuses the previous part records. Documents do not change their
   * source parts until a structural edit or a save, so the app passes `false`
   * between ordinary edits.
   */
  readonly partsChanged?: boolean;
  readonly now?: number;
}

export interface RecoverySaveResult {
  readonly manifest: RecoveryManifest;
  readonly wroteRecords: number;
  readonly bytesWritten: number;
  /** The storage transactions, in order; the manifest is the last one. */
  readonly batches: readonly (readonly string[])[];
}

export class RecoveryCorruptError extends Error {
  constructor(docId: string, detail: string) {
    super(`The recovery copy for "${docId}" is incomplete: ${detail}`);
    this.name = 'RecoveryCorruptError';
  }
}

export class RecoveryQuotaError extends Error {
  constructor(readonly reason: unknown) {
    super('The recovery copy could not be saved because local storage is full.');
    this.name = 'RecoveryQuotaError';
  }
}

export class RecoveryRepository {
  private readonly sheetCache = new Map<
    string,
    { readonly fingerprint: string; readonly json: string; readonly size: number }
  >();

  constructor(
    private readonly storage: OriginStorage,
    private readonly clock: () => number = () => Date.now(),
  ) {}

  get origin(): string {
    return this.storage.origin;
  }

  async save(input: RecoverySaveInput): Promise<RecoverySaveResult> {
    const now = input.now ?? this.clock();
    const previous = await this.readManifest(input.docId);
    const docPrefix = this.docPrefix(input.docId);
    const manifestKey = `${docPrefix}manifest`;

    const sheets: RecoverySheetRef[] = [];
    const sheetJson = new Map<string, string>();
    for (const sheet of input.snapshot.workbook.sheets) {
      const record = this.sheetRecord(input, sheet, previous);
      sheets.push(record.ref);
      sheetJson.set(record.ref.recordKey, record.json);
    }

    const { parts, partBytes } =
      input.partsChanged === false && previous !== null
        ? { parts: previous.parts, partBytes: new Map<string, Uint8Array>() }
        : splitParts(docPrefix, input.snapshot.parts);

    const writes = new Map<string, StorageValue>();
    const previousSheets = new Set((previous?.sheets ?? []).map((ref) => ref.recordKey));
    for (const ref of sheets) {
      if (previousSheets.has(ref.recordKey)) continue;
      const json = sheetJson.get(ref.recordKey);
      if (json !== undefined) writes.set(ref.recordKey, json);
    }
    const previousParts = new Set((previous?.parts ?? []).map((ref) => ref.recordKey));
    for (const ref of parts) {
      if (previousParts.has(ref.recordKey)) continue;
      const bytes = partBytes.get(ref.recordKey);
      if (bytes !== undefined) writes.set(ref.recordKey, bytes);
    }

    const manifest: RecoveryManifest = {
      kind: 'calcplus-recovery',
      schemaVersion: RECOVERY_SCHEMA_VERSION,
      snapshotSchemaVersion: SNAPSHOT_SCHEMA_VERSION,
      docId: input.docId,
      displayName: input.displayName,
      origin: this.storage.origin,
      createdAt: previous?.createdAt ?? now,
      updatedAt: now,
      savedFileModifiedAt: input.savedFileModifiedAt,
      savedFileSize: input.savedFileSize,
      activeSheetId: input.activeSheetId ?? null,
      workbook: { ...input.snapshot.workbook, sheets: [] },
      package: input.snapshot.package,
      sheets,
      parts,
      compatibility: input.compatibility,
    };
    const manifestJson = JSON.stringify(manifest);

    const batches = await this.writeRecords(writes, manifestKey, manifestJson);

    // Only after the new manifest names them can the old records be removed.
    const referenced = new Set<string>([
      manifestKey,
      ...sheets.map((ref) => ref.recordKey),
      ...parts.map((ref) => ref.recordKey),
    ]);
    const stored = await this.storage.keys(docPrefix);
    const stale = stored.filter((key) => !referenced.has(key));
    if (stale.length > 0) {
      await this.storage.deleteMany(stale);
      batches.push(stale.map((key) => `-${key}`));
    }

    const evicted = await this.upsertPending(manifest, input.handle, now);
    for (const docId of evicted) await this.deleteRecords(docId);

    return {
      manifest,
      wroteRecords: writes.size,
      bytesWritten: [...writes.values()].reduce<number>(
        (total, value) => total + valueSize(value),
        0,
      ),
      batches,
    };
  }

  async listPending(): Promise<readonly PendingRecovery[]> {
    const stored = await this.readPendingRaw();
    const pending: PendingRecovery[] = [];
    for (const entry of stored) {
      const manifest = await this.readManifest(entry.docId);
      if (manifest === null || manifest.origin !== this.storage.origin) continue;
      pending.push({
        docId: manifest.docId,
        displayName: manifest.displayName,
        origin: manifest.origin,
        updatedAt: manifest.updatedAt,
        savedFileModifiedAt: manifest.savedFileModifiedAt,
        savedFileSize: manifest.savedFileSize,
        sizeBytes: manifestSize(manifest),
        ...(entry.handle === undefined ? {} : { handle: entry.handle }),
      });
    }
    if (pending.length !== stored.length) await this.writePending(pending);
    return pending.sort((left, right) => right.updatedAt - left.updatedAt);
  }

  /** The newest copy the startup prompt should offer, or `null`. */
  async pendingToOffer(): Promise<PendingRecovery | null> {
    const pending = await this.listPending();
    return pending[0] ?? null;
  }

  async readManifest(docId: string): Promise<RecoveryManifest | null> {
    const value = await this.storage.get(`${this.docPrefix(docId)}manifest`);
    if (typeof value !== 'string') return null;
    try {
      const parsed = JSON.parse(value) as RecoveryManifest;
      if (
        parsed.kind !== 'calcplus-recovery' ||
        parsed.schemaVersion !== RECOVERY_SCHEMA_VERSION ||
        parsed.docId !== docId ||
        !Array.isArray(parsed.sheets) ||
        !Array.isArray(parsed.parts)
      ) {
        return null;
      }
      return parsed;
    } catch {
      return null;
    }
  }

  /**
   * Rebuilds the model and package from the store. It never reads the user's
   * file, so a file changed or deleted while the tab was closed cannot affect
   * the restore (persistence §2).
   */
  async restore(docId: string): Promise<{
    readonly workbook: ReturnType<typeof restoreWorkbook>['workbook'];
    readonly package: ReturnType<typeof restoreWorkbook>['package'];
    readonly compatibility: CompatibilityInfo | null;
    readonly fileName: string;
    readonly savedFileModifiedAt: number | null;
    readonly savedFileSize: number | null;
    readonly activeSheetId: SheetId | null;
    readonly handle: unknown | null;
    readonly manifest: RecoveryManifest;
    /**
     * A sheet record from an older build stored its AutoFilter criteria in a
     * shape that kept none of them. The copy still restores; the caller says
     * what was lost instead of leaving the filter gap unexplained.
     */
    readonly filterCriteriaDropped: boolean;
  } | null> {
    const manifest = await this.readManifest(docId);
    if (manifest === null) return null;

    const sheets: SnapshotSheet[] = [];
    let filterCriteriaDropped = false;
    for (const ref of manifest.sheets) {
      const json = await this.storage.get(ref.recordKey);
      if (typeof json !== 'string') {
        throw new RecoveryCorruptError(docId, `sheet record ${ref.recordKey} is missing`);
      }
      const sheet = JSON.parse(json) as SnapshotSheet;
      if (sheet.autoFilter !== null && !Array.isArray(sheet.autoFilter.columns)) {
        filterCriteriaDropped = true;
      }
      sheets.push(sheet);
    }

    const parts: SnapshotPart[] = [];
    for (const ref of manifest.parts) {
      const bytes = await this.storage.get(ref.recordKey);
      if (!(bytes instanceof Uint8Array)) {
        throw new RecoveryCorruptError(docId, `part record ${ref.recordKey} is missing`);
      }
      parts.push({ path: ref.path, kind: ref.kind, bytes });
    }

    const snapshot: WorkbookSnapshot = {
      schemaVersion: manifest.snapshotSchemaVersion,
      workbook: { ...manifest.workbook, sheets },
      package: manifest.package,
      parts,
    };
    const restored = restoreWorkbook(snapshot);
    const stored = await this.readPendingRaw();
    const handle = stored.find((entry) => entry.docId === docId)?.handle ?? null;
    return {
      workbook: restored.workbook,
      package: restored.package,
      compatibility: manifest.compatibility,
      fileName: manifest.displayName,
      savedFileModifiedAt: manifest.savedFileModifiedAt,
      savedFileSize: manifest.savedFileSize,
      activeSheetId: manifest.activeSheetId ?? null,
      handle,
      manifest,
      filterCriteriaDropped,
    };
  }

  async discard(docId: string): Promise<void> {
    const prefix = this.docPrefix(docId);
    const keys = await this.storage.keys(prefix);
    if (keys.length > 0) await this.storage.deleteMany(keys);
    for (const key of [...this.sheetCache.keys()]) {
      if (key.startsWith(`${docId}\u0000`)) this.sheetCache.delete(key);
    }
    const stored = await this.readPendingRaw();
    const remaining = stored.filter((entry) => entry.docId !== docId);
    if (remaining.length !== stored.length) await this.writePending(remaining);
  }

  /** Drops missing, expired and over-cap copies; returns the removed ids. */
  async prune(now: number = this.clock()): Promise<readonly string[]> {
    const stored = await this.readPendingRaw();
    const removed: string[] = [];
    const alive: PendingRecovery[] = [];
    for (const entry of stored) {
      const manifest = await this.readManifest(entry.docId);
      if (manifest === null || now - manifest.updatedAt > RECOVERY_MAX_AGE_MS) {
        removed.push(entry.docId);
        continue;
      }
      alive.push({
        docId: manifest.docId,
        displayName: manifest.displayName,
        origin: manifest.origin,
        updatedAt: manifest.updatedAt,
        savedFileModifiedAt: manifest.savedFileModifiedAt,
        savedFileSize: manifest.savedFileSize,
        sizeBytes: manifestSize(manifest),
        ...(entry.handle === undefined ? {} : { handle: entry.handle }),
      });
    }
    alive.sort((left, right) => right.updatedAt - left.updatedAt);
    for (const entry of alive.slice(MAX_PENDING_RECOVERY_DOCS)) removed.push(entry.docId);
    for (const docId of removed) await this.deleteRecords(docId);
    await this.writePending(alive.slice(0, MAX_PENDING_RECOVERY_DOCS));
    return removed;
  }

  private sheetRecord(
    input: RecoverySaveInput,
    sheet: SnapshotSheet,
    previous: RecoveryManifest | null,
  ): { readonly ref: RecoverySheetRef; readonly json: string } {
    const cacheKey = `${input.docId}\u0000${sheet.id}`;
    const hint = input.changedSheets ?? 'all';
    const previousRef = previous?.sheets.find((ref) => ref.id === sheet.id);
    const mustRead =
      hint === 'all' ||
      hint.has(sheet.id) ||
      previousRef?.name !== sheet.name ||
      !this.sheetCache.has(cacheKey);
    let cached = this.sheetCache.get(cacheKey);
    if (mustRead || cached === undefined) {
      const json = JSON.stringify(sheet);
      cached = { fingerprint: fingerprintText(json), json, size: byteLength(json) };
      this.sheetCache.set(cacheKey, cached);
    }
    return {
      ref: {
        id: sheet.id,
        name: sheet.name,
        recordKey: `${this.docPrefix(input.docId)}sheet/${cached.fingerprint}`,
        fingerprint: cached.fingerprint,
        size: cached.size,
      },
      json: cached.json,
    };
  }

  private async writeRecords(
    writes: ReadonlyMap<string, StorageValue>,
    manifestKey: string,
    manifestJson: string,
  ): Promise<(readonly string[])[]> {
    const entries = [...writes.entries()].map(([key, value]) => ({ key, value }));
    const totalBytes = entries.reduce((total, entry) => total + valueSize(entry.value), 0);
    const batches: string[][] = [];
    try {
      if (totalBytes <= RECOVERY_SINGLE_TRANSACTION_BYTES) {
        if (entries.length > 0) {
          await this.storage.setMany(entries);
          batches.push(entries.map((entry) => entry.key));
        }
      } else {
        let batch: typeof entries = [];
        let batchBytes = 0;
        for (const entry of entries) {
          const size = valueSize(entry.value);
          if (batch.length > 0 && batchBytes + size > RECOVERY_BATCH_BYTES) {
            await this.storage.setMany(batch);
            batches.push(batch.map((item) => item.key));
            batch = [];
            batchBytes = 0;
          }
          batch.push(entry);
          batchBytes += size;
        }
        if (batch.length > 0) {
          await this.storage.setMany(batch);
          batches.push(batch.map((item) => item.key));
        }
      }
      // The manifest is always its own transaction and always last.
      await this.storage.setMany([{ key: manifestKey, value: manifestJson }]);
      batches.push([manifestKey]);
      return batches;
    } catch (error) {
      if (isQuotaError(error)) throw new RecoveryQuotaError(error);
      throw error;
    }
  }

  private async upsertPending(
    manifest: RecoveryManifest,
    handle: unknown,
    now: number,
  ): Promise<readonly string[]> {
    const stored = await this.readPendingRaw();
    const previous = stored.find((entry) => entry.docId === manifest.docId);
    const rest = stored.filter((entry) => entry.docId !== manifest.docId);
    const next: PendingRecoveryDocument = {
      docId: manifest.docId,
      displayName: manifest.displayName,
      origin: manifest.origin,
      updatedAt: now,
      savedFileModifiedAt: manifest.savedFileModifiedAt,
      savedFileSize: manifest.savedFileSize,
      ...(handle === undefined && previous?.handle === undefined
        ? {}
        : { handle: handle ?? previous?.handle }),
    };
    rest.push(next);
    rest.sort((left, right) => right.updatedAt - left.updatedAt);
    const kept = rest.slice(0, MAX_PENDING_RECOVERY_DOCS);
    try {
      await this.writePending(kept);
    } catch (error) {
      if (isQuotaError(error)) throw new RecoveryQuotaError(error);
      throw error;
    }
    return rest.slice(MAX_PENDING_RECOVERY_DOCS).map((entry) => entry.docId);
  }

  private async readPendingRaw(): Promise<readonly PendingRecoveryDocument[]> {
    const value = await this.storage.get(this.pendingKey());
    if (!Array.isArray(value)) return [];
    return value.filter(
      (entry): entry is PendingRecoveryDocument =>
        typeof entry === 'object' &&
        entry !== null &&
        typeof (entry as PendingRecoveryDocument).docId === 'string',
    );
  }

  private async writePending(entries: readonly PendingRecoveryDocument[]): Promise<void> {
    if (entries.length === 0) {
      await this.storage.deleteMany([this.pendingKey()]);
      return;
    }
    await this.storage.setMany([{ key: this.pendingKey(), value: entries }]);
  }

  private async deleteRecords(docId: string): Promise<void> {
    const keys = await this.storage.keys(this.docPrefix(docId));
    if (keys.length > 0) await this.storage.deleteMany(keys);
  }

  private pendingKey(): string {
    return `${this.storage.prefix}pending`;
  }

  private docPrefix(docId: string): string {
    return `${this.storage.prefix}doc/${encodeURIComponent(docId)}/`;
  }
}

/**
 * Persistence §2: a copy is offered only when it has work newer than the file
 * it was opened from. A copy with no file (a new workbook) is always offered.
 */
export function recoveryIsNewerThanSavedFile(info: {
  readonly updatedAt: number;
  readonly savedFileModifiedAt: number | null;
}): boolean {
  if (info.savedFileModifiedAt === null) return true;
  return info.updatedAt > info.savedFileModifiedAt;
}

export function fingerprintText(text: string): string {
  let first = 0x811c9dc5;
  let second = 5381;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193);
    second = Math.imul(second, 33) ^ code;
  }
  return `${(first >>> 0).toString(16)}${(second >>> 0).toString(16)}-${text.length}`;
}

export function fingerprintBytes(bytes: Uint8Array): string {
  let first = 0x811c9dc5;
  let second = 5381;
  for (let index = 0; index < bytes.length; index += 1) {
    const byte = bytes[index]!;
    first = Math.imul(first ^ byte, 0x01000193);
    second = Math.imul(second, 33) ^ byte;
  }
  return `${(first >>> 0).toString(16)}${(second >>> 0).toString(16)}-${bytes.length}`;
}

function splitParts(
  docPrefix: string,
  parts: readonly SnapshotPart[],
): { readonly parts: RecoveryPartRef[]; readonly partBytes: Map<string, Uint8Array> } {
  const refs: RecoveryPartRef[] = [];
  const bytes = new Map<string, Uint8Array>();
  for (const part of parts) {
    const fingerprint = fingerprintBytes(part.bytes);
    const recordKey = `${docPrefix}part/${fingerprint}`;
    if (!bytes.has(recordKey)) {
      bytes.set(recordKey, part.bytes);
      refs.push({
        path: part.path,
        kind: part.kind,
        recordKey,
        fingerprint,
        size: part.bytes.byteLength,
      });
    }
  }
  return { parts: refs, partBytes: bytes };
}

function manifestSize(manifest: RecoveryManifest): number {
  return (
    manifest.sheets.reduce((total, ref) => total + ref.size, 0) +
    manifest.parts.reduce((total, ref) => total + ref.size, 0)
  );
}

function byteLength(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

function valueSize(value: StorageValue): number {
  if (typeof value === 'string') return byteLength(value);
  if (value instanceof Uint8Array) return value.byteLength;
  if (Array.isArray(value)) return JSON.stringify(value).length;
  if (typeof value === 'object' && value !== null) return JSON.stringify(value).length;
  return 0;
}

function isQuotaError(error: unknown): boolean {
  return (
    typeof error === 'object' &&
    error !== null &&
    'name' in error &&
    (error as { readonly name?: unknown }).name === 'QuotaExceededError'
  );
}
