/**
 * Origin-scoped storage for the two persistence tables (persistence §2, §3).
 *
 * Recovery copies and the recent list live in IndexedDB. The browser scopes a
 * database to the page's origin, so the launcher's stable
 * `http://127.0.0.1:47823` origin is the guarantee that the tables survive a
 * restart (distribution §3). A `--port N` run is a different origin and
 * therefore a separate browser storage area; the launcher documents that.
 *
 * Every key also carries the origin explicitly. IndexedDB already separates
 * origins, so the prefix is redundant in the browser; it keeps the guarantee
 * visible in the key layout, keeps a shared test backend honest, and would keep
 * two origins apart if a future storage medium ever shared one database.
 */

/** The value shape IndexedDB can structured-clone. */
export type StorageValue = unknown;

export interface StorageEntry {
  readonly key: string;
  readonly value: StorageValue;
}

export interface OriginStorage {
  /** The origin this storage belongs to, for diagnostics and key names. */
  readonly origin: string;
  /** The prefix every key in this storage carries. */
  readonly prefix: string;
  get(key: string): Promise<StorageValue | undefined>;
  getAll(prefix: string): Promise<readonly (readonly [string, StorageValue])[]>;
  keys(prefix: string): Promise<readonly string[]>;
  /** One storage transaction; the contract's chunking unit. */
  setMany(entries: readonly StorageEntry[]): Promise<void>;
  deleteMany(keys: readonly string[]): Promise<void>;
}

export const STORAGE_APP_NAME = 'calcplus';
export const STORAGE_SCHEMA_VERSION = 1;
export const STORAGE_DB_NAME = 'calcplus';
export const STORAGE_STORE_NAME = 'records';

/** The origin the app is running on, or a stable placeholder outside a browser. */
export function currentOrigin(): string {
  if (typeof location === 'undefined') return 'no-origin';
  return location.origin;
}

/**
 * Persistence §2: a `--port N` launcher run is a separate origin. Keeping the
 * origin inside every key makes that separation explicit instead of relying on
 * a comment about browser behaviour.
 */
export function originPrefix(origin: string): string {
  return `${STORAGE_APP_NAME}@${origin}/v${STORAGE_SCHEMA_VERSION}/`;
}

/** A transaction-counting in-memory backend, for tests and non-browser hosts. */
export class MemoryOriginStorage implements OriginStorage {
  readonly prefix: string;
  private readonly values = new Map<string, StorageValue>();
  /** The keys written by each transaction, in order. */
  readonly batches: string[][] = [];

  constructor(readonly origin = currentOrigin()) {
    this.prefix = originPrefix(origin);
  }

  async get(key: string): Promise<StorageValue | undefined> {
    return this.values.get(key);
  }

  async getAll(prefix: string): Promise<readonly (readonly [string, StorageValue])[]> {
    return [...this.values.entries()].filter(([key]) => key.startsWith(prefix));
  }

  async keys(prefix: string): Promise<readonly string[]> {
    return [...this.values.keys()].filter((key) => key.startsWith(prefix));
  }

  async setMany(entries: readonly StorageEntry[]): Promise<void> {
    if (entries.length === 0) return;
    for (const entry of entries) this.values.set(entry.key, entry.value);
    this.batches.push(entries.map((entry) => entry.key));
  }

  async deleteMany(keys: readonly string[]): Promise<void> {
    if (keys.length === 0) return;
    for (const key of keys) this.values.delete(key);
    this.batches.push(keys.map((key) => `-${key}`));
  }

  /** Test visibility: the number of stored keys, including the manifest. */
  get size(): number {
    return this.values.size;
  }
}

/** IndexedDB-backed storage. One database per origin is created by the browser. */
export class IndexedDbOriginStorage implements OriginStorage {
  readonly prefix: string;
  private database: Promise<IDBDatabase> | null = null;

  constructor(
    readonly origin = currentOrigin(),
    private readonly factory: IDBFactory = globalThis.indexedDB,
  ) {
    this.prefix = originPrefix(origin);
  }

  async get(key: string): Promise<StorageValue | undefined> {
    const database = await this.open();
    return requestResult(
      database.transaction(STORAGE_STORE_NAME, 'readonly').objectStore(STORAGE_STORE_NAME).get(key),
    );
  }

  async getAll(prefix: string): Promise<readonly (readonly [string, StorageValue])[]> {
    const database = await this.open();
    const store = database
      .transaction(STORAGE_STORE_NAME, 'readonly')
      .objectStore(STORAGE_STORE_NAME);
    // Ask both requests before awaiting so the transaction stays active.
    const [keys, values] = await Promise.all([
      requestResult<IDBValidKey[]>(store.getAllKeys()),
      requestResult<StorageValue[]>(store.getAll()),
    ]);
    const result: (readonly [string, StorageValue])[] = [];
    keys.forEach((key, index) => {
      if (typeof key !== 'string' || !key.startsWith(prefix)) return;
      result.push([key, values[index]]);
    });
    return result;
  }

  async keys(prefix: string): Promise<readonly string[]> {
    const database = await this.open();
    const keys = await requestResult<IDBValidKey[]>(
      database
        .transaction(STORAGE_STORE_NAME, 'readonly')
        .objectStore(STORAGE_STORE_NAME)
        .getAllKeys(),
    );
    return keys.filter((key): key is string => typeof key === 'string' && key.startsWith(prefix));
  }

  async setMany(entries: readonly StorageEntry[]): Promise<void> {
    if (entries.length === 0) return;
    const database = await this.open();
    const transaction = database.transaction(STORAGE_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORAGE_STORE_NAME);
    for (const entry of entries) store.put(entry.value, entry.key);
    await transactionDone(transaction);
  }

  async deleteMany(keys: readonly string[]): Promise<void> {
    if (keys.length === 0) return;
    const database = await this.open();
    const transaction = database.transaction(STORAGE_STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORAGE_STORE_NAME);
    for (const key of keys) store.delete(key);
    await transactionDone(transaction);
  }

  close(): void {
    if (this.database === null) return;
    void this.database.then((database) => database.close());
    this.database = null;
  }

  private open(): Promise<IDBDatabase> {
    this.database ??= new Promise<IDBDatabase>((resolve, reject) => {
      const request = this.factory.open(STORAGE_DB_NAME, 1);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORAGE_STORE_NAME)) {
          database.createObjectStore(STORAGE_STORE_NAME);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error('IndexedDB open failed.'));
      request.onblocked = () => reject(new Error('IndexedDB is blocked by another tab.'));
    });
    return this.database;
  }
}

/**
 * The browser uses IndexedDB under the current origin. A host without
 * IndexedDB (Node tests, an unsupported browser) gets the in-memory backend so
 * the persistence flows still run, but nothing survives the session there.
 */
export function createOriginStorage(origin: string = currentOrigin()): OriginStorage {
  if (typeof indexedDB === 'undefined') return new MemoryOriginStorage(origin);
  return new IndexedDbOriginStorage(origin);
}

function requestResult<T>(request: IDBRequest): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result as T);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed.'));
  });
}

function transactionDone(transaction: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB write failed.'));
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB write aborted.'));
  });
}
