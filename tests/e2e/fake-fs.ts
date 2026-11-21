import { Buffer } from 'node:buffer';
import type { BrowserContext } from '@playwright/test';
import { fixtureBytes } from './grid-helpers';

/**
 * A fake File System Access layer for e2e tests, installed with
 * `addInitScript` before any app script runs.
 *
 * Why it exists: `openFixture` uses `?fsa=0`, so `document.handle` is null
 * and Save falls back to a download. Nothing in the suite ever exercised a
 * real `FileSystemFileHandle`, which is how a save-after-restore bug escaped.
 *
 * How it stays honest:
 *
 * - The file's bytes and stamp live on a fake disk outside any handle, and
 *   are seeded identically in every page, like a real file that does not
 *   change while a tab is closed.
 * - A handle is a class instance whose own properties are data only
 *   (`name` plus a marker), so IndexedDB structured-clones it just like the
 *   browser clones a real `FileSystemFileHandle` into a recovery record.
 *   Prototype methods are not cloned, so reads are rehydrated by a thin
 *   `IDBRequest.result` wrapper; that stands in for the browser's own
 *   deserialisation.
 * - Permission is per-instance state and is never serialised. A handle
 *   revived from a recovery record therefore starts at `prompt`, exactly the
 *   state the maintainer's restored tab was in: `getFile()` throws until
 *   `requestPermission()` grants access again.
 * - `queryPermission`/`requestPermission` are real methods with tracked
 *   calls, and `getFile`/`createWritable` refuse to work without granted
 *   permission, so the tests can prove the app re-asks before touching the
 *   file.
 * - `showSaveFilePicker` enforces Chrome's user-gesture rule: it throws
 *   `SecurityError` unless a trusted gesture is live, and a permission prompt
 *   consumes the gesture. A save flow that stands a native confirm between the
 *   gesture and the picker therefore fails exactly like Chrome fails.
 */

export interface FakeFsInitOptions {
  readonly base64: string;
  readonly fileName: string;
  readonly fileLastModified: number;
}

export interface FakeFsPermissionCall {
  readonly mode: string;
  readonly handle: number;
}

export interface FakeFsCalls {
  openPicker: number;
  savePicker: number;
  getFile: number;
  createWritable: number;
  write: number;
  close: number;
  abort: number;
  /** Handle ids the save picker handed out. */
  savePickerHandles: number[];
  /** Handle ids `createWritable` ran on. */
  createWritableHandles: number[];
  queryPermission: FakeFsPermissionCall[];
  requestPermission: FakeFsPermissionCall[];
}

export interface FakeFsPlan {
  /** What `queryPermission` reports; `null` keeps the handle's own state. */
  query: 'granted' | 'prompt' | 'denied' | null;
  /** What `requestPermission` resolves to; `null` grants. */
  request: 'granted' | 'prompt' | 'denied' | null;
  /** Make `getFile` throw despite permission, like a file deleted on disk. */
  getFileThrows: boolean;
}

export interface FakeFsApi {
  readonly calls: FakeFsCalls;
  plan: FakeFsPlan;
  diskState(): { name: string; size: number; lastModified: number };
  /**
   * Simulate another program saving the file: the stamp always moves, and a
   * `replacement` (base64) also swaps the bytes, like a real external write.
   */
  externalEdit(replacement?: string): void;
}

declare global {
  interface Window {
    __fakeFs: FakeFsApi;
  }
}

/** The fixture every fake handle reads and writes. */
export const FAKE_FILE_NAME = 'm1-chart-unknown.xlsx';

/**
 * An old, fixed stamp for the fake file: the recovery copy is always newer
 * than the file, so startup always offers Restore, and the stamp captured at
 * open still matches after the tab dies (the file did not change on disk).
 */
export const FAKE_FILE_LAST_MODIFIED = 1_700_000_000_000;

/** Installs the fake pickers, fake disk and handle rehydration in every page. */
export async function installFakeFs(
  context: BrowserContext,
  options: { readonly fileName?: string } = {},
): Promise<void> {
  const fileName = options.fileName ?? FAKE_FILE_NAME;
  await context.addInitScript(fakeFsMain, {
    base64: Buffer.from(fixtureBytes(fileName)).toString('base64'),
    fileName,
    fileLastModified: FAKE_FILE_LAST_MODIFIED,
  } satisfies FakeFsInitOptions);
}

/**
 * The init script. It is serialised into the page, so it must be
 * self-contained: no imports, no helpers outside the function, no enums.
 */
export function fakeFsMain(options: FakeFsInitOptions): void {
  const binary = atob(options.base64);
  const seed = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    seed[index] = binary.charCodeAt(index);
  }

  // The fake disk. Every page starts from the same content, so a file that
  // "did not change while the tab was gone" keeps its stamp.
  const disk = {
    name: options.fileName,
    bytes: seed,
    lastModified: options.fileLastModified,
  };

  const calls: FakeFsCalls = {
    openPicker: 0,
    savePicker: 0,
    getFile: 0,
    createWritable: 0,
    write: 0,
    close: 0,
    abort: 0,
    savePickerHandles: [],
    createWritableHandles: [],
    queryPermission: [],
    requestPermission: [],
  };

  const plan: FakeFsPlan = {
    query: null,
    request: null,
    getFileThrows: false,
  };

  // Per-handle permission, never serialised: a revived handle starts at
  // `prompt` and must be asked again (persistence §1/§3).
  const states = new WeakMap<FakeFileHandle, { permission: PermissionState }>();
  let nextHandleId = 0;

  // Transient user activation, as Chrome enforces it for `showSaveFilePicker`.
  // A trusted user gesture marks it live for the gesture's own task (the
  // microtasks behind the picker's call run in it), and a permission prompt
  // consumes it (it needs the gesture to show its own UI). A native confirm
  // never grants it, so a save flow that stands a confirm — or pending async
  // work — between the gesture and the picker throws the same SecurityError
  // Chrome throws.
  let gestureFresh = false;
  for (const type of ['pointerdown', 'mousedown', 'click', 'keydown']) {
    window.addEventListener(
      type,
      () => {
        gestureFresh = true;
        setTimeout(() => {
          gestureFresh = false;
        }, 0);
      },
      true,
    );
  }

  function stateOf(handle: FakeFileHandle): { permission: PermissionState } {
    let state = states.get(handle);
    if (state === undefined) {
      state = { permission: 'prompt' };
      states.set(handle, state);
    }
    return state;
  }

  class FakeFileHandle {
    readonly name: string;
    readonly __fakeFsHandle = true;
    readonly __fakeFsId: number;

    constructor(name: string) {
      this.name = name;
      this.__fakeFsId = nextHandleId;
      nextHandleId += 1;
    }

    async getFile(): Promise<File> {
      calls.getFile += 1;
      if (plan.getFileThrows) {
        throw new DOMException('The fake disk cannot read this file.', 'NotFoundError');
      }
      if (stateOf(this).permission !== 'granted') {
        throw new DOMException(
          'Permission has not been granted to read this file.',
          'NotAllowedError',
        );
      }
      return new File([disk.bytes], disk.name, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        lastModified: disk.lastModified,
      });
    }

    async createWritable(): Promise<{
      write(data: BufferSource): Promise<void>;
      close(): Promise<void>;
      abort(): Promise<void>;
    }> {
      calls.createWritable += 1;
      calls.createWritableHandles.push(this.__fakeFsId);
      if (stateOf(this).permission !== 'granted') {
        throw new DOMException(
          'Permission has not been granted to write this file.',
          'NotAllowedError',
        );
      }
      const chunks: Uint8Array[] = [];
      return {
        write: async (data: BufferSource): Promise<void> => {
          calls.write += 1;
          chunks.push(
            data instanceof ArrayBuffer
              ? new Uint8Array(data)
              : new Uint8Array(data.buffer, data.byteOffset, data.byteLength),
          );
        },
        close: async (): Promise<void> => {
          calls.close += 1;
          let total = 0;
          for (const chunk of chunks) total += chunk.length;
          const merged = new Uint8Array(total);
          let offset = 0;
          for (const chunk of chunks) {
            merged.set(chunk, offset);
            offset += chunk.length;
          }
          disk.bytes = merged;
          // A write moves the file's stamp, like a real save.
          disk.lastModified = Date.now();
        },
        abort: async (): Promise<void> => {
          calls.abort += 1;
        },
      };
    }

    async queryPermission(descriptor: {
      readonly mode: 'read' | 'readwrite';
    }): Promise<PermissionState> {
      calls.queryPermission.push({ mode: descriptor.mode, handle: this.__fakeFsId });
      return plan.query ?? stateOf(this).permission;
    }

    async requestPermission(descriptor: {
      readonly mode: 'read' | 'readwrite';
    }): Promise<PermissionState> {
      calls.requestPermission.push({ mode: descriptor.mode, handle: this.__fakeFsId });
      // The prompt consumed the user's gesture, like the browser's own.
      gestureFresh = false;
      const result = plan.request ?? 'granted';
      stateOf(this).permission = result;
      return result;
    }
  }

  function makeHandle(name: string, permission: PermissionState): FakeFileHandle {
    const handle = new FakeFileHandle(name);
    states.set(handle, { permission });
    return handle;
  }

  // Structured clone keeps only data, so a handle read back from IndexedDB
  // is a plain record; give it its methods back, like the browser does when
  // it deserialises a real FileSystemFileHandle.
  function reviveValue(value: unknown): unknown {
    if (Array.isArray(value)) return value.map(reviveValue);
    if (typeof value !== 'object' || value === null) return value;
    const proto = Object.getPrototypeOf(value);
    if (proto !== Object.prototype && proto !== null) return value;
    const record = value as Record<string, unknown>;
    if (record.__fakeFsHandle === true) return new FakeFileHandle(String(record.name));
    const copy: Record<string, unknown> = {};
    for (const key of Object.keys(record)) copy[key] = reviveValue(record[key]);
    return copy;
  }

  // Wrap the two reads the app uses so revived records carry live handles.
  const transforms = new WeakMap<IDBRequest, (value: unknown) => unknown>();
  const revived = new WeakMap<IDBRequest, unknown>();

  const originalGet = IDBObjectStore.prototype.get;
  IDBObjectStore.prototype.get = function (
    this: IDBObjectStore,
    query: IDBValidKey | IDBKeyRange,
  ): IDBRequest {
    const request = originalGet.call(this, query);
    transforms.set(request, reviveValue);
    return request;
  };

  const originalGetAll = IDBObjectStore.prototype.getAll;
  IDBObjectStore.prototype.getAll = function (
    this: IDBObjectStore,
    query?: IDBValidKey | IDBKeyRange | null,
    count?: number,
  ): IDBRequest {
    const request = originalGetAll.call(this, query, count);
    transforms.set(request, reviveValue);
    return request;
  };

  const resultDescriptor = Object.getOwnPropertyDescriptor(IDBRequest.prototype, 'result');
  const originalResult = resultDescriptor?.get;
  if (originalResult !== undefined) {
    Object.defineProperty(IDBRequest.prototype, 'result', {
      configurable: true,
      get(this: IDBRequest): unknown {
        const value = originalResult.call(this);
        const transform = transforms.get(this);
        if (transform === undefined) return value;
        if (!revived.has(this)) revived.set(this, transform(value));
        return revived.get(this);
      },
    });
  }

  // The pickers. Opening grants access (the user just chose the file);
  // Save As hands out a fresh handle for the new file. Like Chrome, the save
  // picker refuses without live user activation.
  const fsaWindow = window as unknown as {
    showOpenFilePicker?: () => Promise<FakeFileHandle[]>;
    showSaveFilePicker?: (pickerOptions?: { suggestedName?: string }) => Promise<FakeFileHandle>;
  };
  fsaWindow.showOpenFilePicker = async (): Promise<FakeFileHandle[]> => {
    calls.openPicker += 1;
    return [makeHandle(disk.name, 'granted')];
  };
  fsaWindow.showSaveFilePicker = async (pickerOptions?: {
    suggestedName?: string;
  }): Promise<FakeFileHandle> => {
    if (!gestureFresh) {
      throw new DOMException(
        'Must be handling a user gesture to show a file picker.',
        'SecurityError',
      );
    }
    calls.savePicker += 1;
    const handle = makeHandle(pickerOptions?.suggestedName ?? disk.name, 'granted');
    calls.savePickerHandles.push(handle.__fakeFsId);
    return handle;
  };

  window.__fakeFs = {
    calls,
    plan,
    diskState: () => ({
      name: disk.name,
      size: disk.bytes.length,
      lastModified: disk.lastModified,
    }),
    externalEdit: (replacement?: string) => {
      if (replacement !== undefined) {
        const binary = atob(replacement);
        const bytes = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index += 1) {
          bytes[index] = binary.charCodeAt(index);
        }
        disk.bytes = bytes;
      }
      // An external save always moves the stamp, past anything Date.now()
      // could produce for our own writes.
      disk.lastModified = Math.max(Date.now(), disk.lastModified) + 1_000;
    },
  };
}
