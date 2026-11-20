export interface WritableFileLike {
  write(data: BufferSource): Promise<void>;
  close(): Promise<void>;
  abort?(): Promise<void>;
}

export interface FileHandleLike {
  readonly name: string;
  createWritable(): Promise<WritableFileLike>;
  getFile?(): Promise<File>;
  /** File System Access permission methods, absent on a handle stub. */
  queryPermission?(descriptor: { readonly mode: 'read' | 'readwrite' }): Promise<PermissionState>;
  requestPermission?(descriptor: { readonly mode: 'read' | 'readwrite' }): Promise<PermissionState>;
}

/** Persistence §4: the values compared to detect an external edit. */
export interface FileStamp {
  readonly lastModified: number;
  readonly size: number;
}

export interface PickedFile {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly handle: FileHandleLike | null;
  readonly stamp: FileStamp | null;
}

export interface SaveOutcome {
  readonly method: 'handle' | 'picker' | 'download';
  readonly handle: FileHandleLike | null;
  readonly fileName: string;
}

interface OpenFilePickerOptions {
  multiple?: boolean;
  excludeAcceptAllOption?: boolean;
  types?: {
    description: string;
    accept: Record<string, string[]>;
  }[];
}

interface SaveFilePickerOptions {
  suggestedName?: string;
  types?: {
    description: string;
    accept: Record<string, string[]>;
  }[];
}

interface FileSystemAccessWindow {
  showOpenFilePicker?: (options?: OpenFilePickerOptions) => Promise<FileHandleLike[]>;
  showSaveFilePicker?: (options?: SaveFilePickerOptions) => Promise<FileHandleLike>;
}

const WORKBOOK_TYPES = [
  {
    description: 'Excel workbook',
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
    },
  },
];

export function canUseFileSystemAccess(forceFallback: boolean): boolean {
  if (forceFallback) return false;
  const api = window as unknown as FileSystemAccessWindow;
  return typeof api.showOpenFilePicker === 'function';
}

export async function pickWorkbookFile(forceFallback: boolean): Promise<PickedFile | null> {
  const api = window as unknown as FileSystemAccessWindow;
  if (!forceFallback && typeof api.showOpenFilePicker === 'function') {
    try {
      const [handle] = await api.showOpenFilePicker({
        multiple: false,
        excludeAcceptAllOption: true,
        types: WORKBOOK_TYPES,
      });
      if (handle === undefined) return null;
      const file = await handle.getFile?.();
      if (file === undefined) return null;
      return {
        name: handle.name,
        bytes: new Uint8Array(await file.arrayBuffer()),
        handle,
        stamp: stampOf(file),
      };
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return null;
      throw error;
    }
  }
  return null;
}

export async function readBrowserFile(file: File): Promise<PickedFile> {
  return {
    name: file.name,
    bytes: new Uint8Array(await file.arrayBuffer()),
    handle: null,
    stamp: stampOf(file),
  };
}

/**
 * Persistence §4: compare the file's `lastModified` time and size with the
 * values captured at open or last save. A `null` stamp means the browser did
 * not expose one, so no comparison is possible.
 */
export function stampOf(file: {
  readonly lastModified?: number;
  readonly size?: number;
}): FileStamp | null {
  if (typeof file.lastModified !== 'number' || typeof file.size !== 'number') return null;
  return { lastModified: file.lastModified, size: file.size };
}

export function stampChanged(recorded: FileStamp | null, current: FileStamp | null): boolean {
  if (recorded === null || current === null) return false;
  return recorded.lastModified !== current.lastModified || recorded.size !== current.size;
}

/** Reads the current stamp without reading the whole file; `null` when unreadable. */
export async function readFileStamp(handle: FileHandleLike): Promise<FileStamp | null> {
  try {
    const file = await handle.getFile?.();
    return file === undefined ? null : stampOf(file);
  } catch {
    return null;
  }
}

/** Reads a file through its handle, for Reload and for a recent entry. */
export async function readHandleFile(handle: FileHandleLike): Promise<{
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly stamp: FileStamp | null;
} | null> {
  try {
    const file = await handle.getFile?.();
    if (file === undefined) return null;
    return {
      name: file.name.length === 0 ? handle.name : file.name,
      bytes: new Uint8Array(await file.arrayBuffer()),
      stamp: stampOf(file),
    };
  } catch {
    return null;
  }
}

/**
 * Persistence §1/§3: re-request read/write permission after a restart. A
 * handle that exposes no permission API (a stub in a test) is treated as
 * usable.
 */
export async function ensureHandlePermission(
  handle: FileHandleLike,
  mode: 'read' | 'readwrite' = 'readwrite',
): Promise<boolean> {
  if (typeof handle.queryPermission !== 'function') return true;
  try {
    if ((await handle.queryPermission({ mode })) === 'granted') return true;
    if (typeof handle.requestPermission !== 'function') return false;
    return (await handle.requestPermission({ mode })) === 'granted';
  } catch {
    return false;
  }
}

export async function saveWorkbookFile(
  bytes: Uint8Array,
  fileName: string,
  handle: FileHandleLike | null,
  options: { forceFallback: boolean; saveAs: boolean },
): Promise<SaveOutcome> {
  const api = window as unknown as FileSystemAccessWindow;
  if (!options.forceFallback && !options.saveAs && handle !== null) {
    const writable = await handle.createWritable();
    try {
      await writable.write(bufferSource(bytes));
      await writable.close();
      return { method: 'handle', handle, fileName: handle.name };
    } catch (error) {
      await writable.abort?.();
      throw error;
    }
  }

  if (!options.forceFallback && typeof api.showSaveFilePicker === 'function') {
    try {
      const picked = await api.showSaveFilePicker({
        suggestedName: options.saveAs ? fileName : fileName,
        types: WORKBOOK_TYPES,
      });
      const writable = await picked.createWritable();
      try {
        await writable.write(bufferSource(bytes));
        await writable.close();
        return { method: 'picker', handle: picked, fileName: picked.name };
      } catch (error) {
        await writable.abort?.();
        throw error;
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') {
        return { method: 'picker', handle: null, fileName };
      }
      throw error;
    }
  }

  downloadBytes(bytes, fileName);
  return { method: 'download', handle: null, fileName };
}

/**
 * The browser download path. `mimeType` defaults to the workbook type so the
 * existing save flow is unchanged; File → Export PDF passes
 * `application/pdf` (M7 lane 3).
 */
export function downloadBytes(
  bytes: Uint8Array,
  fileName: string,
  mimeType: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
): void {
  const buffer = bufferSource(bytes);
  const blob = new Blob([buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2_000);
}

function bufferSource(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}
