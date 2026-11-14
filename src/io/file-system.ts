export interface WritableFileLike {
  write(data: BufferSource): Promise<void>;
  close(): Promise<void>;
  abort?(): Promise<void>;
}

export interface FileHandleLike {
  readonly name: string;
  createWritable(): Promise<WritableFileLike>;
  getFile?(): Promise<File>;
}

export interface PickedFile {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly handle: FileHandleLike | null;
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
  };
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

export function downloadBytes(bytes: Uint8Array, fileName: string): void {
  const buffer = bufferSource(bytes);
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
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
