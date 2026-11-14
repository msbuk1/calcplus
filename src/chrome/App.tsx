import { useEffect, useMemo, useRef, useState } from 'react';
import type { JSX } from 'react';
import { DocumentSession } from '@commands/session';
import type { CellAddress, SheetId } from '@model/types';
import { formatCellRef, cellKey } from '@model/types';
import { parseUserValue } from '@model/value';
import { openBlankWorkbook } from '@ooxml/blank';
import type { CompatibilityInfo } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import type { WorkbookPackage } from '@ooxml/types';
import { WorkbookSaveError } from '@ooxml/errors';
import { adoptSavedPackage, saveWorkbook } from '@ooxml/save';
import { snapshotWorkbook } from '@ooxml/snapshot';
import { confirmReplaceDocument } from '@io/confirm';
import { createSheetProjection } from '@view/projection';
import { CanvasGrid } from '@grid/CanvasGrid';
import { GRID_METRICS } from '@grid/frame';
import {
  canUseFileSystemAccess,
  pickWorkbookFile,
  readBrowserFile,
  saveWorkbookFile,
  type FileHandleLike,
} from '@io/file-system';

interface LoadedDocument {
  readonly session: DocumentSession;
  readonly pkg: WorkbookPackage;
  readonly fileName: string;
  readonly handle: FileHandleLike | null;
  readonly sheetId: string;
  readonly compatibility: CompatibilityInfo;
}

const FORCE_FALLBACK =
  typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('fsa') === '0';

export function App(): JSX.Element {
  const [document, setDocument] = useState<LoadedDocument | null>(null);
  const [version, setVersion] = useState(0);
  const [active, setActive] = useState<CellAddress>({ row: 0, col: 0 });
  const [status, setStatus] = useState('Open an .xlsx workbook to begin.');
  const [busy, setBusy] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const recoveryRef = useRef<ReturnType<typeof snapshotWorkbook> | null>(null);

  const projection = useMemo(() => {
    if (document === null) return null;
    return createSheetProjection(document.session.workbook, document.sheetId, {
      version,
      defaultRowHeight: GRID_METRICS.rowHeight,
      defaultColumnWidth: GRID_METRICS.columnWidth,
    });
  }, [document, document?.sheetId, version]);

  const loadOpened = (
    opened: ReturnType<typeof openWorkbook>,
    fileName: string,
    handle: FileHandleLike | null,
  ): void => {
    const session = new DocumentSession(opened.workbook);
    setDocument({
      session,
      pkg: opened.package,
      fileName,
      handle,
      sheetId: opened.initialSheetId,
      compatibility: opened.compatibility,
    });
    setActive({ row: 0, col: 0 });
    setVersion((value) => value + 1);
    setStatus(`Opened ${fileName}.`);
    recoveryRef.current = null;
  };

  const loadPickedFile = (picked: {
    name: string;
    bytes: Uint8Array;
    handle: FileHandleLike | null;
  }): void => {
    if (!confirmReplaceDocument(document?.session.hasUnsavedChanges ?? false)) {
      setStatus('Open cancelled. Unsaved changes were kept.');
      return;
    }
    try {
      loadOpened(openWorkbook(picked.bytes, picked.name), picked.name, picked.handle);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  };

  const newWorkbook = (): void => {
    if (busy) return;
    if (!confirmReplaceDocument(document?.session.hasUnsavedChanges ?? false)) {
      setStatus('New workbook cancelled. Unsaved changes were kept.');
      return;
    }
    try {
      loadOpened(openBlankWorkbook(), 'Untitled.xlsx', null);
      setStatus('Created a new blank workbook. Use Save As to choose a file.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  };

  const openFile = async (): Promise<void> => {
    if (busy) return;
    setBusy(true);
    try {
      if (canUseFileSystemAccess(FORCE_FALLBACK)) {
        const picked = await pickWorkbookFile(FORCE_FALLBACK);
        if (picked !== null) loadPickedFile(picked);
      } else {
        fileInputRef.current?.click();
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  const saveFile = async (saveAs: boolean): Promise<void> => {
    if (document === null || busy) return;
    setBusy(true);
    // Snapshot the dirty state before the write await. Edits made while the
    // browser is writing are not in the saved bytes and must stay dirty.
    const savedDirty = document.session.dirtyCellSnapshot();
    try {
      const result = saveWorkbook({
        workbook: document.session.workbook,
        package: document.pkg,
        dirtyCells: dirtyKeysFromSnapshot(savedDirty),
      });
      const outcome = await saveWorkbookFile(result.bytes, document.fileName, document.handle, {
        forceFallback: FORCE_FALLBACK,
        saveAs,
      });
      if (outcome.method === 'picker' && outcome.handle === null) {
        setStatus('Save As cancelled.');
        return;
      }
      adoptSavedPackage(document.session.workbook, result.package);
      document.session.markSaved(savedDirty);
      setDocument({
        ...document,
        pkg: result.package,
        fileName: outcome.fileName,
        handle: outcome.handle ?? document.handle,
      });
      setVersion((value) => value + 1);
      recoveryRef.current = null;
      setStatus(result.noOp ? 'Saved. No changes to write.' : 'Saved.');
    } catch (error) {
      if (error instanceof WorkbookSaveError) {
        recoveryRef.current = snapshotWorkbook(document.session.workbook, document.pkg);
        setStatus(
          `${error.message} An in-memory recovery snapshot was kept; the file was not touched.`,
        );
      } else {
        setStatus(error instanceof Error ? error.message : String(error));
      }
    } finally {
      setBusy(false);
    }
  };

  const runCellSet = (address: CellAddress, rawText: string): void => {
    if (document === null) return;
    const value = parseUserValue(rawText);
    const result = document.session.dispatch('cell.set', {
      sheetId: document.sheetId,
      row: address.row,
      col: address.col,
      value,
    });
    setVersion((current) => current + 1);
    setActive(address);
    if (result.ui?.message !== undefined) {
      setStatus(result.ui.message);
      return;
    }
    if (result.changes.length === 0) {
      setStatus(`${formatCellRef(address.row, address.col)} was unchanged.`);
      return;
    }
    setStatus(`Edited ${formatCellRef(address.row, address.col)}.`);
  };

  const undo = (): void => {
    if (document === null) return;
    document.session.dispatch('edit.undo', undefined);
    setVersion((current) => current + 1);
    setStatus('Undo.');
  };

  const redo = (): void => {
    if (document === null) return;
    document.session.dispatch('edit.redo', undefined);
    setVersion((current) => current + 1);
    setStatus('Redo.');
  };

  const dirty = document?.session.hasUnsavedChanges ?? false;

  useEffect(() => {
    if (!dirty) return;
    const listener = (event: BeforeUnloadEvent): void => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', listener);
    return () => window.removeEventListener('beforeunload', listener);
  }, [dirty]);

  useEffect(() => {
    const listener = (event: KeyboardEvent): void => {
      if (!event.ctrlKey && !event.metaKey) return;
      const target = event.target as HTMLElement | null;
      const inTextInput =
        target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;
      if (event.key === 's' || event.key === 'S') {
        event.preventDefault();
        void saveFile(event.shiftKey);
      } else if (event.key === 'o' || event.key === 'O') {
        event.preventDefault();
        void openFile();
      } else if (!inTextInput && (event.key === 'z' || event.key === 'Z')) {
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
      } else if (!inTextInput && (event.key === 'y' || event.key === 'Y')) {
        event.preventDefault();
        redo();
      }
    };
    window.addEventListener('keydown', listener);
    return () => window.removeEventListener('keydown', listener);
  });

  useEffect(() => {
    const debugWindow = window as unknown as {
      __calcplus?: Record<string, unknown>;
    };
    debugWindow.__calcplus = {
      getState: () => ({
        fileName: document?.fileName ?? null,
        sheetId: document?.sheetId ?? null,
        dirty: document?.session.hasUnsavedChanges ?? false,
        version: document?.session.version ?? 0,
        active,
        canUndo: document?.session.canUndo ?? false,
        canRedo: document?.session.canRedo ?? false,
        status,
      }),
      getCellDisplay: (row: number, col: number) =>
        projection?.cellAt(row, col)?.displayText ?? null,
      getCell: (row: number, col: number) => {
        if (document === null) return null;
        const cell = document.session.workbook.sheets
          .find((sheet) => sheet.id === document.sheetId)
          ?.cells.get(cellKey(row, col));
        return cell ?? null;
      },
      hasRecoverySnapshot: () => recoveryRef.current !== null,
    };
  });

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <span className="brand-name">calcplus</span>
          <span className="brand-phase">M1 walking skeleton</span>
        </div>
        <div className="file-meta">
          <span data-testid="file-name">{document?.fileName ?? 'No workbook open'}</span>
          {dirty ? (
            <span className="dirty-dot" title="Unsaved changes" data-testid="dirty-indicator">
              ●
            </span>
          ) : null}
        </div>
      </header>

      <div className="toolbar" role="toolbar" aria-label="Workbook commands">
        <button type="button" onClick={newWorkbook} disabled={busy} data-testid="new-button">
          New
        </button>
        <button
          type="button"
          onClick={() => void openFile()}
          disabled={busy}
          data-testid="open-button"
        >
          Open
        </button>
        <button
          type="button"
          onClick={() => void saveFile(false)}
          disabled={document === null || busy}
          data-testid="save-button"
        >
          Save
        </button>
        <button
          type="button"
          onClick={() => void saveFile(true)}
          disabled={document === null || busy}
          data-testid="save-as-button"
        >
          Save As
        </button>
        <span className="toolbar-divider" />
        <button
          type="button"
          onClick={undo}
          disabled={document === null || !document.session.canUndo}
          data-testid="undo-button"
        >
          Undo
        </button>
        <button
          type="button"
          onClick={redo}
          disabled={document === null || !document.session.canRedo}
          data-testid="redo-button"
        >
          Redo
        </button>
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12"
        hidden
        data-testid="file-input"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file === undefined) return;
          void readBrowserFile(file).then(loadPickedFile);
        }}
      />

      {document?.compatibility.notes.length ? (
        <div className="compat-banner" data-testid="compat-banner">
          {document.compatibility.notes.join(' ')}
        </div>
      ) : null}

      <main className="grid-area">
        {document !== null && projection !== null ? (
          <CanvasGrid
            projection={projection}
            active={active}
            onSelect={setActive}
            onCommit={runCellSet}
            onClear={(address) => runCellSet(address, '')}
          />
        ) : (
          <div className="empty-state">
            <h1>Open a small workbook</h1>
            <p>
              The M1 walking skeleton opens a real <code>.xlsx</code>, renders its values, edits a
              cell, undoes it, and saves without rewriting parts it does not understand.
            </p>
            <button type="button" onClick={() => void openFile()} data-testid="empty-open-button">
              Open workbook
            </button>
          </div>
        )}
      </main>

      <footer className="status-bar">
        <span data-testid="active-cell">
          {document === null ? '' : formatCellRef(active.row, active.col)}
        </span>
        <span data-testid="active-value">
          {projection?.cellAt(active.row, active.col)?.displayText ?? ''}
        </span>
        <span className="status-spacer" />
        <span data-testid="status-message">{status}</span>
      </footer>
    </div>
  );
}

function dirtyKeysFromSnapshot(
  snapshot: ReadonlyMap<SheetId, ReadonlyMap<string, unknown>>,
): Map<SheetId, Set<string>> {
  return new Map(
    [...snapshot.entries()].map(([sheetId, cells]) => [sheetId, new Set(cells.keys())]),
  );
}
