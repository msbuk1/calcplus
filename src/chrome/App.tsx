import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { JSX } from 'react';
import appIcon from '../../assets/icons/icon-64.png';
import { DocumentSession } from '@commands/session';
import type { SheetId } from '@model/types';
import { formatCellRef, cellKey, createCell } from '@model/types';
import { parseUserValue, BLANK } from '@model/value';
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
import type { GridCellAddress, GridSelection } from '@grid/frame';
import { gridDebugState } from '@grid/debug';
import { cellRect } from '@grid/panes';
import { collapseSelection } from '@grid/selection';
import { FormulaBar } from '@chrome/FormulaBar';
import { NameBox } from '@chrome/NameBox';
import { nextFocusZone, type FocusZone } from '@chrome/focus';
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
  const [selection, setSelection] = useState<GridSelection>(collapseSelection({ row: 0, col: 0 }));
  const [focusZone, setFocusZone] = useState<FocusZone>('grid');
  const [status, setStatus] = useState('Open an .xlsx workbook to begin.');
  const [busy, setBusy] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const nameBoxRef = useRef<HTMLInputElement | null>(null);
  const formulaBarRef = useRef<HTMLInputElement | null>(null);
  const shellRef = useRef<HTMLDivElement | null>(null);
  const recoveryRef = useRef<ReturnType<typeof snapshotWorkbook> | null>(null);

  const active = selection.focus;

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
    setSelection(collapseSelection({ row: 0, col: 0 }));
    setFocusZone('grid');
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
      setVersion((current) => current + 1);
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

  const runCellSet = (address: GridCellAddress, rawText: string): void => {
    if (document === null) return;
    // A leading equals sign means the user authored a formula. The engine
    // arrives with the model-depth milestone; calcplus stores the authored
    // text and lets the saved file ask Excel to recalculate.
    const isFormula = rawText.startsWith('=');
    const formula = isFormula ? rawText.slice(1) : null;
    const value = isFormula ? BLANK : parseUserValue(rawText);
    const result = document.session.dispatch('cell.set', {
      sheetId: document.sheetId,
      row: address.row,
      col: address.col,
      value,
      formula,
    });
    setVersion((current) => current + 1);
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

  const focusZoneRef = useRef<FocusZone>(focusZone);
  focusZoneRef.current = focusZone;

  useEffect(() => {
    const listener = (event: KeyboardEvent): void => {
      if (event.key === 'F6') {
        event.preventDefault();
        setFocusZone((zone) => nextFocusZone(zone, event.shiftKey));
        return;
      }
      if (event.key === 'Escape' && focusZoneRef.current !== 'grid') {
        setFocusZone('grid');
        return;
      }
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

  // One focus owner: the zone decides which control holds the keyboard.
  useEffect(() => {
    if (focusZone === 'grid') {
      shellRef.current?.focus();
      return;
    }
    if (focusZone === 'name-box') {
      nameBoxRef.current?.focus();
      return;
    }
    formulaBarRef.current?.focus();
  }, [focusZone]);

  const navigateToAddress = useCallback(
    (text: string): string | null => {
      const trimmed = text.trim();
      if (trimmed === '') return null;
      const match = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(trimmed);
      if (match === null) {
        setStatus(`"${trimmed}" is not a cell address.`);
        return null;
      }
      const column = lettersToColumn(match[1] ?? '');
      const row = Number(match[2] ?? '0') - 1;
      if (projection === null) return null;
      if (row < 0 || row >= projection.rowCount || column >= projection.columnCount) {
        setStatus(`"${trimmed}" is outside the sheet.`);
        return null;
      }
      setSelection(collapseSelection({ row, col: column }));
      // The name box hands the focus back to the grid (command-system §7).
      setFocusZone('grid');
      setStatus(`Selected ${formatCellRef(row, column)}.`);
      return formatCellRef(row, column);
    },
    [projection],
  );

  const loadBenchmarkSheet = (options: {
    rows: number;
    cols: number;
    seed?: number;
    /** When set, only every nth row holds a cell: a sparse sheet. */
    touchedEvery?: number;
  }): number => {
    const opened = openBlankWorkbook();
    const sheet = opened.workbook.sheets[0];
    if (sheet === undefined) return 0;
    let state = (options.seed ?? 1) >>> 0;
    const random = (): number => {
      state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
      return state / 4_294_967_296;
    };
    for (let index = 1; index < 6; index += 1) {
      opened.workbook.styles.push({
        id: index,
        numFmtId: index === 3 ? 4 : 0,
        fontId: 0,
        fillId: 0,
        borderId: 0,
        xfId: 0,
        rawXml: '',
        font: {
          name: 'Calibri',
          size: index === 2 ? 14 : 11,
          bold: index === 1,
          italic: false,
          underline: false,
          strike: false,
          color: index === 4 ? '#b00020' : null,
        },
        fill:
          index === 5 ? { pattern: 'solid', color: '#fff2cc' } : { pattern: 'none', color: null },
        alignment: {
          horizontal: index === 3 ? 'right' : 'general',
          vertical: 'bottom',
          wrapText: false,
          indent: 0,
          textRotation: 0,
        },
        numFmtCode: index === 3 ? '#,##0.00' : null,
      });
    }
    let count = 0;
    const step = options.touchedEvery ?? 1;
    for (let row = 0; row < options.rows; row += step) {
      for (let col = 0; col < options.cols; col += 1) {
        const value = Math.round(random() * 100_000) / 100;
        sheet.cells.set(`${row}:${col}`, createCell({ type: 'number', value }, (row + col) % 6));
        count += 1;
      }
    }
    loadOpened(opened, 'benchmark.xlsx', null);
    setStatus(`Loaded a benchmark sheet with ${count} touched cells.`);
    return count;
  };

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
        selection,
        focusZone,
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
      loadBenchmarkSheet,
      getCellRect: (row: number, col: number) => {
        if (projection === null) return null;
        const layout = gridDebugState.layout;
        if (layout === null) return null;
        return cellRect(
          projection,
          layout,
          gridDebugState.scrollX,
          gridDebugState.scrollY,
          row,
          col,
        );
      },
      getScroll: () => ({
        x: gridDebugState.scrollX,
        y: gridDebugState.scrollY,
      }),
      getFreeze: () => ({
        rows: projection?.freezeRows ?? 0,
        cols: projection?.freezeCols ?? 0,
      }),
      getCellStyle: (row: number, col: number) => {
        const cell = projection?.cellAt(row, col);
        if (cell === undefined || cell === null) return null;
        return {
          font: cell.style.font,
          fill: cell.style.fill,
          alignment: cell.style.alignment,
          numberFormat: cell.style.numberFormat,
          displayText: cell.displayText,
          merge: cell.merge,
        };
      },
      getFillHandle: () => gridDebugState.fillHandle,
    };
  });

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <img className="brand-icon" src={appIcon} alt="" width={28} height={28} />
          <span className="brand-name">calcplus</span>
          <span className="brand-phase">M2 grid foundation</span>
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

      <div className="formula-row">
        <NameBox
          value={projection === null ? '' : formatCellRef(active.row, active.col)}
          onNavigate={navigateToAddress}
          onCancel={() => setFocusZone('grid')}
          onFocused={() => setFocusZone('name-box')}
          inputRef={nameBoxRef}
        />
        <FormulaBar
          text={projection?.rawTextAt(active.row, active.col) ?? ''}
          onCommit={(text, options) => {
            runCellSet(active, text);
            if (options.moveDown && projection !== null) {
              setSelection(
                collapseSelection({
                  row: Math.min(projection.rowCount - 1, active.row + 1),
                  col: active.col,
                }),
              );
            }
            setFocusZone('grid');
          }}
          onCancel={() => setFocusZone('grid')}
          onFocused={() => setFocusZone('formula-bar')}
          inputRef={formulaBarRef}
        />
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
            selection={selection}
            onSelectionChange={setSelection}
            onCommit={runCellSet}
            onClear={(address) => runCellSet(address, '')}
            shellRef={shellRef}
            onFocusZone={() => setFocusZone('grid')}
          />
        ) : (
          <div className="empty-state">
            <h1>Open a small workbook</h1>
            <p>
              The grid scrolls a full 1,048,576 by 16,384 sheet, freezes panes, and edits cells
              without touching the parts of the package calcplus does not model.
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

function lettersToColumn(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function dirtyKeysFromSnapshot(
  snapshot: ReadonlyMap<SheetId, ReadonlyMap<string, unknown>>,
): Map<SheetId, Set<string>> {
  return new Map(
    [...snapshot.entries()].map(([sheetId, cells]) => [sheetId, new Set(cells.keys())]),
  );
}
