import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { JSX } from 'react';
import type { RibbonTabId } from '@commands/inventory';
import { DocumentSession } from '@commands/session';
import type { CommandResult, FindResultPayload, PainterSnapshot } from '@commands/types';
import { dateSystemLabel } from '@commands/workbook';
import { CalculationController } from '@engine/calculation-controller';
import { functionEntry, functionSkeleton } from '@engine/function-catalog';
import type { PublishedFunction } from '@engine/function-registry';
import type { EngineCellRef } from '@engine/port';
import type { ClipboardMode, ClipboardSnapshot } from '@model/clipboard';
import type { FindMatch } from '@model/find';
import type { GridRange, SheetId, Workbook } from '@model/types';
import { formatCellRef, cellKey, columnToLetters, createCell } from '@model/types';
import { resolveReferenceText } from '@model/references';
import type { StructuralEdit } from '@model/range-transform';
import { parseUserValue, BLANK } from '@model/value';
import { openBlankWorkbook } from '@ooxml/blank';
import type { CompatibilityInfo } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import type { WorkbookPackage } from '@ooxml/types';
import { WorkbookSaveError } from '@ooxml/errors';
import { applyRangeTransformPrototype } from '@ooxml/range-transform';
import { adoptSavedPackage, saveWorkbook } from '@ooxml/save';
import { snapshotWorkbook } from '@ooxml/snapshot';
import { confirmReplaceDocument } from '@io/confirm';
import { createSheetProjection } from '@view/projection';
import { firstVisibleSheetId, isSheetVisible, visibleSheetCount } from '@view/sheet-visibility';
import {
  DEFAULT_SHEET_VIEW,
  clampZoomPercent,
  initialSheetViews,
  sheetTabList,
  viewForSheet,
  type SheetViewState,
} from '@view/sheets';
import type { DimensionPreview, SelectionStatistics, SheetProjection } from '@view/types';
import { CanvasGrid } from '@grid/CanvasGrid';
import type { AutofitTarget } from '@grid/autofit';
import type { GridCellAddress, GridSelection } from '@grid/frame';
import type { DimensionAxis, DimensionResize } from '@grid/resize';
import { gridDebugState } from '@grid/debug';
import { cellRect } from '@grid/panes';
import { collapseSelection, selectionForRange, selectionRange } from '@grid/selection';
import { isPointingBoundary, pointingReference } from '@grid/editor';
import { resolveNavigation } from '@grid/navigation';
import { AboutDialog } from '@chrome/AboutDialog';
import { AppHeader } from '@chrome/AppHeader';
import { CompactBar } from '@chrome/CompactBar';
import { ContextMenu } from '@chrome/ContextMenu';
import type { DialogCloseReason } from '@chrome/Dialog';
import { FindDialog } from '@chrome/FindDialog';
import type { FindMatchView, FindOutcome, FindQuery, FindReplaceQuery } from '@chrome/FindDialog';
import { FormatCellsDialog } from '@chrome/FormatCellsDialog';
import type { FormatCellsOutcome, FormatCellsPayload } from '@chrome/FormatCellsDialog';
import { FormulaBar } from '@chrome/FormulaBar';
import type { FormulaBarKeyModifiers } from '@chrome/FormulaBar';
import { GoToDialog } from '@chrome/GoToDialog';
import type { GoToOutcome } from '@chrome/GoToDialog';
import { InsertFunctionDialog } from '@chrome/InsertFunctionDialog';
import { NameBox } from '@chrome/NameBox';
import { Ribbon } from '@chrome/Ribbon';
import { SheetTabs } from '@chrome/SheetTabs';
import { StatusBar } from '@chrome/StatusBar';
import { resolveCommandState } from '@chrome/command-state';
import type { ChromeCommandContext } from '@chrome/command-state';
import { nextFocusZone, type FocusZone } from '@chrome/focus';
import { contextMenuItems } from '@chrome/ribbon-model';
import { useMediaQuery } from '@chrome/useMediaQuery';
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
  readonly calculation: CalculationController;
}

type AppDialog = 'about' | 'find' | 'replace' | 'goto' | 'format' | 'insert-function';

const FORCE_FALLBACK =
  typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('fsa') === '0';

const FORMULA_BAR_POINT_KEYS: ReadonlySet<string> = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
]);

export function App(): JSX.Element {
  const [document, setDocument] = useState<LoadedDocument | null>(null);
  const [version, setVersion] = useState(0);
  const [selection, setSelection] = useState<GridSelection>(collapseSelection({ row: 0, col: 0 }));
  const [focusZone, setFocusZone] = useState<FocusZone>('grid');
  // The live text of an in-cell edit, so the formula bar mirrors typing before
  // the cell commits. `null` means no edit is open.
  const [liveEditText, setLiveEditText] = useState<string | null>(null);
  // The formula bar's own draft. `null` mirrors the active cell, so clicking
  // away returns the bar to the committed text.
  const [barDraft, setBarDraft] = useState<string | null>(null);
  const [barCaretRequest, setBarCaretRequest] = useState<{ readonly offset: number } | null>(null);
  const [status, setStatus] = useState('Open an .xlsx workbook to begin.');
  // Sheet view settings are split: the imported values live in the document;
  // the live zoom and display toggles are transient view state per sheet.
  const [sheetViews, setSheetViews] = useState<Record<string, SheetViewState>>({});
  // A header drag has not changed the document yet; the projection paints the
  // live size until the pointer release commits one dimension command.
  const [dimensionPreview, setDimensionPreview] = useState<DimensionPreview | null>(null);
  const [calcWarning, setCalcWarning] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [ribbonTab, setRibbonTab] = useState<RibbonTabId>('Home');
  const [ribbonExpanded, setRibbonExpanded] = useState(false);
  const [contextMenu, setContextMenu] = useState<{ readonly x: number; readonly y: number } | null>(
    null,
  );
  const [dialog, setDialog] = useState<AppDialog | null>(null);
  const [formatCellsTab, setFormatCellsTab] = useState<string | null>(null);
  // Functions inserted this session, most recent first. View state only: it
  // feeds the dialog's "Recently used" category and never enters the workbook.
  const [recentFunctions, setRecentFunctions] = useState<readonly PublishedFunction[]>([]);
  const [painter, setPainter] = useState<PainterSnapshot | null>(null);
  const narrow = useMediaQuery('(max-width: 899px)');
  const ribbonRef = useRef<HTMLElement | null>(null);
  const compactRef = useRef<HTMLElement | null>(null);
  const runCommandRef = useRef<((commandId: string, args?: unknown) => void) | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const nameBoxRef = useRef<HTMLInputElement | null>(null);
  const formulaBarRef = useRef<HTMLInputElement | null>(null);
  const barEditAddressRef = useRef<GridCellAddress | null>(null);
  const barPointingRef = useRef<{
    readonly anchor: GridCellAddress;
    readonly focus: GridCellAddress;
    readonly start: number;
    readonly end: number;
  } | null>(null);
  const shellRef = useRef<HTMLDivElement | null>(null);
  const recoveryRef = useRef<ReturnType<typeof snapshotWorkbook> | null>(null);
  const clipboardRef = useRef<ClipboardSnapshot | null>(null);
  const focusGridAfterDialogRef = useRef(false);

  const active = selection.focus;
  const activeView =
    document === null ? DEFAULT_SHEET_VIEW : (sheetViews[document.sheetId] ?? DEFAULT_SHEET_VIEW);

  const projection = useMemo(() => {
    if (document === null) return null;
    // Let the projection use the imported sheetFormatPr defaults, with its
    // grid metrics as the fallback for a file that declares none. The view
    // overrides are passed in; they are not document mutations.
    return createSheetProjection(document.session.workbook, document.sheetId, {
      version,
      zoom: activeView.zoom / 100,
      showGridLines: activeView.showGridLines,
      showHeadings: activeView.showHeadings,
      dimensionPreview,
    });
  }, [
    document,
    document?.sheetId,
    version,
    activeView.zoom,
    activeView.showGridLines,
    activeView.showHeadings,
    dimensionPreview,
  ]);

  const sheetTabs = useMemo(
    () => (document === null ? [] : sheetTabList(document.session.workbook)),
    [document, version],
  );

  const selectionStatistics = useMemo<SelectionStatistics | null>(() => {
    if (projection === null) return null;
    const statistics = projection.selectionStatistics(selectionRange(selection));
    return statistics.count === 0 ? null : statistics;
  }, [projection, selection]);

  const nameBoxValue = projection === null ? '' : formatSelectionAddress(selection, projection);

  const loadOpened = (
    opened: ReturnType<typeof openWorkbook>,
    fileName: string,
    handle: FileHandleLike | null,
  ): void => {
    document?.calculation.dispose();
    const session = new DocumentSession(opened.workbook);
    const initialSheetId = isSheetVisible(opened.workbook, opened.initialSheetId)
      ? opened.initialSheetId
      : (firstVisibleSheetId(opened.workbook) ?? opened.initialSheetId);
    const calculation = new CalculationController(session, {
      onModelChanged: () => setVersion((value) => value + 1),
      onCircular: (circular, sheetNames) => {
        setCalcWarning(circularWarningMessage(circular, sheetNames));
        setStatus(
          circular.length === 0
            ? 'Circular reference resolved.'
            : 'Circular reference detected. Affected cells show #CIRC!.',
        );
      },
      onError: (message) =>
        setStatus(`Formula engine unavailable: ${message}. Cached results are shown.`),
    });
    setDocument({
      session,
      pkg: opened.package,
      fileName,
      handle,
      sheetId: initialSheetId,
      compatibility: opened.compatibility,
      calculation,
    });
    setSheetViews(initialSheetViews(opened.workbook));
    setSelection(collapseSelection({ row: 0, col: 0 }));
    setFocusZone('grid');
    setLiveEditText(null);
    setBarDraft(null);
    setBarCaretRequest(null);
    setDimensionPreview(null);
    setContextMenu(null);
    setDialog(null);
    setPainter(null);
    setRibbonTab('Home');
    barEditAddressRef.current = null;
    barPointingRef.current = null;
    clipboardRef.current = null;
    setVersion((value) => value + 1);
    setStatus(`Opened ${fileName}.`);
    setCalcWarning(null);
    recoveryRef.current = null;
    calculation.start();
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
    const savedDimensions = document.session.dirtyDimensionSnapshot();
    const savedMerges = document.session.dirtyMergeSnapshot();
    const savedStructuralVersion = document.session.structuralVersion;
    try {
      const result = saveWorkbook({
        workbook: document.session.workbook,
        package: document.pkg,
        dirtyCells: dirtyKeysFromSnapshot(savedDirty),
        dirtyDimensions: savedDimensions,
        dirtyMerges: new Set(savedMerges.keys()),
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
      document.session.markSaved(savedDirty, savedDimensions, savedStructuralVersion, savedMerges);
      setDocument((current) => {
        // The file write is asynchronous, so a tab switch or a terminal sheet
        // delete may have changed the view since this save began. Adopt only
        // the saved package fields and keep a still-valid active sheet; when
        // the active sheet was removed, move to the first visible one.
        if (current === null || current.session !== document.session) return current;
        const workbook = current.session.workbook;
        const sheetId = workbook.sheets.some((sheet) => sheet.id === current.sheetId)
          ? current.sheetId
          : (firstVisibleSheetId(workbook) ?? workbook.sheets[0]?.id ?? current.sheetId);
        return {
          ...current,
          pkg: result.package,
          fileName: outcome.fileName,
          handle: outcome.handle ?? current.handle,
          sheetId,
        };
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
    // A leading equals sign means the user authored a formula. The model keeps
    // the exact text; the worker engine recalculates it asynchronously.
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

  const focusGrid = (): void => {
    setFocusZone('grid');
    shellRef.current?.focus();
  };

  const clearPainter = (): void => {
    setPainter(null);
  };

  /**
   * `painter.pickup` captures all supported formatting from the active cell.
   * The command returns the snapshot it will apply; the application keeps it
   * as view state, exactly as the clipboard snapshot is kept.
   */
  const startPainter = (): void => {
    if (document === null) return;
    setPainter(null);
    // The command reads the style at the source range's anchor, so the whole
    // selection is the pickup source; a collapsed selection is the common case.
    const source = selectionRange(selection);
    try {
      const result = document.session.dispatch(
        'painter.pickup',
        { sheetId: document.sheetId, range: source },
        {
          activeSheetId: document.sheetId,
          selection: source,
          activeCell: selection.focus,
        },
      );
      if (result.ui?.message !== undefined) {
        setStatus(result.ui.message);
        return;
      }
      const snapshot = readPainterSnapshot(result);
      if (snapshot === null) {
        setStatus('The Format Painter could not pick up formatting from that cell.');
        return;
      }
      setPainter(snapshot);
      focusGrid();
      setStatus(
        `Format Painter picked up ${formatCellRef(source.startRow, source.startCol)}. Select a cell or drag a range, or press Esc to cancel.`,
      );
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  };

  /**
   * A completed painter gesture is one `painter.apply` command, so the whole
   * target range is one undo transaction. The command owns style copying; the
   * chrome only supplies the source and target.
   */
  const applyPainter = (target: GridRange): void => {
    const snapshot = painter;
    if (document === null || snapshot === null) return;
    const args = { sheetId: document.sheetId, range: target, snapshot };
    setPainter(null);
    try {
      const result = document.session.dispatch('painter.apply', args, {
        activeSheetId: document.sheetId,
        selection: target,
        activeCell: { row: target.startRow, col: target.startCol },
      });
      setVersion((current) => current + 1);
      if (result.selectionAfter !== undefined) {
        setSelection(
          selectionForRange(result.selectionAfter, {
            row: result.selectionAfter.startRow,
            col: result.selectionAfter.startCol,
          }),
        );
      }
      if (result.ui?.message !== undefined) {
        setStatus(result.ui.message);
      } else if (result.changes.length === 0) {
        setStatus('The Format Painter made no change.');
      } else {
        setStatus(`Applied formatting to ${formatRangeAddress(target, false, false)}.`);
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
    focusGrid();
  };

  // A completed dialog command (Go To) returns the keyboard to the grid after
  // the dialog shell has unmounted and restored its invoker.
  useEffect(() => {
    if (dialog !== null || !focusGridAfterDialogRef.current) return;
    focusGridAfterDialogRef.current = false;
    focusGrid();
  }, [dialog]);

  /**
   * Esc is owned by the dialog shell and returns to the grid; a dismissal
   * returns to the invoking control through the shell. When that control was
   * a context-menu item that has unmounted, fall back to the grid so focus
   * never disappears into the body.
   */
  const closeDialog = (reason: DialogCloseReason): void => {
    setDialog(null);
    if (reason === 'escape') {
      focusGrid();
      return;
    }
    requestAnimationFrame(() => {
      if (window.document.activeElement === window.document.body) focusGrid();
    });
  };

  /** A file command must not steal an in-progress cell or formula-bar edit. */
  const focusGridUnlessEditing = (): void => {
    const active = window.document.activeElement;
    if (active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement) return;
    focusGrid();
  };

  /**
   * Apply a command's active-sheet hint. It is view state: undo of an insert
   * returns to the fallback tab, while a terminal sheet delete names the
   * neighbouring tab the chrome should activate. A dialog command keeps its
   * keyboard focus inside the dialog, so it can suppress the grid focus.
   */
  const applyActiveSheetHint = (
    result: CommandResult,
    options: { readonly focus?: boolean } = {},
  ): void => {
    if (document === null) return;
    const workbook = document.session.workbook;
    const sheetId = result.activeSheetId;
    if (sheetId === undefined || sheetId === null || sheetId === document.sheetId) return;
    const hintSheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
    const currentSheet = workbook.sheets.find((sheet) => sheet.id === document.sheetId);
    const targetSheet =
      hintSheet !== undefined && isSheetVisible(workbook, hintSheet.id)
        ? hintSheet
        : currentSheet !== undefined && isSheetVisible(workbook, currentSheet.id)
          ? currentSheet
          : undefined;
    // A hidden target is never activated; fall back to the current visible
    // sheet or the first visible tab. This stays view state only.
    const targetId = targetSheet?.id ?? firstVisibleSheetId(workbook);
    if (targetId === null || targetId === document.sheetId) return;
    setDocument({ ...document, sheetId: targetId });
    setPainter(null);
    setSelection(collapseSelection({ row: 0, col: 0 }));
    setLiveEditText(null);
    setBarDraft(null);
    setBarCaretRequest(null);
    barEditAddressRef.current = null;
    barPointingRef.current = null;
    if (options.focus !== false) focusGrid();
  };

  const undo = (): void => {
    if (document === null) return;
    const result = document.session.dispatch('edit.undo', undefined);
    applyActiveSheetHint(result);
    setVersion((current) => current + 1);
    setStatus('Undo.');
  };

  const redo = (): void => {
    if (document === null) return;
    const result = document.session.dispatch('edit.redo', undefined);
    applyActiveSheetHint(result);
    setVersion((current) => current + 1);
    setStatus('Redo.');
  };

  const copySelection = (mode: ClipboardMode): void => {
    if (document === null) return;
    const range = selectionRange(selection);
    const result = document.session.dispatch(mode === 'copy' ? 'edit.copy' : 'edit.cut', {
      sheetId: document.sheetId,
      range,
    });
    if (result.ui?.message !== undefined) {
      setStatus(result.ui.message);
      focusGrid();
      return;
    }
    if (result.clipboard !== undefined) clipboardRef.current = result.clipboard;
    const label =
      projection === null ? 'the selection' : formatSelectionAddress(selection, projection);
    setStatus(`${mode === 'copy' ? 'Copied' : 'Cut'} ${label}.`);
    focusGrid();
  };

  const pasteClipboard = (): void => {
    if (document === null) return;
    const snapshot = clipboardRef.current;
    if (snapshot === null) {
      setStatus('The clipboard is empty.');
      focusGrid();
      return;
    }
    const result = document.session.dispatch('edit.paste', {
      sheetId: document.sheetId,
      target: selectionRange(selection),
      snapshot,
    });
    setVersion((current) => current + 1);
    if (result.ui?.message !== undefined) {
      setStatus(result.ui.message);
      focusGrid();
      return;
    }
    if (result.selectionAfter !== undefined) {
      setSelection(
        selectionForRange(result.selectionAfter, {
          row: result.selectionAfter.startRow,
          col: result.selectionAfter.startCol,
        }),
      );
    }
    if (snapshot.mode === 'cut') clipboardRef.current = null;
    if (result.changes.length === 0) {
      setStatus('The paste made no change.');
    } else {
      const label =
        result.selectionAfter === undefined || projection === null
          ? 'the pasted range'
          : formatRangeAddress(result.selectionAfter, false, false);
      setStatus(`Pasted ${label}.`);
    }
    focusGrid();
  };

  const dirty = document?.session.hasUnsavedChanges ?? false;
  const calculation = document?.calculation.state() ?? null;
  const staleCount = calculation?.staleCount ?? 0;
  const activeCell = projection?.cellAt(active.row, active.col) ?? null;
  // The bar's draft owns the text while it is editing; otherwise it mirrors an
  // in-cell edit or the active cell.
  const barText = barDraft ?? liveEditText ?? projection?.rawTextAt(active.row, active.col) ?? '';
  const barFormulaActive = barDraft !== null && barDraft.startsWith('=');

  const commandContext = useMemo<ChromeCommandContext>(
    () => ({
      workbook: document?.session.workbook ?? null,
      activeSheetId: document?.sheetId ?? null,
      selection: selectionRange(selection),
      activeCell: selection.focus,
      view: {
        showGridLines: activeView.showGridLines,
        showHeadings: activeView.showHeadings,
        zoom: activeView.zoom,
        freezeRows: projection?.freezeRows ?? 0,
        freezeCols: projection?.freezeCols ?? 0,
      },
      busy,
      canUndo: document?.session.canUndo ?? false,
      canRedo: document?.session.canRedo ?? false,
    }),
    [
      document,
      selection,
      activeView.showGridLines,
      activeView.showHeadings,
      activeView.zoom,
      projection,
      busy,
      version,
    ],
  );

  const resolveState = useCallback(
    (commandId: string) => {
      const state = resolveCommandState(commandContext, commandId);
      if (state.enabled) return state;
      // Format Cells is a chrome-owned dialog command; the dialog itself
      // dispatches the formatting commands. An active sheet is its only
      // enablement condition.
      if (
        commandId === 'format.cellsDialog' &&
        commandContext.activeSheetId !== null &&
        !commandContext.busy
      ) {
        return { enabled: true, checked: state.checked };
      }
      return state;
    },
    [commandContext],
  );

  const controlValues = useMemo(
    () => ({
      'font-family': activeCell?.style.font.name ?? 'Calibri',
      'font-size': activeCell === null ? '11' : String(activeCell.style.font.size),
      'number-format': activeCell?.style.numberFormat ?? 'General',
    }),
    [activeCell],
  );

  useEffect(() => {
    if (!narrow) setRibbonExpanded(false);
  }, [narrow]);

  useEffect(() => {
    if (!dirty) return;
    const listener = (event: BeforeUnloadEvent): void => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', listener);
    return () => window.removeEventListener('beforeunload', listener);
  }, [dirty]);

  // A sheet inserted by a command starts with its imported/default view.
  useEffect(() => {
    if (document === null) return;
    setSheetViews((previous) => {
      let next: Record<string, SheetViewState> | null = null;
      for (const sheet of document.session.workbook.sheets) {
        if (previous[sheet.id] !== undefined) continue;
        next ??= { ...previous };
        next[sheet.id] = viewForSheet(sheet);
      }
      return next ?? previous;
    });
  }, [document, version]);

  const focusZoneRef = useRef<FocusZone>(focusZone);
  focusZoneRef.current = focusZone;

  useEffect(() => {
    const listener = (event: KeyboardEvent): void => {
      // The dialog owns Esc and Tab; do not let global shortcuts move a zone
      // behind the modal while it is open. Keep the browser's own Find/History
      // shortcuts from opening over an in-app dialog.
      if (dialog !== null) {
        if (
          (event.ctrlKey || event.metaKey) &&
          !event.shiftKey &&
          (event.key === 'f' ||
            event.key === 'F' ||
            event.key === 'h' ||
            event.key === 'H' ||
            event.key === 'g' ||
            event.key === 'G')
        ) {
          event.preventDefault();
        }
        return;
      }
      if (event.key === 'Escape' && painter !== null) {
        event.preventDefault();
        setPainter(null);
        setFocusZone('grid');
        setStatus('Format Painter cancelled.');
        return;
      }
      if (event.key === 'F6') {
        event.preventDefault();
        setFocusZone((zone) => nextFocusZone(zone, event.shiftKey));
        return;
      }
      if (event.key === 'Escape' && focusZoneRef.current !== 'grid') {
        setFocusZone('grid');
        return;
      }
      if (event.altKey && event.key === '=') {
        event.preventDefault();
        setFocusZone('grid');
        runCommandRef.current?.('formula.autosum');
        return;
      }
      if (event.shiftKey && event.key === 'F11') {
        event.preventDefault();
        runCommandRef.current?.('sheet.insert');
        return;
      }
      if (!event.ctrlKey && !event.metaKey) return;
      const target = event.target as HTMLElement | null;
      const inTextInput =
        target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;
      if (!event.shiftKey && (event.key === 'f' || event.key === 'F')) {
        if (document === null) return;
        event.preventDefault();
        runCommandRef.current?.('find.search');
      } else if (!event.shiftKey && (event.key === 'h' || event.key === 'H')) {
        if (document === null) return;
        event.preventDefault();
        runCommandRef.current?.('find.replaceOne');
      } else if (!event.shiftKey && (event.key === 'g' || event.key === 'G')) {
        if (document === null) return;
        event.preventDefault();
        runCommandRef.current?.('goto.goto');
      } else if (!inTextInput && !event.shiftKey && event.key === '1') {
        if (document === null) return;
        event.preventDefault();
        runCommandRef.current?.('format.cellsDialog');
      } else if (event.key === 's' || event.key === 'S') {
        event.preventDefault();
        runCommandRef.current?.(event.shiftKey ? 'file.saveAs' : 'file.save');
      } else if (event.key === 'o' || event.key === 'O') {
        event.preventDefault();
        runCommandRef.current?.('file.open');
      } else if (!inTextInput && (event.key === 'z' || event.key === 'Z')) {
        event.preventDefault();
        runCommandRef.current?.(event.shiftKey ? 'edit.redo' : 'edit.undo');
      } else if (!inTextInput && (event.key === 'y' || event.key === 'Y')) {
        event.preventDefault();
        runCommandRef.current?.('edit.redo');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'b' || event.key === 'B')) {
        event.preventDefault();
        runCommandRef.current?.('format.bold');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'i' || event.key === 'I')) {
        event.preventDefault();
        runCommandRef.current?.('format.italic');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'u' || event.key === 'U')) {
        event.preventDefault();
        runCommandRef.current?.('format.underline');
      } else if (!inTextInput && !event.shiftKey && event.key === '5') {
        event.preventDefault();
        runCommandRef.current?.('format.strike');
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
    if (focusZone === 'formula-bar') {
      formulaBarRef.current?.focus();
      return;
    }
    const target = narrow && !ribbonExpanded ? compactRef.current : ribbonRef.current;
    if (target !== null && !target.contains(window.document.activeElement)) target.focus();
  }, [focusZone, narrow, ribbonExpanded]);

  interface NavigationTarget {
    readonly sheetId: string;
    readonly sheetName: string;
    readonly range: GridRange;
    readonly wholeColumn: boolean;
    readonly wholeRow: boolean;
    readonly label: string;
  }

  type NavigationResolution =
    | { readonly kind: 'ok'; readonly target: NavigationTarget }
    | { readonly kind: 'invalid'; readonly message: string };

  /** Resolve A1 addresses, ranges and defined names for the name box and Go To. */
  const resolveNavigationTarget = useCallback(
    (text: string): NavigationResolution => {
      const trimmed = text.trim();
      if (trimmed === '' || document === null) {
        return { kind: 'invalid', message: 'Enter a cell reference or defined name.' };
      }
      const resolved = resolveReferenceText(document.session.workbook, trimmed, document.sheetId);
      if (resolved.kind === 'invalid') {
        return { kind: 'invalid', message: resolved.message };
      }
      const sheet = document.session.workbook.sheets.find(
        (candidate) => candidate.id === resolved.sheetId,
      );
      if (sheet === undefined) {
        return {
          kind: 'invalid',
          message: `"${trimmed}" points at a sheet that is no longer open.`,
        };
      }
      if (!isSheetVisible(document.session.workbook, sheet.id)) {
        return {
          kind: 'invalid',
          message: `Sheet "${sheet.name}" is hidden and cannot be activated in this milestone.`,
        };
      }
      const { range } = resolved;
      if (
        range.startRow < 0 ||
        range.startCol < 0 ||
        range.endRow >= 1_048_576 ||
        range.endCol >= 16_384
      ) {
        return { kind: 'invalid', message: `"${trimmed}" is outside the sheet.` };
      }
      return {
        kind: 'ok',
        target: {
          sheetId: sheet.id,
          sheetName: sheet.name,
          range,
          wholeColumn: resolved.wholeColumn,
          wholeRow: resolved.wholeRow,
          label: formatRangeAddress(range, resolved.wholeColumn, resolved.wholeRow),
        },
      };
    },
    [document],
  );

  const selectNavigationTarget = useCallback(
    (target: NavigationTarget): void => {
      if (document !== null && target.sheetId !== document.sheetId) {
        setDocument({ ...document, sheetId: target.sheetId });
        setPainter(null);
      }
      const single =
        target.range.startRow === target.range.endRow &&
        target.range.startCol === target.range.endCol;
      setSelection(
        single
          ? collapseSelection({ row: target.range.startRow, col: target.range.startCol })
          : selectionForRange(target.range, {
              row: target.range.startRow,
              col: target.range.startCol,
            }),
      );
    },
    [document],
  );

  const navigateToAddress = useCallback(
    (text: string): string | null => {
      const resolution = resolveNavigationTarget(text);
      if (resolution.kind === 'invalid') {
        setStatus(resolution.message);
        return null;
      }
      const { target } = resolution;
      selectNavigationTarget(target);
      // The name box hands the focus back to the grid (command-system §7).
      setFocusZone('grid');
      setStatus(`Selected ${target.sheetName}!${target.label}.`);
      return target.label;
    },
    [resolveNavigationTarget, selectNavigationTarget],
  );

  /** Dispatch a sheet command through the session; the sheet track owns the model. */
  const dispatchSheetCommand = (commandId: string, args: unknown): CommandResult | null => {
    if (document === null) return null;
    try {
      const result = document.session.dispatch(commandId, args, {
        activeSheetId: document.sheetId,
        selection: selectionRange(selection),
        activeCell: selection.focus,
      });
      applyActiveSheetHint(result);
      setVersion((value) => value + 1);
      if (result.selectionAfter !== undefined) {
        setSelection(
          selectionForRange(result.selectionAfter, {
            row: result.selectionAfter.startRow,
            col: result.selectionAfter.startCol,
          }),
        );
      }
      if (result.clipboard !== undefined) clipboardRef.current = result.clipboard;
      if (result.ui?.message !== undefined) setStatus(result.ui.message);
      return result;
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
      return null;
    }
  };

  /**
   * Submit one find/replace command through the session. The dialog owns the
   * form and the view list; the command owns matching and the cell changes.
   * Find Next advances the view selection from the returned match list without
   * touching the document.
   */
  const runFindCommand = (
    commandId: 'find.search' | 'find.replaceOne' | 'find.replaceAll',
    query: FindQuery | FindReplaceQuery,
    options: { readonly selectNext?: boolean } = {},
  ): FindOutcome => {
    if (document === null) return findErrorOutcome('Open a workbook first.');
    try {
      const result = document.session.dispatch(commandId, query, {
        activeSheetId: document.sheetId,
        selection: selectionRange(selection),
        activeCell: selection.focus,
      });
      applyActiveSheetHint(result, { focus: false });
      if (result.changes.length > 0) setVersion((value) => value + 1);
      if (result.selectionAfter !== undefined) {
        setSelection(
          selectionForRange(result.selectionAfter, {
            row: result.selectionAfter.startRow,
            col: result.selectionAfter.startCol,
          }),
        );
      }
      const outcome = findOutcomeFromResult(result, document.session.workbook);
      if (options.selectNext === true) {
        if (outcome.matches.length === 0) {
          setStatus(outcome.message ?? 'No matches found.');
        } else {
          const next = nextFindMatch(
            outcome.matches,
            document.sheetId,
            selection.focus,
            document.session.workbook,
            query.direction,
          );
          if (next !== null) selectFindMatch(next);
        }
      } else if (result.ui?.message !== undefined) {
        setStatus(result.ui.message);
      } else if (outcome.replaced > 0) {
        setStatus(`Replaced ${outcome.replaced} cell${outcome.replaced === 1 ? '' : 's'}.`);
      }
      return outcome;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(message);
      return findErrorOutcome(message);
    }
  };

  /** Move the view selection to one match, keeping the dialog's focus. */
  const selectFindMatch = (match: FindMatchView): void => {
    if (document === null) return;
    const sheet = document.session.workbook.sheets.find(
      (candidate) => candidate.id === match.sheetId,
    );
    if (sheet === undefined) {
      setStatus('That match is no longer in the workbook.');
      return;
    }
    if (!isSheetVisible(document.session.workbook, sheet.id)) {
      setStatus(`Sheet "${sheet.name}" is hidden and cannot be activated in this milestone.`);
      return;
    }
    if (sheet.id !== document.sheetId) {
      setDocument({ ...document, sheetId: sheet.id });
      setPainter(null);
      setLiveEditText(null);
      setBarDraft(null);
      setBarCaretRequest(null);
      barEditAddressRef.current = null;
      barPointingRef.current = null;
    }
    setSelection(collapseSelection({ row: match.row, col: match.col }));
    setStatus(`Selected ${sheet.name}!${match.address}.`);
  };

  /** Select a Find All row without closing the dialog or moving its focus. */
  const activateFindMatch = (match: FindMatchView): void => {
    selectFindMatch(match);
  };

  /**
   * Go To validates the target first, then submits `goto.goto`. A rejection
   * leaves the selection and the active sheet exactly where they were.
   */
  const submitGoTo = (text: string): GoToOutcome => {
    if (document === null) return { ok: false, message: 'Open a workbook first.' };
    const resolution = resolveNavigationTarget(text);
    if (resolution.kind === 'invalid') {
      setStatus(resolution.message);
      return { ok: false, message: resolution.message };
    }
    const target = resolution.target;
    try {
      const result = document.session.dispatch(
        'goto.goto',
        { target: text.trim() },
        {
          activeSheetId: document.sheetId,
          selection: selectionRange(selection),
          activeCell: selection.focus,
        },
      );
      // A toast with no selection is a rejected target, not a completed jump.
      if (result.ui?.type === 'toast' && result.selectionAfter === undefined) {
        const message = result.ui.message ?? 'That reference cannot be selected.';
        setStatus(message);
        return { ok: false, message };
      }
      if (result.changes.length > 0) setVersion((value) => value + 1);
      applyActiveSheetHint(result, { focus: false });
      if (result.selectionAfter !== undefined) {
        setSelection(
          selectionForRange(result.selectionAfter, {
            row: result.selectionAfter.startRow,
            col: result.selectionAfter.startCol,
          }),
        );
      } else if (
        result.activeSheetId === undefined ||
        result.activeSheetId === null ||
        result.activeSheetId === target.sheetId
      ) {
        // The command may only report the target; keep the local resolution as
        // the view selection, never as a model edit.
        selectNavigationTarget(target);
      }
      setStatus(result.ui?.message ?? `Selected ${target.sheetName}!${target.label}.`);
      // A successful jump is a command: hand the keyboard to the grid after
      // the dialog shell has unmounted. A direct dispatch focuses at once.
      if (dialog !== null) {
        focusGridAfterDialogRef.current = true;
        setDialog(null);
      } else {
        focusGrid();
      }
      return { ok: true, message: null };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(message);
      return { ok: false, message };
    }
  };

  /** Switch tabs without touching the document or marking it dirty. */
  const activateSheet = (sheetId: string): void => {
    if (document === null || sheetId === document.sheetId) return;
    const sheet = document.session.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) return;
    if (!isSheetVisible(document.session.workbook, sheet.id)) {
      setStatus(`Sheet "${sheet.name}" is hidden and cannot be activated in this milestone.`);
      return;
    }
    setDocument({ ...document, sheetId: sheet.id });
    setPainter(null);
    setSelection(collapseSelection({ row: 0, col: 0 }));
    setLiveEditText(null);
    setBarDraft(null);
    setBarCaretRequest(null);
    barEditAddressRef.current = null;
    barPointingRef.current = null;
    focusGrid();
    setStatus(`Switched to ${sheet.name}.`);
  };

  const changeZoom = (value: number): void => {
    if (document === null) return;
    const zoom = clampZoomPercent(value);
    const sheetId = document.sheetId;
    setSheetViews((previous) => {
      const current = previous[sheetId] ?? DEFAULT_SHEET_VIEW;
      if (current.zoom === zoom) return previous;
      return { ...previous, [sheetId]: { ...current, zoom } };
    });
  };

  const toggleGridLines = (): void => {
    if (document === null) return;
    const sheetId = document.sheetId;
    setSheetViews((previous) => {
      const current = previous[sheetId] ?? DEFAULT_SHEET_VIEW;
      return {
        ...previous,
        [sheetId]: { ...current, showGridLines: !current.showGridLines },
      };
    });
  };

  const toggleHeadings = (): void => {
    if (document === null) return;
    const sheetId = document.sheetId;
    setSheetViews((previous) => {
      const current = previous[sheetId] ?? DEFAULT_SHEET_VIEW;
      return {
        ...previous,
        [sheetId]: { ...current, showHeadings: !current.showHeadings },
      };
    });
  };

  const insertSheet = (): void => {
    if (document === null) return;
    focusGrid();
    const sheets = document.session.workbook.sheets;
    const activeIndex = sheets.findIndex((sheet) => sheet.id === document.sheetId);
    const result = dispatchSheetCommand('sheet.insert', {
      index: activeIndex < 0 ? sheets.length : activeIndex + 1,
    });
    if (result === null) return;
    if (result.changes.length === 0) {
      setStatus(result.ui?.message ?? 'No sheet was inserted.');
      return;
    }
    const inserted =
      result.activeSheetId === undefined || result.activeSheetId === null
        ? undefined
        : document.session.workbook.sheets.find((sheet) => sheet.id === result.activeSheetId);
    focusGrid();
    setStatus(`Inserted ${inserted?.name ?? 'a sheet'}.`);
  };

  const renameSheet = (sheetId: string, name: string): void => {
    if (document === null) return;
    const trimmed = name.trim();
    if (trimmed === '') {
      focusGrid();
      setStatus('A sheet name cannot be empty.');
      return;
    }
    const target = document.session.workbook.sheets.find((sheet) => sheet.id === sheetId);
    if (target === undefined || target.name === trimmed) {
      focusGrid();
      return;
    }
    const previousName = target.name;
    const result = dispatchSheetCommand('sheet.rename', { sheetId, name: trimmed });
    focusGrid();
    if (result !== null) {
      setStatus(result.ui?.message ?? `Renamed ${previousName} to ${trimmed}.`);
    }
  };

  const deleteSheet = (sheetId: string): void => {
    if (document === null) return;
    focusGrid();
    const sheets = document.session.workbook.sheets;
    if (visibleSheetCount(document.session.workbook) <= 1) {
      setStatus('A workbook must keep at least one visible sheet.');
      return;
    }
    const target = sheets.find((sheet) => sheet.id === sheetId);
    if (target === undefined) return;
    if (!isSheetVisible(document.session.workbook, target.id)) {
      setStatus(`Sheet "${target.name}" is hidden and cannot be deleted from the tab strip.`);
      return;
    }
    const targetName = target.name;

    // The command asks for the permanent-delete warning before it changes the
    // model. The delete is terminal, so the chrome offers no undo afterwards.
    const first = dispatchSheetCommand('sheet.delete', { sheetId });
    if (first === null) return;
    if (first.changes.length > 0) {
      setStatus(first.ui?.message ?? `Deleted ${targetName}.`);
      return;
    }
    if (first.ui?.type !== 'dialog') {
      setStatus(first.ui?.message ?? 'The sheet was not deleted.');
      return;
    }
    const confirmed =
      typeof window === 'undefined'
        ? false
        : window.confirm(
            first.ui.message ??
              `Deleting "${targetName}" cannot be undone, and its data may be lost. Delete the sheet?`,
          );
    if (!confirmed) {
      setStatus('Delete cancelled.');
      return;
    }
    const second = dispatchSheetCommand('sheet.delete', { sheetId, confirm: true });
    if (second === null) return;
    setStatus(
      second.changes.length > 0
        ? (second.ui?.message ?? `Deleted ${targetName}. This cannot be undone.`)
        : (second.ui?.message ?? 'The sheet was not deleted.'),
    );
  };

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
          underlineStyle: 'none',
          strike: false,
          color: index === 4 ? '#b00020' : null,
        },
        fill:
          index === 5
            ? { pattern: 'solid', color: '#fff2cc', backgroundColor: null }
            : { pattern: 'none', color: null, backgroundColor: null },
        border: {
          left: {
            style: index === 5 ? 'thin' : 'none',
            color: null,
            weight: index === 5 ? 'thin' : null,
          },
          right: {
            style: index === 5 ? 'thin' : 'none',
            color: null,
            weight: index === 5 ? 'thin' : null,
          },
          top: {
            style: index === 5 ? 'thin' : 'none',
            color: null,
            weight: index === 5 ? 'thin' : null,
          },
          bottom: {
            style: index === 5 ? 'thin' : 'none',
            color: null,
            weight: index === 5 ? 'thin' : null,
          },
          diagonal: { style: 'none', color: null, weight: null },
          diagonalUp: false,
          diagonalDown: false,
        },
        alignment: {
          horizontal: index === 3 ? 'right' : 'general',
          vertical: 'bottom',
          wrapText: false,
          indent: 0,
          textRotation: 0,
          shrinkToFit: false,
        },
        numFmtCode: index === 3 ? '#,##0.00' : null,
        numFmtKind: index === 3 ? 'builtin' : 'general',
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
    const dispatchCommand = (
      commandId: string,
      args: unknown,
      options?: Parameters<DocumentSession['dispatch']>[2],
    ): unknown => {
      if (document === null) return null;
      const result = document.session.dispatch(commandId, args, {
        activeSheetId: document.sheetId,
        selection: selectionRange(selection),
        activeCell: selection.focus,
        ...options,
      });
      applyActiveSheetHint(result);
      setVersion((value) => value + 1);
      if (result.ui?.message !== undefined) setStatus(result.ui.message);
      else if (commandId === 'workbook.dateSystem') {
        setStatus(`Date system: ${dateSystemLabel(document.session.workbook)}.`);
      } else setStatus(`${commandId} completed.`);
      return result;
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
        calculation: document?.calculation.state() ?? null,
        status,
        dateSystem: document === null ? null : dateSystemLabel(document.session.workbook),
      }),
      getSheets: () => (document === null ? [] : sheetTabList(document.session.workbook)),
      getViewState: () => ({
        activeSheetId: document?.sheetId ?? null,
        zoom: activeView.zoom,
        showGridLines: activeView.showGridLines,
        showHeadings: activeView.showHeadings,
      }),
      getSavedViewState: () => {
        if (document === null) return null;
        const sheet = document.session.workbook.sheets.find(
          (candidate) => candidate.id === document.sheetId,
        );
        return sheet === undefined
          ? null
          : {
              zoom: sheet.view.zoom,
              showGridLines: sheet.view.showGridLines,
              showHeadings: sheet.view.showHeadings,
            };
      },
      getSelectionStatistics: () => selectionStatistics,
      setZoom: changeZoom,
      toggleGridLines,
      toggleHeadings,
      activateSheet,
      insertSheet,
      renameSheet,
      deleteSheet,
      dispatchCommand,
      setDateSystem: (date1904: boolean) => dispatchCommand('workbook.dateSystem', { date1904 }),
      getDateSystem: () =>
        document === null
          ? null
          : {
              label: dateSystemLabel(document.session.workbook),
              date1904: document.session.workbook.date1904,
            },
      getNames: () =>
        document === null
          ? null
          : {
              workbook: document.session.workbook.names.map((name) => ({
                name: name.name,
                formula: name.formula,
              })),
              sheets: document.session.workbook.sheets.map((sheet) => ({
                sheet: sheet.name,
                names: sheet.names.map((name) => ({
                  name: name.name,
                  formula: name.formula,
                })),
              })),
            },
      runRangeTransformPrototype: (
        edit: Omit<StructuralEdit, 'sheetId'> & { sheetId?: string },
      ) => {
        if (document === null) return null;
        const structuralEdit: StructuralEdit = {
          ...edit,
          sheetId: edit.sheetId ?? document.sheetId,
        };
        const result = applyRangeTransformPrototype({
          workbook: document.session.workbook,
          package: document.pkg,
          edit: structuralEdit,
        });
        setStatus(
          `Range transform prototype: ${result.report.formulaRewrites} formula rewrite(s), ${result.report.mergesRewritten} merge(s), ${result.report.namesRewritten} name(s).`,
        );
        return { report: result.report, plan: result.plan };
      },
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
          border: cell.style.border,
          alignment: cell.style.alignment,
          numberFormat: cell.style.numberFormat,
          numberFormatColor: cell.style.numberFormatColor,
          displayText: cell.displayText,
          merge: cell.merge,
        };
      },
      getFillHandle: () => gridDebugState.fillHandle,
    };
  });

  const startFormulaBarEdit = (text: string): void => {
    barEditAddressRef.current = active;
    // The Insert Function dialog may have queued a new formula draft just
    // before focus; keep it instead of replacing it with the text the focus
    // event saw.
    setBarDraft((draft) => draft ?? text);
    setFocusZone('formula-bar');
  };

  const stopFormulaBarEdit = (): void => {
    // Insert Function moves the keyboard into the dialog. The bar keeps its
    // draft and edit address until the dialog inserts a skeleton or closes, so
    // opening and cancelling the dialog never loses typed formula text.
    if (dialog === 'insert-function') return;
    barEditAddressRef.current = null;
    barPointingRef.current = null;
    setBarDraft(null);
    setBarCaretRequest(null);
  };

  const changeFormulaBarText = (text: string): void => {
    // Typing ends the pointing gesture: the next reference starts at the caret
    // instead of replacing the reference the grid inserted.
    barPointingRef.current = null;
    setBarDraft(text);
  };

  /** Insert or replace the pointed range in the formula bar draft. */
  const pointFormulaBarSelection = (anchor: GridCellAddress, focus: GridCellAddress): void => {
    if (projection === null || barDraft === null || !barDraft.startsWith('=')) return;
    const element = formulaBarRef.current;
    const pointing = barPointingRef.current;
    const start = pointing?.start ?? element?.selectionStart ?? barDraft.length;
    const end = pointing?.end ?? element?.selectionEnd ?? start;
    const reference = pointingReference(projection, anchor, focus);
    const nextEnd = start + reference.length;
    barPointingRef.current = { anchor, focus, start, end: nextEnd };
    setBarDraft(`${barDraft.slice(0, start)}${reference}${barDraft.slice(end)}`);
    setBarCaretRequest({ offset: nextEnd });
  };

  /** Move or extend the pointed range from the bar's keyboard. */
  const pointFormulaBarFromKey = (key: string, modifiers: FormulaBarKeyModifiers): boolean => {
    if (projection === null || barDraft === null || !barDraft.startsWith('=')) return false;
    if (!FORMULA_BAR_POINT_KEYS.has(key)) return false;
    const pointing = barPointingRef.current;
    const caret = formulaBarRef.current?.selectionStart ?? barDraft.length;
    if (pointing === null && !isPointingBoundary({ text: barDraft }, caret)) return false;
    const selectionForKey =
      pointing === null
        ? collapseSelection(barEditAddressRef.current ?? active)
        : { anchor: pointing.anchor, focus: pointing.focus, extent: { kind: 'cells' as const } };
    const outcome = resolveNavigation({
      key,
      ctrlKey: modifiers.ctrlKey,
      shiftKey: modifiers.shiftKey,
      altKey: modifiers.altKey,
      selection: selectionForKey,
      projection,
      pageRows: 20,
      pageCols: 10,
    });
    if (outcome.kind === 'none') return false;
    setSelection(outcome.selection);
    pointFormulaBarSelection(outcome.selection.anchor, outcome.selection.focus);
    return true;
  };

  const cancelFormulaBar = (): void => {
    const origin = barEditAddressRef.current;
    barPointingRef.current = null;
    setBarDraft(null);
    setBarCaretRequest(null);
    if (origin !== null) setSelection(collapseSelection(origin));
    setFocusZone('grid');
  };

  const commitFormulaBar = (text: string, options: { readonly moveDown: boolean }): void => {
    // A pointing gesture moves the selection, but the draft belongs to the cell
    // where bar editing started.
    const address = barEditAddressRef.current ?? active;
    barPointingRef.current = null;
    setBarDraft(null);
    setBarCaretRequest(null);
    runCellSet(address, text);
    if (projection !== null) {
      const target = options.moveDown
        ? { row: Math.min(projection.rowCount - 1, address.row + 1), col: address.col }
        : address;
      setSelection(collapseSelection(target));
    }
    setFocusZone('grid');
  };

  /**
   * Insert Function places the chosen skeleton in the formula bar and leaves
   * the caret after the opening parenthesis (ticket 30). This is one editing
   * action: the draft belongs to the active cell but no `cell.set` runs until
   * the user commits the bar, so it is not an undo step of its own.
   */
  const insertFunctionByName = (name: PublishedFunction): void => {
    const entry = functionEntry(name);
    if (entry === undefined) return;
    const skeleton = functionSkeleton(entry.name);
    barEditAddressRef.current = active;
    barPointingRef.current = null;
    setBarDraft(skeleton);
    setBarCaretRequest({ offset: skeleton.length });
    setRecentFunctions((current) => [
      entry.name,
      ...current.filter((candidate) => candidate !== entry.name),
    ]);
    setDialog(null);
    // The shell restores its invoker as it unmounts; hand the keyboard to the
    // bar afterwards so the user types the first argument straight away.
    requestAnimationFrame(() => {
      const input = formulaBarRef.current;
      if (input === null) return;
      input.focus();
      input.setSelectionRange(skeleton.length, skeleton.length);
      setFocusZone('formula-bar');
    });
  };

  /**
   * Header double-click autofit. The grid measured each selected member's own
   * content; storing the results as one dimension batch keeps the imported
   * flags intact, records `customHeight`/`customWidth` and `bestFit`, and makes
   * the whole selection one undo step.
   */
  const autofitDimension = (axis: DimensionAxis, targets: readonly AutofitTarget[]): void => {
    if (document === null || targets.length === 0) return;
    const entries = targets.map(({ index, result }) => {
      const stored =
        axis === 'row'
          ? (projection?.rowDimension(index) ?? null)
          : (projection?.columnDimension(index) ?? null);
      return { index, size: result.modelValue, hidden: stored?.hidden };
    });
    const outcome = document.session.dispatch('dimension.set', {
      sheetId: document.sheetId,
      axis,
      entries,
      custom: true,
      bestFit: axis === 'column' ? true : undefined,
    });
    setVersion((current) => current + 1);
    const labels = targets.map((target) => target.index);
    setStatus(
      outcome.changes.length === 0
        ? `No ${axis === 'row' ? 'row height' : 'column width'} change was needed.`
        : `Autofitted ${dimensionTargetLabel(axis, labels)}.`,
    );
  };

  /** Paint the live header-drag size without touching the document. */
  const previewDimensionResize = (
    axis: DimensionAxis,
    indices: readonly number[],
    result: DimensionResize | null,
  ): void => {
    setDimensionPreview((current) => {
      if (result === null) return null;
      if (
        current !== null &&
        current.axis === axis &&
        sameIndices(current.indices, indices) &&
        current.size === result.modelValue
      ) {
        return current;
      }
      return { axis, indices: [...indices], size: result.modelValue };
    });
  };

  /**
   * A released header drag is one undoable dimension command for every
   * selected visible member. The stored hidden flag and every other imported
   * flag stay with each dimension record; a manual resize clears a previous
   * autofit's `bestFit`.
   */
  const commitDimensionResize = (
    axis: DimensionAxis,
    indices: readonly number[],
    result: DimensionResize,
  ): void => {
    setDimensionPreview(null);
    if (document === null || indices.length === 0) return;
    const entries = indices.map((index) => {
      const stored =
        axis === 'row'
          ? (projection?.rowDimension(index) ?? null)
          : (projection?.columnDimension(index) ?? null);
      return { index, size: result.modelValue, hidden: stored?.hidden };
    });
    const outcome = document.session.dispatch('dimension.set', {
      sheetId: document.sheetId,
      axis,
      entries,
      custom: true,
      bestFit: axis === 'column' ? false : undefined,
    });
    setVersion((current) => current + 1);
    setStatus(
      outcome.changes.length === 0
        ? `No ${axis === 'row' ? 'row height' : 'column width'} change was needed.`
        : `Resized ${dimensionTargetLabel(axis, indices)}.`,
    );
  };

  /**
   * Submit the Format Cells dialog's changed tabs as one atomic batch. The
   * dialog validates values first; the session applies every tab command and
   * records a single undo transaction, rolling the whole batch back if a
   * command rejects (for example the 100,000-cell formatting budget) or throws.
   * The rejection keeps the dialog open and is shown in the status bar and
   * inline.
   */
  const applyFormatCells = (payload: FormatCellsPayload): FormatCellsOutcome => {
    if (document === null) return { ok: false, message: 'Open a workbook first.' };
    const range = selectionRange(selection);
    const entries: { readonly id: string; readonly args: Record<string, unknown> }[] = [];
    if (payload.font !== undefined) entries.push({ id: 'format.font', args: { ...payload.font } });
    if (payload.fill !== undefined) entries.push({ id: 'format.fill', args: { ...payload.fill } });
    if (payload.border !== undefined) {
      entries.push({ id: 'format.border', args: { ...payload.border } });
    }
    if (payload.alignment !== undefined) {
      entries.push({ id: 'format.alignment', args: { ...payload.alignment } });
    }
    if (payload.numberFormat !== undefined) {
      entries.push({ id: 'format.numberFormat', args: { ...payload.numberFormat } });
    }
    if (entries.length === 0) return { ok: true, message: null };
    const batch = entries.map((entry) => ({
      id: entry.id,
      args: { sheetId: document.sheetId, range, ...entry.args },
    }));
    try {
      const result = document.session.dispatchBatch(
        batch,
        {
          activeSheetId: document.sheetId,
          selection: range,
          activeCell: selection.focus,
        },
        // A single-tab Apply keeps its own command id in the transaction log;
        // a multi-tab Apply is one `format.cellsDialog` transaction.
        batch.length === 1 ? batch[0]!.id : 'format.cellsDialog',
      );
      if (result.ui?.message !== undefined) {
        setStatus(result.ui.message);
        return { ok: false, message: result.ui.message };
      }
      if (result.changes.length > 0) setVersion((current) => current + 1);
      const label = formatRangeAddress(range, false, false);
      setStatus(
        result.changes.length === 0 ? `No formatting change for ${label}.` : `Formatted ${label}.`,
      );
      return { ok: true, message: null };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(message);
      return { ok: false, message };
    }
  };

  const runCommand = (commandId: string, args?: unknown): string | null => {
    switch (commandId) {
      case 'file.new':
        newWorkbook();
        focusGrid();
        return null;
      case 'file.open':
        // The picker owns focus until it closes; a successful open returns it to the grid.
        void openFile();
        return null;
      case 'file.save':
        void saveFile(false).finally(focusGridUnlessEditing);
        return null;
      case 'file.saveAs':
        void saveFile(true).finally(focusGridUnlessEditing);
        return null;
      case 'edit.undo':
        undo();
        focusGrid();
        return null;
      case 'edit.redo':
        redo();
        focusGrid();
        return null;
      case 'edit.copy':
        copySelection('copy');
        return null;
      case 'edit.cut':
        copySelection('cut');
        return null;
      case 'edit.paste':
        pasteClipboard();
        return null;
      case 'sheet.insert':
        insertSheet();
        return null;
      case 'view.gridlines':
        toggleGridLines();
        focusGrid();
        return null;
      case 'view.headings':
        toggleHeadings();
        focusGrid();
        return null;
      case 'view.zoom100':
        changeZoom(100);
        focusGrid();
        return null;
      case 'format.cellsDialog': {
        if (document === null || busy) return null;
        clearPainter();
        // A group launcher names its tab (ticket 28); without one the dialog
        // keeps the last tab, like Excel's Ctrl+1.
        const tab = (args as { readonly tab?: unknown } | undefined)?.tab;
        if (typeof tab === 'string') setFormatCellsTab(tab);
        setDialog('format');
        return null;
      }
      case 'painter.pickup':
        if (busy) return null;
        startPainter();
        return null;
      case 'find.search': {
        if (document === null) return null;
        const query = findQueryArgument(args);
        if (query === null) {
          clearPainter();
          setDialog('find');
          return null;
        }
        runFindCommand('find.search', query);
        return null;
      }
      case 'find.replaceOne': {
        if (document === null) return null;
        const query = findQueryArgument(args);
        if (query === null) {
          clearPainter();
          setDialog('replace');
          return null;
        }
        runFindCommand('find.replaceOne', query);
        return null;
      }
      case 'find.replaceAll': {
        if (document === null) return null;
        const query = findReplaceQueryArgument(args);
        if (query === null) {
          clearPainter();
          setDialog('replace');
          return null;
        }
        runFindCommand('find.replaceAll', query);
        return null;
      }
      case 'formula.insertFunction': {
        if (document === null || busy) return null;
        clearPainter();
        // The command requests the dialog; the chrome owns the surface.
        const result = dispatchSheetCommand('formula.insertFunction', undefined);
        if (result?.ui?.type === 'dialog') setDialog('insert-function');
        return null;
      }
      case 'goto.goto': {
        if (document === null) return null;
        const target = goToTargetArgument(args);
        if (target === null) {
          clearPainter();
          setDialog('goto');
          return null;
        }
        submitGoTo(target);
        return null;
      }
      case 'help.about':
        clearPainter();
        setDialog('about');
        return null;
      case 'nameBox.navigate': {
        const text = commandTextArgument(args);
        return text === null ? null : navigateToAddress(text);
      }
      default: {
        const result = dispatchSheetCommand(commandId, args);
        if (result?.ui?.type !== 'dialog') focusGrid();
        return null;
      }
    }
  };
  runCommandRef.current = runCommand;

  return (
    <div className="app">
      <AppHeader
        fileName={document?.fileName ?? null}
        dirty={dirty}
        canSave={resolveState('file.save').enabled}
        canUndo={resolveState('edit.undo').enabled}
        canRedo={resolveState('edit.redo').enabled}
        onCommand={runCommand}
      />

      {narrow ? (
        <>
          <CompactBar
            onCommand={runCommand}
            resolveState={resolveState}
            ribbonExpanded={ribbonExpanded}
            onToggleRibbon={() => setRibbonExpanded((expanded) => !expanded)}
            compactRef={compactRef}
            onFocus={() => setFocusZone('ribbon')}
          />
          {ribbonExpanded ? (
            <Ribbon
              activeTab={ribbonTab}
              onTabChange={setRibbonTab}
              onCommand={runCommand}
              resolveState={resolveState}
              controlValues={controlValues}
              ribbonRef={ribbonRef}
              onFocus={() => setFocusZone('ribbon')}
            />
          ) : null}
        </>
      ) : (
        <Ribbon
          activeTab={ribbonTab}
          onTabChange={setRibbonTab}
          onCommand={runCommand}
          resolveState={resolveState}
          controlValues={controlValues}
          ribbonRef={ribbonRef}
          onFocus={() => setFocusZone('ribbon')}
        />
      )}

      <div className="formula-row">
        <NameBox
          value={nameBoxValue}
          onNavigate={(text) => runCommand('nameBox.navigate', { text })}
          onCancel={() => setFocusZone('grid')}
          onFocused={() => setFocusZone('name-box')}
          inputRef={nameBoxRef}
        />
        <FormulaBar
          text={barText}
          onTextChange={changeFormulaBarText}
          onCommit={commitFormulaBar}
          onCancel={cancelFormulaBar}
          onFocused={startFormulaBarEdit}
          onStopEditing={stopFormulaBarEdit}
          onPointKey={pointFormulaBarFromKey}
          caretRequest={barCaretRequest}
          inputRef={formulaBarRef}
          onInsertFunction={() => runCommand('formula.insertFunction')}
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
        <div
          className="compat-banner"
          title={document.compatibility.notes.join(' ')}
          data-testid="compat-banner"
        >
          {document.compatibility.notes.join(' ')}
        </div>
      ) : null}

      {calcWarning !== null ? (
        <div className="calc-warning" role="status" title={calcWarning} data-testid="calc-warning">
          {calcWarning}
        </div>
      ) : null}

      <main
        className="grid-area"
        onContextMenu={(event) => {
          event.preventDefault();
          if (document !== null) setContextMenu({ x: event.clientX, y: event.clientY });
        }}
      >
        {document !== null && projection !== null ? (
          <CanvasGrid
            projection={projection}
            selection={selection}
            onSelectionChange={setSelection}
            onCommit={runCellSet}
            onClear={(range) => {
              const result = dispatchSheetCommand('edit.clearContents', { range });
              if (result === null) return;
              if (result.ui?.message === undefined) {
                const label = formatRangeAddress(range, false, false);
                setStatus(
                  result.changes.length === 0 ? `${label} was already empty.` : `Cleared ${label}.`,
                );
              }
              if (result.ui?.type !== 'dialog') focusGrid();
            }}
            onCopy={() => copySelection('copy')}
            onCut={() => copySelection('cut')}
            onPaste={pasteClipboard}
            onEditorChange={setLiveEditText}
            formulaBarPointing={barFormulaActive ? { onPoint: pointFormulaBarSelection } : null}
            onAutofit={autofitDimension}
            onResizePreview={previewDimensionResize}
            onResizeCommit={commitDimensionResize}
            onZoomChange={changeZoom}
            painterActive={painter !== null}
            onPainterApply={applyPainter}
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
            <button
              type="button"
              data-command="file.open"
              onClick={() => runCommand('file.open')}
              data-testid="empty-open-button"
            >
              Open workbook
            </button>
          </div>
        )}
      </main>

      {document !== null ? (
        <SheetTabs
          sheets={sheetTabs}
          activeSheetId={document.sheetId}
          onSelectSheet={activateSheet}
          onInsertSheet={() => runCommand('sheet.insert')}
          onRenameSheet={renameSheet}
          onDeleteSheet={deleteSheet}
        />
      ) : null}

      <StatusBar
        status={status}
        activeCell={document === null ? '' : formatCellRef(active.row, active.col)}
        activeValue={projection?.cellAt(active.row, active.col)?.displayText ?? ''}
        stale={activeCell?.stale === true}
        staleCount={staleCount}
        statistics={selectionStatistics}
        showGridLines={resolveState('view.gridlines').checked}
        showHeadings={resolveState('view.headings').checked}
        gridLinesEnabled={resolveState('view.gridlines').enabled}
        headingsEnabled={resolveState('view.headings').enabled}
        zoom={activeView.zoom}
        onToggleGridLines={() => runCommand('view.gridlines')}
        onToggleHeadings={() => runCommand('view.headings')}
        onZoomChange={changeZoom}
      />

      {contextMenu !== null ? (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          items={contextMenuItems(resolveState)}
          onSelect={(item) => {
            setContextMenu(null);
            if (item.commandId !== undefined) runCommand(item.commandId, item.args);
          }}
          onClose={() => {
            setContextMenu(null);
            focusGrid();
          }}
        />
      ) : null}

      {dialog === 'find' || dialog === 'replace' ? (
        <FindDialog
          title={dialog === 'replace' ? 'Find and Replace' : 'Find'}
          onClose={closeDialog}
          onSearch={(query, all) => runFindCommand('find.search', query, { selectNext: !all })}
          onReplace={(query, all) =>
            runFindCommand(all ? 'find.replaceAll' : 'find.replaceOne', query)
          }
          onActivateMatch={activateFindMatch}
        />
      ) : null}

      {dialog === 'goto' ? (
        <GoToDialog initialValue={nameBoxValue} onClose={closeDialog} onSubmit={submitGoTo} />
      ) : null}

      {dialog === 'format' ? (
        <FormatCellsDialog
          initial={activeCell?.style}
          targetLabel={formatRangeAddress(selectionRange(selection), false, false)}
          initialTab={formatCellsTab ?? undefined}
          onClose={closeDialog}
          onApply={applyFormatCells}
        />
      ) : null}

      {dialog === 'about' ? <AboutDialog onClose={closeDialog} /> : null}

      {dialog === 'insert-function' ? (
        <InsertFunctionDialog
          recentlyUsed={recentFunctions}
          onClose={closeDialog}
          onInsert={insertFunctionByName}
        />
      ) : null}
    </div>
  );
}

/** The picked-up style record is view state returned by `painter.pickup`. */
function readPainterSnapshot(result: CommandResult): PainterSnapshot | null {
  return result.painter ?? null;
}

function commandTextArgument(args: unknown): string | null {
  if (typeof args !== 'object' || args === null || !('text' in args)) return null;
  const text = (args as { readonly text?: unknown }).text;
  return typeof text === 'string' ? text : null;
}

function findQueryArgument(args: unknown): FindQuery | null {
  if (typeof args !== 'object' || args === null) return null;
  const value = args as Record<string, unknown>;
  if (typeof value.query !== 'string') return null;
  return {
    query: value.query,
    scope: value.scope === 'workbook' ? 'workbook' : 'sheet',
    matchCase: value.matchCase === true,
    wholeCell: value.wholeCell === true,
    inFormulas: value.inFormulas !== false,
    inValues: value.inValues !== false,
    direction: value.direction === 'up' ? 'up' : 'down',
  };
}

function findReplaceQueryArgument(args: unknown): FindReplaceQuery | null {
  const query = findQueryArgument(args);
  if (query === null) return null;
  const replacement = (args as { readonly replacement?: unknown }).replacement;
  return typeof replacement === 'string' ? { ...query, replacement } : null;
}

function goToTargetArgument(args: unknown): string | null {
  if (typeof args !== 'object' || args === null) return null;
  const target = (args as { readonly target?: unknown }).target;
  return typeof target === 'string' && target.trim() !== '' ? target : null;
}

function findErrorOutcome(message: string): FindOutcome {
  return { error: message, message: null, matches: [], total: 0, replaced: 0 };
}

/**
 * The find commands attach their match list as `result.find` and `result.matches`.
 * The view only reads that payload, the selection and the cell changes; it
 * never inspects workbook cells to search.
 */
function findOutcomeFromResult(result: CommandResult, workbook: Workbook): FindOutcome {
  const payload: FindResultPayload | undefined = result.find;
  const matches = findMatches(payload?.matches ?? result.matches, workbook);
  const total =
    payload?.total !== undefined && Number.isFinite(payload.total) ? payload.total : matches.length;
  const replaced = result.replaced ?? payload?.replaced ?? 0;
  return {
    error: null,
    message: result.ui?.message ?? null,
    matches,
    total,
    replaced,
  };
}

function findMatches(value: readonly FindMatch[] | undefined, workbook: Workbook): FindMatchView[] {
  const matches: FindMatchView[] = [];
  for (const match of value ?? []) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === match.sheetId);
    matches.push({
      key: `${match.sheetId}:${match.row}:${match.col}`,
      sheetId: match.sheetId,
      sheetName: sheet?.name ?? 'Sheet',
      row: match.row,
      col: match.col,
      address: formatCellRef(match.row, match.col),
      preview: match.preview,
    });
  }
  return matches;
}

/** The next match to select: the first after the active cell in scan order. */
function nextFindMatch(
  matches: readonly FindMatchView[],
  activeSheetId: string,
  active: { readonly row: number; readonly col: number },
  workbook: Workbook,
  direction: 'down' | 'up' = 'down',
): FindMatchView | null {
  if (matches.length === 0) return null;
  const sheetIndex = (sheetId: string): number =>
    workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
  const activeSheetIndex = sheetIndex(activeSheetId);
  // The command returns matches in scan order: document order downwards,
  // reverse document order upwards. The test itself works on document
  // position, so the direction only flips which side counts as "next".
  const next = matches.find((match) => {
    const index = sheetIndex(match.sheetId);
    const position =
      index !== activeSheetIndex
        ? Math.sign(index - activeSheetIndex)
        : match.row !== active.row
          ? Math.sign(match.row - active.row)
          : Math.sign(match.col - active.col);
    return direction === 'up' ? position < 0 : position > 0;
  });
  return next ?? matches[0] ?? null;
}

function formatSelectionAddress(selection: GridSelection, projection: SheetProjection): string {
  const range = selectionRange(selection);
  const { startRow, startCol, endRow, endCol } = range;
  if (startRow === endRow && startCol === endCol) {
    return formatCellRef(startRow, startCol);
  }
  // Excel labels a whole column as A:A and a whole row as 1:1.
  if (startRow === 0 && endRow === projection.rowCount - 1) {
    return `${columnToLetters(startCol)}:${columnToLetters(endCol)}`;
  }
  if (startCol === 0 && endCol === projection.columnCount - 1) {
    return `${startRow + 1}:${endRow + 1}`;
  }
  return `${formatCellRef(startRow, startCol)}:${formatCellRef(endRow, endCol)}`;
}

/** One dimension uses its sheet label; a multi-header selection counts members. */
function dimensionTargetLabel(axis: DimensionAxis, indices: readonly number[]): string {
  if (indices.length === 1) {
    const index = indices[0]!;
    return axis === 'row' ? `row ${index + 1}` : `column ${columnToLetters(index)}`;
  }
  return `${indices.length} ${axis === 'row' ? 'rows' : 'columns'}`;
}

function sameIndices(left: readonly number[], right: readonly number[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function formatRangeAddress(
  range: {
    readonly startRow: number;
    readonly startCol: number;
    readonly endRow: number;
    readonly endCol: number;
  },
  wholeColumn: boolean,
  wholeRow: boolean,
): string {
  if (wholeColumn) {
    return `${columnToLetters(range.startCol)}:${columnToLetters(range.endCol)}`;
  }
  if (wholeRow) {
    return `${range.startRow + 1}:${range.endRow + 1}`;
  }
  if (range.startRow === range.endRow && range.startCol === range.endCol) {
    return formatCellRef(range.startRow, range.startCol);
  }
  return `${formatCellRef(range.startRow, range.startCol)}:${formatCellRef(
    range.endRow,
    range.endCol,
  )}`;
}

function dirtyKeysFromSnapshot(
  snapshot: ReadonlyMap<SheetId, ReadonlyMap<string, unknown>>,
): Map<SheetId, Set<string>> {
  return new Map(
    [...snapshot.entries()].map(([sheetId, cells]) => [sheetId, new Set(cells.keys())]),
  );
}

function circularWarningMessage(
  circular: readonly EngineCellRef[],
  sheetNames: readonly string[],
): string | null {
  if (circular.length === 0) return null;
  const labels = circular
    .slice(0, 5)
    .map((ref) => `${sheetNames[ref.sheet] ?? 'Sheet'}!${formatCellRef(ref.row, ref.col)}`);
  const remainder =
    circular.length > labels.length ? ` and ${circular.length - labels.length} more` : '';
  return `Circular reference detected: ${labels.join(', ')}${remainder}. Affected cells show #CIRC!; the file keeps the authored text and cached result.`;
}
