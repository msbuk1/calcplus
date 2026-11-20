import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { JSX } from 'react';
import type { ContextMenuRegion, RibbonTabId } from '@commands/inventory';
import { exportPdfPlan } from '@commands/chrome';
import { DocumentSession } from '@commands/session';
import type {
  CommandResult,
  FilterPayload,
  FindResultPayload,
  PainterSnapshot,
  SortPayload,
} from '@commands/types';
import { dateSystemLabel } from '@commands/workbook';
import { CalculationController } from '@engine/calculation-controller';
import { functionEntry, functionSkeleton } from '@engine/function-catalog';
import type { PublishedFunction } from '@engine/function-registry';
import type { EngineCellRef } from '@engine/port';
import { collectValueChecklist, computeHiddenLayers } from '@model/autofilter';
import type { ClipboardMode, ClipboardSnapshot } from '@model/clipboard';
import type { FindMatch } from '@model/find';
import { MERGE_DATA_WARNING } from '@model/merge';
import type { SortDirection } from '@model/sort';
import type {
  AutoFilter,
  Cell,
  ColumnFilter,
  GridRange,
  HyperlinkRecord,
  Sheet,
  SheetId,
  Workbook,
} from '@model/types';
import { formatCellRef, cellKey, columnToLetters, createCell } from '@model/types';
import { resolveReferenceText } from '@model/references';
import type { StructuralEdit } from '@model/range-transform';
import { parseUserValue, BLANK, formatCellValue } from '@model/value';
import type { CellValue } from '@model/value';
import { openBlankWorkbook } from '@ooxml/blank';
import type { CompatibilityInfo } from '@ooxml/open';
import { openWorkbook } from '@ooxml/open';
import type { WorkbookPackage } from '@ooxml/types';
import { WorkbookSaveError } from '@ooxml/errors';
import { applyRangeTransformPrototype } from '@ooxml/range-transform';
import { adoptSavedPackage, saveWorkbook } from '@ooxml/save';
import { snapshotWorkbook } from '@ooxml/snapshot';
import { listUnsupportedConditionalFormats } from '@ooxml/conditional-format-xml';
import type { PreservedPartReport } from '@ooxml/part-policy';
import { listPreservedParts } from '@ooxml/part-policy';
import { buildCompatibilityReport, buildExportReport } from '@ooxml/report';
import type { CompatibilityReport, ExportReport } from '@ooxml/report';
import { compatibilityReportSummary, exportReportSummary } from '@ooxml/report';
import {
  EXTERNAL_EDIT_UNSUPPORTED_WARNING,
  chooseExternalChange,
  confirmCloseDocument,
  confirmRecoveryRestore,
  confirmReplaceDocument,
  planMacroSave,
  planSignatureSave,
  shouldWarnBeforeUnload,
} from '@io/confirm';
import { createSheetProjection } from '@view/projection';
import { checkExternalLink, internalLocationText, parseInternalLocation } from '@view/hyperlinks';
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
import {
  activeCell,
  collapseSelection,
  moveRange,
  selectionForRange,
  selectionRange,
} from '@grid/selection';
import { isPointingBoundary, pointingReference } from '@grid/editor';
import { resolveNavigation } from '@grid/navigation';
import { AboutDialog } from '@chrome/AboutDialog';
import { AppHeader } from '@chrome/AppHeader';
import { CompactBar } from '@chrome/CompactBar';
import { ContextMenu } from '@chrome/ContextMenu';
import type { DialogCloseReason } from '@chrome/Dialog';
import { FindDialog } from '@chrome/FindDialog';
import type { FindMatchView, FindOutcome, FindQuery, FindReplaceQuery } from '@chrome/FindDialog';
import { FilterDropdown } from '@chrome/FilterDropdown';
import type { FilterApplyOutcome } from '@chrome/FilterDropdown';
import { FormatCellsDialog } from '@chrome/FormatCellsDialog';
import type { FormatCellsOutcome, FormatCellsPayload } from '@chrome/FormatCellsDialog';
import { FormulaBar } from '@chrome/FormulaBar';
import type { FormulaBarKeyModifiers } from '@chrome/FormulaBar';
import { GoToDialog } from '@chrome/GoToDialog';
import type { GoToOutcome } from '@chrome/GoToDialog';
import { InsertFunctionDialog } from '@chrome/InsertFunctionDialog';
import { NameBox } from '@chrome/NameBox';
import { PrintPreviewDialog } from '@chrome/PrintPreviewDialog';
import { RecentDialog } from '@chrome/RecentDialog';
import { ReportDialog } from '@chrome/ReportDialog';
import { compatibilitySections, exportSections } from '@chrome/report-view';
import { Ribbon } from '@chrome/Ribbon';
import { SheetTabs } from '@chrome/SheetTabs';
import { SortDialog } from '@chrome/SortDialog';
import type { SortDialogOutcome, SortSubmit } from '@chrome/SortDialog';
import { StatusBar } from '@chrome/StatusBar';
import {
  detectFilterColumnType,
  filterRecordCount,
  isDateFormatCode,
} from '@chrome/filter-dropdown-state';
import type { SortDialogColumn } from '@chrome/sort-dialog-state';
import { resolveCommandState } from '@chrome/command-state';
import type { ChromeCommandContext } from '@chrome/command-state';
import { nextFocusZone, type FocusZone } from '@chrome/focus';
import { contextMenuItems } from '@chrome/ribbon-model';
import { ribbonControlValues } from '@chrome/ribbon-readouts';
import { useMediaQuery } from '@chrome/useMediaQuery';
import {
  canUseFileSystemAccess,
  ensureHandlePermission,
  pickWorkbookFile,
  readBrowserFile,
  readFileStamp,
  readHandleFile,
  saveWorkbookFile,
  stampChanged,
  type FileHandleLike,
  type FileStamp,
} from '@io/file-system';
import {
  RECOVERY_DEBOUNCE_MS,
  RecoveryQuotaError,
  RecoveryRepository,
  recoveryIsNewerThanSavedFile,
  type PendingRecovery,
} from '@io/recovery';
import { RecentRepository, recentDocumentId, type RecentEntry } from '@io/recent';
import { createOriginStorage, type OriginStorage } from '@io/storage';
import { createPrintPlan } from '../print/planner';
import { createPrintSnapshot } from '../print/adapter';
import type { PrintPlan } from '../print/types';

interface LoadedDocument {
  readonly session: DocumentSession;
  readonly pkg: WorkbookPackage;
  readonly fileName: string;
  readonly handle: FileHandleLike | null;
  readonly sheetId: string;
  readonly compatibility: CompatibilityInfo;
  /** The open-time unsupported conditional-format listing (read-only in v1). */
  readonly unsupportedConditionalFormats: ReturnType<typeof listUnsupportedConditionalFormats>;
  /** Lane 4's preserved-part listing for the compatibility report. */
  readonly preservedParts: readonly PreservedPartReport[];
  readonly calculation: CalculationController;
  /** Persistence §4: the file's stamp at open or last save, for conflict checks. */
  readonly fileStamp: FileStamp | null;
  /**
   * A restored recovery copy holds work the file does not have yet. The
   * session starts clean (there are no commands yet), so this flag keeps the
   * close prompt and the Save state honest until the next save.
   */
  readonly dirtyFromRecovery: boolean;
}

type AppDialog =
  | 'about'
  | 'find'
  | 'replace'
  | 'goto'
  | 'format'
  | 'insert-function'
  | 'sort'
  | 'compatibility-report'
  | 'export-report'
  | 'print-preview'
  | 'recent';

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
  const [contextMenu, setContextMenu] = useState<{
    readonly x: number;
    readonly y: number;
    readonly region: ContextMenuRegion;
  } | null>(null);
  const [dialog, setDialog] = useState<AppDialog | null>(null);
  const [formatCellsTab, setFormatCellsTab] = useState<string | null>(null);
  // The Sort dialog's prefill: the range, the header guess and the sheet the
  // command resolved when the menu asked for keys (sort-autofilter §2).
  const [sortRequest, setSortRequest] = useState<SortPayload | null>(null);
  // The AutoFilter dropdown's open column and viewport anchor. View state only:
  // opening it and typing in its search box never create a transaction (§3.8).
  const [filterDropdown, setFilterDropdown] = useState<{
    readonly column: number;
    readonly anchor: { readonly top: number; readonly left: number; readonly headerTop: number };
  } | null>(null);
  // Functions inserted this session, most recent first. View state only: it
  // feeds the dialog's "Recently used" category and never enters the workbook.
  const [recentFunctions, setRecentFunctions] = useState<readonly PublishedFunction[]>([]);
  // The last save's export report, kept for File → Export Report. It is a
  // read-only view of facts the save already reported; it never enters the model.
  const [lastExportReport, setLastExportReport] = useState<ExportReport | null>(null);
  const [painter, setPainter] = useState<PainterSnapshot | null>(null);
  // The plan the Print preview window renders. File → Print computes it from
  // the live display records and drops it on close; it is plain data and never
  // mutates the model.
  const [printPlan, setPrintPlan] = useState<PrintPlan | null>(null);
  // File → Recent files (persistence §3). The list is read from IndexedDB when
  // the dialog opens; it is display state, never part of the workbook.
  const [recentEntries, setRecentEntries] = useState<readonly RecentEntry[]>([]);
  // Copy mode's marquee, the paste destination while a Ctrl+C or Ctrl+X
  // snapshot is live. View state only, like the painter's target (ticket 31 B).
  const [copyMarquee, setCopyMarquee] = useState<GridRange | null>(null);
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
  const headerRef = useRef<HTMLElement | null>(null);
  const recoveryRef = useRef<ReturnType<typeof snapshotWorkbook> | null>(null);
  const clipboardRef = useRef<ClipboardSnapshot | null>(null);
  const focusGridAfterDialogRef = useRef(false);
  // The persistence tables are opened lazily on the page's origin. The refs
  // keep one repository per tab; the storage itself is shared by the origin.
  const originStorageRef = useRef<OriginStorage | null>(null);
  const recoveryRepositoryRef = useRef<RecoveryRepository | null>(null);
  const recentRepositoryRef = useRef<RecentRepository | null>(null);
  const recoveryTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const recoveryInFlightRef = useRef(false);
  const recoveryQueuedRef = useRef(false);
  const recoveryWriteRef = useRef<() => Promise<void>>(async () => undefined);
  const recoveryOfferCheckedRef = useRef(false);
  const recoveryHintRef = useRef<{
    readonly docId: string;
    readonly pkg: WorkbookPackage;
    readonly structuralVersion: number;
  } | null>(null);
  const persistRequestedRef = useRef(false);
  const externalEditWarningRef = useRef<Set<string>>(new Set());
  /** Bumped when a document is replaced, so a late recovery write cannot revive it. */
  const loadGenerationRef = useRef(0);

  const originStorage = (): OriginStorage => {
    originStorageRef.current ??= createOriginStorage();
    return originStorageRef.current;
  };
  const recoveryStore = (): RecoveryRepository => {
    recoveryRepositoryRef.current ??= new RecoveryRepository(originStorage());
    return recoveryRepositoryRef.current;
  };
  const recentStore = (): RecentRepository => {
    recentRepositoryRef.current ??= new RecentRepository(originStorage());
    return recentRepositoryRef.current;
  };

  const active = activeCell(selection);
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

  /**
   * The on-open compatibility report (file-compat §7). It consumes the open
   * path's `CompatibilityInfo`, the banked unsupported-rule listing and lane
   * 4's preserved-part listing; `version` refreshes the stale-cell items.
   */
  const compatibilityReport = useMemo<CompatibilityReport | null>(() => {
    if (document === null) return null;
    return buildCompatibilityReport({
      workbook: document.session.workbook,
      compatibility: document.compatibility,
      unsupportedConditionalFormats: document.unsupportedConditionalFormats,
      preserved: document.preservedParts,
    });
  }, [document, version]);

  const selectionStatistics = useMemo<SelectionStatistics | null>(() => {
    if (projection === null) return null;
    const statistics = projection.selectionStatistics(selectionRange(selection));
    return statistics.count === 0 ? null : statistics;
  }, [projection, selection]);

  const nameBoxValue = projection === null ? '' : formatSelectionAddress(selection, projection);

  const clearRecoveryTimer = (): void => {
    if (recoveryTimerRef.current === null) return;
    clearTimeout(recoveryTimerRef.current);
    recoveryTimerRef.current = null;
  };

  /**
   * The sheets a recovery write may need to re-serialise. The dirty sets
   * accumulate until the next save, so this is a safe superset of the sheets
   * edited since the previous write.
   */
  const changedSheetIds = (session: DocumentSession): ReadonlySet<string> => {
    const ids = new Set<string>();
    for (const sheetId of session.dirtyCellMap().keys()) ids.add(sheetId);
    for (const sheetId of session.dirtyMergeSheets()) ids.add(sheetId);
    for (const sheetId of session.dirtyDimensionSnapshot().keys()) ids.add(sheetId);
    return ids;
  };

  const writeRecoveryNow = async (): Promise<void> => {
    const current = document;
    if (current === null || !current.session.hasUnsavedChanges) return;
    if (recoveryInFlightRef.current) {
      recoveryQueuedRef.current = true;
      return;
    }
    recoveryInFlightRef.current = true;
    const generation = loadGenerationRef.current;
    try {
      if (!persistRequestedRef.current) {
        persistRequestedRef.current = true;
        const storage = typeof navigator === 'undefined' ? undefined : navigator.storage;
        void storage?.persist?.().catch(() => undefined);
      }
      const docId = current.session.workbook.id;
      const hint = recoveryHintRef.current;
      const sameStructure =
        hint !== null &&
        hint.docId === docId &&
        hint.pkg === current.pkg &&
        hint.structuralVersion === current.session.structuralVersion;
      await recoveryStore().save({
        docId,
        displayName: current.fileName,
        snapshot: snapshotWorkbook(current.session.workbook, current.pkg),
        compatibility: current.compatibility,
        savedFileModifiedAt: current.fileStamp?.lastModified ?? null,
        savedFileSize: current.fileStamp?.size ?? null,
        activeSheetId: current.sheetId,
        handle: current.handle ?? undefined,
        changedSheets: sameStructure ? changedSheetIds(current.session) : 'all',
        partsChanged: !sameStructure,
      });
      if (loadGenerationRef.current !== generation) {
        // The document was replaced while this write was in flight; the copy
        // belongs to a session the user already discarded.
        await recoveryStore().discard(docId);
        return;
      }
      recoveryHintRef.current = {
        docId,
        pkg: current.pkg,
        structuralVersion: current.session.structuralVersion,
      };
      clearRecoveryTimer();
    } catch (error) {
      setStatus(
        error instanceof RecoveryQuotaError
          ? `${error.message} The last good copy was kept.`
          : `The recovery copy could not be written: ${
              error instanceof Error ? error.message : String(error)
            }`,
      );
    } finally {
      recoveryInFlightRef.current = false;
      if (recoveryQueuedRef.current) {
        recoveryQueuedRef.current = false;
        void recoveryWriteRef.current();
      }
    }
  };
  recoveryWriteRef.current = writeRecoveryNow;

  const rememberRecent = (picked: {
    readonly name: string;
    readonly handle: FileHandleLike | null;
    readonly stamp: FileStamp | null;
  }): void => {
    const origin = originStorage().origin;
    void recentStore()
      .remember({
        id: recentDocumentId(origin, picked.name),
        name: picked.name,
        fileSize: picked.stamp?.size ?? null,
        lastModified: picked.stamp?.lastModified ?? null,
        handle: picked.handle,
      })
      .then(setRecentEntries)
      .catch(() => undefined);
  };

  const refreshRecent = async (): Promise<void> => {
    try {
      setRecentEntries(await recentStore().list());
    } catch {
      setRecentEntries([]);
    }
  };

  const offerRecovery = async (manual: boolean): Promise<void> => {
    try {
      const repository = recoveryStore();
      await repository.prune();
      const pending = await repository.listPending();
      let candidate: PendingRecovery | null = null;
      for (const entry of pending) {
        if (manual || recoveryIsNewerThanSavedFile(entry)) {
          candidate = entry;
          break;
        }
        // The saved file is newer than the copy, so there is nothing to
        // restore; the copy is only litter from an earlier crash.
        await repository.discard(entry.docId);
      }
      if (candidate === null) {
        if (manual) setStatus('There is no unsaved work to recover.');
        return;
      }
      if (!confirmRecoveryRestore(candidate.displayName)) {
        await repository.discard(candidate.docId);
        setStatus('Recovery copy discarded.');
        return;
      }
      const restored = await repository.restore(candidate.docId);
      if (restored === null) {
        setStatus('The recovery copy could not be read. It was left in local storage.');
        return;
      }
      const compatibility = restored.compatibility ?? {
        opaquePartPaths: [...restored.workbook.opaqueParts.keys()].sort(),
        externalLinkParts: [...restored.package.externalLinkParts],
        hasMacros: [...restored.package.entries.keys()].some((path) =>
          /vbaproject\.bin$/i.test(path),
        ),
        hasDigitalSignatures: [...restored.package.entries.keys()].some(
          (path) => path.startsWith('_xmlsignatures/') || path.endsWith('.sigs'),
        ),
        notes: [],
      };
      loadOpened(
        {
          workbook: restored.workbook,
          package: restored.package,
          initialSheetId:
            restored.activeSheetId ??
            firstVisibleSheetId(restored.workbook) ??
            restored.workbook.sheets[0]?.id ??
            '',
          compatibility,
          conditionalFormatRules: new Map(),
        },
        restored.fileName,
        asFileHandle(restored.handle),
        restored.savedFileModifiedAt === null && restored.savedFileSize === null
          ? null
          : {
              lastModified: restored.savedFileModifiedAt ?? 0,
              size: restored.savedFileSize ?? 0,
            },
        true,
      );
      setStatus('Restored unsaved changes.');
    } catch (error) {
      setStatus(
        `The recovery copy could not be restored: ${
          error instanceof Error ? error.message : String(error)
        }. It was left in local storage.`,
      );
    }
  };

  const openRecent = async (entry: RecentEntry): Promise<void> => {
    if (busy) return;
    if (entry.handle === null) {
      // Firefox keeps names only (persistence §3); ask for the file again.
      setDialog(null);
      setStatus(`Choose "${entry.name}" again in the file picker.`);
      fileInputRef.current?.click();
      return;
    }
    setBusy(true);
    try {
      if (!(await ensureHandlePermission(entry.handle))) {
        await recentStore().remove(entry.id);
        setRecentEntries(await recentStore().list());
        setStatus(
          `Permission for "${entry.name}" was denied. It was removed from the recent list.`,
        );
        return;
      }
      const file = await readHandleFile(entry.handle);
      if (file === null) {
        await recentStore().remove(entry.id);
        setRecentEntries(await recentStore().list());
        setStatus(`"${entry.name}" is no longer available. It was removed from the recent list.`);
        return;
      }
      if (
        loadPickedFile({
          name: entry.name,
          bytes: file.bytes,
          handle: entry.handle,
          stamp: file.stamp,
        })
      ) {
        setDialog(null);
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  const loadOpened = (
    opened: ReturnType<typeof openWorkbook>,
    fileName: string,
    handle: FileHandleLike | null,
    fileStamp: FileStamp | null = null,
    dirtyFromRecovery = false,
  ): void => {
    document?.calculation.dispose();
    loadGenerationRef.current += 1;
    // A replaced document's recovery copy belongs to the session being
    // discarded; its replacement is either the file just opened or the new
    // unsaved document, both of which start clean.
    if (document !== null) {
      clearRecoveryTimer();
      void recoveryStore()
        .discard(document.session.workbook.id)
        .catch(() => undefined);
    }
    recoveryHintRef.current = null;
    const session = new DocumentSession(opened.workbook);
    // A restored copy carries unsaved commands over the package bytes it kept;
    // mark them so a later Save writes the recovered model (persistence §2).
    if (dirtyFromRecovery) session.markRestored();
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
      unsupportedConditionalFormats: listUnsupportedConditionalFormats(opened.workbook),
      preservedParts: listPreservedParts({
        package: opened.package,
        workbook: opened.workbook,
        format: fileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx',
      }),
      calculation,
      fileStamp,
      dirtyFromRecovery,
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
    setCopyMarquee(null);
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

  const loadPickedFile = (
    picked: {
      name: string;
      bytes: Uint8Array;
      handle: FileHandleLike | null;
      stamp: FileStamp | null;
    },
    replaceConfirmed = false,
  ): boolean => {
    if (
      !replaceConfirmed &&
      !confirmReplaceDocument(
        document !== null && (document.session.hasUnsavedChanges || document.dirtyFromRecovery),
      )
    ) {
      setStatus('Open cancelled. Unsaved changes were kept.');
      return false;
    }
    try {
      loadOpened(openWorkbook(picked.bytes, picked.name), picked.name, picked.handle, picked.stamp);
      rememberRecent(picked);
      return true;
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
      return false;
    }
  };

  const newWorkbook = (): void => {
    if (busy) return;
    if (
      !confirmReplaceDocument(
        document !== null && (document.session.hasUnsavedChanges || document.dirtyFromRecovery),
      )
    ) {
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

  const closeWorkbook = (): void => {
    if (document === null || busy) return;
    // Persistence §4: an unsaved workbook asks before the session is discarded.
    if (!confirmCloseDocument(document.session.hasUnsavedChanges || document.dirtyFromRecovery)) {
      setStatus('Close cancelled. Unsaved changes were kept.');
      return;
    }
    document.calculation.dispose();
    loadGenerationRef.current += 1;
    clearRecoveryTimer();
    // Persistence §2: closing discards the session, so its recovery copy goes
    // with it. The close prompt already protected the unsaved work.
    void recoveryStore()
      .discard(document.session.workbook.id)
      .catch(() => undefined);
    recoveryHintRef.current = null;
    setDocument(null);
    setSheetViews({});
    setSelection(collapseSelection({ row: 0, col: 0 }));
    setFocusZone('grid');
    setLiveEditText(null);
    setBarDraft(null);
    setBarCaretRequest(null);
    setDimensionPreview(null);
    setContextMenu(null);
    setDialog(null);
    setSortRequest(null);
    setFilterDropdown(null);
    setPainter(null);
    setCopyMarquee(null);
    setRibbonTab('Home');
    setLastExportReport(null);
    setVersion((value) => value + 1);
    setStatus('Closed the workbook.');
    setCalcWarning(null);
    recoveryRef.current = null;
    barEditAddressRef.current = null;
    barPointingRef.current = null;
    clipboardRef.current = null;
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
    let effectiveSaveAs = saveAs;
    try {
      // Persistence §4: a plain Save compares the file's stamp with the value
      // captured at open or last save. A newer external file is never
      // overwritten without an explicit Overwrite choice.
      if (!effectiveSaveAs && document.handle !== null) {
        const currentStamp = await readFileStamp(document.handle);
        if (currentStamp === null) {
          if (
            !window.confirm(
              `calcplus could not read "${document.fileName}" to check for external changes. Save as a new file instead?`,
            )
          ) {
            setStatus('Save cancelled. The file was not touched.');
            return;
          }
          effectiveSaveAs = true;
        } else if (stampChanged(document.fileStamp, currentStamp)) {
          const choice = chooseExternalChange(document.fileName);
          if (choice === 'cancel') {
            setStatus('Save cancelled. The newer file on disk was not touched.');
            return;
          }
          if (choice === 'reload') {
            const file = await readHandleFile(document.handle);
            if (file === null) {
              setStatus('The newer file could not be read. Local changes were kept.');
              return;
            }
            if (
              loadPickedFile(
                {
                  name: document.fileName,
                  bytes: file.bytes,
                  handle: document.handle,
                  stamp: file.stamp,
                },
                true,
              )
            ) {
              setStatus('Reloaded the file from disk. Local changes were discarded.');
            }
            return;
          }
          if (choice === 'saveAs') effectiveSaveAs = true;
        }
      }
      // File-compat §6.3: the default Save keeps .xlsm; Save As from a macro
      // workbook converts to .xlsx and only after explicit confirmation.
      const macroPlan = planMacroSave({
        fileName: document.fileName,
        hasMacros: document.compatibility.hasMacros,
        saveAs: effectiveSaveAs,
      });
      if (macroPlan === null) {
        setStatus('Save As cancelled. The workbook was not written.');
        return;
      }
      // File-compat §6.5: a signed package is preserved, but saving invalidates
      // the signature, so the write is preceded by a confirmation. A save with
      // nothing to write returns the original bytes and never asks, and an
      // unsigned workbook has nothing to warn about. Declining writes nothing.
      if (
        !planSignatureSave({
          hasDigitalSignatures: document.compatibility.hasDigitalSignatures,
          changed: document.session.hasUnsavedChanges || document.dirtyFromRecovery,
        })
      ) {
        setStatus('Save cancelled. The workbook was not written; its signature is untouched.');
        return;
      }
      // Snapshot the dirty state before the write await. Edits made while the
      // browser is writing are not in the saved bytes and must stay dirty.
      const savedDirty = document.session.dirtyCellSnapshot();
      const savedDimensions = document.session.dirtyDimensionSnapshot();
      const savedMerges = document.session.dirtyMergeSnapshot();
      const savedStructuralVersion = document.session.structuralVersion;
      const result = saveWorkbook({
        workbook: document.session.workbook,
        package: document.pkg,
        dirtyCells: dirtyKeysFromSnapshot(savedDirty),
        dirtyDimensions: savedDimensions,
        dirtyMerges: new Set(savedMerges.keys()),
        options: { format: macroPlan.format },
      });
      const outcome = await saveWorkbookFile(result.bytes, macroPlan.fileName, document.handle, {
        forceFallback: FORCE_FALLBACK,
        saveAs: effectiveSaveAs,
      });
      if (outcome.method === 'picker' && outcome.handle === null) {
        setStatus('Save As cancelled.');
        return;
      }
      // The export report is assembled from the before/after packages and the
      // save's own part report (file-compat §7). It is shown only when it has
      // something stale, removed or warned about; the model is never touched.
      const transformedParts = result.report.patchedParts.filter(
        (path) => result.report.modeByPart.get(path) === 'tree',
      );
      const formatChange =
        document.fileName.toLowerCase().endsWith('.xlsm') &&
        outcome.fileName.toLowerCase().endsWith('.xlsx')
          ? ({ from: 'xlsm', to: 'xlsx' } as const)
          : null;
      const exportReport = buildExportReport({
        workbook: document.session.workbook,
        before: document.pkg,
        after: result.package,
        report: result.report,
        compatibility: document.compatibility,
        noOp: result.noOp,
        structuralChange: savedStructuralVersion > 0,
        formatChange,
        preserved: listPreservedParts({
          package: document.pkg,
          entries: result.package.entries,
          workbook: document.session.workbook,
          dirtyCells: dirtyKeysFromSnapshot(savedDirty),
          transformedParts,
          format: outcome.fileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx',
        }),
      });
      setLastExportReport(exportReport);
      adoptSavedPackage(document.session.workbook, result.package);
      document.session.markSaved(savedDirty, savedDimensions, savedStructuralVersion, savedMerges);
      // The write itself moves the file's stamp; capture the new one so the
      // next Save compares against our own write, not an external edit.
      const savedStamp = outcome.handle === null ? null : await readFileStamp(outcome.handle);
      const savedDocumentId = document.session.workbook.id;
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
          fileStamp: savedStamp,
          dirtyFromRecovery: false,
          // A confirmed format change removed the macro parts (§6.3), so the
          // next save must not warn about macros that are no longer there.
          compatibility:
            macroPlan.format === 'xlsx' && current.compatibility.hasMacros
              ? { ...current.compatibility, hasMacros: false }
              : current.compatibility,
          preservedParts: listPreservedParts({
            package: result.package,
            workbook,
            format: outcome.fileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx',
          }),
        };
      });
      setVersion((current) => current + 1);
      recoveryRef.current = null;
      recoveryHintRef.current = null;
      // Any recovery write started before this save is now superseded.
      loadGenerationRef.current += 1;
      if (document.session.hasUnsavedChanges) {
        // An edit landed while the browser was writing, so it is not in the
        // saved bytes. Keep a fresh copy for it instead of dropping the old
        // one (persistence §2: never lose unsaved work).
        setTimeout(() => void recoveryWriteRef.current(), 0);
      } else {
        void recoveryStore()
          .discard(savedDocumentId)
          .catch(() => undefined);
      }
      rememberRecent({
        name: outcome.fileName,
        handle: outcome.handle ?? document.handle,
        stamp: savedStamp,
      });
      let savedMessage = result.noOp ? 'Saved. No changes to write.' : 'Saved.';
      if (outcome.method === 'download' && !externalEditWarningRef.current.has(savedDocumentId)) {
        // Persistence §4: a browser without handles cannot detect external
        // edits, so say so once per document.
        externalEditWarningRef.current.add(savedDocumentId);
        savedMessage = `${savedMessage} ${EXTERNAL_EDIT_UNSUPPORTED_WARNING}`;
      }
      setStatus(savedMessage);
      if (exportReport.needsAttention) setDialog('export-report');
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
    // Starting another gesture mode ends copy mode; the marquee would otherwise
    // mark a destination the painter is about to overwrite (ticket 31 B).
    setCopyMarquee(null);
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
          activeCell: active,
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
    setSortRequest(null);
    setPrintPlan(null);
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
    setCopyMarquee(null);
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
    if (result.clipboard !== undefined) {
      clipboardRef.current = result.clipboard;
      // Copy mode: the dashed marquee marks the copied range and becomes the
      // paste destination the arrow keys move (ticket 31 B).
      setCopyMarquee(range);
    }
    const label =
      projection === null ? 'the selection' : formatSelectionAddress(selection, projection);
    setStatus(`${mode === 'copy' ? 'Copied' : 'Cut'} ${label}.`);
    focusGrid();
  };

  /**
   * Paste the in-app snapshot at a destination in one `edit.paste`
   * transaction. Copy mode ends only when the paste is applied, so a rejected
   * paste leaves the marquee in place to try again.
   */
  const pasteAt = (target: GridRange): void => {
    if (document === null) return;
    const snapshot = clipboardRef.current;
    if (snapshot === null) {
      setStatus('The clipboard is empty.');
      setCopyMarquee(null);
      focusGrid();
      return;
    }
    const result = document.session.dispatch('edit.paste', {
      sheetId: document.sheetId,
      target,
      snapshot,
    });
    setVersion((current) => current + 1);
    if (result.ui?.message !== undefined) {
      setStatus(result.ui.message);
      focusGrid();
      return;
    }
    setCopyMarquee(null);
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

  /** Ctrl+V pastes at the marquee while copy mode is on, otherwise at the selection. */
  const pasteClipboard = (): void => {
    pasteAt(copyMarquee ?? selectionRange(selection));
  };

  /** An arrow key moves the marquee one row or column, clamped to the sheet. */
  const moveCopyMarquee = (rowDelta: number, colDelta: number): void => {
    if (copyMarquee === null || projection === null) return;
    setCopyMarquee(
      moveRange(copyMarquee, rowDelta, colDelta, projection.rowCount, projection.columnCount),
    );
  };

  /** Enter pastes at the marquee; Esc and any other action end copy mode. */
  const pasteCopyMarquee = (): void => {
    if (copyMarquee === null) return;
    pasteAt(copyMarquee);
  };

  const endCopyMarquee = (): void => {
    setCopyMarquee(null);
  };

  const dirty =
    document !== null && (document.session.hasUnsavedChanges || document.dirtyFromRecovery);
  const calculation = document?.calculation.state() ?? null;
  const staleCount = calculation?.staleCount ?? 0;
  const activeProjected = projection?.cellAt(active.row, active.col) ?? null;
  // The bar's draft owns the text while it is editing; otherwise it mirrors an
  // in-cell edit or the active cell.
  const barText = barDraft ?? liveEditText ?? projection?.rawTextAt(active.row, active.col) ?? '';
  const barFormulaActive = barDraft !== null && barDraft.startsWith('=');

  const commandContext = useMemo<ChromeCommandContext>(
    () => ({
      workbook: document?.session.workbook ?? null,
      activeSheetId: document?.sheetId ?? null,
      selection: selectionRange(selection),
      activeCell: active,
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

  // The ribbon readouts and the colour pickers' Current swatch all preview the
  // active cell (ticket 31). A merged cell is one cell, so they read its
  // anchor's style (ticket 40).
  const controlValues = useMemo(
    () => ribbonControlValues(projection, active),
    [projection, active],
  );

  useEffect(() => {
    if (!narrow) setRibbonExpanded(false);
  }, [narrow]);

  useEffect(() => {
    if (!shouldWarnBeforeUnload(dirty)) return;
    const listener = (event: BeforeUnloadEvent): void => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', listener);
    return () => window.removeEventListener('beforeunload', listener);
  }, [dirty]);

  // Persistence §2: debounce a recovery write about three seconds after the
  // last document edit. `version` moves on every command and undo, so the
  // effect restarts the timer; view changes do not touch it.
  useEffect(() => {
    if (document === null || !document.session.hasUnsavedChanges) return;
    const timer = setTimeout(() => {
      if (recoveryTimerRef.current === timer) recoveryTimerRef.current = null;
      void recoveryWriteRef.current();
    }, RECOVERY_DEBOUNCE_MS);
    recoveryTimerRef.current = timer;
    return () => {
      if (recoveryTimerRef.current !== timer) return;
      clearTimeout(timer);
      recoveryTimerRef.current = null;
    };
  }, [document, version]);

  // Persistence §2: also write on `visibilitychange` (hidden) and `pagehide`.
  // The listener is registered once and reads the latest writer through a ref.
  useEffect(() => {
    const flush = (): void => {
      void recoveryWriteRef.current();
    };
    const onVisibility = (): void => {
      if (window.document.visibilityState === 'hidden') flush();
    };
    window.addEventListener('pagehide', flush);
    window.document.addEventListener('visibilitychange', onVisibility);
    return () => {
      window.removeEventListener('pagehide', flush);
      window.document.removeEventListener('visibilitychange', onVisibility);
    };
  }, []);

  // Persistence §2: offer the newest recovery copy once, at startup. The guard
  // stops a re-render from asking again in the same session.
  useEffect(() => {
    if (recoveryOfferCheckedRef.current) return;
    recoveryOfferCheckedRef.current = true;
    void offerRecovery(false);
  }, []);

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
      // shortcuts from opening over an in-app dialog. The AutoFilter dropdown
      // is not modal, but it owns the keyboard while it is open for the same
      // reason: its search box and Esc must not trigger a ribbon shortcut.
      if (dialog !== null || filterDropdown !== null) {
        if (filterDropdown !== null && event.key === 'Escape') {
          event.preventDefault();
          setFilterDropdown(null);
          focusGrid();
          return;
        }
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
      // Esc ends copy mode from any zone. The grid owns the key when it has
      // the focus, so this is the fallback (ticket 31 B).
      if (event.key === 'Escape' && copyMarquee !== null) {
        event.preventDefault();
        setCopyMarquee(null);
        return;
      }
      if (event.altKey && event.key === '=') {
        event.preventDefault();
        setFocusZone('grid');
        runCommandRef.current?.('formula.autosum');
        return;
      }
      // Alt+Down opens the AutoFilter dropdown for the active column (§3.8).
      // The grid's navigation resolver leaves Alt combinations alone, so the
      // event reaches this listener even with the grid focused. The dropdown
      // focuses its own first item, so the focus zone is left untouched here.
      if (event.altKey && event.key === 'ArrowDown') {
        event.preventDefault();
        openFilterDropdown(active.col);
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
      } else if (!inTextInput && event.shiftKey && (event.key === 'l' || event.key === 'L')) {
        // Ctrl+Shift+L toggles AutoFilter on the selection (§3.8).
        event.preventDefault();
        runCommandRef.current?.('data.filter');
      } else if (event.key === 's' || event.key === 'S') {
        event.preventDefault();
        runCommandRef.current?.(event.shiftKey ? 'file.saveAs' : 'file.save');
      } else if (event.key === 'o' || event.key === 'O') {
        event.preventDefault();
        runCommandRef.current?.('file.open');
      } else if (!event.shiftKey && (event.key === 'p' || event.key === 'P')) {
        // Ctrl+P opens the in-app preview; the browser's own print then runs
        // from the preview, so the plan's pages are what reaches the printer.
        event.preventDefault();
        runCommandRef.current?.('file.print');
      } else if (!inTextInput && (event.key === 'z' || event.key === 'Z')) {
        event.preventDefault();
        runCommandRef.current?.(event.shiftKey ? 'edit.redo' : 'edit.undo');
      } else if (!inTextInput && (event.key === 'y' || event.key === 'Y')) {
        event.preventDefault();
        runCommandRef.current?.('edit.redo');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'b' || event.key === 'B')) {
        event.preventDefault();
        runCommandRef.current?.('format.bold');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'd' || event.key === 'D')) {
        // Ctrl+D fills the selection down from its first row (keyboard map §6).
        event.preventDefault();
        runCommandRef.current?.('edit.fillDown');
      } else if (!inTextInput && !event.shiftKey && (event.key === 'r' || event.key === 'R')) {
        // Ctrl+R fills the selection right from its first column.
        event.preventDefault();
        runCommandRef.current?.('edit.fillRight');
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
    if (
      target !== null &&
      !target.contains(window.document.activeElement) &&
      !headerRef.current?.contains(window.document.activeElement)
    ) {
      target.focus();
    }
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
        setCopyMarquee(null);
      }
      // A jump changes the selection, so copy mode ends first (ticket 31 B).
      setCopyMarquee(null);
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

  /**
   * A Ctrl+click on a linked cell (feature matrix §8 "followable
   * (Ctrl+click)"). An external address opens in a new tab only when its scheme
   * is on the allowlist; an internal location resolves through the same path as
   * the name box and Go To. Every refusal reports why instead of failing
   * silently, and the app itself never navigates.
   */
  const followHyperlink = useCallback(
    (link: HyperlinkRecord): void => {
      if (link.target !== '') {
        const check = checkExternalLink(link.target);
        if (check.kind === 'refused') {
          setStatus(check.message);
          return;
        }
        window.open(check.url, '_blank', 'noopener,noreferrer');
        setStatus('Opened the hyperlink in a new tab.');
        return;
      }
      const location = link.location ?? '';
      const parsed = parseInternalLocation(location);
      if (parsed === null) {
        setStatus(`The hyperlink target "${location}" could not be resolved.`);
        return;
      }
      const resolution = resolveNavigationTarget(internalLocationText(parsed));
      if (resolution.kind === 'invalid') {
        setStatus(
          `The hyperlink target "${location}" could not be resolved: ${resolution.message}`,
        );
        return;
      }
      selectNavigationTarget(resolution.target);
      setFocusZone('grid');
      setStatus(`Followed the link to ${resolution.target.sheetName}!${resolution.target.label}.`);
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
        activeCell: active,
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
   * A completed fill gesture is one `fill.handle` command: the grid supplies
   * the source block and the whole destination rectangle, the command owns the
   * copy/adjust logic, and the whole gesture is one undo transaction.
   */
  const handleFill = (source: GridRange, destination: GridRange): void => {
    const result = dispatchSheetCommand('fill.handle', { source, destination });
    if (result === null) return;
    if (result.ui?.message !== undefined) {
      setStatus(result.ui.message);
    } else if (result.changes.length === 0) {
      setStatus('The fill made no change.');
    } else {
      setStatus(`Filled ${formatRangeAddress(destination, false, false)}.`);
    }
    if (result.ui?.type !== 'dialog') focusGrid();
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
        activeCell: active,
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
            active,
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
    // A match changes the selection, so copy mode ends first (ticket 31 B).
    setCopyMarquee(null);
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
          activeCell: active,
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
    setCopyMarquee(null);
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
        activeCell: active,
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
      /**
       * The effective hidden rows of the active sheet: the imported manual
       * layer unioned with the filter layer (sort-autofilter §3.3). The grid
       * paints manual rows only until the renderer consumes the filter layer.
       */
      getHiddenRows: () => {
        if (document === null) return [];
        const sheet = document.session.workbook.sheets.find(
          (candidate) => candidate.id === document.sheetId,
        );
        if (sheet === undefined) return [];
        const manual = new Set<number>();
        for (const [row, dimension] of sheet.rowDimensions) {
          if (dimension.hidden === true) manual.add(row);
        }
        const layers = computeHiddenLayers(sheet, sheet.autoFilter, manual);
        return [...layers.effectiveHidden].sort((left, right) => left - right);
      },
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
      /** Test hook: force the debounced recovery write and wait for it. */
      flushRecovery: async () => {
        await writeRecoveryNow();
      },
      getRecoveryPending: () => recoveryStore().listPending(),
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
    // The bar's commit moves the selection, so copy mode ends first (ticket 31 B).
    setCopyMarquee(null);
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
          activeCell: active,
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
      const message =
        result.changes.length === 0 ? `No formatting change for ${label}.` : `Formatted ${label}.`;
      setStatus(message);
      // The dialog shows this same line when Apply keeps the dialog open, so a
      // no-change apply never claims it wrote something (issue 35).
      return { ok: true, message };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(message);
      return { ok: false, message };
    }
  };

  /**
   * The Sort dialog's columns: every column of the target range with its header
   * text when the first-row cell is text, so the labels match the labels the
   * command reports for the applied sort (§2.2).
   */
  const sortColumns = useMemo<readonly SortDialogColumn[]>(() => {
    if (sortRequest === null || projection === null) return [];
    const columns: SortDialogColumn[] = [];
    for (let col = sortRequest.range.startCol; col <= sortRequest.range.endCol; col += 1) {
      const cell = projection.cellAt(sortRequest.range.startRow, col);
      columns.push({
        column: col,
        letter: projection.columnLabel(col),
        header: cell !== null && cell.value.type === 'text' ? cell.displayText : '',
      });
    }
    return columns;
  }, [sortRequest, projection]);

  /** Open the Sort dialog from a `needs-keys` command result. */
  const openSortDialog = (request: SortPayload): void => {
    clearPainter();
    setSortRequest(request);
    setDialog('sort');
  };

  /**
   * Confirming the dialog dispatches `data.sort` with the collected keys. The
   * command owns the transaction, so the whole sort is one undo step (§2.7).
   * A blocked sort keeps the dialog open with Excel's message.
   */
  const submitSort = (request: SortSubmit): SortDialogOutcome => {
    if (document === null || sortRequest === null) {
      return { ok: false, message: 'Open a workbook first.' };
    }
    try {
      const result = document.session.dispatch(
        'data.sort',
        {
          sheetId: sortRequest.sheetId,
          range: sortRequest.range,
          keys: request.keys,
          hasHeaders: request.hasHeaders,
        },
        {
          activeSheetId: document.sheetId,
          selection: selectionRange(selection),
          activeCell: active,
        },
      );
      if (result.changes.length > 0) setVersion((value) => value + 1);
      const payload = result.sort;
      if (payload === undefined) {
        const message = result.ui?.message ?? 'The sort did not run.';
        setStatus(message);
        return { ok: false, message };
      }
      setStatus(sortStatusMessage(payload));
      if (payload.state === 'blocked') {
        return { ok: false, message: payload.message ?? 'The sort was blocked.' };
      }
      return { ok: true, message: null };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(message);
      return { ok: false, message };
    }
  };

  /** Open the column dropdown under the active filter's header cell. */
  const openFilterDropdown = (column: number): void => {
    if (document === null || projection === null) return;
    const sheet = document.session.workbook.sheets.find(
      (candidate) => candidate.id === document.sheetId,
    );
    const filter = sheet?.autoFilter ?? null;
    if (filter === null) {
      setStatus('Turn on AutoFilter (Ctrl+Shift+L) before opening a column filter.');
      return;
    }
    if (column < filter.range.startCol || column > filter.range.endCol) {
      setStatus('That column is outside the AutoFilter range.');
      return;
    }
    const anchor = filterDropdownAnchor(projection, column, filter.range.startRow);
    if (anchor === null) return;
    setFilterDropdown({ column, anchor });
  };

  const closeFilterDropdown = (): void => {
    setFilterDropdown(null);
    focusGrid();
  };

  /**
   * Submit one filter command for the dropdown. The dropdown never touches the
   * model: Apply, Clear and the conditions window all land here, and each one
   * is the command's single transaction (§3.9).
   */
  const dispatchFilterCommand = (
    commandId: string,
    args: Record<string, unknown>,
  ): CommandResult | null => {
    if (document === null) return null;
    try {
      const result = document.session.dispatch(
        commandId,
        { sheetId: document.sheetId, ...args },
        {
          activeSheetId: document.sheetId,
          selection: selectionRange(selection),
          activeCell: active,
        },
      );
      setVersion((value) => value + 1);
      if (result.ui?.message !== undefined) setStatus(result.ui.message);
      else if (result.filter !== undefined) setStatus(filterStatusMessage(result.filter));
      return result;
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
      return null;
    }
  };

  /**
   * Apply or clear one column's criterion. An imported criterion calcplus cannot
   * edit is replaced only after its confirmation (§3.7), the same way the sheet
   * delete asks before its terminal edit.
   */
  const applyFilterCriterion = (
    column: number,
    criterion: ColumnFilter | null,
  ): FilterApplyOutcome => {
    const commandId = criterion === null ? 'data.filterClearColumn' : 'data.filterColumn';
    const first = dispatchFilterCommand(
      commandId,
      criterion === null ? { column } : { column, criterion },
    );
    if (first === null) return { ok: false, message: 'The filter could not be applied.' };
    if (first.ui?.type === 'dialog') {
      const confirmed =
        typeof window === 'undefined' ||
        window.confirm(first.ui.message ?? 'Replace the existing filter?');
      if (!confirmed) {
        setStatus('Filter replacement cancelled.');
        return { ok: false, message: 'The existing filter was kept.' };
      }
      const second = dispatchFilterCommand('data.filterColumn', {
        column,
        criterion,
        confirm: true,
      });
      if (second === null) return { ok: false, message: 'The filter could not be applied.' };
    }
    return { ok: true, message: null };
  };

  /**
   * The dropdown's Sort A to Z / Sort Z to A. The AutoFilter's first row is
   * always the header (§3.1), so the sort keeps it and reorders the whole
   * filter range, not just the current region.
   */
  const sortFilterColumn = (column: number, direction: SortDirection): void => {
    if (document === null) return;
    const sheet = document.session.workbook.sheets.find(
      (candidate) => candidate.id === document.sheetId,
    );
    const filter = sheet?.autoFilter ?? null;
    if (filter === null) return;
    setFilterDropdown(null);
    runCommand('data.sort', {
      sheetId: document.sheetId,
      range: filter.range,
      keys: [{ column, direction }],
      hasHeaders: true,
    });
  };

  /**
   * The dropdown's checklist, criterion and type. The model groups, orders and
   * formats the values; the chrome only ticks entries and builds the criterion.
   */
  const filterPanel = useMemo(() => {
    if (document === null || projection === null || filterDropdown === null) return null;
    const sheet = document.session.workbook.sheets.find(
      (candidate) => candidate.id === document.sheetId,
    );
    if (sheet === undefined) return null;
    const filter = sheet.autoFilter;
    if (filter === null) return null;
    const column = filterDropdown.column;
    if (column < filter.range.startCol || column > filter.range.endCol) return null;
    const offset = column - filter.range.startCol;
    const checklist = collectValueChecklist(sheet, filter, offset, {
      displayText: filterChecklistDisplay(sheet, filter, column, projection),
    });
    const headerFormat =
      projection.cellAt(filter.range.startRow, column)?.style.numberFormat ?? null;
    const bodyFormat =
      projection.cellAt(filter.range.startRow + 1, column)?.style.numberFormat ?? null;
    return {
      column,
      columnLabel: filterColumnLabel(projection, filter, column),
      range: filter.range,
      checklist,
      criterion: filter.columns.get(offset) ?? null,
      filterType: detectFilterColumnType(
        checklist,
        isDateFormatCode(headerFormat) || isDateFormatCode(bodyFormat),
      ),
      date1904: document.session.workbook.date1904,
    };
  }, [document, projection, filterDropdown]);

  /** The status bar's record count, derived from the same layers as the payload. */
  const filterRecords = useMemo(() => {
    if (document === null) return null;
    const sheet = document.session.workbook.sheets.find(
      (candidate) => candidate.id === document.sheetId,
    );
    if (sheet === undefined) return null;
    const filter = sheet.autoFilter;
    if (filter === null) return null;
    return filterRecordCount(sheet, filter);
  }, [document, version]);

  // Undo, a sheet switch or AutoFilter being switched off can remove the column
  // the dropdown was opened on. The overlay closes rather than describing a
  // filter that no longer exists.
  useEffect(() => {
    if (filterDropdown !== null && filterPanel === null) setFilterDropdown(null);
  }, [filterDropdown, filterPanel]);

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
      case 'file.close':
        closeWorkbook();
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
      case 'merge.center':
      case 'merge.across':
      case 'merge.cells':
      case 'merge.unmerge': {
        if (document === null) return null;
        const first = dispatchSheetCommand(commandId, args);
        if (first === null) return null;
        if (first.ui?.type === 'dialog') {
          // Discarding covered data asks first, like the sheet delete. The
          // second dispatch carries the answer into the same one transaction.
          const confirmed =
            typeof window === 'undefined' || window.confirm(first.ui.message ?? MERGE_DATA_WARNING);
          if (!confirmed) {
            setStatus('Merge cancelled. No cells were changed.');
            focusGrid();
            return null;
          }
          const second = dispatchSheetCommand(commandId, {
            ...(typeof args === 'object' && args !== null ? args : {}),
            confirm: true,
          });
          if (second !== null) {
            setStatus(
              second.changes.length > 0
                ? (second.ui?.message ?? mergeStatusMessage(commandId))
                : (second.ui?.message ?? 'The merge made no change.'),
            );
          }
          if (second?.ui?.type !== 'dialog') focusGrid();
          return null;
        }
        if (first.changes.length > 0) {
          setStatus(first.ui?.message ?? mergeStatusMessage(commandId));
        } else if (first.ui?.message === undefined) {
          setStatus('The merge made no change.');
        }
        focusGrid();
        return null;
      }
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
      case 'data.sort': {
        if (document === null) return null;
        const result = dispatchSheetCommand('data.sort', args);
        if (result === null) return null;
        if (result.sort?.state === 'needs-keys') {
          // The menu and the ribbon ask the dialog for keys; the dialog then
          // submits this command again with them.
          openSortDialog(result.sort);
          return null;
        }
        if (result.sort !== undefined) setStatus(sortStatusMessage(result.sort));
        focusGrid();
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
      case 'file.compatibilityReport':
        if (document === null) return null;
        clearPainter();
        setDialog('compatibility-report');
        return null;
      case 'file.exportReport':
        if (document === null) return null;
        clearPainter();
        setDialog('export-report');
        return null;
      case 'file.recent':
        clearPainter();
        void refreshRecent();
        setDialog('recent');
        return null;
      case 'file.recover':
        clearPainter();
        void offerRecovery(true);
        return null;
      case 'file.print': {
        if (document === null || projection === null || busy) return null;
        clearPainter();
        try {
          // The plan is computed once per print (print-pdf §6) from the same
          // display records the grid paints, at 100%, never the view zoom.
          const plan = createPrintPlan(createPrintSnapshot(document.session.workbook, projection));
          setPrintPlan(plan);
          setDialog('print-preview');
        } catch (error) {
          setStatus(error instanceof Error ? error.message : String(error));
          focusGrid();
        }
        return null;
      }
      case 'file.exportPdf': {
        if (document === null || projection === null || busy) return null;
        clearPainter();
        try {
          // The same plan the preview builds, so the PDF matches it (print-pdf §7).
          const plan = createPrintPlan(createPrintSnapshot(document.session.workbook, projection));
          setStatus('Building the PDF…');
          void exportPdfPlan(plan, document.fileName).then(
            (summary) => {
              const raster =
                summary.fallbackPages.length === 0
                  ? ''
                  : `; ${summary.fallbackPages.length} page(s) rasterised at 300 DPI`;
              setStatus(`Exported ${summary.fileName} (${summary.pageCount} pages${raster}).`);
            },
            (error: unknown) => {
              setStatus(error instanceof Error ? error.message : String(error));
            },
          );
        } catch (error) {
          setStatus(error instanceof Error ? error.message : String(error));
        }
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
        resolveState={resolveState}
        onCommand={runCommand}
        onFocus={() => setFocusZone('ribbon')}
        headerRef={headerRef}
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

      {document !== null &&
      compatibilityReport !== null &&
      (compatibilityReport.hasContent || document.compatibility.notes.length > 0) ? (
        <div
          className="compat-banner"
          title={[
            ...document.compatibility.notes,
            compatibilityReportSummary(compatibilityReport),
          ].join(' ')}
          data-testid="compat-banner"
        >
          <span className="compat-banner-text">
            {[
              ...document.compatibility.notes,
              ...(compatibilityReport.hasContent
                ? [compatibilityReportSummary(compatibilityReport)]
                : []),
            ].join(' ')}
          </span>
          {compatibilityReport.hasContent ? (
            <button
              type="button"
              className="compat-banner-action"
              data-testid="compatibility-report-open"
              onClick={() => runCommand('file.compatibilityReport')}
            >
              Review
            </button>
          ) : null}
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
          if (document !== null) {
            setContextMenu({
              x: event.clientX,
              y: event.clientY,
              region: contextMenuRegion(event, resolveState('view.headings').checked),
            });
          }
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
            onFill={handleFill}
            onFollowHyperlink={followHyperlink}
            onFilterHeader={openFilterDropdown}
            marquee={
              copyMarquee === null
                ? undefined
                : {
                    range: copyMarquee,
                    onMove: moveCopyMarquee,
                    onPaste: pasteCopyMarquee,
                    onEnd: endCopyMarquee,
                  }
            }
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
        stale={activeProjected?.stale === true}
        staleCount={staleCount}
        statistics={selectionStatistics}
        records={filterRecords}
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
          items={contextMenuItems(resolveState, contextMenu.region)}
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
          initial={activeProjected?.style}
          targetLabel={formatRangeAddress(selectionRange(selection), false, false)}
          initialTab={formatCellsTab ?? undefined}
          onClose={closeDialog}
          onApply={applyFormatCells}
        />
      ) : null}

      {dialog === 'about' ? <AboutDialog onClose={closeDialog} /> : null}

      {dialog === 'sort' && sortRequest !== null ? (
        <SortDialog
          targetLabel={formatRangeAddress(sortRequest.range, false, false)}
          columns={sortColumns}
          initialHasHeaders={sortRequest.hasHeaders ?? false}
          onClose={closeDialog}
          onSubmit={submitSort}
        />
      ) : null}

      {filterPanel !== null ? (
        <FilterDropdown
          columnLabel={filterPanel.columnLabel}
          checklist={filterPanel.checklist}
          criterion={filterPanel.criterion}
          filterType={filterPanel.filterType}
          date1904={filterPanel.date1904}
          anchor={filterDropdown?.anchor ?? { top: 0, left: 0, headerTop: 0 }}
          onSort={(direction) => sortFilterColumn(filterPanel.column, direction)}
          onCustomSort={() => {
            setFilterDropdown(null);
            runCommand('data.sort', { range: filterPanel.range });
          }}
          onClearColumn={() => {
            applyFilterCriterion(filterPanel.column, null);
          }}
          onApply={(criterion) => applyFilterCriterion(filterPanel.column, criterion)}
          onClose={closeFilterDropdown}
        />
      ) : null}

      {dialog === 'insert-function' ? (
        <InsertFunctionDialog
          recentlyUsed={recentFunctions}
          onClose={closeDialog}
          onInsert={insertFunctionByName}
        />
      ) : null}

      {dialog === 'compatibility-report' && compatibilityReport !== null ? (
        <ReportDialog
          title="Compatibility Report"
          summary={
            compatibilityReport.hasContent
              ? compatibilityReportSummary(compatibilityReport)
              : 'Nothing to report.'
          }
          sections={compatibilitySections(compatibilityReport)}
          notes={compatibilityReport.notes}
          emptyMessage="There is no preserved, unsupported or unenforced content in this workbook."
          copyText={compatibilityReport.text}
          onClose={closeDialog}
        />
      ) : null}

      {dialog === 'export-report' ? (
        <ReportDialog
          title="Export Report"
          summary={
            lastExportReport === null
              ? 'No export has been reported since this workbook was opened.'
              : exportReportSummary(lastExportReport)
          }
          sections={lastExportReport === null ? [] : exportSections(lastExportReport)}
          emptyMessage="No export has been reported since this workbook was opened."
          copyText={lastExportReport?.text ?? 'Export report\n\nNothing to report.'}
          onClose={closeDialog}
        />
      ) : null}

      {dialog === 'print-preview' && printPlan !== null ? (
        <PrintPreviewDialog plan={printPlan} onClose={closeDialog} />
      ) : null}

      {dialog === 'recent' ? (
        <RecentDialog
          entries={recentEntries}
          onOpen={(entry) => void openRecent(entry)}
          onRemove={(entry) => {
            void recentStore()
              .remove(entry.id)
              .then(setRecentEntries)
              .catch(() => undefined);
          }}
          onClose={closeDialog}
        />
      ) : null}
    </div>
  );
}

/** The picked-up style record is view state returned by `painter.pickup`. */
function readPainterSnapshot(result: CommandResult): PainterSnapshot | null {
  return result.painter ?? null;
}

/** The structured-clone value from a recovery record, when it is a handle. */
function asFileHandle(value: unknown): FileHandleLike | null {
  if (typeof value !== 'object' || value === null) return null;
  const candidate = value as Partial<FileHandleLike>;
  if (typeof candidate.name !== 'string' || typeof candidate.createWritable !== 'function') {
    return null;
  }
  return candidate as FileHandleLike;
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

/** The status line a completed merge command reports. */
function mergeStatusMessage(commandId: string): string {
  if (commandId === 'merge.unmerge') return 'Unmerged the selected cells.';
  return 'Merged the selected cells.';
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

/** The status-bar summary of a sort's outcome; the dialog shows blocks inline. */
function sortStatusMessage(payload: SortPayload): string {
  if (payload.state === 'sorted') {
    const keys = payload.keyLabels ?? [];
    const range = formatRangeAddress(payload.range, false, false);
    const by = keys.length === 0 ? '' : ` by ${keys.join(', ')}`;
    const warning = payload.warnings[0];
    return `Sorted ${range}${by}.${warning === undefined ? '' : ` ${warning.message}`}`;
  }
  if (payload.state === 'blocked') return payload.message ?? 'The sort was blocked.';
  switch (payload.reason) {
    case 'no-keys':
      return 'Choose at least one sort key.';
    case 'too-few-rows':
      return 'That range has fewer than two data rows; there is nothing to sort.';
    case 'nothing-to-order':
      return 'The sort made no change.';
    default:
      return 'The sort made no change.';
  }
}

/** The status-bar summary of an AutoFilter command, including `X of Y records`. */
function filterStatusMessage(payload: FilterPayload): string {
  if (!payload.changed) return 'The filter made no change.';
  const records = `${payload.rowsVisible} of ${payload.rowsTotal} records`;
  switch (payload.state) {
    case 'cleared':
      return `Filter cleared. Showing ${records}.`;
    case 'reapplied':
      return `Filter reapplied. Showing ${records}.`;
    case 'none':
      return 'AutoFilter is off.';
    default:
      return `Filter applied. Showing ${records}.`;
  }
}

/**
 * The AutoFilter's own label for a column: the header text, or `Column1`,
 * `Column2`, … when the first row looks like data (§3.1).
 */
function filterColumnLabel(view: SheetProjection, filter: AutoFilter, column: number): string {
  const header = view.cellAt(filter.range.startRow, column);
  if (header !== null && header.value.type === 'text' && header.displayText.trim() !== '') {
    return header.displayText;
  }
  return `Column${column - filter.range.startCol + 1}`;
}

/**
 * The checklist's display seam. The model groups values by their formatted
 * display; the projection owns that formatting, so each cell's painted text is
 * looked up by identity and the model's plain formatting is the fallback for a
 * cell the projection does not hold.
 */
function filterChecklistDisplay(
  sheet: Sheet,
  filter: AutoFilter,
  column: number,
  view: SheetProjection,
): (cell: Cell, value: CellValue) => string {
  const displays = new WeakMap<Cell, string>();
  for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
    const cell = sheet.cells.get(cellKey(row, column));
    if (cell === undefined) continue;
    const display = view.cellAt(row, column)?.displayText;
    if (display !== undefined) displays.set(cell, display);
  }
  return (cell, value) => displays.get(cell) ?? formatCellValue(value);
}

/**
 * Which strip of the grid a right-click landed on. The header strips are
 * painted by the grid canvas, so the chrome measures the same published
 * geometry the grid paints with (`gridDebugState.layout` plus the canvas
 * bounding rect). With headings hidden there is no header region, and the
 * corner above the row numbers counts as a cell: it has no row or column of
 * its own, and Excel offers the cell menu there.
 */
function contextMenuRegion(
  event: { readonly clientX: number; readonly clientY: number },
  showHeadings: boolean,
): ContextMenuRegion {
  if (!showHeadings) return 'cell';
  const canvas = window.document.querySelector<HTMLElement>('[data-testid="grid-canvas"]');
  const layout = gridDebugState.layout;
  if (canvas === null || layout === null) return 'cell';
  const box = canvas.getBoundingClientRect();
  const x = event.clientX - box.left;
  const y = event.clientY - box.top;
  if (x < layout.headerWidth && y >= layout.headerHeight) return 'rowHeader';
  if (y < layout.headerHeight && x >= layout.headerWidth) return 'columnHeader';
  return 'cell';
}

/**
 * Anchor the dropdown under a column's header cell. The header buttons are
 * painted by the grid canvas, which the chrome cannot hit-test yet, so the
 * overlay is placed from the same published geometry the grid paints with
 * (`gridDebugState` plus `cellRect`). It is measured when it opens; a scroll
 * dismisses it rather than leaving it detached from its column.
 */
function filterDropdownAnchor(
  view: SheetProjection,
  column: number,
  headerRow: number,
): { readonly top: number; readonly left: number; readonly headerTop: number } | null {
  const canvas = window.document.querySelector<HTMLElement>('[data-testid="grid-canvas"]');
  const layout = gridDebugState.layout;
  if (canvas === null || layout === null) return null;
  const rect = cellRect(
    view,
    layout,
    gridDebugState.scrollX,
    gridDebugState.scrollY,
    headerRow,
    column,
  );
  const box = canvas.getBoundingClientRect();
  const headerTop = box.top + rect.y;
  return { top: headerTop + rect.height + 1, left: box.left + rect.x, headerTop };
}
