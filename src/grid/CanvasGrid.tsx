import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type {
  JSX,
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react';
import { flushSync } from 'react-dom';
import type { FunctionCatalogEntry } from '@engine/function-catalog';
import type { HyperlinkRecord } from '@model/types';
import type { SheetProjection } from '@view/types';
import { FormulaSuggestions } from './FormulaSuggestions';
import { useFormulaCompletion } from './useFormulaCompletion';
import { autofitColumnWidth, autofitRowHeight, type AutofitTarget } from './autofit';
import { CanvasRenderFrame, cellAtPoint } from './canvas-renderer';
import { gridDebugState } from './debug';
import {
  applyPointing,
  editorCommitText,
  editorHasChanges,
  isFormulaEntry,
  isPointingBoundary,
  pointingRange,
  pointingReference,
  resolveEditorAction,
  startEdit,
  type EditorCommitDirection,
  type EditorSession,
} from './editor';
import { filterButtonRect, pointInFilterButton } from './filter-affordance';
import {
  GRID_METRICS,
  type GridCellAddress,
  type GridRange,
  type GridSelection,
  type RenderFrame,
} from './frame';
import { mergeAnchor, resolveNavigation } from './navigation';
import {
  cellRect,
  clampScroll,
  columnAtHeader,
  computePaneLayout,
  headerResizeTarget,
  headerSelectionIndices,
  headerSelectionRange,
  mergedRect,
  rowAtHeader,
  scrollToReveal,
  type CellRect,
  type PaneLayout,
} from './panes';
import { resizeDimension, type DimensionAxis, type DimensionResize } from './resize';
import {
  activeCell,
  collapseSelection,
  extendSelection,
  selectionContains,
  selectionForRange,
  selectionRange,
} from './selection';

export interface CanvasGridProps {
  readonly projection: SheetProjection;
  readonly selection: GridSelection;
  readonly onSelectionChange: (selection: GridSelection) => void;
  readonly onCommit: (address: GridCellAddress, text: string) => void;
  readonly onClear: (range: GridRange) => void;
  /** In-app clipboard commands; the application owns the snapshot. */
  readonly onCopy?: () => void;
  readonly onCut?: () => void;
  readonly onPaste?: () => void;
  /**
   * Reports the live editor text, or `null` once the edit ends, so the formula
   * bar mirrors in-cell typing before the cell is committed.
   */
  readonly onEditorChange?: (text: string | null) => void;
  /**
   * The formula bar holds a formula draft and points at cells. A press inserts
   * a reference into the bar draft while the bar keeps the keyboard.
   */
  readonly formulaBarPointing?: {
    readonly onPoint: (anchor: GridCellAddress, focus: GridCellAddress) => void;
  } | null;
  /**
   * A header double-click autofits every selected row or column to its own
   * content. The grid owns the canvas metrics; the application stores the
   * results through one `dimension.set` transaction.
   */
  readonly onAutofit?: (axis: DimensionAxis, targets: readonly AutofitTarget[]) => void;
  /**
   * A header-edge drag previews a row height or column width on every selected
   * visible member. The size is view state until the pointer is released;
   * `null` abandons the preview.
   */
  readonly onResizePreview?: (
    axis: DimensionAxis,
    indices: readonly number[],
    result: DimensionResize | null,
  ) => void;
  /**
   * A header-edge drag ended. The application stores the result through the
   * undoable `dimension.set` command as one transaction.
   */
  readonly onResizeCommit?: (
    axis: DimensionAxis,
    indices: readonly number[],
    result: DimensionResize,
  ) => void;
  /**
   * Ctrl+wheel changes the display zoom. The application owns the live zoom
   * because it is view state, not document state.
   */
  readonly onZoomChange?: (zoomPercent: number) => void;
  /**
   * Format Painter mode. The grid paints the live target and turns a click or
   * drag into one apply gesture; the application owns the source snapshot and
   * dispatches the command.
   */
  readonly painterActive?: boolean;
  readonly onPainterApply?: (range: GridRange) => void;
  /**
   * The fill handle was released over a larger rectangle, or double-clicked.
   * The grid supplies the source block and the whole destination rectangle
   * (source included); the application dispatches one `fill.handle` command.
   */
  readonly onFill?: (source: GridRange, destination: GridRange) => void;
  /**
   * Ctrl+click on a cell that carries a hyperlink (file-compat §9 F13). The
   * grid owns the gesture only: the application decides whether an external
   * address may be opened, navigates an internal location and reports the
   * outcome. A plain click selects the cell as usual.
   */
  readonly onFollowHyperlink?: (link: HyperlinkRecord, address: GridCellAddress) => void;
  /**
   * Open the AutoFilter dropdown for a header column (sort-autofilter §3.8).
   * The grid owns the affordance and its hit test; the application owns the
   * dropdown and the filter commands.
   */
  readonly onFilterHeader?: (column: number) => void;
  /**
   * Copy mode (ticket 31 B). The dashed marquee marks the copy's destination:
   * the arrow keys move it without moving the selection, Enter pastes there and
   * Esc — or any other key or pointer action that moves the selection — ends
   * the mode. All state is the application's; it is view state, never document
   * state, exactly like the Format Painter's.
   */
  readonly marquee?: {
    readonly range: GridRange;
    /** One row or one column per press, clamped to the sheet. */
    readonly onMove: (rowDelta: number, colDelta: number) => void;
    /** Enter: paste the snapshot at the marquee and end copy mode. */
    readonly onPaste: () => void;
    /** Esc, another key, or a pointer that moves the selection. */
    readonly onEnd: () => void;
  };
  /**
   * The focus owner hands the grid shell to the focus manager.
   */
  readonly shellRef?: RefObject<HTMLDivElement | null>;
  readonly onFocusZone?: () => void;
}

interface Viewport {
  readonly width: number;
  readonly height: number;
}

interface ScrollOffset {
  readonly x: number;
  readonly y: number;
}

export function CanvasGrid({
  projection,
  selection,
  onSelectionChange,
  onCommit,
  onClear,
  onCopy,
  onCut,
  onPaste,
  onEditorChange,
  formulaBarPointing,
  onAutofit,
  onResizePreview,
  onResizeCommit,
  onZoomChange,
  painterActive,
  onPainterApply,
  onFill,
  onFollowHyperlink,
  onFilterHeader,
  marquee,
  shellRef,
  onFocusZone,
}: CanvasGridProps): JSX.Element {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const frameRef = useRef<RenderFrame | null>(null);
  const editorRef = useRef<HTMLTextAreaElement | null>(null);
  const verticalThumbRef = useRef<HTMLDivElement | null>(null);
  const horizontalThumbRef = useRef<HTMLDivElement | null>(null);
  const pendingPaintRef = useRef(false);
  const scrollRef = useRef<ScrollOffset>({ x: 0, y: 0 });
  const layoutRef = useRef<PaneLayout | null>(null);
  const stateRef = useRef({ projection, selection });
  const editorSessionRef = useRef<EditorSession | null>(null);
  const dragRef = useRef<{ axis: 'x' | 'y'; start: number; origin: number } | null>(null);
  const resizeDragRef = useRef<{
    readonly pointerId: number;
    readonly axis: DimensionAxis;
    readonly index: number;
  } | null>(null);
  const zoomWheelRef = useRef(0);
  const selectionDragRef = useRef<{
    readonly pointerId: number;
    readonly anchor: GridCellAddress;
    focus: GridCellAddress;
  } | null>(null);
  /** Stops the animation-frame loop that owns edge auto-scroll for one drag. */
  const autoScrollCleanupRef = useRef<(() => void) | null>(null);
  const painterTargetRef = useRef<GridRange | null>(null);
  const painterCleanupRef = useRef<(() => void) | null>(null);
  const painterDragRef = useRef<{
    readonly pointerId: number;
    readonly anchor: GridCellAddress;
    focus: GridCellAddress;
  } | null>(null);
  const fillDragRef = useRef<{
    readonly pointerId: number;
    readonly source: GridRange;
    focus: GridCellAddress;
  } | null>(null);
  const fillCleanupRef = useRef<(() => void) | null>(null);
  /** Time of the last fill-handle press, for Excel's double-click fill. */
  const fillPressRef = useRef(0);
  /**
   * The last fill-handle press point, in client coordinates, and its time. The
   * browser's `dblclick` fires after the fill has already moved the handle, so
   * the painted-handle hit test alone cannot recognise the gesture; the
   * recorded press point lets the double-click handler consume it.
   */
  const fillPressOriginRef = useRef<{
    readonly x: number;
    readonly y: number;
    readonly at: number;
  } | null>(null);
  const formulaDragRef = useRef<{
    readonly pointerId: number;
    readonly anchor: GridCellAddress;
  } | null>(null);
  const pointingRangeRef = useRef<GridRange | null>(null);
  const pendingCaretRef = useRef<number | null>(null);
  const [viewport, setViewport] = useState<Viewport>({ width: 800, height: 480 });
  const [editor, setEditor] = useState<EditorSession | null>(null);
  /** The in-cell caret, so the function suggestions follow the typed token. */
  const [editorCaret, setEditorCaret] = useState<number | null>(null);
  const editorCompletion = useFormulaCompletion({
    text: editor?.text ?? '',
    caret: editorCaret,
    formula: editor !== null && isFormulaEntry(editor),
    enabled: editor !== null && !editor.composing,
  });
  /** The live destination outline while the fill handle is dragged. */
  const [fillPreview, setFillPreview] = useState<GridRange | null>(null);
  /** The copy marquee's overlay, kept over its range while the sheet scrolls. */
  const marqueeRef = useRef<HTMLDivElement | null>(null);
  const marqueeRangeRef = useRef<GridRange | null>(null);
  /**
   * The visible filter-header columns. The canvas paints the arrow and funnel;
   * transparent DOM buttons over the same rectangles carry the accessible name
   * and the pointer/keyboard target.
   */
  const [filterHeaderColumns, setFilterHeaderColumns] = useState<readonly number[]>([]);
  const filterHeaderSignatureRef = useRef('');
  const filterButtonRefs = useRef(new Map<number, HTMLButtonElement>());
  /**
   * The transient hover overlays: a comment popup over a noted cell and a plain
   * tooltip over a linked cell. They are view-only and never touch the model or
   * the selection.
   */
  const commentPopupRef = useRef<HTMLDivElement | null>(null);
  const linkTooltipRef = useRef<HTMLDivElement | null>(null);
  const hoverRef = useRef<{
    readonly comment: GridCellAddress | null;
    readonly link: GridCellAddress | null;
  }>({ comment: null, link: null });
  const [hover, setHover] = useState(hoverRef.current);

  /**
   * The filter header columns currently on screen: the frozen strip plus the
   * scrolling columns between the current scroll offset and the right edge.
   * Scanning those two ranges avoids walking a filter range that spans the
   * whole sheet on every paint.
   */
  const syncFilterHeaderColumns = useCallback(
    (
      projection: SheetProjection,
      layout: PaneLayout,
      scroll: { readonly x: number; readonly y: number },
    ): void => {
      const filter = projection.autoFilter;
      const columns: number[] = [];
      if (
        filter !== null &&
        filter.range.startRow >= 0 &&
        filter.range.startRow < projection.rowCount &&
        !projection.isRowHidden(filter.range.startRow)
      ) {
        const frozenEnd = Math.min(filter.range.endCol, projection.freezeCols - 1);
        const scrollingStart = Math.max(filter.range.startCol, projection.columnAt(scroll.x));
        const scrollingEnd = Math.min(
          filter.range.endCol,
          projection.columnAt(scroll.x + layout.scrollRegionWidth),
        );
        for (let col = Math.max(0, filter.range.startCol); col <= frozenEnd; col += 1) {
          if (projection.isColumnHidden(col) || projection.columnWidth(col) <= 0) continue;
          columns.push(col);
        }
        for (let col = scrollingStart; col <= scrollingEnd; col += 1) {
          if (projection.isColumnHidden(col) || projection.columnWidth(col) <= 0) continue;
          columns.push(col);
        }
      }
      const signature = columns.join(',');
      if (signature === filterHeaderSignatureRef.current) return;
      filterHeaderSignatureRef.current = signature;
      setFilterHeaderColumns(columns);
    },
    [],
  );

  /** Keep the DOM buttons over their painted rectangles while the sheet scrolls. */
  const positionFilterButtons = useCallback(
    (
      projection: SheetProjection,
      layout: PaneLayout,
      scroll: { readonly x: number; readonly y: number },
    ): void => {
      const filter = projection.autoFilter;
      if (filter === null) return;
      const bodyRight = layout.headerWidth + layout.frozenWidth + layout.scrollRegionWidth;
      const bodyBottom = layout.headerHeight + layout.frozenHeight + layout.scrollRegionHeight;
      for (const [col, button] of filterButtonRefs.current) {
        if (projection.isColumnHidden(col) || projection.columnWidth(col) <= 0) {
          button.style.visibility = 'hidden';
          continue;
        }
        const cell = cellRect(projection, layout, scroll.x, scroll.y, filter.range.startRow, col);
        const rect = filterButtonRect(cell, projection.zoom);
        const visible =
          rect.width >= 8 &&
          rect.height >= 6 &&
          rect.x >= layout.headerWidth &&
          rect.x + rect.width <= bodyRight + 0.5 &&
          rect.y >= layout.headerHeight &&
          rect.y + rect.height <= bodyBottom + 0.5;
        if (!visible) {
          button.style.visibility = 'hidden';
          continue;
        }
        button.style.visibility = 'visible';
        button.style.transform = `translate(${rect.x}px, ${rect.y}px)`;
        button.style.width = `${rect.width}px`;
        button.style.height = `${rect.height}px`;
      }
    },
    [],
  );

  const paint = useCallback((): void => {
    const frame = frameRef.current;
    if (frame === null) return;
    const current = stateRef.current;
    const layout = computePaneLayout(current.projection, viewport.width, viewport.height);
    layoutRef.current = layout;
    // The scroll offset is a sheet offset; it never goes below the frozen
    // area, so a sheet with frozen panes starts at the first scrolling row.
    const scroll = clampScroll(layout, scrollRef.current.x, scrollRef.current.y);
    scrollRef.current = scroll;
    syncFilterHeaderColumns(current.projection, layout, scroll);
    positionFilterButtons(current.projection, layout, scroll);
    gridDebugState.layout = layout;
    gridDebugState.scrollX = scroll.x;
    gridDebugState.scrollY = scroll.y;
    const dpr = globalThis.devicePixelRatio || 1;
    frame.render({
      projection: current.projection,
      selection: current.selection,
      pointing: pointingRangeRef.current,
      painter: painterTargetRef.current,
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: viewport.width,
      height: viewport.height,
      dpr,
    });
    const session = editorSessionRef.current;
    if (session !== null) {
      positionEditor(editorRef.current, current.projection, layout, scroll, session);
    }
    positionMarquee(
      marqueeRef.current,
      current.projection,
      layout,
      scroll,
      marqueeRangeRef.current,
    );
    // The hover overlays are DOM elements too: the paint places them so they
    // keep tracking the cell while the sheet scrolls.
    positionHoverOverlay(
      commentPopupRef.current,
      current.projection,
      layout,
      scroll,
      hoverRef.current.comment,
      viewport,
      'right',
    );
    positionHoverOverlay(
      linkTooltipRef.current,
      current.projection,
      layout,
      scroll,
      hoverRef.current.link,
      viewport,
      'below',
    );
    updateScrollbarThumbs(
      verticalThumbRef.current,
      horizontalThumbRef.current,
      layout,
      scroll,
      viewport,
    );
  }, [viewport, syncFilterHeaderColumns, positionFilterButtons]);

  const schedulePaint = useCallback((): void => {
    if (pendingPaintRef.current) return;
    pendingPaintRef.current = true;
    requestAnimationFrame(() => {
      pendingPaintRef.current = false;
      paint();
    });
  }, [paint]);

  /**
   * Replace the hover overlays. A pointer moves many times inside one cell, so
   * an unchanged cell keeps the same state object and React never re-renders
   * for it. The paint loop reads the ref, so the placement follows scrolling.
   */
  const setHoverOverlays = useCallback(
    (comment: GridCellAddress | null, link: GridCellAddress | null): void => {
      const current = hoverRef.current;
      if (sameAddress(current.comment, comment) && sameAddress(current.link, link)) return;
      const next = { comment, link };
      hoverRef.current = next;
      setHover(next);
      schedulePaint();
    },
    [schedulePaint],
  );

  const clearHoverOverlays = useCallback((): void => {
    setHoverOverlays(null, null);
  }, [setHoverOverlays]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (canvas === null) return;
    const frame = new CanvasRenderFrame(canvas);
    frameRef.current = frame;
    return () => {
      frame.destroy();
      frameRef.current = null;
    };
  }, []);

  useLayoutEffect(() => {
    const container = containerRef.current;
    if (container === null) return;
    const update = (): void => {
      // The canvas fills the shell's content box, so the viewport is measured
      // the same way.
      setViewport({
        width: Math.max(200, Math.floor(container.clientWidth)),
        height: Math.max(160, Math.floor(container.clientHeight)),
      });
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(container);
    return () => observer.disconnect();
  }, []);

  useLayoutEffect(() => {
    stateRef.current = { projection, selection };
    schedulePaint();
  }, [projection, selection, schedulePaint]);

  // New filter buttons mount after the paint that named their columns; place
  // them once React has committed the nodes.
  useLayoutEffect(() => {
    const layout = layoutRef.current;
    if (layout === null) return;
    positionFilterButtons(projection, layout, scrollRef.current);
  }, [projection, filterHeaderColumns, viewport, positionFilterButtons]);

  // Any gesture that changes the selection — a click, a key, a jump or a sheet
  // switch — dismisses the hover overlays, as Excel does.
  useEffect(() => {
    clearHoverOverlays();
  }, [projection, selection, clearHoverOverlays]);

  // Leaving Format Painter mode drops the live target without a document
  // change. Esc is handled by the application, which flips this prop.
  useEffect(() => {
    if (painterActive === true) return;
    // Esc can end the mode while a drag is in flight; detach its window
    // listeners so the release cannot still apply the old target.
    painterCleanupRef.current?.();
    painterCleanupRef.current = null;
    painterDragRef.current = null;
    if (painterTargetRef.current === null) return;
    painterTargetRef.current = null;
    schedulePaint();
  }, [painterActive, schedulePaint]);

  useEffect(() => {
    return () => {
      autoScrollCleanupRef.current?.();
      autoScrollCleanupRef.current = null;
      painterCleanupRef.current?.();
      painterCleanupRef.current = null;
      fillCleanupRef.current?.();
      fillCleanupRef.current = null;
    };
  }, []);

  useLayoutEffect(() => {
    const frame = frameRef.current;
    if (frame === null) return;
    const dpr = globalThis.devicePixelRatio || 1;
    frame.resize(viewport.width, viewport.height, dpr);
    schedulePaint();
  }, [viewport, schedulePaint]);

  // Keep the selection's extension end visible when it moves or the zoom
  // changes: during a drag or a Shift extension that is the cell under the
  // pointer, and for a collapsed selection it is the active cell. The layout is
  // recomputed here because the paint may not have run with the new projection
  // yet.
  useEffect(() => {
    const layout = computePaneLayout(projection, viewport.width, viewport.height);
    const clamped = clampScroll(layout, scrollRef.current.x, scrollRef.current.y);
    const next = scrollToReveal(projection, layout, clamped.x, clamped.y, selection.focus);
    if (next.x !== scrollRef.current.x || next.y !== scrollRef.current.y) {
      scrollRef.current = next;
      schedulePaint();
    }
  }, [projection, selection, viewport, schedulePaint]);

  // The copy marquee follows the arrow keys; keep its far corner on screen so
  // the paste destination stays visible while it moves (ticket 31 B).
  const marqueeRange = marquee?.range ?? null;
  const marqueeEndRow = marqueeRange?.endRow ?? null;
  const marqueeEndCol = marqueeRange?.endCol ?? null;
  marqueeRangeRef.current = marqueeRange;
  useEffect(() => {
    if (marqueeEndRow === null || marqueeEndCol === null) return;
    const layout = computePaneLayout(projection, viewport.width, viewport.height);
    const clamped = clampScroll(layout, scrollRef.current.x, scrollRef.current.y);
    const next = scrollToReveal(projection, layout, clamped.x, clamped.y, {
      row: marqueeEndRow,
      col: marqueeEndCol,
    });
    if (next.x !== scrollRef.current.x || next.y !== scrollRef.current.y) {
      scrollRef.current = next;
      schedulePaint();
    }
  }, [projection, viewport, marqueeEndRow, marqueeEndCol, schedulePaint]);

  const editorId = editor?.id ?? null;

  // Focus belongs to the DOM commit, not the next paint: an input method that
  // starts a composition must find the textarea already focused.
  useLayoutEffect(() => {
    if (editorId === null) return;
    const element = editorRef.current;
    if (element === null) return;
    element.focus();
    // The caret belongs at the end of the text the editor started with.
    const length = element.value.length;
    element.setSelectionRange(length, length);
  }, [editorId]);

  // Pointing rewrites the formula text, so the caret follows the reference
  // that was just inserted or extended.
  useLayoutEffect(() => {
    const caret = pendingCaretRef.current;
    if (caret === null || editor === null) return;
    pendingCaretRef.current = null;
    const element = editorRef.current;
    if (element === null) return;
    element.focus();
    element.setSelectionRange(caret, caret);
  }, [editor]);

  // The wheel listener is registered natively so the event stays cancelable.
  useEffect(() => {
    const container = containerRef.current;
    if (container === null) return;
    const onWheel = (event: WheelEvent): void => {
      const layout = layoutRef.current;
      if (layout === null) return;
      // Pinch-to-zoom arrives as a ctrl+wheel event. Accumulate small
      // trackpad deltas so one gesture becomes a few zoom steps, not dozens.
      if ((event.ctrlKey || event.metaKey) && onZoomChange !== undefined) {
        event.preventDefault();
        zoomWheelRef.current += event.deltaY;
        if (Math.abs(zoomWheelRef.current) < 30) return;
        const direction = zoomWheelRef.current < 0 ? 1 : -1;
        zoomWheelRef.current = 0;
        const current = Math.round(stateRef.current.projection.zoom * 100);
        onZoomChange(current + direction * 10);
        return;
      }
      event.preventDefault();
      const deltaX = event.shiftKey ? event.deltaY : event.deltaX;
      const deltaY = event.shiftKey ? 0 : event.deltaY;
      const next = clampScroll(layout, scrollRef.current.x + deltaX, scrollRef.current.y + deltaY);
      scrollRef.current = next;
      schedulePaint();
    };
    container.addEventListener('wheel', onWheel, { passive: false });
    return () => container.removeEventListener('wheel', onWheel);
  }, [onZoomChange, schedulePaint]);

  const setShellRef = useCallback(
    (element: HTMLDivElement | null): void => {
      containerRef.current = element;
      if (shellRef !== undefined) shellRef.current = element;
    },
    [shellRef],
  );

  /**
   * The single place the editor session changes. The live text goes to the
   * formula bar on every keystroke and pointing insertion; `null` marks the end
   * of the edit so the bar returns to the committed cell text.
   */
  const updateEditorSession = (session: EditorSession | null): void => {
    editorSessionRef.current = session;
    setEditor(session);
    if (session === null) setEditorCaret(null);
    onEditorChange?.(session?.text ?? null);
  };

  const beginEditing = (address: GridCellAddress, text: string, commitIfChanged: boolean): void => {
    // Editing always happens on the anchor of a merged range.
    const target = mergeAnchor(stateRef.current.projection, address);
    // F2 and double-click edit the cell's content. A formula cell shows the
    // authored text without its leading `=` (as the sheet's own editor does);
    // the formula stays intact until the edit commits.
    const projection = stateRef.current.projection;
    const formula = commitIfChanged
      ? (projection.cellAt(target.row, target.col)?.formula ?? null)
      : null;
    const session = startEdit(target, formula ?? text, commitIfChanged, formula !== null);
    pointingRangeRef.current = null;
    pendingCaretRef.current = null;
    setEditorCaret(session.text.length);
    updateEditorSession(session);
  };

  const updateEditorText = (text: string, caret: number): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    // Typing ends the pointing gesture: the next reference starts at the caret
    // instead of replacing the reference the pointer inserted.
    const next = { ...session, text, pointing: null };
    pointingRangeRef.current = null;
    pendingCaretRef.current = null;
    setEditorCaret(caret);
    updateEditorSession(next);
    schedulePaint();
  };

  const setComposing = (composing: boolean): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    const next = { ...session, composing };
    updateEditorSession(next);
  };

  const clearPointing = (): void => {
    pointingRangeRef.current = null;
    pendingCaretRef.current = null;
    schedulePaint();
  };

  const commitEditor = (direction: EditorCommitDirection): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    updateEditorSession(null);
    clearPointing();
    if (editorHasChanges(session)) onCommit(session.address, editorCommitText(session));
    const target = moveAfterCommit(session.address, direction, projection);
    if (target !== null) onSelectionChange(collapseSelection(target));
    containerRef.current?.focus();
  };

  const cancelEditor = (): void => {
    updateEditorSession(null);
    clearPointing();
    containerRef.current?.focus();
  };

  /** The textarea selection, where a fresh pointing reference is inserted. */
  const editorCaretToken = (): { readonly start: number; readonly end: number } | null => {
    const element = editorRef.current;
    if (element === null) return null;
    return { start: element.selectionStart, end: element.selectionEnd };
  };

  /**
   * Return the keyboard to the in-cell editor. A pointing gesture inserts
   * text into the editor, so focus and the caret must survive the pointer
   * press and release instead of moving to the grid shell.
   */
  const restoreEditorFocus = (session: EditorSession | null): void => {
    const element = editorRef.current;
    if (element === null) return;
    element.focus();
    if (session !== null && session.pointing !== null) {
      element.setSelectionRange(session.pointing.end, session.pointing.end);
    }
  };

  /** Reveal a pointed cell the same way selection navigation does. */
  const revealPointedCell = (cell: GridCellAddress): void => {
    const layout = layoutRef.current;
    if (layout === null) return;
    scrollRef.current = scrollToReveal(
      stateRef.current.projection,
      layout,
      scrollRef.current.x,
      scrollRef.current.y,
      cell,
    );
  };

  /**
   * Insert or replace the reference text and leave the caret after it. The
   * editor keeps the keyboard focus; only the tinted pointing range moves.
   */
  const setPointing = (
    session: EditorSession,
    anchor: GridCellAddress,
    focus: GridCellAddress,
    token: { readonly start: number; readonly end: number } | null,
  ): EditorSession => {
    const reference = pointingReference(stateRef.current.projection, anchor, focus);
    const next = applyPointing(session, reference, anchor, focus, token);
    updateEditorSession(next);
    pointingRangeRef.current = pointingRange(anchor, focus);
    pendingCaretRef.current = next.pointing?.end ?? null;
    setEditorCaret(next.pointing?.end ?? null);
    revealPointedCell(focus);
    schedulePaint();
    return next;
  };

  /**
   * Insert a suggested function name at the caret token and leave the caret
   * after `(`. The session keeps its `formulaEntry` flag, so a formula shown
   * without `=` still commits with one.
   */
  const acceptEditorCompletion = (entry?: FunctionCatalogEntry): boolean => {
    const session = editorSessionRef.current;
    if (session === null) return false;
    const choice = entry === undefined ? editorCompletion.accept() : editorCompletion.choose(entry);
    if (choice === null) return false;
    const next: EditorSession = { ...session, text: choice.text, pointing: null };
    pointingRangeRef.current = null;
    pendingCaretRef.current = choice.caret;
    setEditorCaret(choice.caret);
    updateEditorSession(next);
    schedulePaint();
    return true;
  };

  /** The cell under a viewport pointer position, for a drag gesture. */
  const cellAtClientPoint = (clientX: number, clientY: number): GridCellAddress | null => {
    const canvas = canvasRef.current;
    const layout = layoutRef.current;
    if (canvas === null || layout === null) return null;
    const rect = canvas.getBoundingClientRect();
    return cellAtPoint(
      stateRef.current.projection,
      layout,
      clientX - rect.left,
      clientY - rect.top,
      scrollRef.current.x,
      scrollRef.current.y,
    );
  };

  /**
   * The cell under a drag point, including a point past the grid body's edge.
   * A drag that leaves the body still names a cell: the nearest cell along the
   * scrolling panes, so the selection can extend while edge scrolling reveals
   * more rows or columns. A point over a frozen pane keeps its own frozen cell;
   * a point outside the body targets the scrolling pane, because the frozen
   * strip never scrolls.
   */
  const dragCellAtClientPoint = (clientX: number, clientY: number): GridCellAddress | null => {
    const canvas = canvasRef.current;
    const layout = layoutRef.current;
    if (canvas === null || layout === null) return null;
    const rect = canvas.getBoundingClientRect();
    const pointX = clientX - rect.left;
    const pointY = clientY - rect.top;
    const bodyLeft = layout.headerWidth;
    const bodyTop = layout.headerHeight;
    const bodyRight = bodyLeft + layout.frozenWidth + layout.scrollRegionWidth;
    const bodyBottom = bodyTop + layout.frozenHeight + layout.scrollRegionHeight;
    const x =
      pointX < bodyLeft
        ? bodyLeft + layout.frozenWidth
        : Math.min(pointX, Math.max(bodyLeft, bodyRight - 1));
    const y =
      pointY < bodyTop
        ? bodyTop + layout.frozenHeight
        : Math.min(pointY, Math.max(bodyTop, bodyBottom - 1));
    return cellAtPoint(
      stateRef.current.projection,
      layout,
      x,
      y,
      scrollRef.current.x,
      scrollRef.current.y,
    );
  };

  /**
   * Start the edge auto-scroll loop for one pointer drag. The loop lives for
   * the whole gesture and runs whether or not the pointer moves: while its
   * last position is past a body edge, every animation frame scrolls toward
   * that edge and re-applies the focus cell, so the selection follows the
   * rows or columns the scroll reveals. Speed grows with the distance past
   * the edge, as Excel accelerates near the edge.
   *
   * The body edges include the frozen strips: a pointer over a frozen pane
   * names a frozen cell and must not scroll it. The scroll offset itself is
   * clamped by `clampScroll`, so the sheet stops at its first and last row or
   * column instead of wrapping.
   */
  const beginDragAutoScroll = (
    initialClientX: number,
    initialClientY: number,
    applyFocus: (clientX: number, clientY: number) => void,
  ): { readonly move: (clientX: number, clientY: number) => void; readonly stop: () => void } => {
    // A gesture that never received its release must not keep scrolling.
    autoScrollCleanupRef.current?.();
    let clientX = initialClientX;
    let clientY = initialClientY;
    let lastTime: number | null = null;
    let frame: number | null = null;

    const stop = (): void => {
      if (frame !== null) cancelAnimationFrame(frame);
      frame = null;
      if (autoScrollCleanupRef.current === stop) autoScrollCleanupRef.current = null;
    };

    const tick = (time: number): void => {
      frame = requestAnimationFrame(tick);
      const layout = layoutRef.current;
      const canvas = canvasRef.current;
      if (layout === null || canvas === null) return;
      const rect = canvas.getBoundingClientRect();
      const pointX = clientX - rect.left;
      const pointY = clientY - rect.top;
      const bodyLeft = layout.headerWidth;
      const bodyTop = layout.headerHeight;
      const bodyRight = bodyLeft + layout.frozenWidth + layout.scrollRegionWidth;
      const bodyBottom = bodyTop + layout.frozenHeight + layout.scrollRegionHeight;
      let velocityX = 0;
      let velocityY = 0;
      if (pointX < bodyLeft) velocityX = -autoScrollSpeed(bodyLeft - pointX);
      else if (pointX > bodyRight) velocityX = autoScrollSpeed(pointX - bodyRight);
      if (pointY < bodyTop) velocityY = -autoScrollSpeed(bodyTop - pointY);
      else if (pointY > bodyBottom) velocityY = autoScrollSpeed(pointY - bodyBottom);
      // Cap a long frame (a background tab or a slow paint) so a resume does
      // not jump the sheet by a second's worth of travel at once.
      const elapsed = lastTime === null ? 0 : Math.min(100, time - lastTime);
      lastTime = time;
      let scrolled = false;
      if (elapsed > 0 && (velocityX !== 0 || velocityY !== 0)) {
        const before = scrollRef.current;
        const next = clampScroll(
          layout,
          before.x + (velocityX * elapsed) / 1000,
          before.y + (velocityY * elapsed) / 1000,
        );
        if (next.x !== before.x || next.y !== before.y) {
          scrollRef.current = next;
          schedulePaint();
          scrolled = true;
        }
      }
      // Re-map the held pointer at the new offset: the cell under it changes
      // as the sheet moves, even when the pointer itself does not. A press
      // that has not moved and not scrolled has already named its own cell.
      if (scrolled) applyFocus(clientX, clientY);
    };

    frame = requestAnimationFrame(tick);
    autoScrollCleanupRef.current = stop;
    return {
      move: (nextX: number, nextY: number): void => {
        clientX = nextX;
        clientY = nextY;
        applyFocus(clientX, clientY);
      },
      stop,
    };
  };

  /**
   * Update the hover overlays for a pointer position: the comment popup appears
   * over the corner indicator (Excel's gesture) and the link tooltip over a
   * linked cell. Returns `true` when the pointer is over a hyperlink, so the
   * caller can show the hand cursor.
   */
  const hoverAtClientPoint = (clientX: number, clientY: number): boolean => {
    const canvas = canvasRef.current;
    const layout = layoutRef.current;
    if (canvas === null || layout === null) return false;
    const rect = canvas.getBoundingClientRect();
    const pointX = clientX - rect.left;
    const pointY = clientY - rect.top;
    const projection = stateRef.current.projection;
    if (pointX < layout.headerWidth || pointY < layout.headerHeight) {
      setHoverOverlays(null, null);
      return false;
    }
    const scroll = scrollRef.current;
    const address = cellAtPoint(projection, layout, pointX, pointY, scroll.x, scroll.y);
    if (address === null) {
      setHoverOverlays(null, null);
      return false;
    }
    if (projection.isRowHidden(address.row) || projection.isColumnHidden(address.col)) {
      setHoverOverlays(null, null);
      return false;
    }
    const cell = cellRect(projection, layout, scroll.x, scroll.y, address.row, address.col);
    if (
      projection.commentAt(address.row, address.col) !== null &&
      overCommentIndicator(pointX, pointY, cell, projection.zoom)
    ) {
      setHoverOverlays(address, null);
      return false;
    }
    const merge = projection.mergeAt(address.row, address.col);
    const link =
      projection.hyperlinkAt(address.row, address.col) ??
      (merge === null ? null : projection.hyperlinkAt(merge.anchorRow, merge.anchorCol));
    if (link === null) {
      setHoverOverlays(null, null);
      return false;
    }
    const tooltip = link.tooltip ?? '';
    setHoverOverlays(null, tooltip === '' ? null : address);
    return true;
  };

  /** The rectangle a painter gesture covers between its first and current cell. */
  const painterRange = (anchor: GridCellAddress, focus: GridCellAddress): GridRange => ({
    startRow: Math.min(anchor.row, focus.row),
    startCol: Math.min(anchor.col, focus.col),
    endRow: Math.max(anchor.row, focus.row),
    endCol: Math.max(anchor.col, focus.col),
  });

  const sameGridRange = (left: GridRange | null, right: GridRange | null): boolean =>
    left === right ||
    (left !== null &&
      right !== null &&
      left.startRow === right.startRow &&
      left.startCol === right.startCol &&
      left.endRow === right.endRow &&
      left.endCol === right.endCol);

  /**
   * Copy mode ends as soon as a pointer moves the selection: the ruling ends
   * the mode first and then performs the action (ticket 31 B).
   */
  const endCopyMode = (): void => {
    marquee?.onEnd();
  };

  /** `true` when the press landed on the painted fill handle. */
  const fillHandleAt = (pointX: number, pointY: number): boolean => {
    const handle = gridDebugState.fillHandle;
    if (handle === null) return false;
    return (
      pointX >= handle.x - FILL_HANDLE_HIT_MARGIN_PIXELS &&
      pointX <= handle.x + handle.width + FILL_HANDLE_HIT_MARGIN_PIXELS &&
      pointY >= handle.y - FILL_HANDLE_HIT_MARGIN_PIXELS &&
      pointY <= handle.y + handle.height + FILL_HANDLE_HIT_MARGIN_PIXELS
    );
  };

  /**
   * `true` when this point was a fill-handle press. The painted handle is the
   * first test, but a double-click fill moves the handle before the browser
   * delivers `dblclick`, so the recorded press point covers that gap. Anything
   * that presses the handle therefore consumes the gesture and the in-cell
   * editor never opens from it.
   */
  const fillPressedAt = (
    pointX: number,
    pointY: number,
    clientX: number,
    clientY: number,
  ): boolean => {
    if (fillHandleAt(pointX, pointY)) return true;
    const press = fillPressOriginRef.current;
    if (press === null) return false;
    if (Date.now() - press.at > FILL_HANDLE_DOUBLE_CLICK_MS) return false;
    // Client coordinates, not canvas coordinates: a transient layout shift
    // (the open banner, a status row) can move the canvas between the press and
    // the browser's `dblclick`, while the pointer itself has not moved.
    return (
      Math.abs(clientX - press.x) <= FILL_HANDLE_DOUBLE_CLICK_SLOP_PIXELS &&
      Math.abs(clientY - press.y) <= FILL_HANDLE_DOUBLE_CLICK_SLOP_PIXELS
    );
  };

  /** The rectangle a fill drag covers: the source extended to the hovered cell. */
  const fillTarget = (source: GridRange, focus: GridCellAddress): GridRange => ({
    startRow: Math.min(source.startRow, focus.row),
    startCol: Math.min(source.startCol, focus.col),
    endRow: Math.max(source.endRow, focus.row),
    endCol: Math.max(source.endCol, focus.col),
  });

  /**
   * Pressing the fill handle starts one fill gesture. The destination outline
   * grows with the pointer and the release submits `fill.handle` once, over the
   * whole rectangle. A press and release without movement is not a fill (the
   * source is its own destination), so it leaves the selection and the document
   * untouched, exactly like Excel.
   */
  const beginFillDrag = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    if (event.button !== 0) return;
    const source = selectionRange(stateRef.current.selection);
    const drag = {
      pointerId: event.pointerId,
      source,
      focus: { row: source.endRow, col: source.endCol },
    };
    fillDragRef.current = drag;
    setFillPreview(source);

    const applyFocus = (clientX: number, clientY: number): void => {
      const address = dragCellAtClientPoint(clientX, clientY);
      if (address === null) return;
      const focus = mergeAnchor(stateRef.current.projection, address);
      if (focus.row === drag.focus.row && focus.col === drag.focus.col) return;
      drag.focus = focus;
      setFillPreview(fillTarget(drag.source, focus));
    };
    const autoScroll = beginDragAutoScroll(event.clientX, event.clientY, applyFocus);
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      autoScroll.move(moveEvent.clientX, moveEvent.clientY);
    };
    const end = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId) return;
      stopTracking();
      fillDragRef.current = null;
      setFillPreview(null);
      const destination = fillTarget(drag.source, drag.focus);
      if (sameGridRange(destination, drag.source)) return;
      onFill?.(drag.source, destination);
    };
    const cancel = (cancelEvent: PointerEvent): void => {
      if (cancelEvent.pointerId !== drag.pointerId) return;
      abandon();
    };
    function abandon(): void {
      stopTracking();
      fillDragRef.current = null;
      fillCleanupRef.current = null;
      setFillPreview(null);
    }
    function onWindowBlur(): void {
      abandon();
    }
    function stopTracking(): void {
      autoScroll.stop();
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', cancel);
      window.removeEventListener('blur', onWindowBlur);
    }
    fillCleanupRef.current = stopTracking;
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
    window.addEventListener('pointercancel', cancel);
    window.addEventListener('blur', onWindowBlur);
  };

  /** Extend the ordinary selection from the press cell as the pointer moves. */
  const beginSelectionDrag = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    anchor: GridCellAddress,
    focus: GridCellAddress,
  ): void => {
    if (event.button !== 0) return;
    const drag = { pointerId: event.pointerId, anchor, focus };
    selectionDragRef.current = drag;
    const applyFocus = (clientX: number, clientY: number): void => {
      const address = dragCellAtClientPoint(clientX, clientY);
      if (address === null) return;
      const target = mergeAnchor(stateRef.current.projection, address);
      if (target.row === drag.focus.row && target.col === drag.focus.col) return;
      drag.focus = target;
      onSelectionChange(extendSelection(collapseSelection(drag.anchor), target));
    };
    const autoScroll = beginDragAutoScroll(event.clientX, event.clientY, applyFocus);
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      autoScroll.move(moveEvent.clientX, moveEvent.clientY);
    };
    const end = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId) return;
      if (selectionDragRef.current === drag) selectionDragRef.current = null;
      autoScroll.stop();
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', end);
      window.removeEventListener('blur', onWindowBlur);
    };
    function onWindowBlur(): void {
      // A window that loses focus may never deliver the release, and the loop
      // must not keep scrolling the sheet behind the user's back.
      if (selectionDragRef.current === drag) selectionDragRef.current = null;
      autoScroll.stop();
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', end);
      window.removeEventListener('blur', onWindowBlur);
    }
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
    window.addEventListener('pointercancel', end);
    window.addEventListener('blur', onWindowBlur);
  };

  /** Drag across a row or column header strip to select a contiguous run. */
  const beginHeaderSelectionDrag = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    axis: DimensionAxis,
    index: number,
    focus: GridCellAddress,
  ): void => {
    if (event.button !== 0) return;
    const drag = {
      pointerId: event.pointerId,
      axis,
      startIndex: index,
      focus,
      last: index,
    };
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      const canvas = canvasRef.current;
      const layout = layoutRef.current;
      if (canvas === null || layout === null) return;
      const rect = canvas.getBoundingClientRect();
      const pointX = moveEvent.clientX - rect.left;
      const pointY = moveEvent.clientY - rect.top;
      const projection = stateRef.current.projection;
      const target =
        axis === 'column'
          ? columnAtHeader(projection, layout, pointX, scrollRef.current.x)
          : rowAtHeader(projection, layout, pointY, scrollRef.current.y);
      if (target === null || target === drag.last) return;
      drag.last = target;
      if (axis === 'column') {
        onSelectionChange(
          selectionForRange(
            {
              startRow: 0,
              startCol: Math.min(drag.startIndex, target),
              endRow: projection.rowCount - 1,
              endCol: Math.max(drag.startIndex, target),
            },
            drag.focus,
          ),
        );
        return;
      }
      onSelectionChange(
        selectionForRange(
          {
            startRow: Math.min(drag.startIndex, target),
            startCol: 0,
            endRow: Math.max(drag.startIndex, target),
            endCol: projection.columnCount - 1,
          },
          drag.focus,
        ),
      );
    };
    const end = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId) return;
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', end);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
    window.addEventListener('pointercancel', end);
  };

  /** Drag a formula reference from the pressed cell to the cell under the pointer. */
  const beginFormulaDrag = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    anchor: GridCellAddress,
  ): void => {
    if (event.button !== 0) return;
    const drag = { pointerId: event.pointerId, anchor };
    formulaDragRef.current = drag;
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      const session = editorSessionRef.current;
      const pointing = session?.pointing ?? null;
      if (session === null || !isFormulaEntry(session) || pointing === null) return;
      const address = cellAtClientPoint(moveEvent.clientX, moveEvent.clientY);
      if (address === null) return;
      const target = mergeAnchor(stateRef.current.projection, address);
      if (target.row === pointing.focus.row && target.col === pointing.focus.col) return;
      setPointing(session, drag.anchor, target, {
        start: pointing.start,
        end: pointing.end,
      });
    };
    function end(upEvent: PointerEvent): void {
      if (upEvent.pointerId !== drag.pointerId) return;
      stopTracking();
      // The release ends the gesture, not the edit. A repaint or a focus
      // change during the press must not take the keyboard away from the
      // textarea that just received the reference.
      const current = editorSessionRef.current;
      if (current !== null && current.pointing !== null) restoreEditorFocus(current);
    }
    // A window that loses focus may never deliver the release. Forget the
    // gesture so a later click can still commit the edit.
    function cancel(): void {
      if (formulaDragRef.current === drag) formulaDragRef.current = null;
      stopTracking();
    }
    function stopTracking(): void {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', end);
      window.removeEventListener('blur', cancel);
    }
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
    window.addEventListener('pointercancel', end);
    window.addEventListener('blur', cancel);
  };

  /**
   * Drag a reference from the pressed cell into the formula bar draft. The bar
   * keeps the keyboard; only the selection and the reference text change.
   */
  const beginFormulaBarDrag = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    anchor: GridCellAddress,
    focus: GridCellAddress,
  ): void => {
    if (event.button !== 0) return;
    const drag = { pointerId: event.pointerId, anchor, focus };
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      const address = cellAtClientPoint(moveEvent.clientX, moveEvent.clientY);
      if (address === null) return;
      const target = mergeAnchor(stateRef.current.projection, address);
      if (target.row === drag.focus.row && target.col === drag.focus.col) return;
      drag.focus = target;
      onSelectionChange(extendSelection(collapseSelection(drag.anchor), target));
      formulaBarPointing?.onPoint(drag.anchor, target);
    };
    const end = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId) return;
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
      window.removeEventListener('pointercancel', end);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
    window.addEventListener('pointercancel', end);
  };

  /** The resize cursor belongs to the DOM canvas, not to a repaint. */
  const setCanvasCursor = (axis: DimensionAxis | null): void => {
    const canvas = canvasRef.current;
    if (canvas === null) return;
    canvas.style.cursor = axis === 'column' ? 'col-resize' : axis === 'row' ? 'row-resize' : '';
  };

  /**
   * A header-edge drag changes every selected visible row height or column
   * width to the same size. The live size is previewed through the application
   * so the projection paints it immediately; the release commits one undoable
   * `dimension.set` transaction for the whole selection.
   */
  const beginDimensionResize = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    axis: DimensionAxis,
    index: number,
  ): void => {
    if (event.button !== 0) return;
    const projection = stateRef.current.projection;
    const selected = headerSelectionIndices(projection, stateRef.current.selection, axis);
    let indices: readonly number[] = [index];
    let anchor = index;
    if (selected !== null && selected.length > 1) {
      // A leading edge belongs to the column or row before it; when that edge
      // touches the selection, the whole selection is the resize target. A
      // single header keeps the edge owner, as it always has.
      if (selected.includes(index)) {
        indices = selected;
        anchor = index;
      } else if (selected.includes(index + 1)) {
        indices = selected;
        anchor = selected[0] ?? index;
      }
    }
    const origin =
      axis === 'column' ? projection.columnWidth(anchor) : projection.rowHeight(anchor);
    if (!(origin > 0)) return;
    const drag = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      originPixels: origin / projection.zoom,
      zoom: projection.zoom,
      indices,
      moved: false,
      last: null as DimensionResize | null,
    };
    resizeDragRef.current = { pointerId: drag.pointerId, axis, index };
    setCanvasCursor(axis);

    const update = (clientX: number, clientY: number): void => {
      const delta = axis === 'column' ? clientX - drag.startX : clientY - drag.startY;
      // A double-click on the gutter must reach the autofit path, so small
      // pointer jitter is not a resize at all.
      if (!drag.moved && Math.abs(delta) < RESIZE_DRAG_THRESHOLD_PIXELS) return;
      drag.moved = true;
      drag.last = resizeDimension(axis, drag.originPixels, delta, drag.zoom);
      onResizePreview?.(axis, drag.indices, drag.last);
    };

    const finish = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId || resizeDragRef.current === null) return;
      update(upEvent.clientX, upEvent.clientY);
      stopTracking();
      resizeDragRef.current = null;
      setCanvasCursor(null);
      if (drag.moved && drag.last !== null) {
        if (onResizeCommit !== undefined) onResizeCommit(axis, drag.indices, drag.last);
        else onResizePreview?.(axis, drag.indices, null);
      } else {
        onResizePreview?.(axis, drag.indices, null);
      }
    };

    const cancel = (cancelEvent: PointerEvent): void => {
      if (cancelEvent.pointerId !== drag.pointerId) return;
      abandon();
    };

    function abandon(): void {
      if (resizeDragRef.current === null) return;
      stopTracking();
      resizeDragRef.current = null;
      setCanvasCursor(null);
      onResizePreview?.(axis, drag.indices, null);
    }

    function onWindowBlur(): void {
      abandon();
    }

    function stopTracking(): void {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', finish);
      window.removeEventListener('pointercancel', cancel);
      window.removeEventListener('blur', onWindowBlur);
    }

    function move(moveEvent: PointerEvent): void {
      if (moveEvent.pointerId !== drag.pointerId) return;
      update(moveEvent.clientX, moveEvent.clientY);
    }

    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', finish);
    window.addEventListener('pointercancel', cancel);
    window.addEventListener('blur', onWindowBlur);
  };

  const handlePointerMove = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    if (painterActive === true) {
      setCanvasCursor(null);
      if (painterDragRef.current !== null) return;
      const address = cellAtClientPoint(event.clientX, event.clientY);
      const next =
        address === null
          ? null
          : (() => {
              const target = mergeAnchor(stateRef.current.projection, address);
              return painterRange(target, target);
            })();
      if (sameGridRange(painterTargetRef.current, next)) return;
      painterTargetRef.current = next;
      schedulePaint();
      return;
    }
    if (resizeDragRef.current !== null) return;
    const layout = layoutRef.current;
    if (layout === null) return;
    // A drag owns the pointer: no hover overlay appears mid-gesture.
    const dragging =
      selectionDragRef.current !== null ||
      fillDragRef.current !== null ||
      formulaDragRef.current !== null ||
      painterDragRef.current !== null;
    const overLink = dragging ? false : hoverAtClientPoint(event.clientX, event.clientY);
    const rect = event.currentTarget.getBoundingClientRect();
    const target = headerResizeTarget(
      stateRef.current.projection,
      layout,
      event.clientX - rect.left,
      event.clientY - rect.top,
      scrollRef.current.x,
      scrollRef.current.y,
    );
    setCanvasCursor(target === null ? null : target.axis);
    if (target === null && overLink) event.currentTarget.style.cursor = 'pointer';
  };

  const handlePointerLeave = (): void => {
    // Leaving the canvas is a mouse-away dismissal, exactly like Excel's note.
    clearHoverOverlays();
    if (painterActive === true) {
      setCanvasCursor(null);
      if (painterDragRef.current !== null || painterTargetRef.current === null) return;
      painterTargetRef.current = null;
      schedulePaint();
      return;
    }
    if (resizeDragRef.current === null) setCanvasCursor(null);
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    const layout = layoutRef.current;
    if (layout === null) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const pointX = event.clientX - rect.left;
    const pointY = event.clientY - rect.top;
    const current = stateRef.current.projection;
    const scroll = scrollRef.current;
    const session = editorSessionRef.current;
    // A press supersedes the previous fill-handle record and closes any hover
    // overlay, whatever gesture follows.
    fillPressOriginRef.current = null;
    clearHoverOverlays();

    // The AutoFilter button owns its press: it opens the column dropdown and
    // never moves the selection or starts an edit (sort-autofilter §3.8). The
    // DOM overlay normally takes the press first; this is the canvas fallback.
    if (event.button === 0 && onFilterHeader !== undefined) {
      const filter = current.autoFilter;
      if (filter !== null) {
        const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
        if (
          address !== null &&
          address.row === filter.range.startRow &&
          address.col >= filter.range.startCol &&
          address.col <= filter.range.endCol &&
          !current.isColumnHidden(address.col)
        ) {
          const cell = cellRect(current, layout, scroll.x, scroll.y, address.row, address.col);
          if (pointInFilterButton(cell, current.zoom, pointX, pointY)) {
            event.preventDefault();
            endCopyMode();
            onFilterHeader(address.col);
            return;
          }
        }
      }
    }

    // Format Painter: a press starts one target range, a drag extends it and
    // the release submits a single apply command. It never changes the
    // document selection or opens the in-cell editor.
    if (painterActive === true && event.button === 0) {
      const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
      if (address === null) return;
      event.preventDefault();
      const anchor = mergeAnchor(current, address);
      const drag = { pointerId: event.pointerId, anchor, focus: anchor };
      painterDragRef.current = drag;
      painterTargetRef.current = painterRange(anchor, anchor);
      schedulePaint();

      const applyFocus = (clientX: number, clientY: number): void => {
        const nextAddress = dragCellAtClientPoint(clientX, clientY);
        if (nextAddress === null) return;
        const focus = mergeAnchor(stateRef.current.projection, nextAddress);
        if (focus.row === drag.focus.row && focus.col === drag.focus.col) return;
        drag.focus = focus;
        painterTargetRef.current = painterRange(drag.anchor, focus);
        schedulePaint();
      };
      const autoScroll = beginDragAutoScroll(event.clientX, event.clientY, applyFocus);
      const move = (moveEvent: PointerEvent): void => {
        if (moveEvent.pointerId !== drag.pointerId) return;
        autoScroll.move(moveEvent.clientX, moveEvent.clientY);
      };
      const end = (upEvent: PointerEvent): void => {
        if (upEvent.pointerId !== drag.pointerId) return;
        stopTracking();
        const range = painterRange(drag.anchor, drag.focus);
        painterDragRef.current = null;
        painterCleanupRef.current = null;
        painterTargetRef.current = null;
        schedulePaint();
        onPainterApply?.(range);
      };
      const cancel = (cancelEvent: PointerEvent): void => {
        if (cancelEvent.pointerId !== drag.pointerId) return;
        stopTracking();
        painterDragRef.current = null;
        painterCleanupRef.current = null;
        painterTargetRef.current = null;
        schedulePaint();
      };
      const stopTracking = (): void => {
        autoScroll.stop();
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', end);
        window.removeEventListener('pointercancel', cancel);
        window.removeEventListener('blur', onWindowBlur);
      };
      const onWindowBlur = (): void => {
        stopTracking();
        painterDragRef.current = null;
        painterCleanupRef.current = null;
        painterTargetRef.current = null;
        schedulePaint();
      };
      painterCleanupRef.current = stopTracking;
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', end);
      window.addEventListener('pointercancel', cancel);
      window.addEventListener('blur', onWindowBlur);
      return;
    }

    // A formula edit points at cells: a press inserts a reference and a drag
    // extends it. The textarea keeps the focus, so the edit is not committed.
    if (
      session !== null &&
      isFormulaEntry(session) &&
      event.button === 0 &&
      pointX >= layout.headerWidth &&
      pointY >= layout.headerHeight
    ) {
      const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
      if (address === null) return;
      event.preventDefault();
      const target = mergeAnchor(current, address);
      const anchor = event.shiftKey && session.pointing !== null ? session.pointing.anchor : target;
      const token = session.pointing === null ? editorCaretToken() : null;
      // Track the gesture before the state update, so a focus change during
      // the pointing render cannot commit the half-finished formula.
      beginFormulaDrag(event, anchor);
      setPointing(session, anchor, target, token);
      return;
    }

    // The formula bar owns a formula draft: a press inserts a reference into
    // that draft. The bar keeps the keyboard, so the press must not focus the
    // grid or start an in-cell edit.
    if (
      session === null &&
      formulaBarPointing !== null &&
      formulaBarPointing !== undefined &&
      event.button === 0 &&
      pointX >= layout.headerWidth &&
      pointY >= layout.headerHeight
    ) {
      const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
      if (address === null) return;
      event.preventDefault();
      const target = mergeAnchor(current, address);
      const anchor = event.shiftKey ? stateRef.current.selection.anchor : target;
      endCopyMode();
      onSelectionChange(extendSelection(collapseSelection(anchor), target));
      formulaBarPointing.onPoint(anchor, target);
      beginFormulaBarDrag(event, anchor, target);
      return;
    }

    containerRef.current?.focus();

    // Ctrl+click follows a hyperlink (feature matrix §8 "followable
    // (Ctrl+click)"). It never moves the selection: the cell is only a target.
    // A plain click falls through to the ordinary selection path below.
    if (
      (event.ctrlKey || event.metaKey) &&
      event.button === 0 &&
      onFollowHyperlink !== undefined &&
      pointX >= layout.headerWidth &&
      pointY >= layout.headerHeight
    ) {
      const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
      if (address !== null) {
        const merge = current.mergeAt(address.row, address.col);
        const anchor = merge === null ? address : { row: merge.anchorRow, col: merge.anchorCol };
        const link =
          current.hyperlinkAt(address.row, address.col) ??
          current.hyperlinkAt(anchor.row, anchor.col);
        if (link !== null) {
          event.preventDefault();
          onFollowHyperlink(link, address);
          return;
        }
      }
    }

    // The fill handle owns its press: it never moves the selection or opens
    // the editor. A second press inside the double-click window is Excel's
    // double-click fill down to the adjacent data edge; a drag fills the
    // rectangle the pointer covers. Either way, one command and one
    // transaction. The browser's own `dblclick` is ignored by the double-click
    // handler while the pointer is on the handle.
    if (event.button === 0 && fillHandleAt(pointX, pointY) && onFill !== undefined) {
      event.preventDefault();
      // A fill re-draws the selection, so copy mode ends before it (ticket 31 B).
      endCopyMode();
      const now = Date.now();
      fillPressOriginRef.current = { x: event.clientX, y: event.clientY, at: now };
      const doubleClick = now - fillPressRef.current <= FILL_HANDLE_DOUBLE_CLICK_MS;
      fillPressRef.current = doubleClick ? 0 : now;
      if (doubleClick) {
        const source = selectionRange(stateRef.current.selection);
        const last = adjacentFillRow(current, source);
        if (last !== null && last > source.endRow) {
          onFill(source, { ...source, endRow: last });
        }
        return;
      }
      beginFillDrag(event);
      return;
    }

    if (event.button === 0) {
      const resize = headerResizeTarget(current, layout, pointX, pointY, scroll.x, scroll.y);
      if (resize !== null) {
        beginDimensionResize(event, resize.axis, resize.index);
        return;
      }
    }

    if (pointY < layout.headerHeight && pointX >= layout.headerWidth) {
      const col = columnAtHeader(current, layout, pointX, scroll.x);
      if (col === null) return;
      if (event.button === 2) {
        if (!rightClickKeepsHeaderSelection(current, stateRef.current.selection, 'column', col)) {
          endCopyMode();
          onSelectionChange(
            columnHeaderSelection(current, col, stateRef.current.selection, scroll, false),
          );
        }
        return;
      }
      if (event.button !== 0) return;
      const next = columnHeaderSelection(
        current,
        col,
        stateRef.current.selection,
        scroll,
        event.shiftKey,
      );
      endCopyMode();
      onSelectionChange(next);
      if (!event.shiftKey) {
        beginHeaderSelectionDrag(event, 'column', col, { row: next.focus.row, col });
      }
      return;
    }
    if (pointX < layout.headerWidth && pointY >= layout.headerHeight) {
      const row = rowAtHeader(current, layout, pointY, scroll.y);
      if (row === null) return;
      if (event.button === 2) {
        if (!rightClickKeepsHeaderSelection(current, stateRef.current.selection, 'row', row)) {
          endCopyMode();
          onSelectionChange(
            rowHeaderSelection(current, row, stateRef.current.selection, scroll, false),
          );
        }
        return;
      }
      if (event.button !== 0) return;
      const next = rowHeaderSelection(
        current,
        row,
        stateRef.current.selection,
        scroll,
        event.shiftKey,
      );
      endCopyMode();
      onSelectionChange(next);
      if (!event.shiftKey) {
        beginHeaderSelectionDrag(event, 'row', row, { row, col: next.focus.col });
      }
      return;
    }

    const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
    if (address === null) return;
    const merge = current.mergeAt(address.row, address.col);
    const target = merge === null ? address : { row: merge.anchorRow, col: merge.anchorCol };
    // Excel: a right click inside the selection keeps it for the context
    // menu; a right click outside moves the selection to the clicked cell.
    // Only a left press starts or extends a selection drag.
    if (event.button === 2) {
      if (!selectionContains(stateRef.current.selection, target.row, target.col)) {
        endCopyMode();
        onSelectionChange(selectionForCell(current, address));
      }
      return;
    }
    if (event.button !== 0) return;
    endCopyMode();
    if (event.shiftKey) {
      const base = stateRef.current.selection;
      onSelectionChange(extendSelection(base, target));
      beginSelectionDrag(event, base.anchor, target);
      return;
    }
    onSelectionChange(selectionForCell(current, address));
    beginSelectionDrag(event, target, target);
  };

  const handleDoubleClick = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    // Format Painter applies on a single click; a second click must not open
    // the editor or apply twice.
    if (painterActive === true) return;
    // A double-click while pointing must not abandon the formula edit.
    if (editorSessionRef.current !== null) return;
    if (formulaBarPointing !== null && formulaBarPointing !== undefined) return;
    const layout = layoutRef.current;
    if (layout === null) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const pointX = event.clientX - rect.left;
    const pointY = event.clientY - rect.top;
    const current = stateRef.current.projection;

    // A fill-handle press consumes the whole gesture, so the browser's
    // `dblclick` must not also open the in-cell editor. The second press may
    // already have filled and moved the handle away from the point `dblclick`
    // reports, so the painted hit test is backed by the recorded press point.
    if (onFill !== undefined && fillPressedAt(pointX, pointY, event.clientX, event.clientY)) return;

    // A header double-click autofits every selected visible row or column to
    // its own content. One transaction stores the whole selection.
    if (onAutofit !== undefined) {
      const frame = frameRef.current;
      const measure = frame?.measureText?.bind(frame) ?? null;
      if (measure !== null && pointY < layout.headerHeight && pointX >= layout.headerWidth) {
        const col = columnAtHeader(current, layout, pointX, scrollRef.current.x);
        if (col === null) return;
        const selected = headerSelectionIndices(current, stateRef.current.selection, 'column');
        const indices =
          selected !== null && selected.length > 1 && selected.includes(col) ? selected : [col];
        onAutofit(
          'column',
          indices.map((index) => ({
            index,
            result: autofitColumnWidth(current, index, measure),
          })),
        );
        return;
      }
      if (measure !== null && pointX < layout.headerWidth && pointY >= layout.headerHeight) {
        const row = rowAtHeader(current, layout, pointY, scrollRef.current.y);
        if (row === null) return;
        const selected = headerSelectionIndices(current, stateRef.current.selection, 'row');
        const indices =
          selected !== null && selected.length > 1 && selected.includes(row) ? selected : [row];
        onAutofit(
          'row',
          indices.map((index) => ({
            index,
            result: autofitRowHeight(current, index, measure),
          })),
        );
        return;
      }
    }

    const address = cellAtPoint(
      current,
      layout,
      pointX,
      pointY,
      scrollRef.current.x,
      scrollRef.current.y,
    );
    if (address === null) return;
    const merge = current.mergeAt(address.row, address.col);
    const target = merge === null ? address : { row: merge.anchorRow, col: merge.anchorCol };
    beginEditing(target, current.rawTextAt(target.row, target.col), true);
  };

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (editorSessionRef.current !== null) return;
    // While an input method composes, the keyboard belongs to it: Enter
    // confirms a candidate and must not move the active cell.
    if (event.nativeEvent.isComposing) return;
    // Esc closes a hover overlay first; it makes no other change.
    if (
      event.key === 'Escape' &&
      (hoverRef.current.comment !== null || hoverRef.current.link !== null)
    ) {
      event.preventDefault();
      clearHoverOverlays();
      return;
    }
    const { key, ctrlKey, shiftKey, altKey, metaKey } = event;
    const current = stateRef.current;
    if ((ctrlKey || metaKey) && !shiftKey && !altKey) {
      const shortcut = key.toLowerCase();
      const clipboardAction =
        shortcut === 'c' ? onCopy : shortcut === 'x' ? onCut : shortcut === 'v' ? onPaste : null;
      if (clipboardAction !== undefined && clipboardAction !== null) {
        event.preventDefault();
        clipboardAction();
        return;
      }
    }
    // Copy mode owns the arrow keys, Enter and Escape (ticket 31 B): the
    // marquee moves without the selection, Enter pastes at it and Escape ends
    // the mode. Any other key ends the mode first and then performs its own
    // action below.
    if (marquee !== undefined) {
      if (
        !altKey &&
        (key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight')
      ) {
        event.preventDefault();
        const rowDelta = key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
        const colDelta = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
        marquee.onMove(rowDelta, colDelta);
        return;
      }
      if (key === 'Enter') {
        event.preventDefault();
        marquee.onPaste();
        return;
      }
      if (key === 'Escape') {
        event.preventDefault();
        marquee.onEnd();
        return;
      }
      // A key that only modifies another key does not end copy mode: its
      // companion decides, so Ctrl+V still pastes at the marquee.
      if (!MARQUEE_MODIFIER_KEYS.has(key)) marquee.onEnd();
    }
    const layout = layoutRef.current;
    const rowHeight =
      current.projection.rowHeight(0) || GRID_METRICS.rowHeight * current.projection.zoom;
    const pageRows =
      layout === null ? 20 : Math.max(1, Math.floor(layout.scrollRegionHeight / rowHeight));
    const pageCols =
      layout === null
        ? 10
        : Math.max(
            1,
            Math.floor(
              layout.scrollRegionWidth / (GRID_METRICS.columnWidth * current.projection.zoom),
            ),
          );
    const outcome = resolveNavigation({
      key,
      ctrlKey,
      shiftKey,
      altKey,
      // A plain key starts from the active cell (the anchor), so a selection
      // built by a Shift extension or a drag collapses and then moves from
      // where the gesture began. A Shift extension keeps extending from the
      // focus end (ticket 31 A).
      selection: shiftKey ? current.selection : collapseSelection(activeCell(current.selection)),
      projection: current.projection,
      pageRows,
      pageCols,
    });
    if (outcome.kind !== 'none') {
      event.preventDefault();
      onSelectionChange(outcome.selection);
      return;
    }
    if (isImeProcessKey(event)) {
      // An input method reports the key that starts a composition before the
      // composition begins. Mount and focus the editor in this event so the
      // composed text lands in the textarea instead of on the grid shell. The
      // default action is not prevented: it is what starts the composition.
      flushSync(() => beginEditing(activeCell(current.selection), '', false));
      return;
    }
    if (key === 'F2') {
      event.preventDefault();
      const active = activeCell(current.selection);
      beginEditing(active, current.projection.rawTextAt(active.row, active.col), true);
      return;
    }
    if (key === 'Delete' || key === 'Backspace') {
      event.preventDefault();
      onClear(selectionRange(current.selection));
      return;
    }
    if (key.length === 1 && !ctrlKey && !altKey && !metaKey && /\S/.test(key)) {
      event.preventDefault();
      beginEditing(activeCell(current.selection), key, false);
    }
  };

  const handleEditorKeyDown = (event: ReactKeyboardEvent<HTMLTextAreaElement>): void => {
    event.stopPropagation();
    const session = editorSessionRef.current;
    if (session === null) return;
    // Function suggestions own the navigation keys while their list is open.
    // Enter/Tab insert the highlighted name instead of committing the cell.
    if (!session.composing && !event.nativeEvent.isComposing && editorCompletion.listOpen) {
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        editorCompletion.move(1);
        return;
      }
      if (event.key === 'ArrowUp') {
        event.preventDefault();
        editorCompletion.move(-1);
        return;
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        editorCompletion.dismiss();
        return;
      }
      if (event.key === 'Tab' || (event.key === 'Enter' && !event.ctrlKey)) {
        if (acceptEditorCompletion()) {
          event.preventDefault();
          return;
        }
      }
    }
    const action = resolveEditorAction(session, {
      key: event.key,
      shiftKey: event.shiftKey,
      ctrlKey: event.ctrlKey,
      altKey: event.altKey,
      isComposing: event.nativeEvent.isComposing,
    });
    if (action.type === 'commit') {
      event.preventDefault();
      commitEditor(action.direction);
      return;
    }
    if (action.type === 'cancel') {
      event.preventDefault();
      cancelEditor();
      return;
    }
    if (session.composing || event.nativeEvent.isComposing) return;
    // Typed entry points at cells; F2 starts in edit mode and keeps the arrow
    // keys for moving the caret unless the caret sits where a reference
    // belongs. Once a point gesture starts, both modes move or extend it.
    if (!isFormulaEntry(session)) return;
    if (!POINTING_KEYS.has(event.key)) return;
    const caret = editorRef.current?.selectionStart ?? session.text.length;
    if (session.pointing === null && session.commitIfChanged && !isPointingBoundary(session, caret))
      return;

    const layout = layoutRef.current;
    const projection = stateRef.current.projection;
    const rowHeight = projection.rowHeight(0) || GRID_METRICS.rowHeight * projection.zoom;
    const pageRows =
      layout === null ? 20 : Math.max(1, Math.floor(layout.scrollRegionHeight / rowHeight));
    const pageCols =
      layout === null
        ? 10
        : Math.max(
            1,
            Math.floor(layout.scrollRegionWidth / (GRID_METRICS.columnWidth * projection.zoom)),
          );
    const pointing = session.pointing;
    const outcome = resolveNavigation({
      key: event.key,
      ctrlKey: event.ctrlKey,
      shiftKey: event.shiftKey,
      altKey: event.altKey,
      selection:
        pointing === null
          ? collapseSelection(session.address)
          : {
              anchor: pointing.anchor,
              focus: pointing.focus,
              extent: { kind: 'cells' },
            },
      projection,
      pageRows,
      pageCols,
    });
    if (outcome.kind === 'none') return;
    event.preventDefault();
    const token =
      pointing === null ? editorCaretToken() : { start: pointing.start, end: pointing.end };
    setPointing(session, outcome.selection.anchor, outcome.selection.focus, token);
  };

  const beginThumbDrag = (axis: 'x' | 'y', event: ReactPointerEvent<HTMLDivElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    dragRef.current = {
      axis,
      start: axis === 'x' ? event.clientX : event.clientY,
      origin: axis === 'x' ? scrollRef.current.x : scrollRef.current.y,
    };
    const move = (moveEvent: PointerEvent): void => {
      const drag = dragRef.current;
      const layout = layoutRef.current;
      if (drag === null || layout === null) return;
      const delta = (drag.axis === 'x' ? moveEvent.clientX : moveEvent.clientY) - drag.start;
      if (drag.axis === 'x') {
        const span = layout.maxScrollX - layout.minScrollX;
        const track = Math.max(1, viewport.width - layout.headerWidth - layout.scrollbarSize);
        const next = span <= 0 ? layout.minScrollX : drag.origin + (delta / track) * span;
        scrollRef.current = clampScroll(layout, next, scrollRef.current.y);
      } else {
        const span = layout.maxScrollY - layout.minScrollY;
        const track = Math.max(1, viewport.height - layout.headerHeight - layout.scrollbarSize);
        const next = span <= 0 ? layout.minScrollY : drag.origin + (delta / track) * span;
        scrollRef.current = clampScroll(layout, scrollRef.current.x, next);
      }
      schedulePaint();
    };
    const end = (): void => {
      dragRef.current = null;
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', end);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', end);
  };

  /**
   * A press on the thumb drags it. A press on the track above, below, before or
   * after the thumb pages the scrolling area by one viewport in that direction.
   */
  const handleScrollbarPointerDown =
    (axis: 'x' | 'y') =>
    (event: ReactPointerEvent<HTMLDivElement>): void => {
      const thumb = axis === 'x' ? horizontalThumbRef.current : verticalThumbRef.current;
      if (thumb !== null && event.target instanceof Node && thumb.contains(event.target)) {
        beginThumbDrag(axis, event);
        return;
      }
      const currentLayout = layoutRef.current;
      const rect = thumb === null ? undefined : thumb.getBoundingClientRect();
      let direction = 0;
      if (rect !== undefined) {
        if (axis === 'x') {
          if (event.clientX < rect.left) direction = -1;
          else if (event.clientX > rect.right) direction = 1;
        } else if (event.clientY < rect.top) {
          direction = -1;
        } else if (event.clientY > rect.bottom) {
          direction = 1;
        }
      }
      // A press level with the thumb, but on the track beside it, keeps the
      // previous grab behavior; only a press beyond the thumb pages.
      if (currentLayout === null || direction === 0 || event.button !== 0) {
        beginThumbDrag(axis, event);
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      const page =
        axis === 'x' ? currentLayout.scrollRegionWidth : currentLayout.scrollRegionHeight;
      const step = direction * page;
      scrollRef.current =
        axis === 'x'
          ? clampScroll(currentLayout, scrollRef.current.x + step, scrollRef.current.y)
          : clampScroll(currentLayout, scrollRef.current.x, scrollRef.current.y + step);
      schedulePaint();
    };

  const layout = computePaneLayout(projection, viewport.width, viewport.height);
  const editorMerge =
    editor === null ? null : projection.mergeAt(editor.address.row, editor.address.col);
  const editorRect =
    editor === null
      ? null
      : editorMerge === null
        ? cellRect(
            projection,
            layout,
            scrollRef.current.x,
            scrollRef.current.y,
            editor.address.row,
            editor.address.col,
          )
        : mergedRect(
            projection,
            layout,
            scrollRef.current.x,
            scrollRef.current.y,
            editorMerge.anchorRow,
            editorMerge.anchorCol,
            editorMerge.rowSpan,
            editorMerge.colSpan,
          );
  const editorStyle =
    editorRect === null
      ? undefined
      : {
          left: `${editorRect.x}px`,
          top: `${editorRect.y}px`,
          width: `${editorRect.width}px`,
          height: `${editorRect.height}px`,
          fontSize: `${Math.max(9, 13 * projection.zoom)}px`,
          padding: `0 ${Math.max(2, 4 * projection.zoom)}px`,
        };

  // The suggestion popup opens below the editor unless that would leave the
  // grid viewport, in which case it opens above.
  const suggestionHeight = 232;
  const suggestionTop =
    editorRect === null
      ? 0
      : editorRect.y + editorRect.height + 2 + suggestionHeight <= viewport.height
        ? editorRect.y + editorRect.height + 2
        : Math.max(2, editorRect.y - suggestionHeight - 2);
  const suggestionStyle =
    editorRect === null
      ? undefined
      : {
          position: 'absolute' as const,
          left: `${Math.max(2, Math.min(editorRect.x, viewport.width - 300))}px`,
          top: `${suggestionTop}px`,
          width: `${Math.max(280, Math.min(460, viewport.width - 16))}px`,
        };

  // The fill destination outline follows the pointer; it is view state and
  // never enters the model or the undo log.
  const fillPreviewRect =
    fillPreview === null
      ? null
      : rangeOverlayRect(projection, layout, scrollRef.current.x, scrollRef.current.y, fillPreview);

  // The copy marquee is the paste destination; like the fill preview it is
  // view state and never enters the model or the undo log (ticket 31 B). Its
  // border is clipped to the pane the range belongs to (ticket 46).
  const marqueeRect =
    marqueeRange === null
      ? null
      : rangeOverlayRect(
          projection,
          layout,
          scrollRef.current.x,
          scrollRef.current.y,
          marqueeRange,
        );
  const marqueeClipPath =
    marqueeRect === null || marqueeRange === null
      ? null
      : paneClipPath(marqueeRect, paneClipRect(projection, layout, marqueeRange, marqueeRect));

  // The comment popup and the link tooltip are anchored to their cells; the
  // paint re-places them every frame so they track scrolling.
  const hoveredComment =
    hover.comment === null ? null : projection.commentAt(hover.comment.row, hover.comment.col);
  const commentPopupRect =
    hover.comment === null
      ? null
      : cellRect(
          projection,
          layout,
          scrollRef.current.x,
          scrollRef.current.y,
          hover.comment.row,
          hover.comment.col,
        );
  const hoveredLink =
    hover.link === null ? null : projection.hyperlinkAt(hover.link.row, hover.link.col);
  const linkTooltipText = hoveredLink?.tooltip ?? '';
  const linkTooltipRect =
    hover.link === null || linkTooltipText === ''
      ? null
      : overlayAnchorRect(projection, layout, scrollRef.current, hover.link);

  const activeFilter = projection.autoFilter;

  return (
    <div
      ref={setShellRef}
      className="grid-shell"
      tabIndex={0}
      contentEditable
      suppressContentEditableWarning
      onBeforeInput={(event) => {
        // The shell is editable only so an input method can begin a
        // composition on the focused element. The composed text belongs in
        // the textarea, never in the shell's own content.
        if (event.target === event.currentTarget) event.preventDefault();
      }}
      onKeyDown={handleKeyDown}
      onFocus={() => onFocusZone?.()}
      onCompositionStart={() => {
        // Fallback for an input method that starts composing without first
        // reporting a process key: open the editor on the selected cell. Mount
        // and focus it before this event returns, so the composition text is
        // inserted into the textarea rather than the shell. The keydown path
        // handles the usual case and opens it before composing.
        if (editorSessionRef.current !== null) return;
        const active = activeCell(stateRef.current.selection);
        flushSync(() => beginEditing(active, '', false));
      }}
      data-testid="grid-shell"
    >
      <canvas
        ref={canvasRef}
        className={painterActive === true ? 'grid-canvas painter-active' : 'grid-canvas'}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerLeave={handlePointerLeave}
        onDoubleClick={handleDoubleClick}
        data-testid="grid-canvas"
      />
      {activeFilter !== null
        ? filterHeaderColumns.map((col) => {
            const label = projection.columnLabel(col);
            const filtered =
              activeFilter.filteredColumns.has(col) || activeFilter.unsupportedColumns.has(col);
            const name = filtered ? `Filter column ${label}, filtered` : `Filter column ${label}`;
            return (
              <button
                key={col}
                type="button"
                className="filter-header-button"
                data-testid={`filter-header-${label}`}
                data-filtered={filtered ? 'true' : 'false'}
                aria-label={name}
                aria-haspopup="dialog"
                title={name}
                ref={(element) => {
                  if (element === null) filterButtonRefs.current.delete(col);
                  else filterButtonRefs.current.set(col, element);
                }}
                onClick={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  endCopyMode();
                  onFilterHeader?.(col);
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    onFilterHeader?.(col);
                  } else if (event.key === 'Escape') {
                    event.preventDefault();
                    containerRef.current?.focus();
                  }
                }}
              />
            );
          })
        : null}
      <div
        className="grid-scrollbar grid-scrollbar-vertical"
        style={{
          top: `${layout.headerHeight}px`,
          right: 0,
          bottom: `${layout.scrollbarSize}px`,
          width: `${layout.scrollbarSize}px`,
        }}
        data-testid="grid-vscroll"
        onPointerDown={handleScrollbarPointerDown('y')}
      >
        <div
          className="grid-scrollbar-thumb"
          ref={verticalThumbRef}
          data-testid="grid-vscroll-thumb"
        />
      </div>
      <div
        className="grid-scrollbar grid-scrollbar-horizontal"
        style={{
          left: `${layout.headerWidth}px`,
          right: `${layout.scrollbarSize}px`,
          bottom: 0,
          height: `${layout.scrollbarSize}px`,
        }}
        data-testid="grid-hscroll"
        onPointerDown={handleScrollbarPointerDown('x')}
      >
        <div
          className="grid-scrollbar-thumb"
          ref={horizontalThumbRef}
          data-testid="grid-hscroll-thumb"
        />
      </div>
      {fillPreviewRect !== null ? (
        <div
          data-testid="fill-preview"
          style={{
            position: 'absolute',
            boxSizing: 'border-box',
            pointerEvents: 'none',
            border: '1px dashed var(--accent)',
            background: 'rgba(26, 115, 232, 0.08)',
            left: `${fillPreviewRect.x}px`,
            top: `${fillPreviewRect.y}px`,
            width: `${fillPreviewRect.width}px`,
            height: `${fillPreviewRect.height}px`,
          }}
        />
      ) : null}
      {marqueeRect !== null ? (
        <div
          ref={marqueeRef}
          className="copy-marquee"
          data-testid="copy-marquee"
          style={{
            position: 'absolute',
            boxSizing: 'border-box',
            pointerEvents: 'none',
            border: '2px dashed var(--accent)',
            left: `${marqueeRect.x}px`,
            top: `${marqueeRect.y}px`,
            width: `${marqueeRect.width}px`,
            height: `${marqueeRect.height}px`,
            clipPath: marqueeClipPath ?? undefined,
          }}
        />
      ) : null}
      {commentPopupRect !== null && hoveredComment !== null ? (
        <div
          ref={commentPopupRef}
          className="comment-popup"
          data-testid="comment-popup"
          role="note"
          style={{
            position: 'absolute',
            left: `${commentPopupRect.x + commentPopupRect.width + 6}px`,
            top: `${commentPopupRect.y}px`,
          }}
        >
          <div className="comment-popup-author" data-testid="comment-popup-author">
            {hoveredComment.author ?? 'Unknown author'}
          </div>
          <div className="comment-popup-text" data-testid="comment-popup-text">
            {hoveredComment.text === '' ? '(no text)' : hoveredComment.text}
          </div>
          {hoveredComment.threaded ? (
            <div className="comment-popup-badge" data-testid="comment-popup-badge">
              Threaded comment
            </div>
          ) : null}
        </div>
      ) : null}
      {linkTooltipRect !== null && linkTooltipText !== '' ? (
        <div
          ref={linkTooltipRef}
          className="link-tooltip"
          data-testid="link-tooltip"
          role="tooltip"
          style={{
            position: 'absolute',
            left: `${linkTooltipRect.x}px`,
            top: `${linkTooltipRect.y + linkTooltipRect.height + 4}px`,
          }}
        >
          {linkTooltipText}
        </div>
      ) : null}
      {editor !== null ? (
        <textarea
          ref={editorRef}
          className="cell-editor"
          style={editorStyle}
          value={editor.text}
          onChange={(event) => updateEditorText(event.target.value, event.target.selectionStart)}
          onCompositionStart={() => setComposing(true)}
          onCompositionEnd={() => setComposing(false)}
          onKeyDown={handleEditorKeyDown}
          onSelect={(event) => setEditorCaret(event.currentTarget.selectionStart)}
          onClick={(event) => setEditorCaret(event.currentTarget.selectionStart)}
          onKeyUp={(event) => setEditorCaret(event.currentTarget.selectionStart)}
          onBlur={() => {
            const session = editorSessionRef.current;
            if (session === null) return;
            // A pointer gesture is still inserting or extending a reference.
            // If the press moved focus before the release, take it back
            // instead of committing a half-finished formula.
            if (formulaDragRef.current !== null && session.pointing !== null) {
              restoreEditorFocus(session);
              return;
            }
            commitEditor('none');
          }}
          data-testid="cell-editor"
        />
      ) : null}
      {editorCompletion.completion !== null && suggestionStyle !== undefined ? (
        <FormulaSuggestions
          id="cell-formula-suggestions"
          completion={editorCompletion.completion}
          activeIndex={editorCompletion.activeIndex}
          activeEntry={editorCompletion.activeEntry}
          style={suggestionStyle}
          onChoose={(entry) => acceptEditorCompletion(entry)}
        />
      ) : null}
    </div>
  );
}

/** Movement below this is a click, so a header double-click can autofit. */
const RESIZE_DRAG_THRESHOLD_PIXELS = 3;

/**
 * Edge auto-scroll speed in CSS pixels per second. A drag at the edge starts at
 * the minimum speed; it accelerates toward the maximum over the first
 * `AUTOSCROLL_ACCELERATION_DISTANCE` pixels past the edge, so a pointer held
 * just outside moves gently and one far outside races. Excel has the same
 * near-edge acceleration.
 */
const AUTOSCROLL_MIN_SPEED_PIXELS_PER_SECOND = 120;
const AUTOSCROLL_MAX_SPEED_PIXELS_PER_SECOND = 1800;
const AUTOSCROLL_ACCELERATION_DISTANCE_PIXELS = 160;

function autoScrollSpeed(distancePastEdge: number): number {
  const ratio = Math.min(
    1,
    Math.max(0, distancePastEdge) / AUTOSCROLL_ACCELERATION_DISTANCE_PIXELS,
  );
  return (
    AUTOSCROLL_MIN_SPEED_PIXELS_PER_SECOND +
    (AUTOSCROLL_MAX_SPEED_PIXELS_PER_SECOND - AUTOSCROLL_MIN_SPEED_PIXELS_PER_SECOND) * ratio
  );
}

/** The fill handle is six pixels wide; a couple of pixels of slack help the press. */
const FILL_HANDLE_HIT_MARGIN_PIXELS = 2;

/** Two presses on the handle inside this window are a double-click fill. */
const FILL_HANDLE_DOUBLE_CLICK_MS = 500;

/** How far the browser's `dblclick` point may sit from the recorded press. */
const FILL_HANDLE_DOUBLE_CLICK_SLOP_PIXELS = 6;

/**
 * The last row a double-click on the fill handle fills to. Excel measures the
 * contiguous run of data immediately below the selection in the **adjacent**
 * column — the column to the left of the selection, or the column to its right
 * when the selection starts in column A — because that column holds the length
 * the fill should match. The source's own leftmost column is the fallback when
 * neither neighbour has a run. `null` when there is no run below the source, in
 * which case Excel does nothing.
 */
function adjacentFillRow(projection: SheetProjection, source: GridRange): number | null {
  for (const col of [source.startCol - 1, source.endCol + 1, source.startCol]) {
    if (col < 0 || col >= projection.columnCount) continue;
    const last = contiguousRunEnd(projection, col, source.endRow + 1);
    if (last !== null) return last;
  }
  return null;
}

/** The last row of the unbroken run of data starting at `startRow`. */
function contiguousRunEnd(
  projection: SheetProjection,
  col: number,
  startRow: number,
): number | null {
  let last: number | null = null;
  for (let row = startRow; row < projection.rowCount; row += 1) {
    const cell = projection.cellAt(row, col);
    const holdsData = cell !== null && (cell.formula !== null || cell.value.type !== 'blank');
    if (!holdsData) break;
    last = row;
  }
  return last;
}

/** A range's rectangle in canvas coordinates, for the live fill outline. */
function rangeOverlayRect(
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
  range: GridRange,
): { readonly x: number; readonly y: number; readonly width: number; readonly height: number } {
  const topLeft = cellRect(projection, layout, scrollX, scrollY, range.startRow, range.startCol);
  const bottomRight = cellRect(projection, layout, scrollX, scrollY, range.endRow, range.endCol);
  return {
    x: topLeft.x,
    y: topLeft.y,
    width: bottomRight.x + bottomRight.width - topLeft.x,
    height: bottomRight.y + bottomRight.height - topLeft.y,
  };
}

/**
 * The part of an overlay rectangle its pane can show (ticket 46).
 *
 * The overlay element keeps its whole rectangle — that is the range it
 * addresses — and only the painted edge is clipped. A range that starts in the
 * scrolling rows belongs to the scrolling pane, so its overlay stops at the
 * freeze divider; a range that starts in the frozen rows owns the frozen band.
 * A range is contiguous, so the panes it touches always form one rectangle.
 */
function paneClipRect(
  projection: SheetProjection,
  layout: PaneLayout,
  range: GridRange,
  rect: CellRect,
): CellRect {
  const left =
    range.startCol >= projection.freezeCols
      ? layout.headerWidth + layout.frozenWidth
      : layout.headerWidth;
  const top =
    range.startRow >= projection.freezeRows
      ? layout.headerHeight + layout.frozenHeight
      : layout.headerHeight;
  const right = layout.headerWidth + layout.frozenWidth + layout.scrollRegionWidth;
  const bottom = layout.headerHeight + layout.frozenHeight + layout.scrollRegionHeight;
  const x = Math.max(rect.x, left);
  const y = Math.max(rect.y, top);
  return {
    x,
    y,
    width: Math.max(0, Math.min(rect.x + rect.width, right) - x),
    height: Math.max(0, Math.min(rect.y + rect.height, bottom) - y),
  };
}

/** The `clip-path` inset that cuts an overlay rectangle to its pane's part. */
function paneClipPath(rect: CellRect, clip: CellRect): string {
  const top = Math.max(0, clip.y - rect.y);
  const left = Math.max(0, clip.x - rect.x);
  const right = Math.max(0, rect.x + rect.width - (clip.x + clip.width));
  const bottom = Math.max(0, rect.y + rect.height - (clip.y + clip.height));
  return `inset(${top}px ${right}px ${bottom}px ${left}px)`;
}

const POINTING_KEYS: ReadonlySet<string> = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'Home',
  'End',
  'PageUp',
  'PageDown',
]);

/**
 * Keys that only modify or lock another key. Pressing one on its own must not
 * end copy mode: the key they modify (Ctrl+V, for example) still belongs to the
 * marquee.
 */
const MARQUEE_MODIFIER_KEYS: ReadonlySet<string> = new Set([
  'Shift',
  'Control',
  'Alt',
  'Meta',
  'CapsLock',
  'NumLock',
  'ScrollLock',
  'OS',
  'AltGraph',
]);

/**
 * Windows and Linux input methods report the keys they process with keyCode
 * 229; the first such key starts a composition. The key is usually "Process"
 * and can be the composed character itself when the input method commits
 * directly.
 */
function isImeProcessKey(event: ReactKeyboardEvent<HTMLDivElement>): boolean {
  return event.key === 'Process' || event.keyCode === 229;
}

function moveAfterCommit(
  address: GridCellAddress,
  direction: EditorCommitDirection,
  projection: SheetProjection,
): GridCellAddress | null {
  switch (direction) {
    case 'down':
      return {
        row: Math.min(projection.rowCount - 1, address.row + 1),
        col: address.col,
      };
    case 'up':
      return { row: Math.max(0, address.row - 1), col: address.col };
    case 'right':
      return {
        row: address.row,
        col: Math.min(projection.columnCount - 1, address.col + 1),
      };
    case 'left':
      return { row: address.row, col: Math.max(0, address.col - 1) };
    default:
      return null;
  }
}

function positionEditor(
  element: HTMLTextAreaElement | null,
  projection: SheetProjection,
  layout: PaneLayout,
  scroll: ScrollOffset,
  session: EditorSession,
): void {
  if (element === null) return;
  const merge = projection.mergeAt(session.address.row, session.address.col);
  // The editor covers the whole merged range, as Excel's does.
  const rect =
    merge === null
      ? cellRect(projection, layout, scroll.x, scroll.y, session.address.row, session.address.col)
      : mergedRect(
          projection,
          layout,
          scroll.x,
          scroll.y,
          merge.anchorRow,
          merge.anchorCol,
          merge.rowSpan,
          merge.colSpan,
        );
  element.style.left = `${rect.x}px`;
  element.style.top = `${rect.y}px`;
  element.style.width = `${rect.width}px`;
  element.style.height = `${rect.height}px`;
}

/**
 * Keep the copy marquee over its range while the sheet scrolls. The overlay is
 * a DOM element, so the paint has to place it exactly as it places the editor
 * and the scrollbar thumbs.
 */
function positionMarquee(
  element: HTMLDivElement | null,
  projection: SheetProjection,
  layout: PaneLayout,
  scroll: ScrollOffset,
  range: GridRange | null,
): void {
  if (element === null || range === null) return;
  const rect = rangeOverlayRect(projection, layout, scroll.x, scroll.y, range);
  element.style.left = `${rect.x}px`;
  element.style.top = `${rect.y}px`;
  element.style.width = `${rect.width}px`;
  element.style.height = `${rect.height}px`;
  // Keep the border out of the frozen strips (ticket 46). The element keeps
  // its full rectangle, so the marquee's geometry and the paste target do not
  // move; only the painted edge is cut at the divider.
  element.style.clipPath = paneClipPath(rect, paneClipRect(projection, layout, range, rect));
}

/** The rectangle a hover overlay hangs from; a link uses its merged range. */
function overlayAnchorRect(
  projection: SheetProjection,
  layout: PaneLayout,
  scroll: ScrollOffset,
  address: GridCellAddress,
): CellRect {
  const merge = projection.mergeAt(address.row, address.col);
  if (merge === null) {
    return cellRect(projection, layout, scroll.x, scroll.y, address.row, address.col);
  }
  return mergedRect(
    projection,
    layout,
    scroll.x,
    scroll.y,
    merge.anchorRow,
    merge.anchorCol,
    merge.rowSpan,
    merge.colSpan,
  );
}

/**
 * Place a hover overlay next to its cell and keep it inside the viewport. The
 * comment popup sits to the right of the cell, the link tooltip below it; both
 * flip to the other side when the viewport runs out of room, then clamp.
 */
function positionHoverOverlay(
  element: HTMLDivElement | null,
  projection: SheetProjection,
  layout: PaneLayout,
  scroll: ScrollOffset,
  address: GridCellAddress | null,
  viewport: Viewport,
  placement: 'right' | 'below',
): void {
  if (element === null || address === null) return;
  const rect = overlayAnchorRect(projection, layout, scroll, address);
  const width = element.offsetWidth || 240;
  const height = element.offsetHeight || 60;
  const gap = 6;
  let left = placement === 'right' ? rect.x + rect.width + gap : rect.x;
  let top = placement === 'right' ? rect.y : rect.y + rect.height + 4;
  if (left + width > viewport.width - 4) {
    left = placement === 'right' ? rect.x - width - gap : viewport.width - width - 4;
  }
  left = Math.max(4, Math.min(left, Math.max(4, viewport.width - width - 4)));
  top = Math.max(4, Math.min(top, Math.max(4, viewport.height - height - 4)));
  element.style.left = `${Math.round(left)}px`;
  element.style.top = `${Math.round(top)}px`;
}

/** Whether a canvas point sits on a comment's corner indicator triangle. */
function overCommentIndicator(
  pointX: number,
  pointY: number,
  rect: CellRect,
  zoom: number,
): boolean {
  // Two pixels of slack, matching the painted triangle plus the pointer's aim.
  const size = Math.max(3, Math.min(8, Math.round(5 * zoom))) + 2;
  const right = rect.x + rect.width;
  return pointX >= right - size && pointX <= right && pointY >= rect.y && pointY <= rect.y + size;
}

/**
 * The selection a press on a cell makes. A merged cell is one cell, so a press
 * anywhere in it selects the whole merged range with its top-left active; a
 * plain cell collapses to itself (ticket 38).
 */
function selectionForCell(projection: SheetProjection, cell: GridCellAddress): GridSelection {
  const merge = projection.mergeAt(cell.row, cell.col);
  if (merge === null) return collapseSelection(cell);
  const range: GridRange = {
    startRow: merge.anchorRow,
    startCol: merge.anchorCol,
    endRow: merge.anchorRow + merge.rowSpan - 1,
    endCol: merge.anchorCol + merge.colSpan - 1,
  };
  return selectionForRange(range, { row: range.startRow, col: range.startCol });
}

function sameAddress(left: GridCellAddress | null, right: GridCellAddress | null): boolean {
  if (left === right) return true;
  if (left === null || right === null) return false;
  return left.row === right.row && left.col === right.col;
}

/**
 * A right click on a header keeps a header selection that already covers the
 * clicked axis, so the context menu acts on the whole selection (Excel's
 * context-menu behaviour; ticket 29).
 */
function rightClickKeepsHeaderSelection(
  projection: SheetProjection,
  selection: GridSelection,
  axis: 'row' | 'column',
  index: number,
): boolean {
  const current = headerSelectionRange(projection, selection, axis);
  if (current === null) return false;
  return axis === 'column'
    ? index >= current.startCol && index <= current.endCol
    : index >= current.startRow && index <= current.endRow;
}

/**
 * A column header click selects the whole column. Shift extends the header
 * selection to a contiguous run. A click inside an existing multi-column
 * selection keeps it, so a double-click can autofit every selected column. The
 * active cell stays in the visible part of the sheet, so the grid does not
 * scroll to the far edge.
 */
function columnHeaderSelection(
  projection: SheetProjection,
  col: number,
  selection: GridSelection,
  scroll: ScrollOffset,
  extend: boolean,
): GridSelection {
  const active = activeCell(selection);
  const current = headerSelectionRange(projection, selection, 'column');
  if (extend && current !== null) {
    if (col >= current.startCol && col <= current.endCol) {
      return selectionForRange(current, active);
    }
    return selectionForRange(
      {
        startRow: 0,
        startCol: Math.min(current.startCol, col),
        endRow: projection.rowCount - 1,
        endCol: Math.max(current.endCol, col),
      },
      active,
    );
  }
  if (current !== null) {
    const visible = headerSelectionIndices(projection, selection, 'column');
    if (visible !== null && visible.length > 1 && visible.includes(col)) return selection;
  }
  const firstVisibleRow = Math.max(projection.freezeRows, projection.rowAt(scroll.y));
  const focus = active.col === col ? active : { row: firstVisibleRow, col };
  return selectionForRange(
    {
      startRow: 0,
      startCol: col,
      endRow: projection.rowCount - 1,
      endCol: col,
    },
    focus,
  );
}

/**
 * A row header click selects the whole row, with Shift extending a contiguous
 * run and a click inside an existing multi-row selection keeping it for an
 * autofit double-click. The active cell stays in the visible part of the
 * sheet.
 */
function rowHeaderSelection(
  projection: SheetProjection,
  row: number,
  selection: GridSelection,
  scroll: ScrollOffset,
  extend: boolean,
): GridSelection {
  const active = activeCell(selection);
  const current = headerSelectionRange(projection, selection, 'row');
  if (extend && current !== null) {
    if (row >= current.startRow && row <= current.endRow) {
      return selectionForRange(current, active);
    }
    return selectionForRange(
      {
        startRow: Math.min(current.startRow, row),
        startCol: 0,
        endRow: Math.max(current.endRow, row),
        endCol: projection.columnCount - 1,
      },
      active,
    );
  }
  if (current !== null) {
    const visible = headerSelectionIndices(projection, selection, 'row');
    if (visible !== null && visible.length > 1 && visible.includes(row)) return selection;
  }
  const firstVisibleCol = Math.max(projection.freezeCols, projection.columnAt(scroll.x));
  const focus = active.row === row ? active : { row, col: firstVisibleCol };
  return selectionForRange(
    {
      startRow: row,
      startCol: 0,
      endRow: row,
      endCol: projection.columnCount - 1,
    },
    focus,
  );
}

function updateScrollbarThumbs(
  vertical: HTMLDivElement | null,
  horizontal: HTMLDivElement | null,
  layout: PaneLayout,
  scroll: ScrollOffset,
  viewport: Viewport,
): void {
  if (vertical !== null) {
    const trackTop = layout.headerHeight;
    const trackHeight = Math.max(1, viewport.height - layout.headerHeight - layout.scrollbarSize);
    const ratio = Math.min(1, layout.scrollRegionHeight / Math.max(1, layout.contentHeight));
    const thumbHeight = Math.max(48, Math.round(trackHeight * ratio));
    const span = Math.max(1, layout.maxScrollY - layout.minScrollY);
    const progress = (scroll.y - layout.minScrollY) / span;
    const top = trackTop + progress * Math.max(0, trackHeight - thumbHeight);
    vertical.style.height = `${thumbHeight}px`;
    vertical.style.transform = `translateY(${Math.round(top)}px)`;
  }
  if (horizontal !== null) {
    const trackLeft = layout.headerWidth;
    const trackWidth = Math.max(1, viewport.width - layout.headerWidth - layout.scrollbarSize);
    const ratio = Math.min(1, layout.scrollRegionWidth / Math.max(1, layout.contentWidth));
    const thumbWidth = Math.max(48, Math.round(trackWidth * ratio));
    const span = Math.max(1, layout.maxScrollX - layout.minScrollX);
    const progress = (scroll.x - layout.minScrollX) / span;
    const left = trackLeft + progress * Math.max(0, trackWidth - thumbWidth);
    horizontal.style.width = `${thumbWidth}px`;
    horizontal.style.transform = `translateX(${Math.round(left)}px)`;
  }
}
