import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type {
  JSX,
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react';
import { flushSync } from 'react-dom';
import type { SheetProjection } from '@view/types';
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
import type { GridCellAddress, GridRange, GridSelection, RenderFrame } from './frame';
import { mergeAnchor, resolveNavigation } from './navigation';
import {
  cellRect,
  clampScroll,
  columnAtHeader,
  computePaneLayout,
  mergedRect,
  rowAtHeader,
  scrollToReveal,
  type PaneLayout,
} from './panes';
import { collapseSelection, extendSelection, selectionForRange } from './selection';

export interface CanvasGridProps {
  readonly projection: SheetProjection;
  readonly selection: GridSelection;
  readonly onSelectionChange: (selection: GridSelection) => void;
  readonly onCommit: (address: GridCellAddress, text: string) => void;
  readonly onClear: (address: GridCellAddress) => void;
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
  onEditorChange,
  formulaBarPointing,
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
  const selectionDragRef = useRef<{
    readonly pointerId: number;
    readonly anchor: GridCellAddress;
    focus: GridCellAddress;
  } | null>(null);
  const formulaDragRef = useRef<{
    readonly pointerId: number;
    readonly anchor: GridCellAddress;
  } | null>(null);
  const pointingRangeRef = useRef<GridRange | null>(null);
  const pendingCaretRef = useRef<number | null>(null);
  const [viewport, setViewport] = useState<Viewport>({ width: 800, height: 480 });
  const [editor, setEditor] = useState<EditorSession | null>(null);

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
    gridDebugState.layout = layout;
    gridDebugState.scrollX = scroll.x;
    gridDebugState.scrollY = scroll.y;
    const dpr = globalThis.devicePixelRatio || 1;
    frame.render({
      projection: current.projection,
      selection: current.selection,
      pointing: pointingRangeRef.current,
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
    updateScrollbarThumbs(
      verticalThumbRef.current,
      horizontalThumbRef.current,
      layout,
      scroll,
      viewport,
    );
  }, [viewport]);

  const schedulePaint = useCallback((): void => {
    if (pendingPaintRef.current) return;
    pendingPaintRef.current = true;
    requestAnimationFrame(() => {
      pendingPaintRef.current = false;
      paint();
    });
  }, [paint]);

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

  useLayoutEffect(() => {
    const frame = frameRef.current;
    if (frame === null) return;
    const dpr = globalThis.devicePixelRatio || 1;
    frame.resize(viewport.width, viewport.height, dpr);
    schedulePaint();
  }, [viewport, schedulePaint]);

  // Keep the active cell visible when the selection moves.
  useEffect(() => {
    const layout = layoutRef.current;
    if (layout === null) return;
    const next = scrollToReveal(
      projection,
      layout,
      scrollRef.current.x,
      scrollRef.current.y,
      selection.focus,
    );
    if (next.x !== scrollRef.current.x || next.y !== scrollRef.current.y) {
      scrollRef.current = next;
      schedulePaint();
    }
  }, [projection, selection, schedulePaint]);

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
      event.preventDefault();
      const deltaX = event.shiftKey ? event.deltaY : event.deltaX;
      const deltaY = event.shiftKey ? 0 : event.deltaY;
      const next = clampScroll(layout, scrollRef.current.x + deltaX, scrollRef.current.y + deltaY);
      scrollRef.current = next;
      schedulePaint();
    };
    container.addEventListener('wheel', onWheel, { passive: false });
    return () => container.removeEventListener('wheel', onWheel);
  }, [schedulePaint]);

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
    updateEditorSession(session);
  };

  const updateEditorText = (text: string): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    // Typing ends the pointing gesture: the next reference starts at the caret
    // instead of replacing the reference the pointer inserted.
    const next = { ...session, text, pointing: null };
    pointingRangeRef.current = null;
    pendingCaretRef.current = null;
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
    revealPointedCell(focus);
    schedulePaint();
    return next;
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

  /** Extend the ordinary selection from the press cell as the pointer moves. */
  const beginSelectionDrag = (
    event: ReactPointerEvent<HTMLCanvasElement>,
    anchor: GridCellAddress,
    focus: GridCellAddress,
  ): void => {
    if (event.button !== 0) return;
    const drag = { pointerId: event.pointerId, anchor, focus };
    selectionDragRef.current = drag;
    const move = (moveEvent: PointerEvent): void => {
      if (moveEvent.pointerId !== drag.pointerId) return;
      const address = cellAtClientPoint(moveEvent.clientX, moveEvent.clientY);
      if (address === null) return;
      const target = mergeAnchor(stateRef.current.projection, address);
      if (target.row === drag.focus.row && target.col === drag.focus.col) return;
      drag.focus = target;
      onSelectionChange(extendSelection(collapseSelection(drag.anchor), target));
    };
    const end = (upEvent: PointerEvent): void => {
      if (upEvent.pointerId !== drag.pointerId) return;
      if (selectionDragRef.current === drag) selectionDragRef.current = null;
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

  const handlePointerDown = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    const layout = layoutRef.current;
    if (layout === null) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const pointX = event.clientX - rect.left;
    const pointY = event.clientY - rect.top;
    const current = stateRef.current.projection;
    const scroll = scrollRef.current;
    const session = editorSessionRef.current;

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
      onSelectionChange(extendSelection(collapseSelection(anchor), target));
      formulaBarPointing.onPoint(anchor, target);
      beginFormulaBarDrag(event, anchor, target);
      return;
    }

    containerRef.current?.focus();

    if (pointY < layout.headerHeight && pointX >= layout.headerWidth) {
      const col = columnAtHeader(current, layout, pointX, scroll.x);
      if (col === null) return;
      onSelectionChange(columnHeaderSelection(current, col, stateRef.current.selection, scroll));
      return;
    }
    if (pointX < layout.headerWidth && pointY >= layout.headerHeight) {
      const row = rowAtHeader(current, layout, pointY, scroll.y);
      if (row === null) return;
      onSelectionChange(rowHeaderSelection(current, row, stateRef.current.selection, scroll));
      return;
    }

    const address = cellAtPoint(current, layout, pointX, pointY, scroll.x, scroll.y);
    if (address === null) return;
    const merge = current.mergeAt(address.row, address.col);
    const target = merge === null ? address : { row: merge.anchorRow, col: merge.anchorCol };
    if (event.shiftKey) {
      const base = stateRef.current.selection;
      onSelectionChange(extendSelection(base, target));
      beginSelectionDrag(event, base.anchor, target);
      return;
    }
    onSelectionChange(collapseSelection(target));
    beginSelectionDrag(event, target, target);
  };

  const handleDoubleClick = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    // A double-click while pointing must not abandon the formula edit.
    if (editorSessionRef.current !== null) return;
    if (formulaBarPointing !== null && formulaBarPointing !== undefined) return;
    const layout = layoutRef.current;
    if (layout === null) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const current = stateRef.current.projection;
    const address = cellAtPoint(
      current,
      layout,
      event.clientX - rect.left,
      event.clientY - rect.top,
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
    const { key, ctrlKey, shiftKey, altKey, metaKey } = event;
    const current = stateRef.current;
    const layout = layoutRef.current;
    const pageRows =
      layout === null
        ? 20
        : Math.max(1, Math.floor(layout.scrollRegionHeight / current.projection.rowHeight(0)));
    const pageCols = layout === null ? 10 : Math.max(1, Math.floor(layout.scrollRegionWidth / 64));
    const outcome = resolveNavigation({
      key,
      ctrlKey,
      shiftKey,
      altKey,
      selection: current.selection,
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
      flushSync(() => beginEditing(current.selection.focus, '', false));
      return;
    }
    if (key === 'F2') {
      event.preventDefault();
      const active = current.selection.focus;
      beginEditing(active, current.projection.rawTextAt(active.row, active.col), true);
      return;
    }
    if (key === 'Delete' || key === 'Backspace') {
      event.preventDefault();
      onClear(current.selection.focus);
      return;
    }
    if (key.length === 1 && !ctrlKey && !altKey && !metaKey && /\S/.test(key)) {
      event.preventDefault();
      beginEditing(current.selection.focus, key, false);
    }
  };

  const handleEditorKeyDown = (event: ReactKeyboardEvent<HTMLTextAreaElement>): void => {
    event.stopPropagation();
    const session = editorSessionRef.current;
    if (session === null) return;
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
    const pageRows =
      layout === null
        ? 20
        : Math.max(1, Math.floor(layout.scrollRegionHeight / projection.rowHeight(0)));
    const pageCols = layout === null ? 10 : Math.max(1, Math.floor(layout.scrollRegionWidth / 64));
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
        };

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
        const active = stateRef.current.selection.focus;
        flushSync(() => beginEditing(active, '', false));
      }}
      data-testid="grid-shell"
    >
      <canvas
        ref={canvasRef}
        className="grid-canvas"
        onPointerDown={handlePointerDown}
        onDoubleClick={handleDoubleClick}
        data-testid="grid-canvas"
      />
      <div
        className="grid-scrollbar grid-scrollbar-vertical"
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
        data-testid="grid-hscroll"
        onPointerDown={handleScrollbarPointerDown('x')}
      >
        <div
          className="grid-scrollbar-thumb"
          ref={horizontalThumbRef}
          data-testid="grid-hscroll-thumb"
        />
      </div>
      {editor !== null ? (
        <textarea
          ref={editorRef}
          className="cell-editor"
          style={editorStyle}
          value={editor.text}
          onChange={(event) => updateEditorText(event.target.value)}
          onCompositionStart={() => setComposing(true)}
          onCompositionEnd={() => setComposing(false)}
          onKeyDown={handleEditorKeyDown}
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
    </div>
  );
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
 * A column header click selects the whole column. The active cell stays in
 * the visible part of the sheet, so the grid does not scroll to the far edge.
 */
function columnHeaderSelection(
  projection: SheetProjection,
  col: number,
  selection: GridSelection,
  scroll: ScrollOffset,
): GridSelection {
  const firstVisibleRow = Math.max(projection.freezeRows, projection.rowAt(scroll.y));
  const focus = selection.focus.col === col ? selection.focus : { row: firstVisibleRow, col };
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
 * A row header click selects the whole row, keeping the active cell in the
 * visible part of the sheet.
 */
function rowHeaderSelection(
  projection: SheetProjection,
  row: number,
  selection: GridSelection,
  scroll: ScrollOffset,
): GridSelection {
  const firstVisibleCol = Math.max(projection.freezeCols, projection.columnAt(scroll.x));
  const focus = selection.focus.row === row ? selection.focus : { row, col: firstVisibleCol };
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
    const thumbHeight = Math.max(24, Math.round(trackHeight * ratio));
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
    const thumbWidth = Math.max(24, Math.round(trackWidth * ratio));
    const span = Math.max(1, layout.maxScrollX - layout.minScrollX);
    const progress = (scroll.x - layout.minScrollX) / span;
    const left = trackLeft + progress * Math.max(0, trackWidth - thumbWidth);
    horizontal.style.width = `${thumbWidth}px`;
    horizontal.style.transform = `translateX(${Math.round(left)}px)`;
  }
}
