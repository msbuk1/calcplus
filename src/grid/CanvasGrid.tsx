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
  editorHasChanges,
  resolveEditorAction,
  startEdit,
  type EditorCommitDirection,
  type EditorSession,
} from './editor';
import type { GridCellAddress, GridSelection, RenderFrame } from './frame';
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

  const beginEditing = (address: GridCellAddress, text: string, commitIfChanged: boolean): void => {
    // Editing always happens on the anchor of a merged range.
    const target = mergeAnchor(stateRef.current.projection, address);
    const session = startEdit(target, text, commitIfChanged);
    editorSessionRef.current = session;
    setEditor(session);
  };

  const updateEditorText = (text: string): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    const next = { ...session, text };
    editorSessionRef.current = next;
    setEditor(next);
  };

  const setComposing = (composing: boolean): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    const next = { ...session, composing };
    editorSessionRef.current = next;
    setEditor(next);
  };

  const commitEditor = (direction: EditorCommitDirection): void => {
    const session = editorSessionRef.current;
    if (session === null) return;
    editorSessionRef.current = null;
    setEditor(null);
    if (editorHasChanges(session)) onCommit(session.address, session.text);
    const target = moveAfterCommit(session.address, direction, projection);
    if (target !== null) onSelectionChange(collapseSelection(target));
    containerRef.current?.focus();
  };

  const cancelEditor = (): void => {
    editorSessionRef.current = null;
    setEditor(null);
    containerRef.current?.focus();
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    const layout = layoutRef.current;
    if (layout === null) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const pointX = event.clientX - rect.left;
    const pointY = event.clientY - rect.top;
    containerRef.current?.focus();
    const current = stateRef.current.projection;
    const scroll = scrollRef.current;

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
    if (event.shiftKey) {
      onSelectionChange(extendSelection(stateRef.current.selection, target));
      return;
    }
    onSelectionChange(collapseSelection(target));
  };

  const handleDoubleClick = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
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
    }
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
      onKeyDown={handleKeyDown}
      onFocus={() => onFocusZone?.()}
      onCompositionStart={() => {
        // Fallback for an input method that starts composing without first
        // reporting a process key: open the editor on the selected cell. The
        // keydown path handles the usual case and opens it before composing.
        if (editorSessionRef.current !== null) return;
        const active = stateRef.current.selection.focus;
        beginEditing(active, '', false);
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
            if (editorSessionRef.current !== null) commitEditor('none');
          }}
          data-testid="cell-editor"
        />
      ) : null}
    </div>
  );
}

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
