import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type {
  JSX,
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
} from 'react';
import type { SheetProjection } from '@view/types';
import { CanvasRenderFrame, cellAtPoint } from './canvas-renderer';
import { GRID_METRICS, type GridCellAddress, type RenderFrame } from './frame';

export interface CanvasGridProps {
  readonly projection: SheetProjection;
  readonly active: GridCellAddress;
  readonly onSelect: (address: GridCellAddress) => void;
  readonly onCommit: (address: GridCellAddress, text: string) => void;
  readonly onClear: (address: GridCellAddress) => void;
}

interface EditorState {
  readonly address: GridCellAddress;
  readonly text: string;
  readonly initialText: string;
  /** F2 starts editing the existing content; an unchanged commit is a no-op. */
  readonly commitIfChanged: boolean;
}

export function editorHasChanges(editor: {
  readonly text: string;
  readonly initialText: string;
  readonly commitIfChanged: boolean;
}): boolean {
  return !editor.commitIfChanged || editor.text !== editor.initialText;
}

export function CanvasGrid({
  projection,
  active,
  onSelect,
  onCommit,
  onClear,
}: CanvasGridProps): JSX.Element {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const frameRef = useRef<RenderFrame | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const committingRef = useRef(false);
  const [size, setSize] = useState({ width: 800, height: 480 });
  const [editor, setEditor] = useState<EditorState | null>(null);

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
      const rect = container.getBoundingClientRect();
      setSize({
        width: Math.max(200, Math.floor(rect.width)),
        height: Math.max(160, Math.floor(rect.height)),
      });
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(container);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    const frame = frameRef.current;
    if (frame === null) return;
    const dpr = globalThis.devicePixelRatio || 1;
    frame.resize(size.width, size.height, dpr);
    frame.render({
      projection,
      active,
      scrollX: 0,
      scrollY: 0,
      width: size.width,
      height: size.height,
      dpr,
    });
  }, [projection, active, size]);

  useEffect(() => {
    if (editor !== null) inputRef.current?.focus();
  }, [editor]);

  const startEditing = (address: GridCellAddress, text: string, commitIfChanged: boolean): void => {
    committingRef.current = false;
    setEditor({ address, text, initialText: text, commitIfChanged });
  };

  const commitEditor = (moveDown: boolean): void => {
    if (editor === null || committingRef.current) return;
    committingRef.current = true;
    const { address } = editor;
    setEditor(null);
    if (editorHasChanges(editor)) onCommit(address, editor.text);
    if (moveDown) {
      onSelect({
        row: Math.min(projection.rowCount - 1, address.row + 1),
        col: address.col,
      });
    }
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLCanvasElement>): void => {
    const rect = event.currentTarget.getBoundingClientRect();
    const address = cellAtPoint(
      projection,
      event.clientX - rect.left,
      event.clientY - rect.top,
      0,
      0,
    );
    if (address !== null) {
      onSelect(address);
      containerRef.current?.focus();
    }
  };

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (editor !== null) return;
    if (event.ctrlKey || event.metaKey || event.altKey) return;
    const { key } = event;
    if (key === 'ArrowUp') {
      event.preventDefault();
      onSelect({ row: Math.max(0, active.row - 1), col: active.col });
    } else if (key === 'ArrowDown' || key === 'Enter') {
      event.preventDefault();
      onSelect({
        row: Math.min(projection.rowCount - 1, active.row + 1),
        col: active.col,
      });
    } else if (key === 'ArrowLeft') {
      event.preventDefault();
      onSelect({ row: active.row, col: Math.max(0, active.col - 1) });
    } else if (key === 'ArrowRight' || key === 'Tab') {
      event.preventDefault();
      onSelect({
        row: active.row,
        col: Math.min(projection.columnCount - 1, active.col + 1),
      });
    } else if (key === 'F2') {
      event.preventDefault();
      startEditing(active, projection.cellAt(active.row, active.col)?.displayText ?? '', true);
    } else if (key === 'Delete' || key === 'Backspace') {
      event.preventDefault();
      onClear(active);
    } else if (key.length === 1) {
      event.preventDefault();
      startEditing(active, key, false);
    }
  };

  const handleEditorKeyDown = (event: ReactKeyboardEvent<HTMLInputElement>): void => {
    event.stopPropagation();
    if (event.key === 'Enter') {
      event.preventDefault();
      commitEditor(true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      committingRef.current = true;
      setEditor(null);
      containerRef.current?.focus();
    } else if (event.key === 'Tab') {
      event.preventDefault();
      const address = editor?.address;
      commitEditor(false);
      if (address !== undefined) {
        onSelect({ row: address.row, col: Math.min(projection.columnCount - 1, address.col + 1) });
      }
    }
  };

  const editorStyle =
    editor === null
      ? undefined
      : {
          left: `${GRID_METRICS.headerWidth + projection.columnX(editor.address.col)}px`,
          top: `${GRID_METRICS.headerHeight + projection.rowY(editor.address.row)}px`,
          width: `${projection.columnWidth(editor.address.col)}px`,
          height: `${projection.rowHeight(editor.address.row)}px`,
        };

  return (
    <div
      ref={containerRef}
      className="grid-shell"
      tabIndex={0}
      onKeyDown={handleKeyDown}
      data-testid="grid-shell"
    >
      <canvas
        ref={canvasRef}
        className="grid-canvas"
        onPointerDown={handlePointerDown}
        data-testid="grid-canvas"
      />
      {editor !== null ? (
        <input
          ref={inputRef}
          className="cell-editor"
          style={editorStyle}
          value={editor.text}
          onChange={(event) => setEditor({ ...editor, text: event.target.value })}
          onKeyDown={handleEditorKeyDown}
          onBlur={() => {
            if (editor !== null) commitEditor(false);
          }}
          data-testid="cell-editor"
        />
      ) : null}
    </div>
  );
}
