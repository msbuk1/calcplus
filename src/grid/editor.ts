import type { SheetProjection } from '@view/types';
import type { GridCellAddress, GridRange } from './frame';

/**
 * The in-cell editor. Editing a cell is a DOM overlay; this module holds the
 * rules that decide what a key does, so they can be tested without a browser.
 */
export type EditorCommitDirection = 'down' | 'up' | 'right' | 'left' | 'none';

/**
 * The reference the pointer or keyboard is currently pointing at. The token
 * offsets identify the reference text in the editor, so moving the selection
 * replaces that reference instead of appending a second one.
 */
export interface EditorPointing {
  readonly anchor: GridCellAddress;
  readonly focus: GridCellAddress;
  readonly start: number;
  readonly end: number;
}

export interface EditorSession {
  /** Identifies the editing session, so the view can focus it once. */
  readonly id: number;
  readonly address: GridCellAddress;
  readonly text: string;
  readonly initialText: string;
  /**
   * F2 starts editing the existing content and commits only when the text
   * changed; typing starts a fresh value and always commits.
   */
  readonly commitIfChanged: boolean;
  /**
   * True when the text is a formula shown without the leading `=` in the
   * editor. The authored formula text is kept exactly and the `=` is restored
   * only when the edit actually commits.
   */
  readonly formulaEntry: boolean;
  /** True while an input method is composing. */
  readonly composing: boolean;
  /** The reference the formula is currently pointing at, or `null`. */
  readonly pointing: EditorPointing | null;
}

export interface EditorKeyEvent {
  readonly key: string;
  readonly shiftKey: boolean;
  readonly ctrlKey: boolean;
  readonly altKey: boolean;
  /** True while the browser reports an active input-method composition. */
  readonly isComposing: boolean;
}

export type EditorAction =
  | { readonly type: 'commit'; readonly direction: EditorCommitDirection }
  | { readonly type: 'cancel' }
  | { readonly type: 'ignore' };

let sessionCounter = 0;

export function startEdit(
  address: GridCellAddress,
  text: string,
  commitIfChanged: boolean,
  formulaEntry = false,
): EditorSession {
  sessionCounter += 1;
  return {
    id: sessionCounter,
    address,
    text,
    initialText: text,
    commitIfChanged,
    formulaEntry,
    composing: false,
    pointing: null,
  };
}

export function editorHasChanges(session: {
  readonly text: string;
  readonly initialText: string;
  readonly commitIfChanged: boolean;
}): boolean {
  return !session.commitIfChanged || session.text !== session.initialText;
}

/** True while the editor holds a formula that can point at cells. */
export function isFormulaEntry(session: EditorSession): boolean {
  return session.formulaEntry || session.text.startsWith('=');
}

/**
 * The text a commit hands to the command layer. A formula entry shown without
 * its leading `=` gets it back here, so the authored formula stays a formula.
 */
export function editorCommitText(session: EditorSession): string {
  if (session.formulaEntry && !session.text.startsWith('=')) return `=${session.text}`;
  return session.text;
}

/**
 * True when the caret sits where a reference belongs, so an arrow key starts
 * pointing instead of moving the caret. This is the Enter-mode entry point
 * (`=`, an operator, an opening parenthesis or an argument comma).
 */
export function isPointingBoundary(session: { readonly text: string }, caret: number): boolean {
  const before = session.text.slice(0, caret).trimEnd();
  return /[=+\-*/^&<>:,([{]$/.test(before);
}

/** The ordered rectangle covered by a pointed anchor and focus. */
export function pointingRange(anchor: GridCellAddress, focus: GridCellAddress): GridRange {
  return {
    startRow: Math.min(anchor.row, focus.row),
    startCol: Math.min(anchor.col, focus.col),
    endRow: Math.max(anchor.row, focus.row),
    endCol: Math.max(anchor.col, focus.col),
  };
}

/**
 * The reference text a formula pointing gesture inserts: one cell for a click,
 * the ordered rectangle for a drag or a Shift-extended range.
 */
export function pointingReference(
  projection: SheetProjection,
  anchor: GridCellAddress,
  focus: GridCellAddress,
): string {
  const range = pointingRange(anchor, focus);
  const start = `${projection.columnLabel(range.startCol)}${projection.rowLabel(range.startRow)}`;
  if (range.startRow === range.endRow && range.startCol === range.endCol) return start;
  return `${start}:${projection.columnLabel(range.endCol)}${projection.rowLabel(range.endRow)}`;
}

/**
 * Replace the text occupied by the previous pointing token, or insert at the
 * caret when the gesture starts. The caret lands after the inserted reference.
 */
export function applyPointing(
  session: EditorSession,
  reference: string,
  anchor: GridCellAddress,
  focus: GridCellAddress,
  token: { readonly start: number; readonly end: number } | null,
): EditorSession {
  const start =
    token?.start ??
    session.pointing?.start ??
    // A new gesture inserts at the caret. The caller may supply a token when
    // the caret moved since the last gesture.
    session.text.length;
  const end = token?.end ?? session.pointing?.end ?? start;
  const text = `${session.text.slice(0, start)}${reference}${session.text.slice(end)}`;
  const nextStart = start;
  return {
    ...session,
    text,
    pointing: { anchor, focus, start: nextStart, end: nextStart + reference.length },
  };
}

const ARROW_DIRECTIONS: Readonly<Record<string, EditorCommitDirection>> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
};

/**
 * Decide what a key does inside the editor. While an input method composes,
 * Enter confirms the candidate and must not commit the cell.
 *
 * Direct typing is Excel's enter mode, so an arrow key commits and moves in
 * that direction. F2 starts in edit mode and keeps the arrows for the caret;
 * formula pointing resolves the arrows in the view after this returns
 * "ignore".
 */
export function resolveEditorAction(session: EditorSession, event: EditorKeyEvent): EditorAction {
  if (session.composing || event.isComposing) return { type: 'ignore' };
  switch (event.key) {
    case 'Enter':
      return { type: 'commit', direction: event.shiftKey ? 'up' : 'down' };
    case 'Tab':
      return { type: 'commit', direction: event.shiftKey ? 'left' : 'right' };
    case 'ArrowUp':
    case 'ArrowDown':
    case 'ArrowLeft':
    case 'ArrowRight':
      if (session.commitIfChanged || isFormulaEntry(session)) return { type: 'ignore' };
      return { type: 'commit', direction: ARROW_DIRECTIONS[event.key] ?? 'none' };
    case 'Escape':
      return { type: 'cancel' };
    default:
      return { type: 'ignore' };
  }
}
