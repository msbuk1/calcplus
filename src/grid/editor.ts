import type { GridCellAddress } from './frame';

/**
 * The in-cell editor. Editing a cell is a DOM overlay; this module holds the
 * rules that decide what a key does, so they can be tested without a browser.
 */
export type EditorCommitDirection = 'down' | 'up' | 'right' | 'left' | 'none';

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
  /** True while an input method is composing. */
  readonly composing: boolean;
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
): EditorSession {
  sessionCounter += 1;
  return {
    id: sessionCounter,
    address,
    text,
    initialText: text,
    commitIfChanged,
    composing: false,
  };
}

export function editorHasChanges(session: {
  readonly text: string;
  readonly initialText: string;
  readonly commitIfChanged: boolean;
}): boolean {
  return !session.commitIfChanged || session.text !== session.initialText;
}

/**
 * Decide what a key does inside the editor. While an input method composes,
 * Enter confirms the candidate and must not commit the cell.
 */
export function resolveEditorAction(session: EditorSession, event: EditorKeyEvent): EditorAction {
  if (session.composing || event.isComposing) return { type: 'ignore' };
  switch (event.key) {
    case 'Enter':
      return { type: 'commit', direction: event.shiftKey ? 'up' : 'down' };
    case 'Tab':
      return { type: 'commit', direction: event.shiftKey ? 'left' : 'right' };
    case 'Escape':
      return { type: 'cancel' };
    default:
      return { type: 'ignore' };
  }
}
