import { useCallback, useLayoutEffect, useRef } from 'react';
import type { JSX, RefObject } from 'react';

export interface FormulaBarKeyModifiers {
  readonly shiftKey: boolean;
  readonly ctrlKey: boolean;
  readonly altKey: boolean;
}

export interface FormulaBarProps {
  /**
   * The draft text. It mirrors the live text while an in-cell edit is open or
   * the bar is editing, otherwise the active cell's raw formula or value.
   */
  readonly text: string;
  /** Replaces the draft while the user types. */
  readonly onTextChange: (text: string) => void;
  /**
   * Commits the edited text to the cell the bar started editing. Enter commits
   * and moves down; Ctrl+Enter commits and stays (command-system §7).
   */
  readonly onCommit: (text: string, options: { readonly moveDown: boolean }) => void;
  readonly onCancel: () => void;
  /** Editing started with this text; the bar now owns the draft. */
  readonly onFocused?: (text: string) => void;
  /** Editing ended without a commit or cancel, for example a click elsewhere. */
  readonly onStopEditing?: () => void;
  /**
   * Handles navigation keys while the draft is a formula, so the grid can
   * extend the selection and insert a reference at the caret. Returns true
   * when the key was used for pointing.
   */
  readonly onPointKey?: (key: string, modifiers: FormulaBarKeyModifiers) => boolean;
  /**
   * Places the caret after a reference the grid inserted. A new object asks for
   * a new placement, so typing afterwards does not move the caret again.
   */
  readonly caretRequest?: { readonly offset: number } | null;
  readonly inputRef?: RefObject<HTMLInputElement | null>;
}

/**
 * The formula bar. It mirrors the active cell and edits it: the raw value or
 * formula, never the formatted display string (number-formats §5).
 */
export function FormulaBar({
  text,
  onTextChange,
  onCommit,
  onCancel,
  onFocused,
  onStopEditing,
  onPointKey,
  caretRequest = null,
  inputRef,
}: FormulaBarProps): JSX.Element {
  const editingRef = useRef(false);
  const localInputRef = useRef<HTMLInputElement | null>(null);
  const pendingCaretRef = useRef<number | null>(null);

  // The `=` is inserted before the draft render, so the caret follows the new
  // value instead of being clamped to the old one.
  useLayoutEffect(() => {
    const caret = pendingCaretRef.current;
    if (caret === null) return;
    pendingCaretRef.current = null;
    localInputRef.current?.setSelectionRange(caret, caret);
  }, [text]);

  // A reference insertion rewrites the draft, so the caret follows the text
  // that was just inserted or replaced.
  useLayoutEffect(() => {
    if (caretRequest === null) return;
    localInputRef.current?.setSelectionRange(caretRequest.offset, caretRequest.offset);
  }, [caretRequest]);

  const setInputRef = useCallback(
    (element: HTMLInputElement | null): void => {
      localInputRef.current = element;
      if (inputRef !== undefined) inputRef.current = element;
    },
    [inputRef],
  );

  /**
   * The Insert Function dialog is a V2 feature. In v1 the `fx` control starts
   * or continues formula entry in the bar: it inserts the leading `=` when the
   * current text is not already a formula, then leaves the keyboard there.
   */
  const startFormula = (): void => {
    const element = localInputRef.current;
    const wasEditing = editingRef.current;
    const current = text;
    const formulaText = current.startsWith('=') ? current : `=${current}`;
    const inserted = formulaText !== current;
    editingRef.current = true;
    onTextChange(formulaText);
    element?.focus();
    // A fresh formula entry starts at the end; continuing one keeps the caret
    // where the user left it.
    if (element !== null && (!wasEditing || inserted)) {
      if (inserted) pendingCaretRef.current = formulaText.length;
      else element.setSelectionRange(formulaText.length, formulaText.length);
    }
  };

  return (
    <div className="formula-bar-group">
      <button
        type="button"
        className="fx-button"
        aria-label="Formula entry"
        title="Start or continue formula entry"
        data-testid="fx-button"
        // Keep the bar's in-progress draft and caret when it already owns the
        // keyboard; the click handler below starts the formula instead of a
        // blur resetting the draft first.
        onMouseDown={(event) => event.preventDefault()}
        onClick={startFormula}
      >
        fx
      </button>
      <input
        ref={setInputRef}
        type="text"
        className="formula-bar"
        value={text}
        spellCheck={false}
        aria-label="Formula bar"
        data-testid="formula-bar"
        onFocus={() => {
          editingRef.current = true;
          onFocused?.(text);
        }}
        onChange={(event) => onTextChange(event.target.value)}
        onBlur={() => {
          editingRef.current = false;
          onStopEditing?.();
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            editingRef.current = false;
            onCommit(text, { moveDown: !event.ctrlKey });
            return;
          }
          if (event.key === 'Escape') {
            event.preventDefault();
            editingRef.current = false;
            onCancel();
            return;
          }
          if (
            onPointKey?.(event.key, {
              shiftKey: event.shiftKey,
              ctrlKey: event.ctrlKey,
              altKey: event.altKey,
            }) === true
          ) {
            event.preventDefault();
          }
        }}
      />
      <button
        type="button"
        className="formula-commit"
        title="Cancel (Esc)"
        aria-label="Cancel edit"
        data-testid="formula-cancel"
        // The click commits or cancels the draft; the press must not blur the
        // input and reset the draft before the click runs.
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          editingRef.current = false;
          onCancel();
        }}
      >
        ✕
      </button>
      <button
        type="button"
        className="formula-commit"
        title="Enter"
        aria-label="Commit edit"
        data-testid="formula-commit"
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          editingRef.current = false;
          onCommit(text, { moveDown: true });
        }}
      >
        ✓
      </button>
    </div>
  );
}
