import { useCallback, useLayoutEffect, useRef, useState } from 'react';
import type { JSX, RefObject } from 'react';
import type { FunctionCatalogEntry } from '@engine/function-catalog';
import { FormulaSuggestions } from '@grid/FormulaSuggestions';
import { useFormulaCompletion } from '@grid/useFormulaCompletion';

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
  /** Opens the Insert Function dialog; the `fx` button never edits text itself. */
  readonly onInsertFunction: () => void;
}

/**
 * The formula bar. It mirrors the active cell and edits it: the raw value or
 * formula, never the formatted display string (number-formats §5).
 *
 * While the draft is formula entry, the bar shares the function autocomplete
 * and signature hint with the in-cell editor. The popup never takes focus:
 * arrows choose a row, Enter/Tab insert it, and Esc closes the list before a
 * second Esc cancels the edit.
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
  onInsertFunction,
}: FormulaBarProps): JSX.Element {
  const localInputRef = useRef<HTMLInputElement | null>(null);
  const [focused, setFocused] = useState(false);
  const [caret, setCaret] = useState<number | null>(null);
  const [acceptCaret, setAcceptCaret] = useState<number | null>(null);
  const autocomplete = useFormulaCompletion({
    text,
    caret,
    formula: text.startsWith('='),
    enabled: focused,
  });

  // A reference insertion rewrites the draft, so the caret follows the text
  // that was just inserted or replaced.
  useLayoutEffect(() => {
    if (caretRequest === null) return;
    localInputRef.current?.setSelectionRange(caretRequest.offset, caretRequest.offset);
    setCaret(caretRequest.offset);
  }, [caretRequest]);

  // An accepted function name rewrites the draft and places the caret after
  // the opening parenthesis.
  useLayoutEffect(() => {
    if (acceptCaret === null) return;
    localInputRef.current?.setSelectionRange(acceptCaret, acceptCaret);
    setCaret(acceptCaret);
    setAcceptCaret(null);
  }, [acceptCaret, text]);

  const setInputRef = useCallback(
    (element: HTMLInputElement | null): void => {
      localInputRef.current = element;
      if (inputRef !== undefined) inputRef.current = element;
    },
    [inputRef],
  );

  const choose = (entry?: FunctionCatalogEntry): void => {
    const choice = entry === undefined ? autocomplete.accept() : autocomplete.choose(entry);
    if (choice === null) return;
    onTextChange(choice.text);
    setAcceptCaret(choice.caret);
  };

  const trackCaret = (element: HTMLInputElement): void => {
    setCaret(element.selectionStart ?? element.value.length);
  };

  /**
   * `fx` opens the Insert Function dialog (ticket 30). It no longer starts
   * formula entry itself: typing `=` in the grid or the bar still does, and the
   * dialog's Insert places the chosen skeleton in this bar.
   */
  return (
    <div className="formula-bar-group">
      <button
        type="button"
        className="fx-button"
        aria-label="Insert Function"
        title="Insert a function"
        data-testid="fx-button"
        // Keep the bar's in-progress draft and caret until the dialog opens.
        onMouseDown={(event) => event.preventDefault()}
        onClick={onInsertFunction}
      >
        fx
      </button>
      <div className="formula-bar-field">
        <input
          ref={setInputRef}
          type="text"
          className="formula-bar"
          value={text}
          spellCheck={false}
          aria-label="Formula bar"
          data-testid="formula-bar"
          onFocus={(event) => {
            setFocused(true);
            trackCaret(event.currentTarget);
            onFocused?.(text);
          }}
          onChange={(event) => {
            onTextChange(event.target.value);
            trackCaret(event.currentTarget);
          }}
          onSelect={(event) => trackCaret(event.currentTarget)}
          onClick={(event) => trackCaret(event.currentTarget)}
          onKeyUp={(event) => trackCaret(event.currentTarget)}
          onBlur={() => {
            setFocused(false);
            onStopEditing?.();
          }}
          onKeyDown={(event) => {
            if (autocomplete.listOpen) {
              if (event.key === 'ArrowDown') {
                event.preventDefault();
                autocomplete.move(1);
                return;
              }
              if (event.key === 'ArrowUp') {
                event.preventDefault();
                autocomplete.move(-1);
                return;
              }
              if (event.key === 'Escape') {
                event.preventDefault();
                autocomplete.dismiss();
                return;
              }
              if (event.key === 'Tab') {
                event.preventDefault();
                choose();
                return;
              }
              if (event.key === 'Enter' && !event.ctrlKey) {
                event.preventDefault();
                choose();
                return;
              }
            }
            if (event.key === 'Enter') {
              event.preventDefault();
              onCommit(text, { moveDown: !event.ctrlKey });
              return;
            }
            if (event.key === 'Escape') {
              event.preventDefault();
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
        {autocomplete.completion === null ? null : (
          <FormulaSuggestions
            id="formula-suggestions"
            completion={autocomplete.completion}
            activeIndex={autocomplete.activeIndex}
            activeEntry={autocomplete.activeEntry}
            style={{ position: 'absolute', top: 'calc(100% + 2px)', left: 0 }}
            onChoose={(entry) => choose(entry)}
          />
        )}
      </div>
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
          onCommit(text, { moveDown: true });
        }}
      >
        ✓
      </button>
    </div>
  );
}
