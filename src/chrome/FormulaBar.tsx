import { useEffect, useRef, useState } from 'react';
import type { JSX, RefObject } from 'react';

export interface FormulaBarProps {
  /** The raw content of the active cell: its formula, or its value. */
  readonly text: string;
  /**
   * Commits the edited text to the active cell. Enter commits and moves down;
   * Ctrl+Enter commits and stays (command-system §7).
   */
  readonly onCommit: (text: string, options: { readonly moveDown: boolean }) => void;
  readonly onCancel: () => void;
  readonly onFocused?: () => void;
  readonly inputRef?: RefObject<HTMLInputElement | null>;
}

/**
 * The formula bar. It mirrors the active cell and edits it: the raw value or
 * formula, never the formatted display string (number-formats §5).
 */
export function FormulaBar({
  text,
  onCommit,
  onCancel,
  onFocused,
  inputRef,
}: FormulaBarProps): JSX.Element {
  const [draft, setDraft] = useState(text);
  const editingRef = useRef(false);

  useEffect(() => {
    if (!editingRef.current) setDraft(text);
  }, [text]);

  return (
    <div className="formula-bar-group">
      <span className="fx-button" aria-hidden="true">
        fx
      </span>
      <input
        ref={inputRef}
        type="text"
        className="formula-bar"
        value={draft}
        spellCheck={false}
        aria-label="Formula bar"
        data-testid="formula-bar"
        onFocus={() => {
          editingRef.current = true;
          onFocused?.();
        }}
        onChange={(event) => setDraft(event.target.value)}
        onBlur={() => {
          editingRef.current = false;
          setDraft(text);
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            editingRef.current = false;
            onCommit(draft, { moveDown: !event.ctrlKey });
            return;
          }
          if (event.key === 'Escape') {
            event.preventDefault();
            editingRef.current = false;
            setDraft(text);
            onCancel();
          }
        }}
      />
      <button
        type="button"
        className="formula-commit"
        title="Cancel (Esc)"
        aria-label="Cancel edit"
        data-testid="formula-cancel"
        onClick={() => {
          editingRef.current = false;
          setDraft(text);
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
        onClick={() => {
          editingRef.current = false;
          onCommit(draft, { moveDown: true });
        }}
      >
        ✓
      </button>
    </div>
  );
}
