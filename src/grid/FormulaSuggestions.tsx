import type { CSSProperties, JSX } from 'react';
import type { FunctionCatalogEntry } from '@engine/function-catalog';
import { signatureSegments } from '@engine/formula-suggestions';
import type { FormulaCompletion } from '@engine/formula-suggestions';

export interface FormulaSuggestionsProps {
  readonly id: string;
  readonly completion: FormulaCompletion;
  readonly activeIndex: number;
  readonly activeEntry: FunctionCatalogEntry | null;
  readonly style?: CSSProperties;
  /** A clicked row; the popup never takes focus away from the editor. */
  readonly onChoose: (entry: FunctionCatalogEntry) => void;
}

/**
 * The shared popup for the cell editor and the formula bar. A name list shows
 * every function that matches the token being typed with its signature; once
 * the caret is inside a call, the signature hint shows the call and emphasises
 * the argument being edited. The popup is pointer-only for choosing: the
 * keyboard belongs to the input, which routes arrows and Enter to the
 * completion controller.
 */
export function FormulaSuggestions({
  id,
  completion,
  activeIndex,
  activeEntry,
  style,
  onChoose,
}: FormulaSuggestionsProps): JSX.Element {
  // Keep the click from blurring the editor before its onClick runs.
  const keepFocus = (event: { preventDefault: () => void }): void => event.preventDefault();

  if (completion.kind === 'signature') {
    const segments = signatureSegments(completion);
    return (
      <div
        id={id}
        className="formula-suggestions formula-signature"
        role="status"
        data-testid="formula-signature"
        style={style}
        onMouseDown={keepFocus}
      >
        <div className="formula-signature-name">{completion.name}</div>
        <code className="formula-signature-code">
          {segments.map((segment, index) => (
            <span
              key={index}
              className={segment.active ? 'formula-signature-arg-active' : undefined}
            >
              {segment.text}
            </span>
          ))}
        </code>
      </div>
    );
  }

  return (
    <div
      id={id}
      className="formula-suggestions"
      role="listbox"
      aria-label="Function suggestions"
      data-testid="formula-suggestions"
      style={style}
      onMouseDown={keepFocus}
    >
      {completion.matches.map((entry, index) => (
        <div
          key={entry.name}
          role="option"
          aria-selected={index === activeIndex}
          className={
            index === activeIndex
              ? 'formula-suggestion formula-suggestion-active'
              : 'formula-suggestion'
          }
          data-testid={`formula-suggestion-${entry.name}`}
          onMouseDown={keepFocus}
          onClick={() => onChoose(entry)}
        >
          <span className="formula-suggestion-name">{entry.name}</span>
          <span className="formula-suggestion-signature">{entry.signature}</span>
        </div>
      ))}
      {activeEntry === null ? null : (
        <div className="formula-suggestion-detail" data-testid="formula-suggestion-signature">
          {activeEntry.signature}
        </div>
      )}
    </div>
  );
}
