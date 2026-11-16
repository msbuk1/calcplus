import { useEffect, useRef, useState } from 'react';
import type { JSX, RefObject } from 'react';

export interface NameBoxProps {
  /** The address of the selection: one cell or a range such as `A1:B12`. */
  readonly value: string;
  /** Returns the normalised address, or `null` when the address is invalid. */
  readonly onNavigate: (text: string) => string | null;
  readonly onCancel: () => void;
  readonly onFocused?: () => void;
  readonly inputRef?: RefObject<HTMLInputElement | null>;
}

/**
 * The name box. It shows the address of the selection and accepts an address
 * to jump to. The Go To dialog itself is a chrome-milestone feature.
 */
export function NameBox({
  value,
  onNavigate,
  onCancel,
  onFocused,
  inputRef,
}: NameBoxProps): JSX.Element {
  const [draft, setDraft] = useState(value);
  const editingRef = useRef(false);

  useEffect(() => {
    if (!editingRef.current) setDraft(value);
  }, [value]);

  return (
    <input
      ref={inputRef}
      type="text"
      className="name-box"
      value={draft}
      spellCheck={false}
      aria-label="Name box"
      data-testid="name-box"
      onFocus={() => {
        editingRef.current = true;
        onFocused?.();
      }}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={() => {
        editingRef.current = false;
        setDraft(value);
      }}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          const accepted = onNavigate(draft);
          if (accepted === null) return;
          editingRef.current = false;
          setDraft(accepted);
          return;
        }
        if (event.key === 'Escape') {
          event.preventDefault();
          editingRef.current = false;
          setDraft(value);
          onCancel();
        }
      }}
    />
  );
}
