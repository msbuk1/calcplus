import { useEffect, useId, useRef, useState } from 'react';
import type { JSX } from 'react';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';

export interface GoToOutcome {
  /** `false` keeps the dialog open and leaves the selection unchanged. */
  readonly ok: boolean;
  readonly message: string | null;
}

export interface GoToDialogProps {
  /** The selection address shown when the dialog opens. */
  readonly initialValue: string;
  readonly onClose: (reason: DialogCloseReason) => void;
  readonly onSubmit: (target: string) => GoToOutcome;
}

/**
 * Go To accepts an A1 address, a range or a defined name. The dialog validates
 * and submits `goto.goto`; a rejected target shows an error and changes neither
 * the selection nor the active sheet. The Dialog shell owns the focus trap and
 * Esc; a successful jump hands focus back to the grid.
 */
export function GoToDialog({ initialValue, onClose, onSubmit }: GoToDialogProps): JSX.Element {
  const [target, setTarget] = useState(initialValue);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const inputId = useId();
  const hintId = useId();

  useEffect(() => {
    const handle = requestAnimationFrame(() => {
      const input = inputRef.current;
      if (input === null) return;
      input.focus();
      input.select();
    });
    return () => cancelAnimationFrame(handle);
  }, []);

  const submit = (): void => {
    if (target.trim() === '') {
      setError('Enter a cell reference or defined name.');
      return;
    }
    const outcome = onSubmit(target);
    if (!outcome.ok) {
      setError(outcome.message ?? 'That reference cannot be selected.');
    }
  };

  return (
    <Dialog
      title="Go to"
      onClose={onClose}
      defaultAction={{ label: 'OK', onSelect: submit }}
      secondaryAction={{ label: 'Cancel', onSelect: () => onClose('dismiss') }}
      className="goto-dialog"
    >
      <div className="goto-form">
        <label className="find-label" htmlFor={inputId}>
          Reference
        </label>
        <input
          ref={inputRef}
          id={inputId}
          type="text"
          className="find-input"
          value={target}
          spellCheck={false}
          autoComplete="off"
          aria-describedby={hintId}
          aria-invalid={error !== null}
          data-testid="goto-input"
          onChange={(event) => {
            setTarget(event.target.value);
            setError(null);
          }}
        />
        <p id={hintId} className="dialog-hint">
          Accepts A1 addresses, ranges such as A1:B4, and defined names. The current selection is{' '}
          {initialValue}.
        </p>
        {error === null ? null : (
          <p className="dialog-error" role="alert" data-testid="goto-error">
            {error}
          </p>
        )}
      </div>
    </Dialog>
  );
}
