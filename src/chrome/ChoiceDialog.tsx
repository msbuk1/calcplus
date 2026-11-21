import { useRef } from 'react';
import type { JSX } from 'react';
import { Dialog } from './Dialog';

export interface ChoiceAction {
  /** Stable id handed back to `onChoose`. */
  readonly id: string;
  readonly label: string;
  /** Plain statement of what this action does and what it loses. */
  readonly statement?: string;
}

export interface ChoiceDialogProps {
  readonly title: string;
  readonly message: string;
  /** The outcome buttons, in the order the owning contract lists them. */
  readonly actions: readonly ChoiceAction[];
  /**
   * The write-nothing action. It sits under the outcome buttons, so the buttons
   * stay in the contract's order, and it owns Enter, Esc and the close button.
   */
  readonly cancelAction: ChoiceAction;
  readonly onChoose: (id: string) => void;
}

/**
 * The prompt for a choice where one button is followed by a file picker
 * (persistence §4, file-compat §6.3/§6.5). A browser confirm cannot carry the
 * user gesture `showSaveFilePicker` needs, so the picker is opened from this
 * dialog's own button click. Each outcome states plainly what it destroys;
 * Enter, Esc and the close button all choose the write-nothing action.
 */
export function ChoiceDialog({
  title,
  message,
  actions,
  cancelAction,
  onChoose,
}: ChoiceDialogProps): JSX.Element {
  const settledRef = useRef(false);
  const settle = (id: string): void => {
    if (settledRef.current) return;
    settledRef.current = true;
    onChoose(id);
  };

  return (
    <Dialog
      title={title}
      onClose={() => settle(cancelAction.id)}
      defaultAction={{ label: cancelAction.label, onSelect: () => settle(cancelAction.id) }}
      className="choice-dialog"
    >
      <p className="dialog-hint choice-message" data-testid="choice-message">
        {message}
      </p>
      <div className="choice-actions" data-testid="choice-actions">
        {actions.map((action) => (
          <div key={action.id} className="choice-action">
            <button
              type="button"
              className="dialog-button"
              data-testid={`choice-${action.id}`}
              onClick={() => settle(action.id)}
            >
              {action.label}
            </button>
            {action.statement === undefined ? null : (
              <span className="choice-statement">{action.statement}</span>
            )}
          </div>
        ))}
      </div>
    </Dialog>
  );
}
