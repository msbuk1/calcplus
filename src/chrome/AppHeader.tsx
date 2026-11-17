import type { JSX } from 'react';
import appIcon from '../../assets/calcplus.svg';

export interface AppHeaderProps {
  readonly fileName: string | null;
  readonly dirty: boolean;
  readonly canSave: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly onCommand: (commandId: string) => void;
}

/** Title strip: document name, dirty marker and the quick-access commands. */
export function AppHeader({
  fileName,
  dirty,
  canSave,
  canUndo,
  canRedo,
  onCommand,
}: AppHeaderProps): JSX.Element {
  return (
    <header className="app-header">
      <div className="brand">
        <img className="brand-icon" src={appIcon} alt="" width={26} height={26} />
        <span className="brand-name">calcplus</span>
        <span className="brand-phase">M4C chrome</span>
      </div>
      <div className="file-meta">
        <span className="file-name" data-testid="file-name">
          {fileName ?? 'No workbook open'}
        </span>
        {dirty ? (
          <span
            className="dirty-dot"
            title="Unsaved changes"
            aria-label="Unsaved changes"
            data-testid="dirty-indicator"
          >
            ●
          </span>
        ) : null}
      </div>
      <div className="quick-access" role="toolbar" aria-label="Quick access">
        <button
          type="button"
          className="quick-access-button"
          title="Save (Ctrl+S)"
          aria-label="Save"
          disabled={!canSave}
          data-command="file.save"
          data-testid="save-button"
          onClick={() => onCommand('file.save')}
        >
          Save
        </button>
        <button
          type="button"
          className="quick-access-button"
          title="Undo (Ctrl+Z)"
          aria-label="Undo"
          disabled={!canUndo}
          data-command="edit.undo"
          data-testid="undo-button"
          onClick={() => onCommand('edit.undo')}
        >
          Undo
        </button>
        <button
          type="button"
          className="quick-access-button"
          title="Redo (Ctrl+Y)"
          aria-label="Redo"
          disabled={!canRedo}
          data-command="edit.redo"
          data-testid="redo-button"
          onClick={() => onCommand('edit.redo')}
        >
          Redo
        </button>
      </div>
    </header>
  );
}
