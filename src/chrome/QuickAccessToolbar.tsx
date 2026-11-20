import type { JSX } from 'react';
import type { CommandStateResolver } from './ribbon-model';
import { quickAccessControls } from './ribbon-model';
import { RibbonIcon } from './RibbonIcon';

export interface QuickAccessToolbarProps {
  readonly dirty: boolean;
  readonly resolveState: CommandStateResolver;
  readonly onCommand: (commandId: string, args?: unknown) => void;
}

/** Stable test hooks for the three oracle quick-access commands. */
const QUICK_ACCESS_TEST_IDS: Readonly<Record<string, string>> = {
  'file.save': 'save-button',
  'edit.undo': 'undo-button',
  'edit.redo': 'redo-button',
};

/**
 * The Excel 365 Quick Access Toolbar: Save, Undo and Redo as small icon
 * buttons with command-and-shortcut tooltips, anchored at the top left above
 * the ribbon. Save As stays in the File menu. The unsaved state shows on the
 * Save button itself; no separate dirty dot sits beside the filename.
 */
export function QuickAccessToolbar({
  dirty,
  resolveState,
  onCommand,
}: QuickAccessToolbarProps): JSX.Element {
  return (
    <div className="quick-access-toolbar" role="toolbar" aria-label="Quick access">
      {quickAccessControls().map((control) => {
        const unsaved = dirty && control.commandId === 'file.save';
        const label = unsaved ? `${control.title}, unsaved changes` : control.title;
        return (
          <button
            key={control.key}
            type="button"
            className="quick-access-button"
            title={label}
            aria-label={label}
            disabled={!resolveState(control.commandId).enabled}
            data-command={control.commandId}
            data-testid={QUICK_ACCESS_TEST_IDS[control.commandId]}
            data-unsaved={unsaved ? 'true' : undefined}
            onClick={() => onCommand(control.commandId)}
          >
            {control.icon === undefined ? null : (
              <span className="ribbon-icon" aria-hidden="true">
                <RibbonIcon name={control.icon} />
              </span>
            )}
            {unsaved ? (
              <span
                className="quick-access-unsaved"
                data-testid="save-unsaved-dot"
                aria-hidden="true"
              />
            ) : null}
          </button>
        );
      })}
    </div>
  );
}
