import { useRef, useState } from 'react';
import type { JSX } from 'react';
import { createPortal } from 'react-dom';
import { Menu } from './Menu';
import type { CommandStateResolver, ResolvedQuickAccessControl } from './ribbon-model';
import { quickAccessControls, ribbonMenuItems } from './ribbon-model';
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
 * the ribbon. Save As stays in the File menu and on the Save button's split
 * arrow; it is never a quick-access button of its own. The unsaved state shows
 * on the Save button itself; no separate dirty dot sits beside the filename.
 */
export function QuickAccessToolbar({
  dirty,
  resolveState,
  onCommand,
}: QuickAccessToolbarProps): JSX.Element {
  return (
    <div className="quick-access-toolbar" role="toolbar" aria-label="Quick access">
      {quickAccessControls().map((control) => (
        <QuickAccessControl
          key={control.key}
          control={control}
          dirty={dirty}
          resolveState={resolveState}
          onCommand={onCommand}
        />
      ))}
    </div>
  );
}

interface QuickAccessControlProps {
  readonly control: ResolvedQuickAccessControl;
  readonly dirty: boolean;
  readonly resolveState: CommandStateResolver;
  readonly onCommand: (commandId: string, args?: unknown) => void;
}

/**
 * One quick-access control. A control with an authored drop-down renders the
 * Excel split shape: the icon button plus a narrow arrow that opens its menu.
 * Every row still dispatches a command ID; the map owns the list.
 */
function QuickAccessControl({
  control,
  dirty,
  resolveState,
  onCommand,
}: QuickAccessControlProps): JSX.Element {
  const [open, setOpen] = useState(false);
  const [menuPosition, setMenuPosition] = useState<{ left: number; top: number } | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const toggleRef = useRef<HTMLButtonElement | null>(null);
  const enabled = resolveState(control.commandId).enabled;
  const items = ribbonMenuItems(control.menu ?? [], resolveState);
  const unsaved = dirty && control.commandId === 'file.save';
  const label = unsaved ? `${control.title}, unsaved changes` : control.title;

  const toggleMenu = (): void => {
    if (open) {
      setOpen(false);
      return;
    }
    const rect = containerRef.current?.getBoundingClientRect();
    if (rect === undefined) return;
    // The menu hangs under the arrow and stays inside the window.
    const width = 210;
    setMenuPosition({
      left: Math.max(6, Math.min(rect.left, window.innerWidth - width)),
      top: rect.bottom + 2,
    });
    setOpen(true);
  };

  const main = (
    <button
      type="button"
      className="quick-access-button"
      title={label}
      aria-label={label}
      disabled={!enabled}
      data-command={control.commandId}
      data-testid={QUICK_ACCESS_TEST_IDS[control.commandId]}
      data-unsaved={unsaved ? 'true' : undefined}
      onClick={() => {
        setOpen(false);
        onCommand(control.commandId);
      }}
    >
      {control.icon === undefined ? null : (
        <span className="ribbon-icon" aria-hidden="true">
          <RibbonIcon name={control.icon} />
        </span>
      )}
      {unsaved ? (
        <span className="quick-access-unsaved" data-testid="save-unsaved-dot" aria-hidden="true" />
      ) : null}
    </button>
  );

  if (control.menu === undefined) return main;

  return (
    <div ref={containerRef} className="quick-access-split">
      {main}
      <button
        ref={toggleRef}
        type="button"
        className="quick-access-split-toggle"
        title={`${control.label} options`}
        aria-label={`${control.label} options`}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={!enabled || items.length === 0}
        onClick={toggleMenu}
      >
        ▾
      </button>
      {open && menuPosition !== null
        ? createPortal(
            <Menu
              items={items}
              label={`${control.label} options`}
              className="ribbon-menu quick-access-menu"
              style={{
                position: 'fixed',
                left: `${menuPosition.left}px`,
                top: `${menuPosition.top}px`,
              }}
              anchorRef={containerRef}
              onSelect={(item) => {
                setOpen(false);
                // The row disappears with the menu; the arrow is the stable
                // invoking control for a dialog's focus restoration.
                toggleRef.current?.focus();
                if (item.commandId !== undefined) onCommand(item.commandId, item.args);
              }}
              onClose={() => {
                setOpen(false);
                toggleRef.current?.focus();
              }}
            />,
            window.document.body,
          )
        : null}
    </div>
  );
}
