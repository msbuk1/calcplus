import { useRef, useState } from 'react';
import type { JSX, RefObject } from 'react';
import { Menu } from './Menu';
import { RibbonIcon } from './RibbonIcon';
import { compactControls, menuItemsFor } from './ribbon-model';
import type { CommandStateResolver } from './ribbon-model';

export interface CompactBarProps {
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly resolveState: CommandStateResolver;
  readonly ribbonExpanded: boolean;
  readonly onToggleRibbon: () => void;
  readonly compactRef?: RefObject<HTMLElement | null>;
  readonly onFocus?: () => void;
}

/**
 * Variant B. Below 900 px the ribbon collapses to one fixed-height row of short
 * commands and a control that expands the full ribbon in place.
 */
export function CompactBar({
  onCommand,
  resolveState,
  ribbonExpanded,
  onToggleRibbon,
  compactRef,
  onFocus,
}: CompactBarProps): JSX.Element {
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const fileButtonRef = useRef<HTMLButtonElement | null>(null);
  const controls = compactControls();
  const fileItems = menuItemsFor('File', resolveState);

  return (
    <section
      className="compact-bar"
      ref={compactRef}
      tabIndex={-1}
      aria-label="Compact command bar"
      onFocusCapture={onFocus}
    >
      <div className="compact-file">
        <button
          ref={fileButtonRef}
          type="button"
          className="compact-button compact-file-button"
          aria-haspopup="menu"
          aria-expanded={fileMenuOpen}
          onClick={() => setFileMenuOpen((open) => !open)}
          onPointerDown={(event) => event.stopPropagation()}
          data-testid="compact-file-menu-button"
        >
          File
        </button>
        {fileMenuOpen ? (
          <Menu
            items={fileItems}
            label="File"
            className="file-menu"
            anchorRef={fileButtonRef}
            onSelect={(item) => {
              setFileMenuOpen(false);
              // The chosen item disappears with the menu; the trigger is the
              // stable invoking control for focus restoration.
              fileButtonRef.current?.focus();
              if (item.commandId !== undefined) onCommand(item.commandId, item.args);
            }}
            onClose={() => {
              setFileMenuOpen(false);
              fileButtonRef.current?.focus();
            }}
          />
        ) : null}
      </div>
      <div className="compact-commands" role="toolbar" aria-label="Common commands">
        {controls.map((control) => {
          const state = resolveState(control.commandId);
          const checked = control.kind === 'toggle' ? state.checked : undefined;
          const iconOnly = control.iconOnly && control.icon !== undefined;
          return (
            <button
              key={control.key}
              type="button"
              className={[
                'compact-button',
                control.kind === 'toggle' ? 'compact-toggle' : '',
                checked === true ? 'active' : '',
                iconOnly ? 'compact-button-icon-only' : '',
              ]
                .filter(Boolean)
                .join(' ')}
              title={control.title}
              aria-label={control.title}
              aria-pressed={checked}
              disabled={!state.enabled}
              data-command={control.commandId}
              onClick={() => onCommand(control.commandId, control.args)}
            >
              {control.icon === undefined ? null : (
                <span className="ribbon-icon" aria-hidden="true">
                  <RibbonIcon name={control.icon} />
                </span>
              )}
              {iconOnly ? null : control.label}
            </button>
          );
        })}
      </div>
      <button
        type="button"
        className="compact-button compact-expand"
        aria-expanded={ribbonExpanded}
        onClick={onToggleRibbon}
        data-testid="ribbon-expand-button"
      >
        {ribbonExpanded ? 'Hide ribbon' : 'Expand ribbon'}
      </button>
    </section>
  );
}
