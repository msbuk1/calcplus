import { useEffect, useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import { createPortal } from 'react-dom';
import type { RibbonSelectOption, RibbonTabId } from '@commands/inventory';
import type { PublishedFunction } from '@engine/function-registry';
import { ColorPickerPopup } from './ColorPicker';
import { Menu } from './Menu';
import { RibbonIcon } from './RibbonIcon';
import { RIBBON_TABS, menuItemsFor, ribbonGroupsFor, ribbonMenuItems } from './ribbon-model';
import type {
  CommandStateResolver,
  ResolvedRibbonControl,
  ResolvedRibbonGroup,
} from './ribbon-model';
import { useMediaQuery } from './useMediaQuery';

/**
 * The one width that flips the Formulas › Function Library group between its
 * two states (see the prototype on the prototype/function-library branch):
 * at this width and above the nine drop-downs sit in one row under 30px
 * icons; below it they stack into three columns beside 18px icons. The 940
 * is the prototype's breakpoint and a guess, not a decision — change it here
 * and both states move together.
 */
export const FUNCTION_LIBRARY_WIDE_MIN_WIDTH = 940;

export interface RibbonProps {
  readonly activeTab: RibbonTabId;
  readonly onTabChange: (tab: RibbonTabId) => void;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly resolveState: CommandStateResolver;
  readonly controlValues: Readonly<Record<string, string>>;
  /** Functions inserted this session, most recent first, for Recently Used. */
  readonly recentFunctions?: readonly PublishedFunction[];
  readonly ribbonRef?: RefObject<HTMLElement | null>;
  readonly onFocus?: () => void;
}

/**
 * Variant A. Every control dispatches a command ID; checked and enabled state
 * comes from the registry through `resolveState`. Groups with no implemented
 * IN command are absent, and OUT commands are never rendered; the one ruled
 * exception is a present-but-disabled placement (ticket 27).
 *
 * Each group band packs up to two rows of controls over its label, as Excel
 * does. A control's row comes from the command map (`row`, `rowSpan`).
 */
export function Ribbon({
  activeTab,
  onTabChange,
  onCommand,
  resolveState,
  controlValues,
  recentFunctions = [],
  ribbonRef,
  onFocus,
}: RibbonProps): JSX.Element {
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const fileButtonRef = useRef<HTMLButtonElement | null>(null);
  const tabRefs = useRef<(HTMLButtonElement | null)[]>([]);
  const fileItems = menuItemsFor('File', resolveState);
  const groups = ribbonGroupsFor(activeTab, resolveState, recentFunctions);

  const focusTab = (index: number): void => {
    const next = (index + RIBBON_TABS.length) % RIBBON_TABS.length;
    const tab = RIBBON_TABS[next];
    if (tab === undefined) return;
    onTabChange(tab.id);
    tabRefs.current[next]?.focus();
  };

  const onTabKeyDown = (event: ReactKeyboardEvent<HTMLButtonElement>, index: number): void => {
    if (event.key === 'ArrowRight') {
      event.preventDefault();
      focusTab(index + 1);
      return;
    }
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      focusTab(index - 1);
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      focusTab(0);
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      focusTab(RIBBON_TABS.length - 1);
    }
  };

  return (
    <section
      className="ribbon"
      ref={ribbonRef}
      tabIndex={-1}
      aria-label="Ribbon"
      onFocusCapture={onFocus}
    >
      <div className="ribbon-tabs">
        <div className="ribbon-file">
          <button
            ref={fileButtonRef}
            type="button"
            className={
              fileMenuOpen ? 'ribbon-tab ribbon-file-tab active' : 'ribbon-tab ribbon-file-tab'
            }
            aria-haspopup="menu"
            aria-expanded={fileMenuOpen}
            onClick={() => setFileMenuOpen((open) => !open)}
            onPointerDown={(event) => event.stopPropagation()}
            data-testid="file-menu-button"
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
        <div className="ribbon-tablist" role="tablist" aria-label="Ribbon tabs">
          {RIBBON_TABS.map((tab, index) => (
            <button
              key={tab.id}
              ref={(element) => {
                tabRefs.current[index] = element;
              }}
              type="button"
              role="tab"
              aria-selected={tab.id === activeTab}
              tabIndex={tab.id === activeTab ? 0 : -1}
              className={tab.id === activeTab ? 'ribbon-tab active' : 'ribbon-tab'}
              onClick={() => onTabChange(tab.id)}
              onKeyDown={(event) => onTabKeyDown(event, index)}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>
      <div className="ribbon-body" role="tabpanel" aria-label={activeTab}>
        {groups.length === 0 ? (
          <div className="ribbon-empty">No commands in this tab yet.</div>
        ) : (
          groups.map((group) => (
            <RibbonGroup
              key={group.id}
              group={group}
              resolveState={resolveState}
              controlValues={controlValues}
              onCommand={onCommand}
            />
          ))
        )}
      </div>
    </section>
  );
}

interface RibbonGroupProps {
  readonly group: ResolvedRibbonGroup;
  readonly resolveState: CommandStateResolver;
  readonly controlValues: Readonly<Record<string, string>>;
  readonly onCommand: (commandId: string, args?: unknown) => void;
}

/**
 * Non-spanning controls grouped into their map rows. Each row is an
 * independent left-to-right flow, so a wide control in row 1 never widens
 * row 2's first slot (the oracle's rows do not share columns).
 */
function controlRows(controls: readonly ResolvedRibbonControl[]): ResolvedRibbonControl[][] {
  const rows: ResolvedRibbonControl[][] = [];
  for (const control of controls) {
    if (control.rowSpan > 1) continue;
    while (rows.length < control.row) rows.push([]);
    rows[control.row - 1]!.push(control);
  }
  return rows;
}

function RibbonGroup({
  group,
  resolveState,
  controlValues,
  onCommand,
}: RibbonGroupProps): JSX.Element {
  const launcher = group.launcher;
  // The Formulas › Function Library group has its own responsive design (the
  // approved prototype); every other group packs the standard band rows.
  if (group.id === 'Function Library') {
    return (
      <FunctionLibraryGroup
        group={group}
        resolveState={resolveState}
        controlValues={controlValues}
        onCommand={onCommand}
      />
    );
  }
  return (
    <div
      className={['ribbon-group', launcher === undefined ? '' : 'ribbon-group-with-launcher']
        .filter(Boolean)
        .join(' ')}
      role="group"
      aria-label={group.label}
      data-group={group.label}
    >
      <div className="ribbon-group-body">
        {group.controls
          .filter((control) => control.rowSpan > 1)
          .map((control) => (
            <div key={control.key} className="ribbon-cell ribbon-cell-large">
              <RibbonControl
                control={control}
                resolveState={resolveState}
                controlValues={controlValues}
                onCommand={onCommand}
              />
            </div>
          ))}
        <div className="ribbon-group-rows">
          {controlRows(group.controls).map((rowControls, index) => (
            <div className="ribbon-group-row" key={index}>
              {rowControls.map((control) => (
                <div key={control.key} className="ribbon-cell">
                  <RibbonControl
                    control={control}
                    resolveState={resolveState}
                    controlValues={controlValues}
                    onCommand={onCommand}
                  />
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
      {launcher === undefined ? null : (
        <button
          type="button"
          className="ribbon-group-launcher"
          title={launcher.title}
          aria-label={launcher.title}
          disabled={!resolveState(launcher.commandId).enabled}
          data-command={launcher.commandId}
          onClick={() => onCommand(launcher.commandId, launcher.args)}
        >
          <svg viewBox="0 0 8 8" aria-hidden="true" focusable="false">
            <path d="M1 1h2M1 1v2M7 7H5M7 7V5M1 7h2M1 7V5" />
          </svg>
        </button>
      )}
      <div className="ribbon-group-label">{group.label}</div>
    </div>
  );
}

/**
 * Formulas › Function Library, built exactly as the approved prototype
 * (prototype/function-library branch): a tall Insert Function hero on the left
 * behind a thin divider, the nine drop-downs beside it, and the group label
 * centred at the bottom of that column. WIDE puts a 30px icon over each
 * label with the caret under it, all nine in one row. NARROW puts an 18px
 * icon beside the text with the caret hugging it, three columns by three
 * rows. The hero is the only control without a caret.
 */
function FunctionLibraryGroup({
  group,
  resolveState,
  controlValues,
  onCommand,
}: RibbonGroupProps): JSX.Element {
  const wide = useMediaQuery(`(min-width: ${FUNCTION_LIBRARY_WIDE_MIN_WIDTH}px)`);
  const [hero, ...dropDowns] = group.controls;
  return (
    <div
      className={`ribbon-group ribbon-function-library ${wide ? 'wide' : 'narrow'}`}
      role="group"
      aria-label={group.label}
      data-group={group.label}
    >
      <div className="ribbon-library-body">
        <div className="ribbon-library-hero">
          {hero === undefined ? null : (
            <RibbonControl
              control={hero}
              resolveState={resolveState}
              controlValues={controlValues}
              onCommand={onCommand}
            />
          )}
        </div>
        <div className="ribbon-library-divider" />
        <div className="ribbon-library-main">
          <div className="ribbon-library-items">
            {dropDowns.map((control) => (
              <div key={control.key} className="ribbon-library-item">
                <RibbonControl
                  control={control}
                  resolveState={resolveState}
                  controlValues={controlValues}
                  onCommand={onCommand}
                  caret
                />
              </div>
            ))}
          </div>
          <div className="ribbon-group-label">{group.label}</div>
        </div>
      </div>
    </div>
  );
}

interface RibbonLabelProps {
  readonly control: ResolvedRibbonControl;
  /** Draw the drop-down caret beside the label. */
  readonly caret?: boolean;
}

/**
 * A control's label, drawn with its authored line breaks where it has any.
 * The drop-down caret is a sibling of the label, not part of its last line,
 * so a flex row puts it beside the text and a flex column puts it under.
 */
function RibbonLabel({ control, caret = false }: RibbonLabelProps): JSX.Element {
  const lines = control.labelLines ?? [control.label];
  return (
    <>
      <span className="ribbon-label">
        {lines.map((line, index) => (
          <span className="ribbon-label-line" key={`${line}:${index}`}>
            {line}
          </span>
        ))}
      </span>
      {caret ? (
        <span className="ribbon-menu-caret" aria-hidden="true">
          ▾
        </span>
      ) : null}
    </>
  );
}

interface RibbonControlProps {
  readonly control: ResolvedRibbonControl;
  readonly resolveState: CommandStateResolver;
  readonly controlValues: Readonly<Record<string, string>>;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  /** Draw the drop-down caret; the Function Library carries one on every drop-down. */
  readonly caret?: boolean;
}

function RibbonControl({
  control,
  resolveState,
  controlValues,
  onCommand,
  caret = false,
}: RibbonControlProps): JSX.Element {
  const state = resolveState(control.commandId);
  const enabled = !control.disabled && state.enabled;
  const wideClass = control.wide ? ' ribbon-control-wide' : '';
  if (control.kind === 'select') {
    return (
      <RibbonSelect
        control={control}
        value={control.controlId === undefined ? '' : (controlValues[control.controlId] ?? '')}
        enabled={enabled}
        onCommand={onCommand}
      />
    );
  }
  if (control.kind === 'split' || control.kind === 'menu') {
    if (control.picker !== undefined) {
      return (
        <RibbonColorControl
          control={control}
          value={control.controlId === undefined ? '' : (controlValues[control.controlId] ?? '')}
          enabled={enabled}
          onCommand={onCommand}
          wideClass={wideClass}
        />
      );
    }
    return (
      <RibbonMenuControl
        control={control}
        enabled={enabled}
        resolveState={resolveState}
        onCommand={onCommand}
        wideClass={wideClass}
        dropdownOnly={control.kind === 'menu'}
      />
    );
  }
  const checked = control.kind === 'toggle' ? state.checked : undefined;
  const iconOnly = control.iconOnly && control.icon !== undefined;
  return (
    <button
      type="button"
      className={[
        'ribbon-button',
        control.kind === 'toggle' ? 'ribbon-toggle' : '',
        checked === true ? 'active' : '',
        iconOnly ? 'ribbon-button-icon-only' : '',
        wideClass,
      ]
        .filter(Boolean)
        .join(' ')}
      title={control.title}
      aria-label={control.title}
      aria-pressed={checked}
      disabled={!enabled}
      data-command={control.commandId}
      onClick={() => onCommand(control.commandId, control.args)}
    >
      {control.icon === undefined ? null : (
        <span className="ribbon-icon" aria-hidden="true">
          <RibbonIcon name={control.icon} />
        </span>
      )}
      {iconOnly ? null : <RibbonLabel control={control} caret={caret} />}
    </button>
  );
}

interface RibbonColorControlProps {
  readonly control: ResolvedRibbonControl;
  readonly value: string;
  readonly enabled: boolean;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly wideClass: string;
}

/**
 * The text-colour and fill split buttons. The main button still applies the
 * last preset; the drop-down opens the shared colour picker, which commits a
 * swatch, the reset row or a hex value through the existing `format.font` or
 * `format.fill` command. One choice is one undo transaction.
 */
function RibbonColorControl({
  control,
  value,
  enabled,
  onCommand,
  wideClass,
}: RibbonColorControlProps): JSX.Element {
  const [open, setOpen] = useState(false);
  const [popupPosition, setPopupPosition] = useState<{ left: number; top: number } | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const toggleRef = useRef<HTMLButtonElement | null>(null);
  const popupRef = useRef<HTMLDivElement | null>(null);
  const iconOnly = control.iconOnly && control.icon !== undefined;
  const isFill = control.picker === 'fill';

  const closePicker = (restoreFocus: boolean): void => {
    setOpen(false);
    if (restoreFocus) toggleRef.current?.focus();
  };

  const togglePicker = (): void => {
    if (open) {
      closePicker(true);
      return;
    }
    const rect = containerRef.current?.getBoundingClientRect();
    if (rect === undefined) return;
    const width = 336;
    setPopupPosition({
      left: Math.max(6, Math.min(rect.left, window.innerWidth - width)),
      top: rect.bottom + 2,
    });
    setOpen(true);
  };

  useEffect(() => {
    if (!open) return;
    const onDocumentPointerDown = (event: PointerEvent): void => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (containerRef.current?.contains(target) === true) return;
      if (popupRef.current?.contains(target) === true) return;
      setOpen(false);
    };
    document.addEventListener('pointerdown', onDocumentPointerDown, true);
    return () => document.removeEventListener('pointerdown', onDocumentPointerDown, true);
  }, [open]);

  const applyColor = (color: string): void => {
    setOpen(false);
    onCommand(
      control.commandId,
      isFill
        ? { pattern: color === '' ? 'none' : 'solid', color: color === '' ? null : color }
        : { color: color === '' ? null : color },
    );
  };

  return (
    <div
      ref={containerRef}
      className={`ribbon-split${wideClass}`}
      onKeyDown={(event) => {
        if (!open) return;
        // The open picker owns Tab the way the ribbon menus do: the press
        // closes it and returns to its trigger rather than leaving it behind.
        if (event.key === 'Tab') {
          event.preventDefault();
          event.stopPropagation();
          closePicker(true);
          return;
        }
        if (event.key === 'F6') closePicker(false);
      }}
    >
      <button
        type="button"
        className={['ribbon-button', 'ribbon-split-main', iconOnly ? 'ribbon-button-icon-only' : '']
          .filter(Boolean)
          .join(' ')}
        title={control.title}
        aria-label={control.title}
        disabled={!enabled}
        data-command={control.commandId}
        onClick={() => {
          setOpen(false);
          onCommand(control.commandId, control.args);
        }}
      >
        {control.icon === undefined ? null : (
          <span className="ribbon-icon" aria-hidden="true">
            <RibbonIcon name={control.icon} />
          </span>
        )}
        {iconOnly ? null : <span>{control.label}</span>}
      </button>
      <button
        ref={toggleRef}
        type="button"
        className="ribbon-split-toggle"
        title={`${control.title} options`}
        aria-label={`${control.title} options`}
        aria-haspopup="dialog"
        aria-controls={open ? `${control.controlId ?? control.commandId}-popup` : undefined}
        aria-expanded={open}
        disabled={!enabled}
        onClick={togglePicker}
      >
        ▾
      </button>
      {open && popupPosition !== null
        ? createPortal(
            <ColorPickerPopup
              id={control.controlId ?? control.commandId}
              testId={control.controlId ?? control.commandId}
              label={control.title}
              value={value}
              emptyColor={isFill ? null : '#000000'}
              resetLabel={isFill ? 'No Fill' : 'Automatic'}
              mode="apply"
              popupRef={popupRef}
              style={{
                position: 'fixed',
                left: `${popupPosition.left}px`,
                top: `${popupPosition.top}px`,
              }}
              onCommit={applyColor}
              onCancel={() => closePicker(true)}
            />,
            window.document.body,
          )
        : null}
    </div>
  );
}

interface RibbonSelectProps {
  readonly control: ResolvedRibbonControl;
  readonly value: string;
  readonly enabled: boolean;
  readonly onCommand: (commandId: string, args?: unknown) => void;
}

function RibbonSelect({ control, value, enabled, onCommand }: RibbonSelectProps): JSX.Element {
  const options = control.options ?? [];
  const safeValue =
    value === '' ? (options[0] === undefined ? '' : optionValue(options[0])) : value;
  const hasCurrent = options.some((option) => optionValue(option) === safeValue);
  return (
    <label
      className={control.wide ? 'ribbon-select ribbon-control-wide' : 'ribbon-select'}
      data-disabled={enabled ? undefined : 'true'}
    >
      {control.icon === undefined ? null : (
        <span className="ribbon-icon ribbon-select-icon" aria-hidden="true">
          <RibbonIcon name={control.icon} />
        </span>
      )}
      {control.label === '' ? null : <span className="ribbon-select-label">{control.label}</span>}
      <select
        className="ribbon-select-input"
        aria-label={control.title}
        value={safeValue}
        disabled={!enabled}
        data-command={control.commandId}
        onChange={(event) => {
          const option = options.find((candidate) => optionValue(candidate) === event.target.value);
          if (option === undefined) return;
          onCommand(control.commandId, option.args);
        }}
      >
        {hasCurrent || safeValue === '' ? null : <option value={safeValue}>{safeValue}</option>}
        {options.map((option) => (
          <option key={optionValue(option)} value={optionValue(option)}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
  );
}

interface RibbonMenuControlProps {
  readonly control: ResolvedRibbonControl;
  readonly enabled: boolean;
  readonly resolveState: CommandStateResolver;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly wideClass: string;
  /** A `menu` control opens its list only; a `split` control has a main action. */
  readonly dropdownOnly: boolean;
}

/**
 * The split button (main action + dropdown) and the dropdown-only button share
 * one implementation so the menu's placement, keyboard handling and focus
 * restoration stay in one place.
 */
function RibbonMenuControl({
  control,
  enabled,
  resolveState,
  onCommand,
  wideClass,
  dropdownOnly,
}: RibbonMenuControlProps): JSX.Element {
  const [open, setOpen] = useState(false);
  const [menuPosition, setMenuPosition] = useState<{ left: number; top: number } | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const toggleRef = useRef<HTMLButtonElement | null>(null);
  const items = ribbonMenuItems(control.menu ?? [], resolveState);
  const iconOnly = control.iconOnly && control.icon !== undefined;

  const toggleMenu = (): void => {
    if (open) {
      setOpen(false);
      return;
    }
    const rect = containerRef.current?.getBoundingClientRect();
    if (rect === undefined) return;
    const width = 210;
    setMenuPosition({
      left: Math.max(6, Math.min(rect.left, window.innerWidth - width)),
      top: rect.bottom + 2,
    });
    setOpen(true);
  };

  return (
    <div ref={containerRef} className={`ribbon-split${wideClass}`}>
      {dropdownOnly ? null : (
        <button
          type="button"
          className={[
            'ribbon-button',
            'ribbon-split-main',
            iconOnly ? 'ribbon-button-icon-only' : '',
          ]
            .filter(Boolean)
            .join(' ')}
          title={control.title}
          aria-label={control.title}
          disabled={!enabled}
          data-command={control.commandId}
          onClick={() => {
            setOpen(false);
            onCommand(control.commandId, control.args);
          }}
        >
          {control.icon === undefined ? null : (
            <span className="ribbon-icon" aria-hidden="true">
              <RibbonIcon name={control.icon} />
            </span>
          )}
          {iconOnly ? null : <RibbonLabel control={control} />}
        </button>
      )}
      <button
        ref={toggleRef}
        type="button"
        className={dropdownOnly ? 'ribbon-button ribbon-menu-button' : 'ribbon-split-toggle'}
        title={dropdownOnly ? control.title : `${control.title} options`}
        aria-label={dropdownOnly ? control.title : `${control.title} options`}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={!enabled || control.menu === undefined || items.length === 0}
        data-command={dropdownOnly ? control.commandId : undefined}
        onClick={toggleMenu}
      >
        {dropdownOnly ? (
          <>
            {control.icon === undefined ? null : (
              <span className="ribbon-icon" aria-hidden="true">
                <RibbonIcon name={control.icon} />
              </span>
            )}
            <RibbonLabel control={control} caret />
          </>
        ) : (
          '▾'
        )}
      </button>
      {open && menuPosition !== null
        ? createPortal(
            <Menu
              items={items}
              label={`${control.title} options`}
              heading={control.menuHeading}
              className="ribbon-menu"
              style={{
                position: 'fixed',
                left: `${menuPosition.left}px`,
                top: `${menuPosition.top}px`,
              }}
              anchorRef={containerRef}
              onSelect={(item) => {
                setOpen(false);
                // The menu item disappears with the menu; the trigger is the
                // stable invoking control for a dialog's focus restoration.
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

function optionValue(option: RibbonSelectOption): string {
  return option.value ?? option.label;
}
