import { useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import { createPortal } from 'react-dom';
import type { RibbonSelectOption, RibbonTabId } from '@commands/inventory';
import { Menu } from './Menu';
import { RIBBON_TABS, menuItemsFor, ribbonGroupsFor, ribbonMenuItems } from './ribbon-model';
import type {
  CommandStateResolver,
  ResolvedRibbonControl,
  ResolvedRibbonGroup,
} from './ribbon-model';

export interface RibbonProps {
  readonly activeTab: RibbonTabId;
  readonly onTabChange: (tab: RibbonTabId) => void;
  readonly onCommand: (commandId: string, args?: unknown) => void;
  readonly resolveState: CommandStateResolver;
  readonly controlValues: Readonly<Record<string, string>>;
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
  ribbonRef,
  onFocus,
}: RibbonProps): JSX.Element {
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const fileButtonRef = useRef<HTMLButtonElement | null>(null);
  const tabRefs = useRef<(HTMLButtonElement | null)[]>([]);
  const fileItems = menuItemsFor('File', resolveState);
  const groups = ribbonGroupsFor(activeTab);

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

function RibbonGroup({
  group,
  resolveState,
  controlValues,
  onCommand,
}: RibbonGroupProps): JSX.Element {
  const launcher = group.launcher;
  return (
    <div
      className={
        launcher === undefined ? 'ribbon-group' : 'ribbon-group ribbon-group-with-launcher'
      }
      role="group"
      aria-label={group.label}
      data-group={group.label}
    >
      <div
        className="ribbon-group-body"
        style={{ gridTemplateRows: `repeat(${group.rowCount}, 24px)` }}
      >
        {group.controls.map((control) => (
          <div
            key={control.key}
            className={control.rowSpan > 1 ? 'ribbon-cell ribbon-cell-large' : 'ribbon-cell'}
            style={{ gridRow: `${control.row} / span ${control.rowSpan}` }}
          >
            <RibbonControl
              control={control}
              resolveState={resolveState}
              controlValues={controlValues}
              onCommand={onCommand}
            />
          </div>
        ))}
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

interface RibbonControlProps {
  readonly control: ResolvedRibbonControl;
  readonly resolveState: CommandStateResolver;
  readonly controlValues: Readonly<Record<string, string>>;
  readonly onCommand: (commandId: string, args?: unknown) => void;
}

function RibbonControl({
  control,
  resolveState,
  controlValues,
  onCommand,
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
  const glyph = alignmentGlyphKind(control);
  return (
    <button
      type="button"
      className={[
        'ribbon-button',
        control.kind === 'toggle' ? 'ribbon-toggle' : '',
        checked === true ? 'active' : '',
        glyph === null ? '' : 'ribbon-button-glyph',
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
      {glyph === null ? (
        control.icon === undefined ? null : (
          <span className="ribbon-icon" aria-hidden="true">
            {control.icon}
          </span>
        )
      ) : (
        <AlignmentGlyph kind={glyph} />
      )}
      {glyph === null ? <span>{control.label}</span> : null}
    </button>
  );
}

type AlignmentGlyphKind =
  | 'align-left'
  | 'align-center'
  | 'align-right'
  | 'align-top'
  | 'align-middle'
  | 'align-bottom'
  | 'wrap-text'
  | 'indent';

/**
 * The six placed alignment commands carry their direction in `args`. The wrap
 * and indent cases are ready for the same command's remaining arguments, which
 * currently have no ribbon placement; the icon set stays in one place.
 */
function alignmentGlyphKind(control: ResolvedRibbonControl): AlignmentGlyphKind | null {
  if (control.commandId !== 'format.alignment' || control.args === undefined) return null;
  const args = control.args;
  if (args.horizontal === 'left') return 'align-left';
  if (args.horizontal === 'center') return 'align-center';
  if (args.horizontal === 'right') return 'align-right';
  if (args.vertical === 'top') return 'align-top';
  if (args.vertical === 'center') return 'align-middle';
  if (args.vertical === 'bottom') return 'align-bottom';
  if (args.wrapText === true) return 'wrap-text';
  if (typeof args.indent === 'number') return 'indent';
  return null;
}

const ALIGNMENT_GLYPH_PATHS: Readonly<Record<AlignmentGlyphKind, readonly string[]>> = {
  'align-left': ['M2.5 3.5h11M2.5 6.5h7M2.5 9.5h11M2.5 12.5h7'],
  'align-center': ['M2.5 3.5h11M4.5 6.5h7M2.5 9.5h11M4.5 12.5h7'],
  'align-right': ['M2.5 3.5h11M6.5 6.5h7M2.5 9.5h11M6.5 12.5h7'],
  'align-top': ['M2.5 3.5h11', 'M8 13V6.5', 'M5.5 8.5 8 6l2.5 2.5'],
  'align-middle': ['M2.5 8h11', 'M5.5 5.5 8 3l2.5 2.5', 'M5.5 10.5 8 13l2.5-2.5'],
  'align-bottom': ['M2.5 12.5h11', 'M8 3v6.5', 'M5.5 7.5 8 10l2.5-2.5'],
  'wrap-text': ['M2.5 3.5h11', 'M2.5 7.5h7.5a2.5 2.5 0 0 1 0 5H7', 'M8.5 10.5 7 12l1.5 1.5'],
  indent: ['M2.5 3.5h11M2.5 7.5h11M2.5 11.5h11', 'M2.5 5.5 5 7.5 2.5 9.5'],
};

function AlignmentGlyph({ kind }: { readonly kind: AlignmentGlyphKind }): JSX.Element {
  return (
    <svg className="ribbon-glyph" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
      {ALIGNMENT_GLYPH_PATHS[kind].map((path) => (
        <path key={path} d={path} />
      ))}
    </svg>
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
    <label className={control.wide ? 'ribbon-select ribbon-control-wide' : 'ribbon-select'}>
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
          className="ribbon-button ribbon-split-main"
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
              {control.icon}
            </span>
          )}
          <span>{control.label}</span>
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
                {control.icon}
              </span>
            )}
            <span>{control.label}</span>
            <span className="ribbon-menu-caret" aria-hidden="true">
              ▾
            </span>
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
