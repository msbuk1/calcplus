import { useEffect, useRef, useState } from 'react';
import type { CSSProperties, JSX, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import { INDEXED_PALETTE, THEME_PALETTE } from '@model/style';
import { DialogTabs } from './Dialog';
import type { DialogTab } from './Dialog';

/**
 * The shared Excel-style colour picker. Format Cells opens it in draft mode:
 * a choice updates the popup's draft until OK hands the value back. The ribbon
 * opens it in apply mode: a swatch, the reset row or the hex Apply button
 * commits immediately through the owning command. Both callers reuse this one
 * picker; the focused control owns focus restoration and placement.
 */

export const COLOR_POPUP_TABS: readonly DialogTab[] = [
  { id: 'standard', label: 'Standard' },
  { id: 'custom', label: 'Custom' },
];

export const THEME_COLORS: readonly string[] = THEME_PALETTE.map(
  (value) => `#${value.toUpperCase()}`,
);

export const STANDARD_COLORS: readonly string[] = Array.from(
  new Set(INDEXED_PALETTE.map((value) => `#${value.toUpperCase()}`)),
).filter((color) => !THEME_COLORS.includes(color));

interface ColorPaletteProps {
  readonly label: string;
  readonly colors: readonly string[];
  readonly selected: string | null;
  readonly disabled?: boolean;
  readonly onSelect: (color: string) => void;
}

/**
 * A roving-tabindex swatch grid. Only one swatch is in the Tab order; the
 * arrow keys walk the palette, so a large standard palette stays keyboard
 * usable instead of adding every swatch to the dialog's tab cycle.
 */
export function ColorPalette({
  label,
  colors,
  selected,
  disabled = false,
  onSelect,
}: ColorPaletteProps): JSX.Element {
  const [focusIndex, setFocusIndex] = useState(0);
  const buttonsRef = useRef<Array<HTMLButtonElement | null>>([]);
  const selectedIndex = selected === null ? -1 : colors.indexOf(selected);

  useEffect(() => {
    if (selectedIndex >= 0) setFocusIndex(selectedIndex);
  }, [selectedIndex]);

  const moveFocus = (offset: number): void => {
    if (colors.length === 0) return;
    const next = (focusIndex + offset + colors.length) % colors.length;
    setFocusIndex(next);
    buttonsRef.current[next]?.focus();
  };

  return (
    <div className="format-color-palette" role="group" aria-label={label}>
      {colors.map((color, index) => (
        <button
          key={`${color}-${index}`}
          ref={(node) => {
            buttonsRef.current[index] = node;
          }}
          type="button"
          className="format-color-swatch"
          aria-label={`Colour ${color}`}
          aria-pressed={selected === color}
          tabIndex={index === focusIndex ? 0 : -1}
          title={color}
          disabled={disabled}
          onClick={() => {
            setFocusIndex(index);
            onSelect(color);
          }}
          onKeyDown={(event) => {
            if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
              event.preventDefault();
              moveFocus(1);
              return;
            }
            if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
              event.preventDefault();
              moveFocus(-1);
            }
          }}
        >
          <span
            className="format-color-swatch-fill"
            style={{ backgroundColor: color }}
            aria-hidden="true"
          />
        </button>
      ))}
    </div>
  );
}

const CLEAR_COLOR_ERROR =
  'Enter a colour as #RGB, #RRGGBB or #AARRGGBB (alpha is ignored), or leave it blank.';

export function parseColor(value: string): {
  readonly color: string | null;
  readonly error?: string;
} {
  const trimmed = value.trim();
  if (trimmed === '') return { color: null };
  const digits = trimmed.startsWith('#') ? trimmed.slice(1) : trimmed;
  if (!/^(?:[0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(digits)) {
    return { color: null, error: CLEAR_COLOR_ERROR };
  }
  // Alpha is presentation-only for cell styles; the workbook stores RGB.
  const rgb = digits.length === 8 ? digits.slice(2) : digits;
  const expanded =
    rgb.length === 3
      ? rgb
          .split('')
          .map((digit) => `${digit}${digit}`)
          .join('')
      : rgb;
  return { color: `#${expanded}`.toUpperCase() };
}

export interface ColorPickerPopupProps {
  /** The owning control's id; the custom input and its hint derive from it. */
  readonly id: string;
  /** The prefix for `data-testid` hooks; the popup, custom input and error use it. */
  readonly testId: string;
  readonly label?: string;
  /** The current committed value; a blank string means automatic or no fill. */
  readonly value: string;
  /** The swatch shown for an empty value; `null` renders the no-colour mark. */
  readonly emptyColor: string | null;
  readonly resetLabel: string;
  readonly errorTestId?: string;
  readonly disabled?: boolean;
  /**
   * `draft` keeps a local value until OK. `apply` commits every choice at once:
   * a swatch, the reset row or the hex Apply button.
   */
  readonly mode?: 'draft' | 'apply';
  readonly popupRef?: RefObject<HTMLDivElement | null>;
  readonly style?: CSSProperties;
  readonly onCommit: (color: string) => void;
  readonly onCancel: () => void;
}

export function ColorPickerPopup({
  id,
  testId,
  label = 'Colours',
  value,
  emptyColor,
  resetLabel,
  errorTestId,
  disabled = false,
  mode = 'draft',
  popupRef,
  style,
  onCommit,
  onCancel,
}: ColorPickerPopupProps): JSX.Element {
  const [popupTab, setPopupTab] = useState('standard');
  const [draftColor, setDraftColor] = useState(value);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const customInputRef = useRef<HTMLInputElement | null>(null);

  const draft = parseColor(draftColor);
  const current = parseColor(value);
  const draftColorValue = draft.color ?? emptyColor;
  const currentColor = current.color ?? emptyColor;

  const popupId = `${id}-popup`;
  const customId = `${id}-custom`;
  const popupErrorId = `${id}-popup-error`;
  const resolvedErrorTestId = errorTestId ?? `${testId}-error`;

  const setRoot = (node: HTMLDivElement | null): void => {
    rootRef.current = node;
    if (popupRef !== undefined) popupRef.current = node;
  };

  useEffect(() => {
    rootRef.current?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.focus();
  }, []);

  const commit = (): void => {
    if (draft.error !== undefined) {
      setPopupTab('custom');
      requestAnimationFrame(() => customInputRef.current?.focus());
      return;
    }
    onCommit(draft.color ?? '');
  };

  const choose = (color: string): void => {
    if (mode === 'apply') {
      onCommit(color);
      return;
    }
    setDraftColor(color);
  };

  const reset = (): void => {
    if (mode === 'apply') {
      onCommit('');
      return;
    }
    setDraftColor('');
  };

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onCancel();
      return;
    }
    // Enter in the hex field commits the picker; it must not reach a dialog's
    // default OK action behind the popup.
    if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
      event.preventDefault();
      event.stopPropagation();
      commit();
    }
  };

  return (
    <div
      ref={setRoot}
      id={popupId}
      className="format-colors-popup"
      role="dialog"
      aria-label={label}
      data-testid={`${testId}-popup`}
      style={style}
      onKeyDown={onKeyDown}
    >
      <DialogTabs tabs={COLOR_POPUP_TABS} activeId={popupTab} onChange={setPopupTab} />

      {popupTab === 'standard' ? (
        <div className="format-colors-panel" role="tabpanel" aria-label="Standard colours">
          <p className="format-colors-caption">Theme colours</p>
          <ColorPalette
            label="Theme colours"
            colors={THEME_COLORS}
            selected={draft.color}
            disabled={disabled}
            onSelect={choose}
          />
          <p className="format-colors-caption">Standard colours</p>
          <ColorPalette
            label="Standard colours"
            colors={STANDARD_COLORS}
            selected={draft.color}
            disabled={disabled}
            onSelect={choose}
          />
          <button
            type="button"
            className="format-colors-more"
            disabled={disabled}
            onClick={() => setPopupTab('custom')}
          >
            More Colours
          </button>
        </div>
      ) : (
        <div className="format-colors-panel" role="tabpanel" aria-label="Custom colour">
          <div className="format-row format-colors-custom-row">
            <label className="find-label" htmlFor={customId}>
              Hex
            </label>
            <input
              ref={customInputRef}
              id={customId}
              type="text"
              className="find-input"
              value={draftColor}
              spellCheck={false}
              autoComplete="off"
              placeholder="#RGB, #RRGGBB or #AARRGGBB"
              aria-invalid={draft.error !== undefined}
              aria-describedby={draft.error === undefined ? `${customId}-hint` : popupErrorId}
              data-testid={`${testId}-custom`}
              onChange={(event) => setDraftColor(event.target.value)}
            />
          </div>
          {draft.error === undefined ? null : (
            <p
              id={popupErrorId}
              className="format-field-error format-colors-error"
              role="alert"
              data-testid={resolvedErrorTestId}
            >
              {draft.error}
            </p>
          )}
          <p className="format-colors-hint" id={`${customId}-hint`}>
            Accepts #RGB, #RRGGBB or #AARRGGBB. Alpha is ignored.
          </p>
        </div>
      )}

      <div className="format-colors-preview">
        <div className="format-colors-preview-item">
          <span className="format-colors-preview-label">New</span>
          <span
            className="format-colors-preview-swatch"
            data-testid={`${testId}-new`}
            data-empty={draft.error === undefined && draftColorValue === null ? 'true' : undefined}
            data-invalid={draft.error === undefined ? undefined : 'true'}
            style={
              draft.error === undefined && draftColorValue !== null
                ? { backgroundColor: draftColorValue }
                : undefined
            }
            role="img"
            aria-label={
              draft.error === undefined
                ? `New colour ${draft.color ?? resetLabel}`
                : 'New colour, invalid'
            }
          />
          <span className="format-colors-preview-value">
            {draft.error === undefined ? (draft.color ?? resetLabel) : 'Invalid'}
          </span>
        </div>
        <div className="format-colors-preview-item">
          <span className="format-colors-preview-label">Current</span>
          <span
            className="format-colors-preview-swatch"
            data-testid={`${testId}-current`}
            data-empty={current.error === undefined && currentColor === null ? 'true' : undefined}
            data-invalid={current.error === undefined ? undefined : 'true'}
            style={
              current.error === undefined && currentColor !== null
                ? { backgroundColor: currentColor }
                : undefined
            }
            role="img"
            aria-label={
              current.error === undefined
                ? `Current colour ${current.color ?? resetLabel}`
                : 'Current colour, invalid'
            }
          />
          <span className="format-colors-preview-value">
            {current.error === undefined ? (current.color ?? resetLabel) : 'Invalid'}
          </span>
        </div>
      </div>

      <div className="format-colors-actions">
        <button type="button" className="format-colors-clear" disabled={disabled} onClick={reset}>
          {resetLabel}
        </button>
        <div className="format-colors-action-buttons">
          <button type="button" className="dialog-button" onClick={onCancel}>
            Cancel
          </button>
          <button
            type="button"
            className="dialog-button dialog-button-primary"
            disabled={disabled}
            onClick={commit}
          >
            {mode === 'apply' ? 'Apply' : 'OK'}
          </button>
        </div>
      </div>
    </div>
  );
}
