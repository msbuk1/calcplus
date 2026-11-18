import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent } from 'react';
import { getFormatInfo, isValidFormat } from 'numfmt';
import {
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
  INDEXED_PALETTE,
  THEME_PALETTE,
  borderEdge,
} from '@model/style';
import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  HorizontalAlignment,
  UnderlineStyle,
  VerticalAlignment,
} from '@model/types';
import { normalizeFormatCode } from '@view/number-format';
import { Dialog, DialogTabs } from './Dialog';
import type { DialogCloseReason, DialogTab } from './Dialog';

/** The resolved style of the active cell, as the sheet projection reports it. */
export interface FormatCellsInitial {
  readonly font: FontStyle;
  readonly fill: FillStyle;
  readonly border: BorderStyle;
  readonly alignment: AlignmentStyle;
  readonly numberFormat: string | null;
}

type BorderEdgeName = 'left' | 'right' | 'top' | 'bottom';
type BorderPreset = 'custom' | 'outline' | 'inside' | 'all' | 'none';

/**
 * The per-tab payloads the dialog submits through the fixed formatting command
 * IDs. A section is present only when the user changed one of its values.
 */
export interface FormatCellsPayload {
  readonly font?: {
    readonly name: string;
    readonly size: number;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly underlineStyle: UnderlineStyle;
    readonly strike: boolean;
    readonly color: string | null;
  };
  readonly fill?: {
    readonly pattern: string;
    readonly color: string | null;
    readonly backgroundColor: string | null;
  };
  readonly border?: {
    readonly preset?: BorderPreset;
    readonly style?: string;
    readonly color?: string | null;
    readonly edges?: Partial<Record<BorderEdgeName | 'diagonal', BorderEdgeStyle>>;
    readonly diagonalUp?: boolean;
    readonly diagonalDown?: boolean;
  };
  readonly alignment?: {
    readonly horizontal: HorizontalAlignment;
    readonly vertical: VerticalAlignment;
    readonly wrapText: boolean;
    readonly indent: number;
  };
  readonly numberFormat?: {
    readonly code: string | null;
  };
}

export interface FormatCellsOutcome {
  readonly ok: boolean;
  readonly message: string | null;
}

export interface FormatCellsDialogProps {
  /** The active cell's current style; omitted means the workbook defaults. */
  readonly initial?: FormatCellsInitial;
  /** The range the formatting will apply to, shown as a hint. */
  readonly targetLabel?: string;
  readonly onClose: (reason: DialogCloseReason) => void;
  readonly onApply: (payload: FormatCellsPayload) => FormatCellsOutcome;
}

interface DraftBase {
  readonly font: {
    readonly name: string;
    readonly size: string;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly underlineStyle: UnderlineStyle;
    readonly strike: boolean;
    readonly color: string;
  };
  readonly fill: {
    readonly pattern: string;
    readonly color: string;
    readonly backgroundColor: string;
  };
  readonly border: {
    readonly preset: BorderPreset;
    readonly style: string;
    readonly color: string;
    readonly edges: Readonly<Record<BorderEdgeName, boolean>>;
    readonly diagonalUp: boolean;
    readonly diagonalDown: boolean;
  };
  readonly alignment: {
    readonly horizontal: HorizontalAlignment;
    readonly vertical: VerticalAlignment;
    readonly wrapText: boolean;
    readonly indent: string;
  };
  readonly number: {
    readonly category: string;
    readonly code: string;
  };
}

type Draft = DraftBase;

interface DraftErrors {
  readonly fontName?: string;
  readonly fontSize?: string;
  readonly fontColor?: string;
  readonly fillColor?: string;
  readonly fillBackground?: string;
  readonly borderColor?: string;
  readonly borderEdges?: string;
  readonly indent?: string;
  readonly numberCode?: string;
}

const FORMAT_TABS: readonly DialogTab[] = [
  { id: 'font', label: 'Font' },
  { id: 'fill', label: 'Fill' },
  { id: 'border', label: 'Border' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'number', label: 'Number' },
];

const DEFAULT_INITIAL: FormatCellsInitial = {
  font: DEFAULT_FONT,
  fill: DEFAULT_FILL,
  border: DEFAULT_BORDER,
  alignment: DEFAULT_ALIGNMENT,
  numberFormat: null,
};

const FONT_FAMILIES: readonly string[] = [
  'Calibri',
  'Arial',
  'Times New Roman',
  'Courier New',
  'Verdana',
  'Tahoma',
  'Segoe UI',
  'Georgia',
  'Trebuchet MS',
  'Comic Sans MS',
];

/** The family list narrows as the field is typed; opening it shows every family. */
function filterFontFamilies(query: string): readonly string[] {
  const trimmed = query.trim().toLowerCase();
  if (trimmed === '') return FONT_FAMILIES;
  return FONT_FAMILIES.filter((family) => family.toLowerCase().includes(trimmed));
}

function fontOptionId(listId: string, index: number): string {
  return `${listId}-option-${index}`;
}

const UNDERLINE_OPTIONS: readonly { readonly value: UnderlineStyle; readonly label: string }[] = [
  { value: 'none', label: 'None' },
  { value: 'single', label: 'Single' },
  { value: 'double', label: 'Double' },
  { value: 'singleAccounting', label: 'Single accounting' },
  { value: 'doubleAccounting', label: 'Double accounting' },
];

const FILL_PATTERNS: readonly { readonly value: string; readonly label: string }[] = [
  { value: 'none', label: 'None' },
  { value: 'solid', label: 'Solid' },
  { value: 'gray125', label: '12.5% gray' },
  { value: 'gray0625', label: '6.25% gray' },
  { value: 'lightGray', label: 'Light gray' },
  { value: 'mediumGray', label: 'Medium gray' },
  { value: 'darkGray', label: 'Dark gray' },
  { value: 'lightHorizontal', label: 'Light horizontal' },
  { value: 'darkHorizontal', label: 'Dark horizontal' },
  { value: 'lightVertical', label: 'Light vertical' },
  { value: 'darkVertical', label: 'Dark vertical' },
  { value: 'lightDown', label: 'Light downward diagonal' },
  { value: 'darkDown', label: 'Dark downward diagonal' },
  { value: 'lightUp', label: 'Light upward diagonal' },
  { value: 'darkUp', label: 'Dark upward diagonal' },
  { value: 'lightGrid', label: 'Light grid' },
  { value: 'darkGrid', label: 'Dark grid' },
  { value: 'lightTrellis', label: 'Light trellis' },
  { value: 'darkTrellis', label: 'Dark trellis' },
];

const BORDER_STYLES: readonly { readonly value: string; readonly label: string }[] = [
  { value: 'thin', label: 'Thin' },
  { value: 'medium', label: 'Medium' },
  { value: 'thick', label: 'Thick' },
  { value: 'double', label: 'Double' },
  { value: 'dashed', label: 'Dashed' },
  { value: 'dotted', label: 'Dotted' },
  { value: 'dashDot', label: 'Dash dot' },
  { value: 'dashDotDot', label: 'Dash dot dot' },
  { value: 'mediumDashed', label: 'Medium dashed' },
  { value: 'mediumDashDot', label: 'Medium dash dot' },
  { value: 'mediumDashDotDot', label: 'Medium dash dot dot' },
  { value: 'slantDashDot', label: 'Slant dash dot' },
  { value: 'hair', label: 'Hairline' },
];

const BORDER_PRESETS: readonly { readonly value: BorderPreset; readonly label: string }[] = [
  { value: 'custom', label: 'Custom edges' },
  { value: 'outline', label: 'Outline' },
  { value: 'inside', label: 'Inside' },
  { value: 'all', label: 'All' },
  { value: 'none', label: 'None' },
];

const HORIZONTAL_OPTIONS: readonly {
  readonly value: HorizontalAlignment;
  readonly label: string;
}[] = [
  { value: 'general', label: 'General' },
  { value: 'left', label: 'Left' },
  { value: 'center', label: 'Center' },
  { value: 'right', label: 'Right' },
  { value: 'fill', label: 'Fill' },
  { value: 'justify', label: 'Justify' },
  { value: 'centerContinuous', label: 'Center across selection' },
  { value: 'distributed', label: 'Distributed' },
];

const VERTICAL_OPTIONS: readonly {
  readonly value: VerticalAlignment;
  readonly label: string;
}[] = [
  { value: 'top', label: 'Top' },
  { value: 'center', label: 'Center' },
  { value: 'bottom', label: 'Bottom' },
  { value: 'justify', label: 'Justify' },
  { value: 'distributed', label: 'Distributed' },
];

interface NumberCategoryDefinition {
  readonly id: string;
  readonly label: string;
  readonly code: string;
}

const NUMBER_CATEGORIES: readonly NumberCategoryDefinition[] = [
  { id: 'general', label: 'General', code: '' },
  { id: 'number', label: 'Number', code: '#,##0.00' },
  { id: 'currency', label: 'Currency', code: '$#,##0.00_);($#,##0.00)' },
  {
    id: 'accounting',
    label: 'Accounting',
    code: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  },
  { id: 'date', label: 'Date', code: 'm/d/yy' },
  { id: 'time', label: 'Time', code: 'h:mm:ss AM/PM' },
  { id: 'percentage', label: 'Percentage', code: '0.00%' },
  { id: 'fraction', label: 'Fraction', code: '# ??/??' },
  { id: 'scientific', label: 'Scientific', code: '0.00E+00' },
  { id: 'text', label: 'Text', code: '@' },
  { id: 'custom', label: 'Custom', code: '' },
];

const COLOR_POPUP_TABS: readonly DialogTab[] = [
  { id: 'standard', label: 'Standard' },
  { id: 'custom', label: 'Custom' },
];

const THEME_COLORS: readonly string[] = THEME_PALETTE.map((value) => `#${value.toUpperCase()}`);

const STANDARD_COLORS: readonly string[] = Array.from(
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
function ColorPalette({
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

interface ColorFieldProps {
  readonly id: string;
  readonly label: string;
  readonly value: string;
  readonly testId: string;
  readonly error?: string;
  readonly errorTestId: string;
  readonly clearLabel: string;
  /** The swatch shown for an empty value; `null` renders the no-colour mark. */
  readonly emptyColor: string | null;
  readonly disabled?: boolean;
  readonly onChange: (value: string) => void;
}

/**
 * The Excel-style colour control: a swatch and drop-down button that opens a
 * separate Colours popup. The popup owns a draft, so Cancel and Esc leave the
 * dialog value untouched; OK hands the canonical hex (or a blank automatic
 * value) back to the form. The Dialog shell still owns the modal focus trap.
 */
function ColorField({
  id,
  label,
  value,
  testId,
  error,
  errorTestId,
  clearLabel,
  emptyColor,
  disabled = false,
  onChange,
}: ColorFieldProps): JSX.Element {
  const [open, setOpen] = useState(false);
  const [popupTab, setPopupTab] = useState('standard');
  const [draftColor, setDraftColor] = useState('');
  const [openedValue, setOpenedValue] = useState('');
  const [popupPosition, setPopupPosition] = useState<{
    readonly top: number;
    readonly left: number;
  } | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const controlRef = useRef<HTMLButtonElement | null>(null);
  const popupRef = useRef<HTMLDivElement | null>(null);
  const customInputRef = useRef<HTMLInputElement | null>(null);

  const current = parseColor(value);
  const currentColor = current.color ?? emptyColor;
  const draft = parseColor(draftColor);
  const draftColorValue = draft.color ?? emptyColor;
  const opened = parseColor(openedValue);
  const openedColor = opened.color ?? emptyColor;

  const popupId = `${id}-popup`;
  const customId = `${id}-custom`;
  const rowErrorId = `${id}-error`;
  const popupErrorId = `${id}-popup-error`;

  // The popup is not a second modal: focus enters it on open, returns to the
  // control on OK/Cancel/Esc, and the shell's trap still contains everything.
  useEffect(() => {
    if (!open) return;
    const popup = popupRef.current;
    if (popup === null) return;
    popup.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.focus();
  }, [open]);

  // The popup floats over the form, so opening a colour never reflows the
  // field below it. Place it before paint, and follow dialog scrolling.
  useLayoutEffect(() => {
    if (!open) return;
    const place = (): void => {
      const control = controlRef.current;
      const popup = popupRef.current;
      if (control === null || popup === null) return;
      const anchor = control.getBoundingClientRect();
      const gap = 4;
      const edge = 8;
      const width = popup.offsetWidth;
      const height = popup.offsetHeight;
      const maxLeft = Math.max(edge, window.innerWidth - width - edge);
      const left = Math.min(Math.max(anchor.left, edge), maxLeft);
      let top = anchor.bottom + gap;
      if (top + height > window.innerHeight - edge) {
        const above = anchor.top - height - gap;
        top = above >= edge ? above : Math.max(edge, window.innerHeight - height - edge);
      }
      setPopupPosition((current) =>
        current !== null && current.top === top && current.left === left ? current : { top, left },
      );
    };
    place();
    window.addEventListener('resize', place);
    document.addEventListener('scroll', place, true);
    return () => {
      window.removeEventListener('resize', place);
      document.removeEventListener('scroll', place, true);
    };
  }, [open, popupTab]);

  useEffect(() => {
    if (!open) return;
    const onDocumentPointerDown = (event: PointerEvent): void => {
      const root = rootRef.current;
      if (root !== null && event.target instanceof Node && root.contains(event.target)) return;
      setOpen(false);
    };
    document.addEventListener('pointerdown', onDocumentPointerDown, true);
    return () => document.removeEventListener('pointerdown', onDocumentPointerDown, true);
  }, [open]);

  const closePopup = (restoreFocus: boolean): void => {
    setOpen(false);
    if (restoreFocus) controlRef.current?.focus();
  };

  const openPopup = (): void => {
    setOpenedValue(value);
    setDraftColor(value);
    setPopupTab('standard');
    setPopupPosition(null);
    setOpen(true);
  };

  const commit = (): void => {
    if (draft.error !== undefined) {
      setPopupTab('custom');
      requestAnimationFrame(() => customInputRef.current?.focus());
      return;
    }
    onChange(draft.color ?? '');
    closePopup(true);
  };

  return (
    <div
      ref={rootRef}
      className="format-color-field"
      data-dialog-popup-open={open ? 'true' : undefined}
      onBlur={(event) => {
        const next = event.relatedTarget;
        if (next instanceof Node && event.currentTarget.contains(next)) return;
        if (open) closePopup(false);
      }}
      onKeyDown={(event) => {
        if (!open) return;
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          closePopup(true);
          return;
        }
        // Enter in the hex field commits the popup; it must not reach the
        // shell's default OK and apply the whole Format Cells dialog.
        if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
          event.preventDefault();
          event.stopPropagation();
          commit();
        }
      }}
    >
      <div className="format-row">
        <label className="find-label" htmlFor={id}>
          {label}
        </label>
        <button
          ref={controlRef}
          id={id}
          type="button"
          className="format-color-control"
          aria-haspopup="dialog"
          aria-expanded={open}
          aria-controls={open ? popupId : undefined}
          aria-describedby={!open && error !== undefined ? rowErrorId : undefined}
          data-testid={testId}
          disabled={disabled}
          onClick={() => {
            if (open) {
              closePopup(true);
            } else {
              openPopup();
            }
          }}
        >
          <span
            className="format-color-preview"
            data-empty={currentColor === null ? 'true' : undefined}
            style={currentColor === null ? undefined : { backgroundColor: currentColor }}
            aria-hidden="true"
          />
          <span className="format-color-value">{current.color ?? clearLabel}</span>
          <span className="format-color-arrow" aria-hidden="true">
            ▾
          </span>
        </button>
      </div>

      {open ? (
        <div
          ref={popupRef}
          id={popupId}
          className="format-colors-popup"
          role="dialog"
          aria-label="Colours"
          data-testid={`${testId}-popup`}
          style={
            popupPosition === null
              ? { visibility: 'hidden' }
              : { top: popupPosition.top, left: popupPosition.left }
          }
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
                onSelect={(color) => setDraftColor(color)}
              />
              <p className="format-colors-caption">Standard colours</p>
              <ColorPalette
                label="Standard colours"
                colors={STANDARD_COLORS}
                selected={draft.color}
                disabled={disabled}
                onSelect={(color) => setDraftColor(color)}
              />
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
                  data-testid={errorTestId}
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
                data-empty={
                  draft.error === undefined && draftColorValue === null ? 'true' : undefined
                }
                data-invalid={draft.error === undefined ? undefined : 'true'}
                style={
                  draft.error === undefined && draftColorValue !== null
                    ? { backgroundColor: draftColorValue }
                    : undefined
                }
                role="img"
                aria-label={
                  draft.error === undefined
                    ? `New colour ${draft.color ?? clearLabel}`
                    : 'New colour, invalid'
                }
              />
              <span className="format-colors-preview-value">
                {draft.error === undefined ? (draft.color ?? clearLabel) : 'Invalid'}
              </span>
            </div>
            <div className="format-colors-preview-item">
              <span className="format-colors-preview-label">Current</span>
              <span
                className="format-colors-preview-swatch"
                data-empty={opened.error === undefined && openedColor === null ? 'true' : undefined}
                data-invalid={opened.error === undefined ? undefined : 'true'}
                style={
                  opened.error === undefined && openedColor !== null
                    ? { backgroundColor: openedColor }
                    : undefined
                }
                role="img"
                aria-label={
                  opened.error === undefined
                    ? `Current colour ${opened.color ?? clearLabel}`
                    : 'Current colour, invalid'
                }
              />
              <span className="format-colors-preview-value">
                {opened.error === undefined ? (opened.color ?? clearLabel) : 'Invalid'}
              </span>
            </div>
          </div>

          <div className="format-colors-actions">
            <button
              type="button"
              className="format-colors-clear"
              disabled={disabled}
              onClick={() => setDraftColor('')}
            >
              {clearLabel}
            </button>
            <div className="format-colors-action-buttons">
              <button type="button" className="dialog-button" onClick={() => closePopup(true)}>
                Cancel
              </button>
              <button
                type="button"
                className="dialog-button dialog-button-primary"
                onClick={commit}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {open || error === undefined ? null : (
        <p id={rowErrorId} className="format-field-error" role="alert" data-testid={errorTestId}>
          {error}
        </p>
      )}
    </div>
  );
}

/**
 * Format Cells owns the form and validation only. It submits every changed tab
 * through the fixed formatting IDs; the application commits them with the
 * target range as one atomic undo transaction. The Dialog shell owns the focus
 * trap, Esc and the restored invoker focus.
 */
export function FormatCellsDialog({
  initial = DEFAULT_INITIAL,
  targetLabel,
  onClose,
  onApply,
}: FormatCellsDialogProps): JSX.Element {
  const [draft, setDraft] = useState<Draft>(() => draftFromInitial(initial));
  const [committed, setCommitted] = useState<Draft>(() => draftFromInitial(initial));
  const [activeTab, setActiveTab] = useState('font');
  const [errors, setErrors] = useState<DraftErrors>({});
  const [dispatchError, setDispatchError] = useState<string | null>(null);
  const [appliedMessage, setAppliedMessage] = useState<string | null>(null);
  const [fontListOpen, setFontListOpen] = useState(false);
  const [fontFilter, setFontFilter] = useState('');
  const [fontActiveIndex, setFontActiveIndex] = useState(-1);

  const ids = {
    fontName: useId(),
    fontSize: useId(),
    underline: useId(),
    fontColor: useId(),
    fillPattern: useId(),
    fillColor: useId(),
    fillBackground: useId(),
    borderPreset: useId(),
    borderStyle: useId(),
    borderColor: useId(),
    horizontal: useId(),
    vertical: useId(),
    indent: useId(),
    numberCategory: useId(),
    numberCode: useId(),
    fontList: useId(),
  };
  const fontNameRef = useRef<HTMLInputElement | null>(null);
  const fontComboRef = useRef<HTMLDivElement | null>(null);

  // The shell focuses its default action first. A feature dialog puts the caret
  // in its primary field instead, after that effect has run.
  useEffect(() => {
    const handle = requestAnimationFrame(() => {
      const input = fontNameRef.current;
      if (input === null) return;
      input.focus();
      input.select();
    });
    return () => cancelAnimationFrame(handle);
  }, []);

  // An open font list closes on a pointer press outside the combobox. This is
  // popup-local state; it never closes the dialog itself.
  useEffect(() => {
    if (!fontListOpen) return;
    const onDocumentPointerDown = (event: PointerEvent): void => {
      const root = fontComboRef.current;
      if (root !== null && event.target instanceof Node && root.contains(event.target)) return;
      setFontListOpen(false);
    };
    document.addEventListener('pointerdown', onDocumentPointerDown, true);
    return () => document.removeEventListener('pointerdown', onDocumentPointerDown, true);
  }, [fontListOpen]);

  const update = (patch: Partial<Draft>): void => {
    setDraft((current) => ({ ...current, ...patch }));
    setErrors({});
    setDispatchError(null);
    setAppliedMessage(null);
  };

  const filteredFontFamilies = filterFontFamilies(fontFilter);
  const activeFontFamily =
    fontActiveIndex >= 0 && fontActiveIndex < filteredFontFamilies.length
      ? filteredFontFamilies[fontActiveIndex]
      : undefined;

  const openFontList = (filterCurrent = false): void => {
    const query = filterCurrent ? draft.font.name : '';
    setFontFilter(query);
    const families = filterFontFamilies(query);
    setFontActiveIndex(families.findIndex((family) => family === draft.font.name));
    setFontListOpen(true);
  };

  const chooseFontFamily = (family: string): void => {
    update({ font: { ...draft.font, name: family } });
    setFontFilter('');
    setFontListOpen(false);
  };

  const onFontKeyDown = (event: ReactKeyboardEvent<HTMLInputElement>): void => {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      if (!fontListOpen) {
        openFontList(true);
        return;
      }
      const count = filteredFontFamilies.length;
      if (count === 0) return;
      setFontActiveIndex((current) => {
        if (current < 0) return event.key === 'ArrowDown' ? 0 : count - 1;
        const offset = event.key === 'ArrowDown' ? 1 : -1;
        return (current + offset + count) % count;
      });
      return;
    }
    if (event.key === 'Enter' && fontListOpen) {
      // Picking a family must not reach the shell's default OK action.
      event.preventDefault();
      event.stopPropagation();
      const family = activeFontFamily ?? filteredFontFamilies[0];
      if (family === undefined) {
        setFontListOpen(false);
        return;
      }
      chooseFontFamily(family);
      return;
    }
    if (event.key === 'Escape' && fontListOpen) {
      event.preventDefault();
      event.stopPropagation();
      setFontListOpen(false);
      return;
    }
    if (event.key === 'Tab') setFontListOpen(false);
  };

  const changeTab = (id: string): void => {
    setFontListOpen(false);
    setActiveTab(id);
  };

  const submit = (closeAfter: boolean): void => {
    const result = buildPayload(draft, committed);
    if ('errors' in result) {
      setErrors(result.errors);
      setDispatchError(null);
      setAppliedMessage(null);
      return;
    }
    setErrors({});
    setDispatchError(null);
    if (Object.keys(result.payload).length === 0) {
      if (closeAfter) onClose('dismiss');
      return;
    }
    const outcome = onApply(result.payload);
    if (!outcome.ok) {
      setDispatchError(outcome.message ?? 'The formatting could not be applied.');
      setAppliedMessage(null);
      return;
    }
    setCommitted(draft);
    if (closeAfter) {
      onClose('dismiss');
      return;
    }
    setAppliedMessage('Formatting applied.');
  };

  const fieldError = (message: string | undefined, testId: string): JSX.Element | null =>
    message === undefined ? null : (
      <p className="format-field-error" role="alert" data-testid={testId}>
        {message}
      </p>
    );

  return (
    <Dialog
      title="Format Cells"
      onClose={onClose}
      defaultAction={{ label: 'OK', onSelect: () => submit(true) }}
      secondaryAction={{ label: 'Cancel', onSelect: () => onClose('dismiss') }}
      className="format-dialog"
    >
      <DialogTabs tabs={FORMAT_TABS} activeId={activeTab} onChange={changeTab} />

      {targetLabel === undefined ? null : (
        <p className="dialog-hint format-target">Applies to {targetLabel}.</p>
      )}

      {activeTab === 'font' ? (
        <div className="format-panel" role="tabpanel" aria-label="Font options">
          <div
            className="format-combo-field"
            ref={fontComboRef}
            data-dialog-popup-open={fontListOpen ? 'true' : undefined}
            onBlur={(event) => {
              const next = event.relatedTarget;
              if (next instanceof Node && event.currentTarget.contains(next)) return;
              setFontListOpen(false);
            }}
          >
            <div className="format-row">
              <label className="find-label" htmlFor={ids.fontName}>
                Font
              </label>
              <div className="format-combobox">
                <input
                  ref={fontNameRef}
                  id={ids.fontName}
                  type="text"
                  className="find-input format-combobox-input"
                  role="combobox"
                  aria-expanded={fontListOpen}
                  aria-controls={ids.fontList}
                  aria-autocomplete="list"
                  aria-haspopup="listbox"
                  aria-activedescendant={
                    fontListOpen && activeFontFamily !== undefined
                      ? fontOptionId(ids.fontList, fontActiveIndex)
                      : undefined
                  }
                  value={draft.font.name}
                  spellCheck={false}
                  autoComplete="off"
                  data-testid="format-font-name"
                  aria-invalid={errors.fontName !== undefined}
                  onClick={() => {
                    if (!fontListOpen) openFontList();
                  }}
                  onChange={(event) => {
                    const value = event.target.value;
                    update({ font: { ...draft.font, name: value } });
                    // Typing narrows an already-open list; typing with the
                    // list closed never opens it or shifts the form layout.
                    if (fontListOpen) {
                      setFontFilter(value);
                      setFontActiveIndex(-1);
                    }
                  }}
                  onKeyDown={onFontKeyDown}
                />
                <button
                  type="button"
                  className="format-combobox-toggle"
                  tabIndex={-1}
                  aria-label={fontListOpen ? 'Hide font families' : 'Show font families'}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => {
                    if (fontListOpen) {
                      setFontListOpen(false);
                    } else {
                      openFontList();
                    }
                    fontNameRef.current?.focus();
                  }}
                >
                  <span aria-hidden="true">▾</span>
                </button>
              </div>
            </div>
            {fontListOpen ? (
              <div className="format-combobox-popup">
                <div
                  id={ids.fontList}
                  className="format-combobox-list"
                  role="listbox"
                  aria-label="Font families"
                >
                  {filteredFontFamilies.length === 0 ? (
                    <p className="format-combobox-empty">No matching fonts.</p>
                  ) : (
                    filteredFontFamilies.map((family, index) => (
                      <div
                        key={family}
                        id={fontOptionId(ids.fontList, index)}
                        className={
                          index === fontActiveIndex
                            ? 'format-combobox-option format-combobox-option-active'
                            : 'format-combobox-option'
                        }
                        role="option"
                        aria-selected={family === draft.font.name}
                        onMouseDown={(event) => event.preventDefault()}
                        onClick={() => chooseFontFamily(family)}
                      >
                        {family}
                      </div>
                    ))
                  )}
                </div>
              </div>
            ) : null}
          </div>
          {fieldError(errors.fontName, 'format-error-font-name')}

          <div className="format-row">
            <label className="find-label" htmlFor={ids.fontSize}>
              Size
            </label>
            <input
              id={ids.fontSize}
              type="number"
              className="find-input format-number-input"
              min={1}
              max={409}
              step={0.5}
              value={draft.font.size}
              data-testid="format-font-size"
              aria-invalid={errors.fontSize !== undefined}
              onChange={(event) => update({ font: { ...draft.font, size: event.target.value } })}
            />
          </div>
          {fieldError(errors.fontSize, 'format-error-font-size')}

          <fieldset className="format-fieldset">
            <legend>Style</legend>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.font.bold}
                data-testid="format-font-bold"
                onChange={(event) =>
                  update({ font: { ...draft.font, bold: event.target.checked } })
                }
              />
              Bold
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.font.italic}
                data-testid="format-font-italic"
                onChange={(event) =>
                  update({ font: { ...draft.font, italic: event.target.checked } })
                }
              />
              Italic
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.font.strike}
                data-testid="format-font-strike"
                onChange={(event) =>
                  update({ font: { ...draft.font, strike: event.target.checked } })
                }
              />
              Strikethrough
            </label>
          </fieldset>

          <div className="format-row">
            <label className="find-label" htmlFor={ids.underline}>
              Underline
            </label>
            <select
              id={ids.underline}
              className="find-select"
              value={draft.font.underlineStyle}
              data-testid="format-font-underline"
              onChange={(event) =>
                update({
                  font: {
                    ...draft.font,
                    underlineStyle: underlineFromSelect(event.target.value),
                  },
                })
              }
            >
              {UNDERLINE_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </div>

          <ColorField
            id={ids.fontColor}
            label="Font colour"
            value={draft.font.color}
            testId="format-font-color"
            error={errors.fontColor}
            errorTestId="format-error-font-color"
            clearLabel="Automatic"
            emptyColor="#000000"
            onChange={(color) => update({ font: { ...draft.font, color } })}
          />
        </div>
      ) : null}

      {activeTab === 'fill' ? (
        <div className="format-panel" role="tabpanel" aria-label="Fill options">
          <div className="format-row">
            <label className="find-label" htmlFor={ids.fillPattern}>
              Pattern
            </label>
            <select
              id={ids.fillPattern}
              className="find-select"
              value={draft.fill.pattern}
              data-testid="format-fill-pattern"
              onChange={(event) => update({ fill: { ...draft.fill, pattern: event.target.value } })}
            >
              {FILL_PATTERNS.map((pattern) => (
                <option key={pattern.value} value={pattern.value}>
                  {pattern.label}
                </option>
              ))}
            </select>
          </div>

          <ColorField
            id={ids.fillColor}
            label="Fill colour"
            value={draft.fill.color}
            testId="format-fill-color"
            error={errors.fillColor}
            errorTestId="format-error-fill-color"
            clearLabel="No colour"
            emptyColor={null}
            onChange={(color) => update({ fill: { ...draft.fill, color } })}
          />

          <ColorField
            id={ids.fillBackground}
            label="Pattern background"
            value={draft.fill.backgroundColor}
            testId="format-fill-background"
            error={errors.fillBackground}
            errorTestId="format-error-fill-background"
            clearLabel="No colour"
            emptyColor={null}
            onChange={(color) => update({ fill: { ...draft.fill, backgroundColor: color } })}
          />
          <p className="dialog-hint">
            A solid fill uses the fill colour. Pattern fills also use the pattern background.
          </p>
        </div>
      ) : null}

      {activeTab === 'border' ? (
        <div className="format-panel" role="tabpanel" aria-label="Border options">
          <div className="format-row">
            <label className="find-label" htmlFor={ids.borderPreset}>
              Preset
            </label>
            <select
              id={ids.borderPreset}
              className="find-select"
              value={draft.border.preset}
              data-testid="format-border-preset"
              onChange={(event) =>
                update({
                  border: {
                    ...draft.border,
                    preset: borderPresetFromSelect(event.target.value),
                  },
                })
              }
            >
              {BORDER_PRESETS.map((preset) => (
                <option key={preset.value} value={preset.value}>
                  {preset.label}
                </option>
              ))}
            </select>
          </div>

          <div className="format-row">
            <label className="find-label" htmlFor={ids.borderStyle}>
              Line style
            </label>
            <select
              id={ids.borderStyle}
              className="find-select"
              value={draft.border.style}
              data-testid="format-border-style"
              disabled={draft.border.preset === 'none'}
              onChange={(event) =>
                update({ border: { ...draft.border, style: event.target.value } })
              }
            >
              {BORDER_STYLES.map((style) => (
                <option key={style.value} value={style.value}>
                  {style.label}
                </option>
              ))}
            </select>
          </div>

          <ColorField
            id={ids.borderColor}
            label="Border colour"
            value={draft.border.color}
            testId="format-border-color"
            error={errors.borderColor}
            errorTestId="format-error-border-color"
            clearLabel="Automatic"
            emptyColor="#000000"
            disabled={draft.border.preset === 'none'}
            onChange={(color) => update({ border: { ...draft.border, color } })}
          />

          <fieldset className="format-fieldset">
            <legend>Edges</legend>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.edges.left}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-left"
                onChange={(event) =>
                  update({
                    border: {
                      ...draft.border,
                      edges: { ...draft.border.edges, left: event.target.checked },
                    },
                  })
                }
              />
              Left
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.edges.right}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-right"
                onChange={(event) =>
                  update({
                    border: {
                      ...draft.border,
                      edges: { ...draft.border.edges, right: event.target.checked },
                    },
                  })
                }
              />
              Right
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.edges.top}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-top"
                onChange={(event) =>
                  update({
                    border: {
                      ...draft.border,
                      edges: { ...draft.border.edges, top: event.target.checked },
                    },
                  })
                }
              />
              Top
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.edges.bottom}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-bottom"
                onChange={(event) =>
                  update({
                    border: {
                      ...draft.border,
                      edges: { ...draft.border.edges, bottom: event.target.checked },
                    },
                  })
                }
              />
              Bottom
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.diagonalUp}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-diagonal-up"
                onChange={(event) =>
                  update({ border: { ...draft.border, diagonalUp: event.target.checked } })
                }
              />
              Diagonal up
            </label>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.border.diagonalDown}
                disabled={draft.border.preset !== 'custom'}
                data-testid="format-border-diagonal-down"
                onChange={(event) =>
                  update({ border: { ...draft.border, diagonalDown: event.target.checked } })
                }
              />
              Diagonal down
            </label>
          </fieldset>
          {fieldError(errors.borderEdges, 'format-error-border-edges')}
          <p className="dialog-hint">
            Outline, inside and all apply to the whole target. Custom edges set each side
            independently; a checked diagonal uses the same line style.
          </p>
        </div>
      ) : null}

      {activeTab === 'alignment' ? (
        <div className="format-panel" role="tabpanel" aria-label="Alignment options">
          <div className="format-row">
            <label className="find-label" htmlFor={ids.horizontal}>
              Horizontal
            </label>
            <select
              id={ids.horizontal}
              className="find-select"
              value={draft.alignment.horizontal}
              data-testid="format-alignment-horizontal"
              onChange={(event) =>
                update({
                  alignment: {
                    ...draft.alignment,
                    horizontal: horizontalFromSelect(event.target.value),
                  },
                })
              }
            >
              {HORIZONTAL_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </div>

          <div className="format-row">
            <label className="find-label" htmlFor={ids.vertical}>
              Vertical
            </label>
            <select
              id={ids.vertical}
              className="find-select"
              value={draft.alignment.vertical}
              data-testid="format-alignment-vertical"
              onChange={(event) =>
                update({
                  alignment: {
                    ...draft.alignment,
                    vertical: verticalFromSelect(event.target.value),
                  },
                })
              }
            >
              {VERTICAL_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </div>

          <div className="format-row">
            <label className="find-label" htmlFor={ids.indent}>
              Indent
            </label>
            <input
              id={ids.indent}
              type="number"
              className="find-input format-number-input"
              min={0}
              max={250}
              step={1}
              value={draft.alignment.indent}
              data-testid="format-alignment-indent"
              aria-invalid={errors.indent !== undefined}
              onChange={(event) =>
                update({ alignment: { ...draft.alignment, indent: event.target.value } })
              }
            />
          </div>
          {fieldError(errors.indent, 'format-error-indent')}

          <fieldset className="format-fieldset">
            <legend>Text control</legend>
            <label className="format-check">
              <input
                type="checkbox"
                checked={draft.alignment.wrapText}
                data-testid="format-alignment-wrap"
                onChange={(event) =>
                  update({
                    alignment: { ...draft.alignment, wrapText: event.target.checked },
                  })
                }
              />
              Wrap text
            </label>
          </fieldset>
          <p className="dialog-hint">
            Imported shrink-to-fit and text rotation are preserved; authoring those controls is v2.
          </p>
        </div>
      ) : null}

      {activeTab === 'number' ? (
        <div className="format-panel" role="tabpanel" aria-label="Number options">
          <div className="format-row">
            <label className="find-label" htmlFor={ids.numberCategory}>
              Category
            </label>
            <select
              id={ids.numberCategory}
              className="find-select"
              value={draft.number.category}
              data-testid="format-number-category"
              onChange={(event) => {
                const category = numberCategoryFor(event.target.value);
                update({
                  number: {
                    category: category.id,
                    code: category.id === 'custom' ? draft.number.code : category.code,
                  },
                });
              }}
            >
              {NUMBER_CATEGORIES.map((category) => (
                <option key={category.id} value={category.id}>
                  {category.label}
                </option>
              ))}
            </select>
          </div>

          <div className="format-row">
            <label className="find-label" htmlFor={ids.numberCode}>
              Format code
            </label>
            <input
              id={ids.numberCode}
              type="text"
              className="find-input"
              value={draft.number.code}
              spellCheck={false}
              autoComplete="off"
              placeholder="0.00"
              data-testid="format-number-code"
              aria-invalid={errors.numberCode !== undefined}
              onChange={(event) => {
                const code = event.target.value;
                update({
                  number: { category: categoryForCode(code), code },
                });
              }}
            />
          </div>
          {fieldError(errors.numberCode, 'format-error-number-code')}
          <p className="dialog-hint">
            Built-in categories fill the code; the field accepts Excel's full custom format grammar.
            An invalid code is rejected before any command is sent.
          </p>
        </div>
      ) : null}

      {Object.keys(errors).length > 0 ? (
        <p className="dialog-error" role="alert" data-testid="format-error-summary">
          Fix the highlighted fields before applying.
        </p>
      ) : null}
      {dispatchError === null ? null : (
        <p className="dialog-error" role="alert" data-testid="format-dispatch-error">
          {dispatchError}
        </p>
      )}
      {appliedMessage === null ? null : (
        <p className="dialog-message" role="status" data-testid="format-applied">
          {appliedMessage}
        </p>
      )}

      <div className="format-actions">
        <button
          type="button"
          className="dialog-button"
          data-testid="format-apply"
          onClick={() => submit(false)}
        >
          Apply
        </button>
      </div>
    </Dialog>
  );
}

function draftFromInitial(initial: FormatCellsInitial): Draft {
  const borderSample = borderSampleFor(initial.border);
  const code = initial.numberFormat ?? '';
  return {
    font: {
      name: initial.font.name,
      size: String(initial.font.size),
      bold: initial.font.bold,
      italic: initial.font.italic,
      underlineStyle: initial.font.underlineStyle,
      strike: initial.font.strike,
      color: initial.font.color ?? '',
    },
    fill: {
      pattern: initial.fill.pattern === '' ? 'none' : initial.fill.pattern,
      color: initial.fill.color ?? '',
      backgroundColor: initial.fill.backgroundColor ?? '',
    },
    border: {
      preset: 'custom',
      style: borderSample.style,
      color: borderSample.color ?? '',
      edges: {
        left: initial.border.left.style !== 'none',
        right: initial.border.right.style !== 'none',
        top: initial.border.top.style !== 'none',
        bottom: initial.border.bottom.style !== 'none',
      },
      diagonalUp: initial.border.diagonalUp,
      diagonalDown: initial.border.diagonalDown,
    },
    alignment: {
      horizontal: initial.alignment.horizontal,
      vertical: initial.alignment.vertical,
      wrapText: initial.alignment.wrapText,
      indent: String(initial.alignment.indent),
    },
    number: {
      category: categoryForCode(code),
      code,
    },
  };
}

function borderSampleFor(border: BorderStyle): BorderEdgeStyle {
  const edges = [border.left, border.right, border.top, border.bottom];
  return (
    edges.find((edge) => edge.style !== 'none' && edge.style !== '') ?? {
      style: 'thin',
      color: null,
      weight: 'thin',
    }
  );
}

type PayloadResult = { readonly payload: FormatCellsPayload } | { readonly errors: DraftErrors };

function buildPayload(draft: Draft, baseline: Draft): PayloadResult {
  const errors: {
    fontName?: string;
    fontSize?: string;
    fontColor?: string;
    fillColor?: string;
    fillBackground?: string;
    borderColor?: string;
    borderEdges?: string;
    indent?: string;
    numberCode?: string;
  } = {};

  const fontName = draft.font.name.trim();
  if (fontName === '') errors.fontName = 'Enter a font name.';
  const fontSize = parseNumber(draft.font.size);
  if (fontSize === null || fontSize < 1 || fontSize > 409) {
    errors.fontSize = 'Enter a font size between 1 and 409.';
  }
  const fontColor = parseColor(draft.font.color);
  if (fontColor.error !== undefined) errors.fontColor = fontColor.error;
  const fillColor = parseColor(draft.fill.color);
  if (fillColor.error !== undefined) errors.fillColor = fillColor.error;
  const fillBackground = parseColor(draft.fill.backgroundColor);
  if (fillBackground.error !== undefined) errors.fillBackground = fillBackground.error;
  const borderColor = parseColor(draft.border.color);
  if (borderColor.error !== undefined) errors.borderColor = borderColor.error;

  const indent = parseNumber(draft.alignment.indent);
  if (indent === null || !Number.isInteger(indent) || indent < 0 || indent > 250) {
    errors.indent = 'Enter an indent between 0 and 250.';
  }

  // The active cell can carry an imported code that `numfmt` does not support;
  // the grid already renders it as General (number-formats §8). Only validate
  // the Number tab when its code actually changed, so formatting another tab
  // never destroys or rejects that imported code.
  const numberChanged =
    normalizedNumberCode(draft.number.code) !== normalizedNumberCode(baseline.number.code);
  let numberCode: string | null = null;
  if (numberChanged) {
    const parsedNumberCode = parseNumberFormatCode(draft.number.code);
    if (parsedNumberCode.error !== undefined) errors.numberCode = parsedNumberCode.error;
    numberCode = parsedNumberCode.code;
  }

  if (
    draft.border.preset === 'custom' &&
    !borderDraftEqual(draft.border, baseline.border) &&
    !hasAnyBorderEdge(draft.border)
  ) {
    errors.borderEdges = 'Choose at least one edge, or use the None preset.';
  }

  if (Object.keys(errors).length > 0) return { errors };

  const font = {
    name: fontName,
    size: fontSize!,
    bold: draft.font.bold,
    italic: draft.font.italic,
    underlineStyle: draft.font.underlineStyle,
    strike: draft.font.strike,
    color: fontColor.color,
  };
  const fill = {
    pattern: draft.fill.pattern,
    color: fillColor.color,
    backgroundColor: fillBackground.color,
  };
  const border = borderPayloadFor(
    draft.border,
    baseline.border,
    borderColor.color,
    parseColor(baseline.border.color).color,
  );
  const alignment = {
    horizontal: draft.alignment.horizontal,
    vertical: draft.alignment.vertical,
    wrapText: draft.alignment.wrapText,
    indent: indent!,
  };

  const baselineResult = buildBaselinePayload(baseline);
  const payload: {
    font?: FormatCellsPayload['font'];
    fill?: FormatCellsPayload['fill'];
    border?: FormatCellsPayload['border'];
    alignment?: FormatCellsPayload['alignment'];
    numberFormat?: FormatCellsPayload['numberFormat'];
  } = {};
  if (!sameJson(font, baselineResult.font)) payload.font = font;
  if (!sameJson(fill, baselineResult.fill)) payload.fill = fill;
  if (Object.keys(border).length > 0 && !sameJson(border, baselineResult.border)) {
    payload.border = border;
  }
  if (!sameJson(alignment, baselineResult.alignment)) payload.alignment = alignment;
  // An untouched imported code is not submitted at all, so the style keeps it
  // even when another tab is applied.
  if (numberChanged) payload.numberFormat = { code: numberCode };
  return { payload };
}

function buildBaselinePayload(baseline: Draft): {
  readonly font: NonNullable<FormatCellsPayload['font']>;
  readonly fill: NonNullable<FormatCellsPayload['fill']>;
  readonly border: NonNullable<FormatCellsPayload['border']>;
  readonly alignment: NonNullable<FormatCellsPayload['alignment']>;
} {
  const fontColor = parseColor(baseline.font.color);
  const fillColor = parseColor(baseline.fill.color);
  const fillBackground = parseColor(baseline.fill.backgroundColor);
  const borderColor = parseColor(baseline.border.color);
  const border =
    baseline.border.preset === 'none'
      ? { preset: 'none' as const }
      : baseline.border.preset === 'custom'
        ? {}
        : {
            preset: baseline.border.preset,
            style: baseline.border.style,
            color: borderColor.color,
          };
  return {
    font: {
      name: baseline.font.name.trim(),
      size: parseNumber(baseline.font.size) ?? 0,
      bold: baseline.font.bold,
      italic: baseline.font.italic,
      underlineStyle: baseline.font.underlineStyle,
      strike: baseline.font.strike,
      color: fontColor.color,
    },
    fill: {
      pattern: baseline.fill.pattern,
      color: fillColor.color,
      backgroundColor: fillBackground.color,
    },
    border,
    alignment: {
      horizontal: baseline.alignment.horizontal,
      vertical: baseline.alignment.vertical,
      wrapText: baseline.alignment.wrapText,
      indent: parseNumber(baseline.alignment.indent) ?? 0,
    },
  };
}

/**
 * Build only the custom border edges the user touched. An edge whose checkbox
 * and the shared style/colour are unchanged is omitted, so the command leaves
 * that side's imported style and colour alone. A style or colour edit applies
 * to the checked edges the user sees; a changed checkbox applies only to that
 * edge.
 */
function borderPayloadFor(
  draft: Draft['border'],
  baseline: Draft['border'],
  color: string | null,
  baselineColor: string | null,
): NonNullable<FormatCellsPayload['border']> {
  if (draft.preset === 'none') return { preset: 'none' };
  if (draft.preset !== 'custom') {
    return { preset: draft.preset, style: draft.style, color };
  }
  const payload: {
    preset?: BorderPreset;
    style?: string;
    color?: string | null;
    edges?: Partial<Record<BorderEdgeName | 'diagonal', BorderEdgeStyle>>;
    diagonalUp?: boolean;
    diagonalDown?: boolean;
  } = {};
  const edges = changedCustomBorderEdges(draft, baseline, color, baselineColor);
  if (Object.keys(edges).length > 0) payload.edges = edges;
  if (draft.diagonalUp !== baseline.diagonalUp) payload.diagonalUp = draft.diagonalUp;
  if (draft.diagonalDown !== baseline.diagonalDown) payload.diagonalDown = draft.diagonalDown;
  return payload;
}

function changedCustomBorderEdges(
  draft: Draft['border'],
  baseline: Draft['border'],
  color: string | null,
  baselineColor: string | null,
): Partial<Record<BorderEdgeName | 'diagonal', BorderEdgeStyle>> {
  const styleChanged = draft.style !== baseline.style || color !== baselineColor;
  const styledEdge = (checked: boolean): BorderEdgeStyle =>
    checked ? borderEdge(draft.style, color) : { style: 'none', color: null, weight: null };
  const edges: Partial<Record<BorderEdgeName | 'diagonal', BorderEdgeStyle>> = {};
  for (const edge of ['left', 'right', 'top', 'bottom'] as const) {
    if (draft.edges[edge] !== baseline.edges[edge] || (styleChanged && draft.edges[edge])) {
      edges[edge] = styledEdge(draft.edges[edge]);
    }
  }
  const diagonal = draft.diagonalUp || draft.diagonalDown;
  const baselineDiagonal = baseline.diagonalUp || baseline.diagonalDown;
  if (diagonal !== baselineDiagonal || (styleChanged && diagonal)) {
    edges.diagonal = styledEdge(diagonal);
  }
  return edges;
}

function hasAnyBorderEdge(border: Draft['border']): boolean {
  return (
    border.edges.left ||
    border.edges.right ||
    border.edges.top ||
    border.edges.bottom ||
    border.diagonalUp ||
    border.diagonalDown
  );
}

function borderDraftEqual(left: Draft['border'], right: Draft['border']): boolean {
  return sameJson(left, right);
}

function parseNumber(value: string): number | null {
  const trimmed = value.trim();
  if (trimmed === '') return null;
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseColor(value: string): { readonly color: string | null; readonly error?: string } {
  const trimmed = value.trim();
  if (trimmed === '') return { color: null };
  const digits = trimmed.startsWith('#') ? trimmed.slice(1) : trimmed;
  if (!/^(?:[0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(digits)) {
    return {
      color: null,
      error: 'Enter a colour as #RGB, #RRGGBB or #AARRGGBB (alpha is ignored), or leave it blank.',
    };
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

/** A stable comparison key: blank and General are the same format state. */
function normalizedNumberCode(value: string): string | null {
  const trimmed = value.trim();
  if (trimmed === '' || trimmed.toLowerCase() === 'general') return null;
  return trimmed;
}

function parseNumberFormatCode(value: string): {
  readonly code: string | null;
  readonly error?: string;
} {
  const trimmed = value.trim();
  if (trimmed === '' || trimmed.toLowerCase() === 'general') return { code: null };
  try {
    const normalized = normalizeFormatCode(value);
    if (normalized.trim() === '' || !isValidFormat(normalized)) {
      return {
        code: null,
        error: 'Enter a valid Excel number format code, for example 0.00.',
      };
    }
  } catch {
    return {
      code: null,
      error: 'Enter a valid Excel number format code, for example 0.00.',
    };
  }
  return { code: value };
}

function categoryForCode(code: string): string {
  if (code.trim() === '' || code.trim().toLowerCase() === 'general') return 'general';
  const known = NUMBER_CATEGORIES.find(
    (category) => category.id !== 'custom' && category.code === code,
  );
  if (known !== undefined) return known.id;
  if (code.startsWith('_("$"')) return 'accounting';
  try {
    const info = getFormatInfo(normalizeFormatCode(code));
    switch (info.type) {
      case 'currency':
        return 'currency';
      case 'date':
      case 'datetime':
        return 'date';
      case 'time':
        return 'time';
      case 'percent':
        return 'percentage';
      case 'fraction':
        return 'fraction';
      case 'scientific':
        return 'scientific';
      case 'text':
        return 'text';
      case 'general':
        return 'general';
      case 'grouped':
      case 'number':
        return 'number';
      default:
        return 'custom';
    }
  } catch {
    return 'custom';
  }
}

function numberCategoryFor(id: string): NumberCategoryDefinition {
  return NUMBER_CATEGORIES.find((category) => category.id === id) ?? NUMBER_CATEGORIES[0]!;
}

function underlineFromSelect(value: string): UnderlineStyle {
  switch (value) {
    case 'single':
    case 'double':
    case 'singleAccounting':
    case 'doubleAccounting':
      return value;
    default:
      return 'none';
  }
}

function borderPresetFromSelect(value: string): BorderPreset {
  switch (value) {
    case 'outline':
    case 'inside':
    case 'all':
    case 'none':
      return value;
    default:
      return 'custom';
  }
}

function horizontalFromSelect(value: string): HorizontalAlignment {
  switch (value) {
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value;
    default:
      return 'general';
  }
}

function verticalFromSelect(value: string): VerticalAlignment {
  switch (value) {
    case 'top':
    case 'center':
    case 'bottom':
    case 'justify':
    case 'distributed':
      return value;
    default:
      return 'bottom';
  }
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}
