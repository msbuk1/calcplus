import type { PageMarginsRecord } from '@model/types';
import { normalisePrintAreaText } from '@commands/print';
import type { PageSetupPatch, PageSetupPayload, PageSetupValues } from '@commands/print';
import {
  MAX_FIT_PAGES,
  MAX_PRINT_MARGIN_INCHES,
  MAX_PRINT_SCALE_PERCENT,
  MIN_PRINT_SCALE_PERCENT,
} from '@commands/print';

/**
 * View state for the Page Setup dialog (print-pdf §4, ui-chrome §4.1).
 *
 * Pure helpers only: the dialog owns React state, and these functions own the
 * draft shape, its validation and the changed-fields payload. Keeping them
 * here means the write model is unit-tested without a browser, and the dialog
 * only renders and calls `onApply`.
 *
 * One Apply is one payload and therefore one `print.applySetup` transaction:
 * every changed field travels together, so undo restores the whole page setup
 * in one step.
 */

export interface PageSetupDraft {
  readonly orientation: 'portrait' | 'landscape';
  readonly scaling: 'fit' | 'scale';
  readonly scale: string;
  readonly fitToWidth: string;
  readonly fitToHeight: string;
  readonly paperSize: string;
  readonly margins: Readonly<Record<keyof PageMarginsRecord, string>>;
  readonly printArea: string;
  readonly printGridlines: boolean;
  readonly printHeadings: boolean;
  readonly horizontalCentered: boolean;
  readonly verticalCentered: boolean;
}

export type PageSetupField =
  'scale' | 'fitToWidth' | 'fitToHeight' | 'margins' | 'printArea' | 'paperSize';

export type PageSetupDraftErrors = Partial<Record<PageSetupField, string>>;

export interface PageSetupDraftResult {
  readonly payload: PageSetupPayload;
  readonly errors: PageSetupDraftErrors;
}

/** What the dialog's live plan line reports: the candidate plan's shape. */
export type PageSetupPlanCheck =
  | {
      readonly ok: true;
      readonly pages: number;
      readonly paperName: string;
      readonly orientation: 'portrait' | 'landscape';
    }
  | { readonly ok: false; readonly message: string };

export function pageSetupDraftFrom(initial: PageSetupValues): PageSetupDraft {
  return {
    orientation: initial.orientation,
    scaling: initial.fitToPage ? 'fit' : 'scale',
    scale: String(initial.scale),
    fitToWidth: String(initial.fitToWidth),
    fitToHeight: String(initial.fitToHeight),
    paperSize: String(initial.paperSize ?? 1),
    margins: {
      top: String(initial.margins.top),
      right: String(initial.margins.right),
      bottom: String(initial.margins.bottom),
      left: String(initial.margins.left),
      header: String(initial.margins.header),
      footer: String(initial.margins.footer),
    },
    printArea: initial.printAreaText,
    printGridlines: initial.printGridlines,
    printHeadings: initial.printHeadings,
    horizontalCentered: initial.horizontalCentered,
    verticalCentered: initial.verticalCentered,
  };
}

/**
 * The changed fields of the draft as one payload. Only fields that differ from
 * the committed values travel, so an untouched dialog writes nothing and a
 * one-field edit produces a one-field transaction.
 */
export function pageSetupDraftPayload(
  draft: PageSetupDraft,
  committed: PageSetupValues,
  sheetName: string,
): PageSetupDraftResult {
  const errors: PageSetupDraftErrors = {};
  const patch: {
    orientation?: 'portrait' | 'landscape';
    fitToPage?: boolean;
    scale?: number;
    fitToWidth?: number;
    fitToHeight?: number;
    paperSize?: number;
    margins?: PageMarginsRecord;
    printGridlines?: boolean;
    printHeadings?: boolean;
    horizontalCentered?: boolean;
    verticalCentered?: boolean;
  } = {};

  const fitToPage = draft.scaling === 'fit';
  if (draft.orientation !== committed.orientation) patch.orientation = draft.orientation;
  if (fitToPage !== committed.fitToPage) patch.fitToPage = fitToPage;
  if (fitToPage) {
    const width = wholeNumber(draft.fitToWidth);
    const height = wholeNumber(draft.fitToHeight);
    if (width === null || width < 0 || width > MAX_FIT_PAGES) {
      errors.fitToWidth = `Fit to width must be a whole number from 0 to ${MAX_FIT_PAGES}.`;
    } else if (width !== committed.fitToWidth) {
      patch.fitToWidth = width;
    }
    if (height === null || height < 0 || height > MAX_FIT_PAGES) {
      errors.fitToHeight = `Fit to height must be a whole number from 0 to ${MAX_FIT_PAGES}.`;
    } else if (height !== committed.fitToHeight) {
      patch.fitToHeight = height;
    }
  } else {
    const scale = wholeNumber(draft.scale);
    if (scale === null || scale < MIN_PRINT_SCALE_PERCENT || scale > MAX_PRINT_SCALE_PERCENT) {
      errors.scale = `Scale must be a whole number from ${MIN_PRINT_SCALE_PERCENT} to ${MAX_PRINT_SCALE_PERCENT}.`;
    } else if (scale !== committed.scale) {
      patch.scale = scale;
    }
  }

  const paperSize = wholeNumber(draft.paperSize);
  if (paperSize === null || paperSize <= 0) {
    errors.paperSize = 'Choose a paper size from the list.';
  } else if (paperSize !== (committed.paperSize ?? 1)) {
    patch.paperSize = paperSize;
  }

  const margins = marginsOf(draft.margins);
  if (margins === null) {
    errors.margins = `Margins must be numbers between 0 and ${MAX_PRINT_MARGIN_INCHES} inches.`;
  } else if (!sameMargins(margins, committed.margins)) {
    patch.margins = margins;
  }

  if (draft.printGridlines !== committed.printGridlines) {
    patch.printGridlines = draft.printGridlines;
  }
  if (draft.printHeadings !== committed.printHeadings) {
    patch.printHeadings = draft.printHeadings;
  }
  if (draft.horizontalCentered !== committed.horizontalCentered) {
    patch.horizontalCentered = draft.horizontalCentered;
  }
  if (draft.verticalCentered !== committed.verticalCentered) {
    patch.verticalCentered = draft.verticalCentered;
  }

  let printArea: string | null | undefined;
  const normalised = normalisePrintAreaText(draft.printArea, sheetName);
  if (!normalised.ok) {
    errors.printArea = normalised.message;
  } else {
    const before = normalisePrintAreaText(committed.printAreaText, sheetName);
    const beforeFormula = before.ok ? before.formula : null;
    if (normalised.formula !== beforeFormula) {
      printArea = normalised.formula === null ? null : draft.printArea.trim();
    }
  }

  const hasPatch = Object.keys(patch).length > 0;
  return {
    errors,
    payload: {
      ...(hasPatch ? { setup: patch as PageSetupPatch } : {}),
      ...(printArea === undefined ? {} : { printArea }),
    },
  };
}

/**
 * The committed values a successful Apply leaves for the next diff. The draft
 * is valid at that point, so every field parses; `null` means it was not.
 */
export function pageSetupValuesFromDraft(
  draft: PageSetupDraft,
  sheetName: string,
): PageSetupValues | null {
  const scale = wholeNumber(draft.scale);
  const fitToWidth = wholeNumber(draft.fitToWidth);
  const fitToHeight = wholeNumber(draft.fitToHeight);
  const paperSize = wholeNumber(draft.paperSize);
  const margins = marginsOf(draft.margins);
  if (
    scale === null ||
    fitToWidth === null ||
    fitToHeight === null ||
    paperSize === null ||
    margins === null
  ) {
    return null;
  }
  if (!normalisePrintAreaText(draft.printArea, sheetName).ok) return null;
  return {
    orientation: draft.orientation,
    fitToPage: draft.scaling === 'fit',
    fitToWidth,
    fitToHeight,
    scale,
    paperSize,
    margins,
    printAreaText: draft.printArea.trim(),
    printGridlines: draft.printGridlines,
    printHeadings: draft.printHeadings,
    horizontalCentered: draft.horizontalCentered,
    verticalCentered: draft.verticalCentered,
  };
}

export function hasPageSetupErrors(errors: PageSetupDraftErrors): boolean {
  return Object.values(errors).some((message) => message !== undefined);
}

export function firstPageSetupError(errors: PageSetupDraftErrors): string | null {
  for (const message of Object.values(errors)) {
    if (message !== undefined) return message;
  }
  return null;
}

/** One line under the panels: the candidate plan's page count and paper. */
export function pageSetupPlanText(check: PageSetupPlanCheck | null): string {
  if (check === null) return '';
  if (!check.ok) return check.message;
  if (check.pages === 0) return 'The plan has no pages to print.';
  const pages = check.pages === 1 ? '1 page' : `${check.pages} pages`;
  return `Plan: ${pages} on ${check.paperName} ${check.orientation}.`;
}

function wholeNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const value = Number(trimmed);
  return Number.isInteger(value) ? value : null;
}

function marginsOf(margins: PageSetupDraft['margins']): PageMarginsRecord | null {
  const value = (text: string): number | null => {
    const trimmed = text.trim();
    if (trimmed === '') return null;
    const parsed = Number(trimmed);
    if (!Number.isFinite(parsed) || parsed < 0 || parsed > MAX_PRINT_MARGIN_INCHES) return null;
    return parsed;
  };
  const top = value(margins.top);
  const right = value(margins.right);
  const bottom = value(margins.bottom);
  const left = value(margins.left);
  const header = value(margins.header);
  const footer = value(margins.footer);
  if (
    top === null ||
    right === null ||
    bottom === null ||
    left === null ||
    header === null ||
    footer === null
  ) {
    return null;
  }
  return { top, right, bottom, left, header, footer };
}

function sameMargins(left: PageMarginsRecord, right: PageMarginsRecord): boolean {
  return (
    left.top === right.top &&
    left.right === right.right &&
    left.bottom === right.bottom &&
    left.left === right.left &&
    left.header === right.header &&
    left.footer === right.footer
  );
}
