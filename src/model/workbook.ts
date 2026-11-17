import type {
  AlignmentStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  Sheet,
  SheetId,
  SheetView,
  StyleRecord,
  Workbook,
  WorkbookId,
} from './types';
import {
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
  builtinNumberFormat,
  builtinNumberFormatId,
} from './style';

/** A style record as it is appended; the visual parts default to the base style. */
export type StyleRecordInput = Omit<
  StyleRecord,
  'id' | 'font' | 'fill' | 'border' | 'alignment' | 'numFmtCode' | 'numFmtKind'
> & {
  readonly id?: number;
  readonly font?: FontStyle;
  readonly fill?: FillStyle;
  readonly border?: BorderStyle;
  readonly alignment?: AlignmentStyle;
  readonly numFmtCode?: string | null;
  readonly numFmtKind?: StyleRecord['numFmtKind'];
};

export const DEFAULT_ROW_HEIGHT_POINTS = 15;
export const DEFAULT_COLUMN_WIDTH_CHARS = 8.43;

let fallbackId = 0;

function newId(prefix: string): string {
  const cryptoApi = globalThis.crypto;
  if (cryptoApi !== undefined && typeof cryptoApi.randomUUID === 'function') {
    return `${prefix}:${cryptoApi.randomUUID()}`;
  }
  fallbackId += 1;
  return `${prefix}:${Date.now().toString(36)}-${fallbackId.toString(36)}`;
}

export function createWorkbookId(): WorkbookId {
  return newId('workbook');
}

export function createSheetId(partPath: string): SheetId {
  return `sheet:${partPath}`;
}

/**
 * A deterministic stable id for a sheet created in calcplus. It includes the
 * new part path, so an id is never handed to a sheet with a different part;
 * the counter keeps successive inserts distinct. LIFO undo means an id can be
 * reused only after the older sheet has left the workbook and cannot be
 * restored on top of the newer one.
 */
export function nextNewSheetId(workbook: Workbook, partPath: string): SheetId {
  const used = new Set(workbook.sheets.map((sheet) => sheet.id));
  let number = 1;
  let candidate = `sheet:new:${partPath}:${number}`;
  while (used.has(candidate)) {
    number += 1;
    candidate = `sheet:new:${partPath}:${number}`;
  }
  return candidate;
}

export function createDefaultView(): SheetView {
  return {
    freezeRows: 0,
    freezeCols: 0,
    showGridLines: true,
    showHeadings: true,
    zoom: 100,
  };
}

export function createSheet(name: string, partPath: string, id = createSheetId(partPath)): Sheet {
  return {
    id,
    name,
    partPath,
    cells: new Map(),
    rowDimensions: new Map(),
    columnDimensions: new Map(),
    merges: [],
    names: [],
    comments: new Map(),
    hyperlinks: new Map(),
    validation: [],
    conditionalFormats: [],
    protection: null,
    view: createDefaultView(),
    defaultRowHeight: null,
    defaultColumnWidth: null,
    opaqueParts: new Map(),
  };
}

export function createDefaultStyle(id = 0): StyleRecord {
  return {
    id,
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    rawXml: '',
    font: DEFAULT_FONT,
    fill: DEFAULT_FILL,
    border: DEFAULT_BORDER,
    alignment: DEFAULT_ALIGNMENT,
    numFmtCode: null,
    numFmtKind: 'general',
  };
}

export function createEmptyWorkbook(): Workbook {
  return {
    id: createWorkbookId(),
    sheets: [],
    names: [],
    styles: [createDefaultStyle(0)],
    theme: { path: null, bytes: null },
    namedStyles: [],
    protection: null,
    opaqueParts: new Map(),
    sourceParts: new Map(),
    indexedPalette: null,
    date1904: false,
  };
}

/**
 * Append-only style table. Existing ids are never renumbered or replaced.
 * Returns the new style id.
 */
export function appendStyle(workbook: Workbook, style: StyleRecordInput): number {
  const id = style.id ?? workbook.styles.length;
  if (workbook.styles.some((existing) => existing.id === id)) {
    throw new Error(`Style id ${id} already exists; the style table is append-only`);
  }
  if (id !== workbook.styles.length) {
    throw new Error('Appended styles must use the next contiguous id');
  }
  workbook.styles.push({
    ...style,
    id,
    font: style.font ?? DEFAULT_FONT,
    fill: style.fill ?? DEFAULT_FILL,
    border: style.border ?? DEFAULT_BORDER,
    alignment: style.alignment ?? DEFAULT_ALIGNMENT,
    numFmtCode: style.numFmtCode ?? null,
    numFmtKind: style.numFmtKind ?? inferNumberFormatKind(style),
  });
  return id;
}

function inferNumberFormatKind(style: StyleRecordInput): StyleRecord['numFmtKind'] {
  if (style.numFmtCode === null || style.numFmtCode === undefined) {
    if (style.numFmtId === 0) return 'general';
    const builtin = builtinNumberFormat(style.numFmtId);
    return builtin === null ? 'unknown' : 'builtin';
  }
  return builtinNumberFormatId(style.numFmtCode) !== null ? 'builtin' : 'custom';
}

export function getStyle(workbook: Workbook, styleId: number): StyleRecord {
  const style = workbook.styles.find((candidate) => candidate.id === styleId);
  if (style === undefined) {
    // A cell can reference a style id we did not model (for example a deleted
    // style in a malformed producer file). Treat it as the default style
    // without renumbering anything on save.
    return workbook.styles[0] ?? createDefaultStyle(0);
  }
  return style;
}

export function findSheet(workbook: Workbook, sheetId: SheetId): Sheet | undefined {
  return workbook.sheets.find((sheet) => sheet.id === sheetId);
}

/** Excel's sheet-name limit, in characters. */
export const MAX_SHEET_NAME_LENGTH = 31;

/**
 * The next Excel-style default name (`Sheet1`, `Sheet4`, ...) that the
 * workbook does not already use. Names are unique case-insensitively.
 */
export function nextSheetName(workbook: Workbook): string {
  const used = new Set(workbook.sheets.map((sheet) => sheet.name.toUpperCase()));
  let number = 1;
  while (used.has(`SHEET${number}`)) number += 1;
  return `Sheet${number}`;
}

/**
 * Validate a sheet name against Excel's rules and workbook uniqueness.
 * Returns a message for the user, or `null` when the name is usable.
 */
export function validateSheetName(
  workbook: Workbook,
  name: string,
  exceptSheetId: SheetId | null = null,
): string | null {
  if (name.trim() === '') return 'A sheet name cannot be blank.';
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    return `A sheet name cannot be longer than ${MAX_SHEET_NAME_LENGTH} characters.`;
  }
  if (/[:\\/?*[\]]/.test(name)) {
    return 'A sheet name cannot contain : \\ / ? * [ or ].';
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    return 'A sheet name cannot begin or end with an apostrophe.';
  }
  if (name.trim() !== name) {
    return 'A sheet name cannot begin or end with a space.';
  }
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001F\u007F]/.test(name)) {
    return 'A sheet name cannot contain control characters.';
  }
  const upper = name.toUpperCase();
  const existing = workbook.sheets.find(
    (sheet) => sheet.id !== exceptSheetId && sheet.name.toUpperCase() === upper,
  );
  if (existing !== undefined) return `A sheet named "${name}" already exists.`;
  return null;
}

/**
 * A worksheet part path that no current sheet, retained source part or opaque
 * part already uses. A saved terminal delete releases its part, so a later
 * insert may reuse the number; that is safe because the old part is gone.
 */
export function nextWorksheetPartPath(workbook: Workbook): string {
  const used = new Set<string>();
  for (const sheet of workbook.sheets) used.add(sheet.partPath);
  for (const path of workbook.sourceParts.keys()) used.add(path);
  for (const path of workbook.opaqueParts.keys()) used.add(path);
  let highest = 0;
  for (const path of used) {
    const match = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(path);
    if (match !== null) highest = Math.max(highest, Number(match[1]));
  }
  let number = highest + 1;
  while (used.has(`xl/worksheets/sheet${number}.xml`)) number += 1;
  return `xl/worksheets/sheet${number}.xml`;
}

/**
 * Whether this sheet holds any document content beyond an empty grid: cells,
 * formatting, merges or side tables. The terminal delete warning no longer
 * depends on it, because deleting any sheet is permanent; the helper remains
 * for callers that need to tell an empty sheet from a populated one.
 */
export function sheetHasContent(sheet: Sheet): boolean {
  return (
    sheet.cells.size > 0 ||
    sheet.rowDimensions.size > 0 ||
    sheet.columnDimensions.size > 0 ||
    sheet.merges.length > 0 ||
    sheet.names.length > 0 ||
    sheet.comments.size > 0 ||
    sheet.hyperlinks.size > 0 ||
    sheet.validation.length > 0 ||
    sheet.conditionalFormats.length > 0 ||
    sheet.protection !== null ||
    sheet.defaultRowHeight !== null ||
    sheet.defaultColumnWidth !== null ||
    sheet.opaqueParts.size > 0
  );
}

export function firstSheet(workbook: Workbook): Sheet {
  const sheet = workbook.sheets[0];
  if (sheet === undefined) throw new Error('Workbook has no sheets');
  return sheet;
}
