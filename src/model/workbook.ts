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

export function firstSheet(workbook: Workbook): Sheet {
  const sheet = workbook.sheets[0];
  if (sheet === undefined) throw new Error('Workbook has no sheets');
  return sheet;
}
