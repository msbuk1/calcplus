import type { Sheet, SheetId, SheetView, StyleRecord, Workbook, WorkbookId } from './types';

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
    date1904: false,
  };
}

/**
 * Append-only style table. Existing ids are never renumbered or replaced.
 * Returns the new style id.
 */
export function appendStyle(
  workbook: Workbook,
  style: Omit<StyleRecord, 'id'> & { id?: number },
): number {
  const id = style.id ?? workbook.styles.length;
  if (workbook.styles.some((existing) => existing.id === id)) {
    throw new Error(`Style id ${id} already exists; the style table is append-only`);
  }
  if (id !== workbook.styles.length) {
    throw new Error('Appended styles must use the next contiguous id');
  }
  workbook.styles.push({ ...style, id });
  return id;
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
