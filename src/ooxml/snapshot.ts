import type {
  BorderStyle,
  Cell,
  ColumnDimension,
  CommentRecord,
  ConditionalFormatRecord,
  DefinedName,
  FillStyle,
  FontStyle,
  HyperlinkRecord,
  MergedRange,
  NamedStyleRecord,
  RowDimension,
  Sheet,
  SheetId,
  StyleRecord,
  ValidationRecord,
  Workbook,
  WorkbookProtectionRecord,
} from '@model/types';
import type { SharedStringsTable, WorkbookPackage } from './types';
import {
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
  builtinNumberFormat,
} from '@model/style';

export const SNAPSHOT_SCHEMA_VERSION = 1;

export interface SnapshotCellEntry {
  readonly key: string;
  readonly cell: Cell;
}

export interface SnapshotSheet {
  readonly id: SheetId;
  readonly name: string;
  readonly partPath: string;
  readonly cells: SnapshotCellEntry[];
  readonly rowDimensions: [number, RowDimension][];
  readonly columnDimensions: [number, ColumnDimension][];
  readonly merges: MergedRange[];
  readonly names: DefinedName[];
  readonly comments: [string, CommentRecord][];
  readonly hyperlinks: [string, HyperlinkRecord][];
  readonly validation: ValidationRecord[];
  readonly conditionalFormats: ConditionalFormatRecord[];
  readonly protection: { readonly rawXml: string } | null;
  readonly view: Sheet['view'];
  readonly defaultRowHeight?: number | null;
  readonly defaultColumnWidth?: number | null;
  /** Side parts attached to this sheet, if the snapshot came from an open workbook. */
  readonly opaquePartPaths?: readonly string[];
}
export interface SnapshotWorkbookMetadata {
  readonly id: string;
  readonly sheets: SnapshotSheet[];
  readonly names: DefinedName[];
  readonly styles: StyleRecord[];
  readonly themePath: string | null;
  readonly namedStyles: NamedStyleRecord[];
  readonly protection: WorkbookProtectionRecord | null;
  readonly date1904: boolean;
  readonly indexedPalette?: readonly string[] | null;
}

export interface SnapshotPart {
  readonly path: string;
  readonly kind: 'source' | 'opaque';
  readonly bytes: Uint8Array;
}

export interface SnapshotPackageMetadata {
  readonly contentTypesPath: string;
  readonly rootRelsPath: string;
  readonly workbookPartPath: string;
  readonly workbookRelsPath: string;
  readonly sharedStringsPath: string | null;
  readonly stylesPath: string | null;
  readonly themePath: string | null;
  readonly sheetPartPaths: [SheetId, string][];
  /** Missing on snapshots written before sheet-structure edits existed. */
  readonly sheetRelationshipIds?: [SheetId, string][];
  readonly sharedStrings: SharedStringsTable | null;
  readonly externalLinkParts: readonly string[];
}

export interface WorkbookSnapshot {
  readonly schemaVersion: number;
  readonly workbook: SnapshotWorkbookMetadata;
  readonly package: SnapshotPackageMetadata;
  /** Binary part records. Never base64 inside the metadata. */
  readonly parts: SnapshotPart[];
}

export function snapshotWorkbook(workbook: Workbook, pkg: WorkbookPackage): WorkbookSnapshot {
  const parts: SnapshotPart[] = [];
  for (const [path, bytes] of workbook.sourceParts) {
    parts.push({ path, kind: 'source', bytes: cloneBytes(bytes) });
  }
  for (const [path, part] of workbook.opaqueParts) {
    parts.push({ path, kind: 'opaque', bytes: cloneBytes(part.bytes) });
  }
  const known = new Set(parts.map((part) => part.path));
  for (const [path, part] of workbook.opaqueParts) {
    if (known.has(path)) continue;
    parts.push({ path, kind: 'opaque', bytes: cloneBytes(part.bytes) });
    known.add(path);
  }
  // Package infrastructure can be present in entries but not classified on a
  // restored model; keep the union lossless.
  for (const [path, bytes] of pkg.entries) {
    if (!known.has(path)) parts.push({ path, kind: 'opaque', bytes: cloneBytes(bytes) });
  }

  return {
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    workbook: {
      id: workbook.id,
      sheets: workbook.sheets.map(snapshotSheet),
      names: workbook.names.map(cloneDefinedName),
      styles: workbook.styles.map(cloneStyleForSnapshot),
      themePath: workbook.theme.path,
      namedStyles: workbook.namedStyles.map((style) => ({ ...style })),
      protection: workbook.protection === null ? null : { ...workbook.protection },
      date1904: workbook.date1904,
      indexedPalette: workbook.indexedPalette === null ? null : [...workbook.indexedPalette],
    },
    package: {
      contentTypesPath: pkg.contentTypesPath,
      rootRelsPath: pkg.rootRelsPath,
      workbookPartPath: pkg.workbookPartPath,
      workbookRelsPath: pkg.workbookRelsPath,
      sharedStringsPath: pkg.sharedStringsPath,
      stylesPath: pkg.stylesPath,
      themePath: pkg.themePath,
      sheetPartPaths: [...pkg.sheetPartPaths.entries()],
      sheetRelationshipIds: [...(pkg.sheetRelationshipIds ?? new Map()).entries()],
      sharedStrings:
        pkg.sharedStrings === null
          ? null
          : {
              ...pkg.sharedStrings,
              entries: pkg.sharedStrings.entries.map((entry) => ({ ...entry })),
            },
      externalLinkParts: [...pkg.externalLinkParts],
    },
    parts,
  };
}

export function restoreWorkbook(snapshot: WorkbookSnapshot): {
  workbook: Workbook;
  package: WorkbookPackage;
} {
  if (snapshot.schemaVersion !== SNAPSHOT_SCHEMA_VERSION) {
    throw new Error(`Unsupported workbook snapshot version ${snapshot.schemaVersion}`);
  }
  const sourceParts = new Map<string, Uint8Array>();
  const opaqueParts = new Map<string, { path: string; bytes: Uint8Array; anchor: null }>();
  const entries = new Map<string, Uint8Array>();
  for (const part of snapshot.parts) {
    const bytes = cloneBytes(part.bytes);
    entries.set(part.path, bytes);
    if (part.kind === 'source') sourceParts.set(part.path, bytes);
    else opaqueParts.set(part.path, { path: part.path, bytes, anchor: null });
  }

  const themePath = snapshot.workbook.themePath;
  const workbook: Workbook = {
    id: snapshot.workbook.id,
    sheets: snapshot.workbook.sheets.map((sheet) => restoreSheet(sheet, entries)),
    names: snapshot.workbook.names.map(cloneDefinedName),
    styles: snapshot.workbook.styles.map(normalizeStyle),
    theme: {
      path: themePath,
      bytes: themePath === null ? null : (entries.get(themePath) ?? null),
    },
    namedStyles: snapshot.workbook.namedStyles.map((style) => ({ ...style })),
    protection: snapshot.workbook.protection === null ? null : { ...snapshot.workbook.protection },
    opaqueParts,
    sourceParts,
    indexedPalette:
      snapshot.workbook.indexedPalette === undefined || snapshot.workbook.indexedPalette === null
        ? null
        : [...snapshot.workbook.indexedPalette],
    date1904: snapshot.workbook.date1904,
  };

  const pkg: WorkbookPackage = {
    originalBytes: null,
    entries,
    contentTypesPath: snapshot.package.contentTypesPath,
    rootRelsPath: snapshot.package.rootRelsPath,
    workbookPartPath: snapshot.package.workbookPartPath,
    workbookRelsPath: snapshot.package.workbookRelsPath,
    sharedStringsPath: snapshot.package.sharedStringsPath,
    stylesPath: snapshot.package.stylesPath,
    themePath: snapshot.package.themePath,
    sheetPartPaths: new Map(snapshot.package.sheetPartPaths),
    sheetRelationshipIds: new Map(snapshot.package.sheetRelationshipIds ?? []),
    relationships: new Map(),
    sharedStrings:
      snapshot.package.sharedStrings === null
        ? null
        : {
            ...snapshot.package.sharedStrings,
            entries: snapshot.package.sharedStrings.entries.map((entry) => ({ ...entry })),
          },
    externalLinkParts: [...snapshot.package.externalLinkParts],
  };
  return { workbook, package: pkg };
}

function snapshotSheet(sheet: Sheet): SnapshotSheet {
  return {
    id: sheet.id,
    name: sheet.name,
    partPath: sheet.partPath,
    cells: [...sheet.cells.entries()].map(([key, cell]) => ({
      key,
      cell: cloneCellForSnapshot(cell),
    })),
    rowDimensions: [...sheet.rowDimensions.entries()].map(([row, dimension]) => [
      row,
      { ...dimension },
    ]),
    columnDimensions: [...sheet.columnDimensions.entries()].map(([column, dimension]) => [
      column,
      { ...dimension },
    ]),
    merges: sheet.merges.map((merge) => ({ range: { ...merge.range } })),
    names: sheet.names.map(cloneDefinedName),
    comments: [...sheet.comments.entries()].map(([key, comment]) => [key, { ...comment }]),
    hyperlinks: [...sheet.hyperlinks.entries()].map(([key, hyperlink]) => [key, { ...hyperlink }]),
    validation: sheet.validation.map((record) => ({
      range: { ...record.range },
      rawXml: record.rawXml,
    })),
    conditionalFormats: sheet.conditionalFormats.map((record) => ({
      range: { ...record.range },
      rawXml: record.rawXml,
    })),
    protection: sheet.protection === null ? null : { ...sheet.protection },
    view: { ...sheet.view },
    defaultRowHeight: sheet.defaultRowHeight,
    defaultColumnWidth: sheet.defaultColumnWidth,
    opaquePartPaths: [...sheet.opaqueParts.keys()],
  };
}

function restoreSheet(sheet: SnapshotSheet, parts: ReadonlyMap<string, Uint8Array>): Sheet {
  return {
    id: sheet.id,
    name: sheet.name,
    partPath: sheet.partPath,
    cells: new Map(sheet.cells.map((entry) => [entry.key, cloneCellForSnapshot(entry.cell)])),
    rowDimensions: new Map(sheet.rowDimensions.map(([row, dimension]) => [row, { ...dimension }])),
    columnDimensions: new Map(
      sheet.columnDimensions.map(([column, dimension]) => [column, { ...dimension }]),
    ),
    merges: sheet.merges.map((merge) => ({ range: { ...merge.range } })),
    names: sheet.names.map(cloneDefinedName),
    comments: new Map(sheet.comments.map(([key, comment]) => [key, { ...comment }])),
    hyperlinks: new Map(sheet.hyperlinks.map(([key, hyperlink]) => [key, { ...hyperlink }])),
    validation: sheet.validation.map((record) => ({
      range: { ...record.range },
      rawXml: record.rawXml,
    })),
    conditionalFormats: sheet.conditionalFormats.map((record) => ({
      range: { ...record.range },
      rawXml: record.rawXml,
    })),
    protection: sheet.protection === null ? null : { ...sheet.protection },
    view: { ...sheet.view },
    defaultRowHeight: sheet.defaultRowHeight ?? null,
    defaultColumnWidth: sheet.defaultColumnWidth ?? null,
    opaqueParts: new Map(
      (sheet.opaquePartPaths ?? []).flatMap((path) => {
        const bytes = parts.get(path);
        return bytes === undefined ? [] : [[path, { path, bytes, anchor: null }] as const];
      }),
    ),
  };
}

function cloneDefinedName(name: DefinedName): DefinedName {
  return { ...name };
}

/**
 * Fill in the style parts a snapshot written before the style table gained
 * font, fill, alignment and number-format data.
 */
function normalizeStyle(style: StyleRecord): StyleRecord {
  return {
    ...style,
    font: normalizeFont(style.font),
    fill: normalizeFill(style.fill),
    border: normalizeBorder(style.border),
    alignment: { ...DEFAULT_ALIGNMENT, ...(style.alignment ?? {}) },
    numFmtCode: style.numFmtCode ?? null,
    numFmtKind: style.numFmtKind ?? inferNumberFormatKind(style),
  };
}

function inferNumberFormatKind(style: StyleRecord): StyleRecord['numFmtKind'] {
  if (style.numFmtCode === null) {
    if (style.numFmtId === 0) return 'general';
    return builtinNumberFormat(style.numFmtId) === null ? 'unknown' : 'builtin';
  }
  const builtin = builtinNumberFormat(style.numFmtId);
  return builtin !== null && builtin === style.numFmtCode ? 'builtin' : 'custom';
}

function normalizeFont(font: FontStyle | undefined): FontStyle {
  if (font === undefined) return DEFAULT_FONT;
  return {
    ...DEFAULT_FONT,
    ...font,
    underlineStyle:
      font.underlineStyle ?? (font.underline === true ? 'single' : DEFAULT_FONT.underlineStyle),
  };
}

function normalizeFill(fill: FillStyle | undefined): FillStyle {
  if (fill === undefined) return DEFAULT_FILL;
  return { ...DEFAULT_FILL, ...fill, backgroundColor: fill.backgroundColor ?? null };
}

function normalizeBorder(border: BorderStyle | undefined): BorderStyle {
  if (border === undefined) return DEFAULT_BORDER;
  return {
    ...DEFAULT_BORDER,
    ...border,
    left: { ...DEFAULT_BORDER.left, ...border.left },
    right: { ...DEFAULT_BORDER.right, ...border.right },
    top: { ...DEFAULT_BORDER.top, ...border.top },
    bottom: { ...DEFAULT_BORDER.bottom, ...border.bottom },
    diagonal: { ...DEFAULT_BORDER.diagonal, ...border.diagonal },
  };
}

function cloneStyleForSnapshot(style: StyleRecord): StyleRecord {
  return {
    ...style,
    font: { ...style.font },
    fill: { ...style.fill },
    border: {
      ...style.border,
      left: { ...style.border.left },
      right: { ...style.border.right },
      top: { ...style.border.top },
      bottom: { ...style.border.bottom },
      diagonal: { ...style.border.diagonal },
    },
    alignment: { ...style.alignment },
  };
}

function cloneCellForSnapshot(cell: Cell): Cell {
  return {
    value: { ...cell.value },
    formula: cell.formula,
    formulaDerived: cell.formulaDerived,
    savedResult: cell.savedResult === null ? null : { ...cell.savedResult },
    styleId: cell.styleId,
    richText: cell.richText === null ? null : { ...cell.richText },
    stale: cell.stale === true,
    ...(cell.formulaSource === undefined ? {} : { formulaSource: { ...cell.formulaSource } }),
  };
}

function cloneBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}
