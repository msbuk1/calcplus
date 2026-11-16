import type {
  Cell,
  ColumnDimension,
  CommentRecord,
  ConditionalFormatRecord,
  DefinedName,
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
import { DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT } from '@model/style';

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
  // Package infrastructure can be present in entries but not classified on a
  // restored model; keep the union lossless.
  const known = new Set(parts.map((part) => part.path));
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
    sheets: snapshot.workbook.sheets.map(restoreSheet),
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
  };
}

function restoreSheet(sheet: SnapshotSheet): Sheet {
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
    opaqueParts: new Map(),
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
    font: style.font ?? DEFAULT_FONT,
    fill: style.fill ?? DEFAULT_FILL,
    border: style.border ?? DEFAULT_BORDER,
    alignment: style.alignment ?? DEFAULT_ALIGNMENT,
    numFmtCode: style.numFmtCode ?? null,
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
