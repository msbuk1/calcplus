import type { DefinedName, GridRange, SheetId, Workbook } from './types';
import { columnToLetters, lettersToColumn } from './types';

/**
 * A1 reference parsing for the model edge (formula semantics §4).
 *
 * Supported: `A1`, `$A$1`, `$A1`, `A$1`, `Sheet2!A1`, `'Sheet name'!A1`,
 * whole-column `A:A`, whole-row `1:1`, and defined names. R1C1 notation, 3D
 * ranges, structured table references, union/intersection and external links
 * are deliberately unparseable here; their authored text is preserved.
 */

export const MAX_ROW = 1_048_576;
export const MAX_COLUMN = 16_384;

export type AreaReferenceKind = 'cells' | 'columns' | 'rows';

export interface AreaReference {
  /** Start index in the scanned formula. */
  readonly start: number;
  /** End index, exclusive. */
  readonly end: number;
  /** The raw sheet prefix including quotes, without the `!`. */
  readonly prefix: string | null;
  /** The decoded sheet name, or `null` for an unqualified reference. */
  readonly sheetName: string | null;
  readonly kind: AreaReferenceKind;
  readonly startRow: number;
  readonly endRow: number;
  readonly startCol: number;
  readonly endCol: number;
  readonly startRowAbs: boolean;
  readonly endRowAbs: boolean;
  readonly startColAbs: boolean;
  readonly endColAbs: boolean;
  /** True when the token names one cell/column/row rather than a `:` range. */
  readonly single: boolean;
}

export interface ResolvedArea {
  readonly kind: 'area';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly wholeColumn: boolean;
  readonly wholeRow: boolean;
}

export type ReferenceResolution =
  ResolvedArea | { readonly kind: 'invalid'; readonly message: string };

/** Find every A1 area in a formula, skipping strings and quoted sheet names. */
export function findAreaReferences(formula: string): AreaReference[] {
  const references: AreaReference[] = [];
  let index = 0;
  while (index < formula.length) {
    const character = formula[index] ?? '';
    if (character === '"') {
      index = skipQuoted(formula, index, '"');
      continue;
    }
    if (character === "'") {
      const prefix = parseSheetPrefixAt(formula, index);
      if (prefix !== null) {
        const area = parseAreaAt(formula, prefix.end);
        if (area !== null) {
          references.push({
            ...area,
            start: prefix.start,
            prefix: prefix.raw,
            sheetName: prefix.name,
          });
          index = area.end;
          continue;
        }
        index = prefix.end;
        continue;
      }
      index = skipQuoted(formula, index, "'");
      continue;
    }
    if (!isReferenceBoundary(formula, index)) {
      index += 1;
      continue;
    }
    if (isIdentifierStart(character)) {
      const prefix = parseSheetPrefixAt(formula, index);
      if (prefix !== null) {
        const area = parseAreaAt(formula, prefix.end);
        if (area !== null) {
          references.push({
            ...area,
            start: prefix.start,
            prefix: prefix.raw,
            sheetName: prefix.name,
          });
          index = area.end;
          continue;
        }
        index = prefix.end;
        continue;
      }
    }
    const area = parseAreaAt(formula, index);
    if (area !== null) {
      references.push({ ...area, prefix: null, sheetName: null });
      index = area.end;
      continue;
    }
    index += 1;
  }
  return references;
}

/** Parse one whole reference, as the Go To box and defined names use it. */
export function parseWholeArea(text: string): AreaReference | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const start = text.indexOf(trimmed);
  const prefix = parseSheetPrefixAt(text, start);
  const areaStart = prefix === null ? start : prefix.end;
  const area = parseAreaAt(text, areaStart);
  if (area === null) return null;
  if (text.slice(area.end).trim() !== '') return null;
  return {
    ...area,
    start: prefix === null ? start : prefix.start,
    prefix: prefix?.raw ?? null,
    sheetName: prefix?.name ?? null,
  };
}

/**
 * Resolve text typed into the Go To box: a defined name, a cross-sheet
 * reference, a quoted-sheet reference, a cell, a range, a whole column or a
 * whole row.
 */
export function resolveReferenceText(
  workbook: Workbook,
  text: string,
  hostSheetId: SheetId | null,
): ReferenceResolution {
  const trimmed = text.trim().replace(/^=/, '');
  if (trimmed === '') return invalid('Type a cell address or defined name.');

  const named = resolveDefinedName(workbook, trimmed, hostSheetId);
  if (named !== null) return named;

  const area = parseWholeArea(trimmed);
  if (area === null) {
    const qualified = splitQualifiedName(trimmed);
    if (qualified !== null) {
      const sheet = findSheetByName(workbook, qualified.sheetName);
      if (sheet !== null) {
        const scoped = resolveDefinedName(workbook, qualified.name, sheet.id);
        if (scoped !== null) return scoped;
      }
    }
    return invalid(`"${trimmed}" is not a cell, range or defined name.`);
  }
  const sheet =
    area.sheetName === null
      ? hostSheet(workbook, hostSheetId)
      : findSheetByName(workbook, area.sheetName);
  if (sheet === null) return invalid(`Unknown sheet "${area.sheetName ?? ''}".`);
  return {
    kind: 'area',
    sheetId: sheet.id,
    range: rangeOfArea(area),
    wholeColumn: area.kind === 'columns',
    wholeRow: area.kind === 'rows',
  };
}

/** Resolve a defined name to a sheet and range, if it names a simple area. */
export function resolveDefinedName(
  workbook: Workbook,
  name: string,
  hostSheetId: SheetId | null,
): ResolvedArea | null {
  const defined = findDefinedName(workbook, name, hostSheetId);
  if (defined === null) return null;
  const formula = defined.formula.trim().replace(/^=/, '');
  const area = parseWholeArea(formula);
  if (area === null) return null;
  const sheet =
    area.sheetName === null
      ? defined.sheetId === null
        ? hostSheet(workbook, hostSheetId)
        : findSheetById(workbook, defined.sheetId)
      : findSheetByName(workbook, area.sheetName);
  if (sheet === null) return null;
  return {
    kind: 'area',
    sheetId: sheet.id,
    range: rangeOfArea(area),
    wholeColumn: area.kind === 'columns',
    wholeRow: area.kind === 'rows',
  };
}

export function findDefinedName(
  workbook: Workbook,
  name: string,
  hostSheetId: SheetId | null,
): DefinedName | null {
  const upper = name.toUpperCase();
  if (hostSheetId !== null) {
    const sheet = findSheetById(workbook, hostSheetId);
    const local = sheet?.names.find((candidate) => candidate.name.toUpperCase() === upper);
    if (local !== undefined) return local;
  }
  const workbookName = workbook.names.find((candidate) => candidate.name.toUpperCase() === upper);
  if (workbookName !== undefined) return workbookName;
  for (const sheet of workbook.sheets) {
    const local = sheet.names.find((candidate) => candidate.name.toUpperCase() === upper);
    if (local !== undefined) return local;
  }
  return null;
}

export function rangeOfArea(area: AreaReference): GridRange {
  return {
    startRow: area.startRow,
    endRow: area.endRow,
    startCol: area.startCol,
    endCol: area.endCol,
  };
}

/** Format an area with absolute markers, for an engine-only name expression. */
export function absolutizeFormulaReferences(formula: string): string {
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  let output = '';
  let cursor = 0;
  for (const reference of references) {
    output += formula.slice(cursor, reference.start);
    output += formatAreaReference(reference, true);
    cursor = reference.end;
  }
  return output + formula.slice(cursor);
}

/** Format an area reference, optionally removing or adding `$` markers. */
export function formatAreaReference(area: AreaReference, absolute = false): string {
  const prefix = area.prefix === null ? '' : `${area.prefix}!`;
  const rowAbs = (value: boolean): string => (absolute || value ? '$' : '');
  const colAbs = (value: boolean): string => (absolute || value ? '$' : '');
  if (area.kind === 'columns') {
    return `${prefix}${colAbs(area.startColAbs)}${columnToLetters(area.startCol)}:${colAbs(
      area.endColAbs,
    )}${columnToLetters(area.endCol)}`;
  }
  if (area.kind === 'rows') {
    return `${prefix}${rowAbs(area.startRowAbs)}${area.startRow + 1}:${rowAbs(
      area.endRowAbs,
    )}${area.endRow + 1}`;
  }
  const start = `${colAbs(area.startColAbs)}${columnToLetters(area.startCol)}${rowAbs(
    area.startRowAbs,
  )}${area.startRow + 1}`;
  if (area.single) return `${prefix}${start}`;
  const end = `${colAbs(area.endColAbs)}${columnToLetters(area.endCol)}${rowAbs(
    area.endRowAbs,
  )}${area.endRow + 1}`;
  return `${prefix}${start}:${end}`;
}

function parseAreaAt(
  text: string,
  index: number,
): Omit<AreaReference, 'prefix' | 'sheetName'> | null {
  const rest = text.slice(index);
  if (rest === '') return null;

  const columnRange = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.])/.exec(rest);
  if (columnRange !== null) {
    const startCol = lettersToColumn(columnRange[2] ?? '');
    const endCol = lettersToColumn(columnRange[4] ?? '');
    if (startCol < MAX_COLUMN && endCol < MAX_COLUMN) {
      const end = index + columnRange[0].length;
      return {
        start: index,
        end,
        kind: 'columns',
        startRow: 0,
        endRow: MAX_ROW - 1,
        startCol: Math.min(startCol, endCol),
        endCol: Math.max(startCol, endCol),
        startRowAbs: false,
        endRowAbs: false,
        startColAbs: columnRange[1] === '$',
        endColAbs: columnRange[3] === '$',
        single: false,
      };
    }
  }

  const cellRange =
    /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/.exec(rest);
  if (cellRange !== null) {
    const end = index + cellRange[0].length;
    const after = text[end] ?? '';
    const afterTrim = firstNonSpace(text, end);
    const startCol = lettersToColumn(cellRange[2] ?? '');
    const startRow = Number(cellRange[4]) - 1;
    const single = cellRange[6] === undefined;
    if (
      startCol < MAX_COLUMN &&
      startRow >= 0 &&
      startRow < MAX_ROW &&
      !/[A-Za-z0-9_.]/.test(after) &&
      (single ? afterTrim !== '(' : true)
    ) {
      const endCol = single ? startCol : lettersToColumn(cellRange[6] ?? '');
      const endRow = single ? startRow : Number(cellRange[8]) - 1;
      if (endCol < MAX_COLUMN && endRow >= 0 && endRow < MAX_ROW) {
        return {
          start: index,
          end,
          kind: 'cells',
          startRow: Math.min(startRow, endRow),
          endRow: Math.max(startRow, endRow),
          startCol: Math.min(startCol, endCol),
          endCol: Math.max(startCol, endCol),
          startRowAbs: cellRange[3] === '$',
          endRowAbs: single ? cellRange[3] === '$' : cellRange[7] === '$',
          startColAbs: cellRange[1] === '$',
          endColAbs: single ? cellRange[1] === '$' : cellRange[5] === '$',
          single,
        };
      }
    }
  }

  const rowRange = /^(\$?)([1-9][0-9]{0,6}):(\$?)([1-9][0-9]{0,6})(?![0-9])/.exec(rest);
  if (rowRange !== null) {
    const startRow = Number(rowRange[2]) - 1;
    const endRow = Number(rowRange[4]) - 1;
    if (startRow < MAX_ROW && endRow < MAX_ROW) {
      return {
        start: index,
        end: index + rowRange[0].length,
        kind: 'rows',
        startRow: Math.min(startRow, endRow),
        endRow: Math.max(startRow, endRow),
        startCol: 0,
        endCol: MAX_COLUMN - 1,
        startRowAbs: rowRange[1] === '$',
        endRowAbs: rowRange[3] === '$',
        startColAbs: false,
        endColAbs: false,
        single: false,
      };
    }
  }
  return null;
}

function parseSheetPrefixAt(
  text: string,
  index: number,
): {
  readonly start: number;
  readonly name: string;
  readonly raw: string;
  readonly end: number;
} | null {
  const character = text[index] ?? '';
  if (character === "'") {
    let cursor = index + 1;
    let name = '';
    while (cursor < text.length) {
      if (text[cursor] === "'") {
        if (text[cursor + 1] === "'") {
          name += "'";
          cursor += 2;
          continue;
        }
        cursor += 1;
        if (text[cursor] !== '!') return null;
        return {
          start: index,
          name,
          raw: text.slice(index, cursor),
          end: cursor + 1,
        };
      }
      name += text[cursor];
      cursor += 1;
    }
    return null;
  }
  if (!isIdentifierStart(character)) return null;
  let cursor = index + 1;
  while (cursor < text.length && isIdentifierPart(text[cursor] ?? '')) cursor += 1;
  if (text[cursor] !== '!') return null;
  return {
    start: index,
    name: text.slice(index, cursor),
    raw: text.slice(index, cursor),
    end: cursor + 1,
  };
}

function splitQualifiedName(
  text: string,
): { readonly sheetName: string; readonly name: string } | null {
  const bang = text.lastIndexOf('!');
  if (bang <= 0) return null;
  const sheetPart = text.slice(0, bang);
  const name = text.slice(bang + 1).trim();
  if (name === '') return null;
  const sheetName =
    sheetPart.startsWith("'") && sheetPart.endsWith("'")
      ? sheetPart.slice(1, -1).replaceAll("''", "'")
      : sheetPart;
  return { sheetName, name };
}

function firstNonSpace(text: string, index: number): string {
  let cursor = index;
  while (cursor < text.length && /\s/.test(text[cursor] ?? '')) cursor += 1;
  return text[cursor] ?? '';
}

function skipQuoted(text: string, start: number, quote: string): number {
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === quote) {
      if (text[index + 1] === quote) {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return text.length;
}

/**
 * A reference token starts only at a name boundary. `Rate1` and `myName2`
 * contain text that looks like a cell address but is part of a defined name;
 * the scanner must not parse a suffix as a reference.
 */
function isReferenceBoundary(text: string, index: number): boolean {
  if (index <= 0) return true;
  const previous = text[index - 1] ?? '';
  return !/[A-Za-z0-9_.\u00C0-\u02AF]/.test(previous);
}

function isIdentifierStart(character: string): boolean {
  return /[A-Za-z_\u00C0-\u02AF]/.test(character);
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.\u00C0-\u02AF]/.test(character);
}

function hostSheet(workbook: Workbook, hostSheetId: SheetId | null) {
  if (hostSheetId !== null) {
    const sheet = findSheetById(workbook, hostSheetId);
    if (sheet !== null) return sheet;
  }
  return workbook.sheets[0] ?? null;
}

function findSheetById(workbook: Workbook, sheetId: SheetId) {
  return workbook.sheets.find((sheet) => sheet.id === sheetId) ?? null;
}

function findSheetByName(workbook: Workbook, name: string) {
  const upper = name.toUpperCase();
  return workbook.sheets.find((sheet) => sheet.name.toUpperCase() === upper) ?? null;
}

function invalid(message: string): ReferenceResolution {
  return { kind: 'invalid', message };
}
