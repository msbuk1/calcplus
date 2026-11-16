import type { GridRange, SheetId, Workbook } from './types';
import { cellKey, formatCellRef } from './types';
import {
  MAX_COLUMN,
  MAX_ROW,
  findAreaReferences,
  formatAreaReference,
  type AreaReference,
} from './references';

/**
 * The single structural-edit range transform (workbook-model §8).
 *
 * M3C builds the prototype for the constructs the gate names: formulas,
 * merged ranges and defined names, across sheets and quoted sheet references.
 * Cell movement and the remaining range-bearing constructs arrive with the
 * full transform in M5/M6.
 */

export type StructureAxis = 'row' | 'column';
export type StructureKind = 'insert' | 'delete';

export interface StructuralEdit {
  readonly kind: StructureKind;
  readonly sheetId: SheetId;
  readonly axis: StructureAxis;
  /** Zero-based first row or column. */
  readonly index: number;
  /** Number of rows or columns inserted or deleted. Always at least 1. */
  readonly count: number;
}

export interface FormulaRewrite {
  readonly sheetId: SheetId;
  readonly row: number;
  readonly col: number;
  readonly before: string;
  readonly after: string;
}

export interface MergeRewrite {
  readonly sheetId: SheetId;
  readonly before: GridRange;
  readonly after: GridRange | null;
}

export interface NameRewrite {
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId: SheetId | null;
  readonly name: string;
  readonly before: string;
  readonly after: string;
}

export interface RangeTransformPlan {
  readonly edit: StructuralEdit;
  readonly formulaRewrites: readonly FormulaRewrite[];
  readonly mergeRewrites: readonly MergeRewrite[];
  readonly nameRewrites: readonly NameRewrite[];
  /** Formulas in shared/array groups, which stay with their group for now. */
  readonly groupedFormulaCount: number;
  readonly changed: boolean;
}

export interface RangeTransformOptions {
  /** Name lookup is case-insensitive, as Excel sheet names are. */
  readonly workbook: Workbook;
}

/** Shift or shrink a range for one axis. `null` means every cell was deleted. */
export function transformRange(range: GridRange, edit: StructuralEdit): GridRange | null {
  if (edit.count <= 0) return range;
  if (edit.axis === 'row') {
    const rows = transformSpan(range.startRow, range.endRow, edit.index, edit.count, edit.kind);
    if (rows === null) return null;
    if (rows[0] === range.startRow && rows[1] === range.endRow) return range;
    return { ...range, startRow: rows[0], endRow: rows[1] };
  }
  const columns = transformSpan(range.startCol, range.endCol, edit.index, edit.count, edit.kind);
  if (columns === null) return null;
  if (columns[0] === range.startCol && columns[1] === range.endCol) return range;
  return { ...range, startCol: columns[0], endCol: columns[1] };
}

/**
 * Rewrite the A1 references in one authored formula. Unsupported constructs
 * (R1C1, 3D ranges, structured/external references, union and intersection)
 * are left byte-for-byte as authored; the calculation layer already marks
 * those cells stale (formula semantics §4 and §9).
 */
export function transformFormulaReferences(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId | null,
): string {
  if (formula.length === 0 || hasPreservingConstruct(formula)) return formula;
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  for (let index = 1; index < references.length; index += 1) {
    const previous = references[index - 1]!;
    const current = references[index]!;
    if (formula.slice(previous.end, current.start).trim() === '') {
      // Excel's intersection operator is v2; keep the authored text.
      return formula;
    }
  }

  let output = '';
  let cursor = 0;
  for (const reference of references) {
    output += formula.slice(cursor, reference.start);
    const transformed = transformReference(reference, edit, options, hostSheetId);
    output += transformed ?? formula.slice(reference.start, reference.end);
    cursor = reference.end;
  }
  return output + formula.slice(cursor);
}

/** Rewrite a defined name's formula. A sheet-scoped name is relative to its sheet. */
export function transformDefinedNameFormula(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  ownerSheetId: SheetId | null,
): string {
  return transformFormulaReferences(formula, edit, options, ownerSheetId);
}

/** Plan every reference rewrite without mutating the workbook. */
export function planRangeTransform(
  workbook: Workbook,
  edit: StructuralEdit,
  options: RangeTransformOptions = { workbook },
): RangeTransformPlan {
  const targetSheet = workbook.sheets.find((sheet) => sheet.id === edit.sheetId);
  if (targetSheet === undefined) {
    throw new Error(`Unknown sheet for structural edit: ${edit.sheetId}`);
  }
  if (!Number.isInteger(edit.index) || edit.index < 0) {
    throw new RangeError('The structural edit needs a non-negative integer index');
  }
  if (!Number.isInteger(edit.count) || edit.count < 1) {
    throw new RangeError('The structural edit needs a positive integer count');
  }
  const limit = edit.axis === 'row' ? MAX_ROW : MAX_COLUMN;
  if (edit.index + edit.count > limit) {
    throw new RangeError('The structural edit is outside the grid');
  }

  const formulaRewrites: FormulaRewrite[] = [];
  let groupedFormulaCount = 0;
  for (const sheet of workbook.sheets) {
    for (const [key, cell] of sheet.cells) {
      if (cell.formula === null) continue;
      if (cell.formulaSource !== undefined) {
        groupedFormulaCount += 1;
        continue;
      }
      const separator = key.indexOf(':');
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      const after = transformFormulaReferences(cell.formula, edit, options, sheet.id);
      if (after !== cell.formula) {
        formulaRewrites.push({ sheetId: sheet.id, row, col, before: cell.formula, after });
      }
    }
  }

  const mergeRewrites: MergeRewrite[] = [];
  for (const merge of targetSheet.merges) {
    const after = transformMergeRange(merge.range, edit);
    mergeRewrites.push({ sheetId: targetSheet.id, before: merge.range, after });
  }

  const nameRewrites: NameRewrite[] = [];
  for (const name of workbook.names) {
    const after = transformDefinedNameFormula(name.formula, edit, options, null);
    if (after !== name.formula) {
      nameRewrites.push({
        scope: 'workbook',
        sheetId: null,
        name: name.name,
        before: name.formula,
        after,
      });
    }
  }
  for (const sheet of workbook.sheets) {
    for (const name of sheet.names) {
      const after = transformDefinedNameFormula(name.formula, edit, options, sheet.id);
      if (after !== name.formula) {
        nameRewrites.push({
          scope: 'sheet',
          sheetId: sheet.id,
          name: name.name,
          before: name.formula,
          after,
        });
      }
    }
  }

  return {
    edit,
    formulaRewrites,
    mergeRewrites,
    nameRewrites,
    groupedFormulaCount,
    changed:
      formulaRewrites.length > 0 ||
      mergeRewrites.some((rewrite) => rewrite.after !== rewrite.before) ||
      nameRewrites.length > 0,
  };
}

export function transformMergeRange(range: GridRange, edit: StructuralEdit): GridRange | null {
  const transformed = transformRange(range, edit);
  if (transformed === null) return null;
  // A merge that collapses to one cell is no longer a merged range.
  if (transformed.startRow === transformed.endRow && transformed.startCol === transformed.endCol) {
    return null;
  }
  return transformed;
}

/**
 * Apply a plan to the model, for the future structural-edit command. The M3C
 * prototype uses the plan only to produce rewritten OOXML bytes.
 */
export function applyPlanToWorkbook(workbook: Workbook, plan: RangeTransformPlan): void {
  for (const rewrite of plan.formulaRewrites) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    if (sheet === undefined) continue;
    const cell = sheet.cells.get(cellKey(rewrite.row, rewrite.col));
    if (cell === undefined) continue;
    sheet.cells.set(cellKey(rewrite.row, rewrite.col), {
      ...cell,
      formula: rewrite.after,
      stale: true,
    });
  }
  for (const rewrite of plan.mergeRewrites) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    if (sheet === undefined) continue;
    const index = sheet.merges.findIndex((merge) => rangeEquals(merge.range, rewrite.before));
    if (index < 0) continue;
    if (rewrite.after === null) sheet.merges.splice(index, 1);
    else sheet.merges[index] = { range: rewrite.after };
  }
  for (const rewrite of plan.nameRewrites) {
    if (rewrite.scope === 'workbook') {
      const index = workbook.names.findIndex((name) => name.name === rewrite.name);
      if (index >= 0) workbook.names[index] = { ...workbook.names[index]!, formula: rewrite.after };
      continue;
    }
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    const index = sheet?.names.findIndex((name) => name.name === rewrite.name) ?? -1;
    if (sheet !== undefined && index >= 0 && sheet.names[index] !== undefined) {
      sheet.names[index] = { ...sheet.names[index]!, formula: rewrite.after };
    }
  }
}

function transformReference(
  reference: AreaReference,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId | null,
): string | null {
  const targetSheetId =
    reference.sheetName === null
      ? hostSheetId
      : findSheetIdByName(options.workbook, reference.sheetName);
  if (targetSheetId !== edit.sheetId) return null;
  const transformed = transformArea(reference, edit);
  if (transformed === null) return '#REF!';
  if (areasEqual(transformed, reference)) return null;
  return formatAreaReference(transformed);
}

function transformArea(reference: AreaReference, edit: StructuralEdit): AreaReference | null {
  if (reference.kind === 'columns') {
    if (edit.axis === 'row') return reference;
    const columns = transformSpan(
      reference.startCol,
      reference.endCol,
      edit.index,
      edit.count,
      edit.kind,
    );
    if (columns === null) return null;
    return {
      ...reference,
      startCol: columns[0],
      endCol: columns[1],
      single: false,
    };
  }
  if (reference.kind === 'rows') {
    if (edit.axis === 'column') return reference;
    const rows = transformSpan(
      reference.startRow,
      reference.endRow,
      edit.index,
      edit.count,
      edit.kind,
    );
    if (rows === null) return null;
    return {
      ...reference,
      startRow: rows[0],
      endRow: rows[1],
      single: false,
    };
  }
  const rows =
    edit.axis === 'row'
      ? transformSpan(reference.startRow, reference.endRow, edit.index, edit.count, edit.kind)
      : ([reference.startRow, reference.endRow] as const);
  if (rows === null) return null;
  const columns =
    edit.axis === 'column'
      ? transformSpan(reference.startCol, reference.endCol, edit.index, edit.count, edit.kind)
      : ([reference.startCol, reference.endCol] as const);
  if (columns === null) return null;
  return {
    ...reference,
    startRow: rows[0],
    endRow: rows[1],
    startCol: columns[0],
    endCol: columns[1],
  };
}

function transformSpan(
  start: number,
  end: number,
  index: number,
  count: number,
  kind: StructureKind,
): readonly [number, number] | null {
  const deleteEnd = index + count - 1;
  if (kind === 'insert') {
    return [start >= index ? start + count : start, end >= index ? end + count : end];
  }
  if (end < index) return [start, end];
  if (start > deleteEnd) return [start - count, end - count];
  if (start >= index && end <= deleteEnd) return null;
  if (start < index && end > deleteEnd) return [start, end - count];
  if (start < index && end <= deleteEnd) return [start, index - 1];
  return [index, end - count];
}

function hasPreservingConstruct(formula: string): boolean {
  const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'/g, "''");
  if (/[[\]@]/.test(stripped)) return true;
  if (/_xlfn\.|_xlws\./i.test(stripped)) return true;
  // R1C1 notation is v2 and is preserved as stale authored text.
  if (/\bR\d+C\d+\b/i.test(stripped)) return true;
  // 3D ranges: Sheet1:Sheet3!A1 or 'Sheet 1':'Sheet 3'!A1.
  if (/(?:''|[A-Za-z_][\w.]*)\s*:\s*(?:''|[A-Za-z_][\w.]*)\s*!/.test(stripped)) return true;
  return hasUnionComma(formula);
}

/**
 * A comma is Excel's union operator when it is not a function argument
 * separator. Grouping parentheses do not make a comma a function separator:
 * `(A1:A2,C1:C2)` is a union and must stay stale.
 */
function hasUnionComma(formula: string): boolean {
  const functionParens: boolean[] = [];
  let index = 0;
  while (index < formula.length) {
    const character = formula[index] ?? '';
    if (character === '"') {
      index = skipQuoted(formula, index, '"');
      continue;
    }
    if (character === "'") {
      index = skipQuoted(formula, index, "'");
      continue;
    }
    if (character === '(') {
      functionParens.push(isFunctionParen(formula, index));
      index += 1;
      continue;
    }
    if (character === ')') {
      functionParens.pop();
      index += 1;
      continue;
    }
    if (character === ',') {
      const inFunction = functionParens[functionParens.length - 1];
      if (inFunction !== true) return true;
    }
    index += 1;
  }
  return false;
}

function isFunctionParen(formula: string, openParen: number): boolean {
  let index = openParen - 1;
  while (index >= 0 && /\s/.test(formula[index] ?? '')) index -= 1;
  if (index < 0) return false;
  return isIdentifierPart(formula[index] ?? '');
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.\u00C0-\u02AF]/.test(character);
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

function findSheetIdByName(workbook: Workbook, name: string): SheetId | null {
  const upper = name.toUpperCase();
  return workbook.sheets.find((sheet) => sheet.name.toUpperCase() === upper)?.id ?? null;
}

function rangeEquals(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

function areasEqual(left: AreaReference, right: AreaReference): boolean {
  return (
    left.kind === right.kind &&
    left.startRow === right.startRow &&
    left.endRow === right.endRow &&
    left.startCol === right.startCol &&
    left.endCol === right.endCol
  );
}

/** The address a plan's formula rewrite points at, for reports and tests. */
export function formulaRewriteAddress(rewrite: FormulaRewrite): string {
  return formatCellRef(rewrite.row, rewrite.col);
}
