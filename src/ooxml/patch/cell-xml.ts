import type { Cell, FormulaStorageKind, SheetId, Workbook } from '@model/types';
import { formatCellRef, parseCellKey } from '@model/types';
import { arrayFormulaForCell } from '@model/formula-ranges';
import { textValue, type CellValue } from '@model/value';
import { formulaRefusedByExcel } from '@engine/formula-text';
import { WorkbookSaveError } from '../errors';
import { escapeXmlAttribute, escapeXmlText } from '../xml';
import type { SharedStringsAppender } from './shared-strings-patch';

/**
 * A formula Excel refuses at entry: a call outside its function's argument
 * range (formula semantics §2 and §9), or an empty parenthesised group
 * (ticket 48). `cell-set.ts` refuses it when the user types it; this is the
 * same screen at the write boundary, because a formula
 * that arrives *inside* a file is loaded into the model untouched and would
 * otherwise be written straight back out. Excel then repairs the file and
 * strips the formula, which is the data loss file-compat §1 forbids and the
 * reason M6 exit criterion 3 failed.
 *
 * Two kinds of cell are deliberately not screened:
 *
 * - A shared or legacy array group. They are read-only this milestone and
 *   `serializeFormula` refuses to rewrite them, so screening here would turn a
 *   preserved group into a blocked save.
 * - A cell holding rich text. `dirtyCellsChangeFormula` refuses to rewrite rich
 *   runs, for the same reason. Neither can coexist with a formula in a file
 *   Excel produced.
 */
export function isRefusedFormulaCell(cell: Cell | null): boolean {
  if (cell === null || cell.formula === null) return false;
  if (cell.formulaSource !== undefined) return false;
  if (cell.richText !== null && cell.richText !== undefined) return false;
  return formulaRefusedByExcel(cell.formula);
}

/**
 * The cell exactly as it is written. A formula Excel refuses becomes a plain
 * value cell: the `<f>` is dropped and the cached result becomes the cell's
 * value, which is what Excel's own repair leaves behind — it removes the
 * Formula record and keeps `<v>`. The model is untouched, so the formula stays
 * visible in the app and can be corrected; only the file omits it.
 *
 * Every other cell comes back unchanged, so callers can apply this
 * unconditionally; `null` passes through so it composes with a null check.
 */
export function writtenCell(cell: Cell | null): Cell | null {
  if (cell === null || !isRefusedFormulaCell(cell)) return cell;
  return {
    ...cell,
    formula: null,
    formulaDerived: false,
    value: cell.savedResult ?? cell.value,
    savedResult: null,
  };
}

/** Where a refused formula sits, for the two reports (file-compat §7). */
export interface RefusedFormulaRef {
  readonly sheetId: SheetId;
  readonly sheetName: string;
  /** The sparse-cell key, `row:col`. */
  readonly key: string;
  /** The A1 reference, `F5`. */
  readonly ref: string;
  /** The authored formula text, `SUM()`. */
  readonly formula: string;
}

/**
 * Every cell in the workbook whose formula Excel refuses and whose write this
 * save skips. Shared by the save path and both reports, so a report entry can
 * never name a cell the writer does not actually handle.
 *
 * A cell inside a legacy array range is left out: the range is read-only this
 * milestone, so its bytes are preserved as they stand and reporting the formula
 * as dropped would be false.
 */
export function findRefusedFormulaCells(workbook: Workbook): RefusedFormulaRef[] {
  const found: RefusedFormulaRef[] = [];
  for (const sheet of workbook.sheets) {
    for (const [key, cell] of sheet.cells) {
      if (!isRefusedFormulaCell(cell)) continue;
      const address = parseCellKey(key);
      if (arrayFormulaForCell(sheet, address.row, address.col) !== null) continue;
      found.push({
        sheetId: sheet.id,
        sheetName: sheet.name,
        key,
        ref: formatCellRef(address.row, address.col),
        formula: cell.formula ?? '',
      });
    }
  }
  return found;
}

export interface CellWriteOptions {
  readonly sharedStrings: SharedStringsAppender | null;
  readonly original: CellOriginal | null;
}

export interface CellOriginal {
  readonly attributes: readonly { readonly name: string; readonly value: string }[];
  readonly unknownChildren: readonly string[];
  readonly formulaXml: string | null;
  readonly formulaText: string | null;
  readonly formulaKind: FormulaStorageKind | null;
}

export function cellToXml(
  prefix: string,
  ref: string,
  input: Cell | null,
  options: CellWriteOptions,
): string | null {
  const cell = writtenCell(input);
  if (cell === null) return serializeBlankOrNull(options.original, ref);
  const formula = cell.formula;
  const hasVisibleValue = cell.value.type !== 'blank';
  const hasStyle = cell.styleId !== 0;
  if (formula === null && !hasVisibleValue && !hasStyle) return null;

  const tag = (local: string): string => (prefix === '' ? local : `${prefix}:${local}`);
  const attributes: string[] = [`r="${ref}"`];
  if (hasStyle) attributes.push(`s="${cell.styleId}"`);
  if (options.original !== null) {
    for (const attribute of options.original.attributes) {
      const local = attribute.name.includes(':')
        ? attribute.name.slice(attribute.name.lastIndexOf(':') + 1)
        : attribute.name;
      if (local === 'r' || local === 's' || local === 't') continue;
      attributes.push(`${attribute.name}="${escapeXmlAttribute(attribute.value)}"`);
    }
  }

  const formulaXml = serializeFormula(cell, options.original, tag);
  // Only a true literal can become text. A formula-backed cell with no authored
  // text (a shared follower) must keep its formula storage and cached result.
  const follower = cell.formulaSource?.kind === 'shared-follower';
  const literalCell = formula === null && cell.formulaSource === undefined;
  const storedValue = literalCell ? storedLiteral(cell.value) : cell.value;
  const type = follower
    ? originalCellType(options.original)
    : cellType(storedValue, cell.savedResult, formula, options.sharedStrings);
  if (type !== null) attributes.push(`t="${type}"`);

  const children: string[] = [];
  if (formulaXml !== null) children.push(formulaXml);
  const value = formula !== null ? cell.savedResult : storedValue;
  if (
    formula === null &&
    value !== null &&
    value.type === 'text' &&
    options.sharedStrings === null
  ) {
    children.push(
      `<${tag('is')}><${tag('t')} xml:space="preserve">${escapeXmlText(
        value.value,
      )}</${tag('t')}></${tag('is')}>`,
    );
  } else if (value !== null && value.type !== 'blank') {
    const serialized =
      formula !== null || follower
        ? serializedFormulaValue(value)
        : serializedValue(value, options.sharedStrings);
    if (serialized !== null) {
      children.push(`<${tag('v')}>${escapeXmlText(serialized)}</${tag('v')}>`);
    }
  }
  if (options.original !== null) {
    children.push(...options.original.unknownChildren);
  }
  const open = `<${tag('c')} ${attributes.join(' ')}`;
  if (children.length === 0) return `${open}/>`;
  return `${open}>${children.join('')}</${tag('c')}>`;
}

/**
 * Formula semantics §6: #CIRC! is an in-app-only error. A literal cell cannot
 * store it as an error value, so the authored literal is written as text.
 * Formula cells keep their authored text and cached <v> instead.
 */
function storedLiteral(value: CellValue): CellValue {
  return isCircularError(value) ? textValue('#CIRC!') : value;
}

function isCircularError(value: CellValue): boolean {
  return value.type === 'error' && value.value === '#CIRC!';
}

function serializeBlankOrNull(original: CellOriginal | null, ref: string): string | null {
  if (original === null) return null;
  // A dirty blank cell with a style must keep its style reference; a dirty
  // blank with no style is removed by the caller. The model has no cell to
  // write here, so only formula-group edits can reach this branch.
  if (original.formulaKind === null) return null;
  throw new WorkbookSaveError(
    'unsupported-edit',
    `The ${original.formulaKind === 'array' ? 'legacy array formula' : 'shared formula group'} at ${ref} cannot be removed in this milestone. The file was not written.`,
  );
}

function serializeFormula(
  cell: Cell,
  original: CellOriginal | null,
  tag: (local: string) => string,
): string | null {
  const originalKind = original?.formulaKind ?? null;
  const currentKind = cell.formulaSource?.kind ?? null;

  if (cell.formula === null) {
    const unchangedFollower =
      originalKind === 'shared-follower' &&
      cell.formulaDerived &&
      currentKind === 'shared-follower' &&
      original?.formulaXml !== null &&
      original?.formulaXml !== undefined;
    if (unchangedFollower) return original.formulaXml;
    if (originalKind === 'array') {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'Legacy array formulas are read-only in this milestone; the cell was not changed.',
      );
    }
    if (originalKind === 'shared-master' || originalKind === 'shared-follower') {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'Shared formulas are read-only in this milestone; the file was not written.',
      );
    }
    return null;
  }

  if (currentKind === 'array') {
    if (
      originalKind !== 'array' ||
      original?.formulaXml === null ||
      original?.formulaXml === undefined
    ) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'Legacy array formulas are read-only in this milestone; the cell was not changed.',
      );
    }
    if (cell.formula !== original.formulaText) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'Legacy array formulas are read-only in this milestone; the cell was not changed.',
      );
    }
    return original.formulaXml;
  }

  if (currentKind === 'shared-master') {
    if (
      originalKind !== 'shared-master' ||
      cell.formula !== original?.formulaText ||
      original?.formulaXml === null ||
      original?.formulaXml === undefined
    ) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'Shared formulas are read-only in this milestone; the file was not written.',
      );
    }
    return original.formulaXml;
  }

  if (currentKind === 'shared-follower') {
    // The model cannot synthesise follower text yet, so a follower that
    // becomes a plain formula is refused rather than written with a broken group.
    throw new WorkbookSaveError(
      'unsupported-edit',
      'Shared formulas are read-only in this milestone; the file was not written.',
    );
  }

  if (
    originalKind === 'array' ||
    originalKind === 'shared-master' ||
    originalKind === 'shared-follower'
  ) {
    throw new WorkbookSaveError(
      'unsupported-edit',
      'Shared and array formulas are read-only in this milestone; the file was not written.',
    );
  }

  return `<${tag('f')}>${escapeXmlText(cell.formula)}</${tag('f')}>`;
}

function originalCellType(original: CellOriginal | null): string | null {
  return original?.attributes.find((attribute) => localName(attribute.name) === 't')?.value ?? null;
}

function localName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

function cellType(
  value: CellValue,
  savedResult: CellValue | null,
  formula: string | null,
  sharedStrings: SharedStringsAppender | null,
): string | null {
  if (formula !== null) {
    if (savedResult === null) return null;
    if (isCircularError(savedResult)) return null;
    switch (savedResult.type) {
      case 'text':
        return 'str';
      case 'boolean':
        return 'b';
      case 'error':
        return 'e';
      case 'number':
      case 'blank':
        return null;
    }
  }
  switch (value.type) {
    case 'text':
      return sharedStrings === null ? 'inlineStr' : 's';
    case 'boolean':
      return 'b';
    case 'error':
      return isCircularError(value) ? null : 'e';
    case 'number':
    case 'blank':
      return null;
  }
}

function serializedValue(
  value: CellValue,
  sharedStrings: SharedStringsAppender | null,
): string | null {
  switch (value.type) {
    case 'number':
      if (!Number.isFinite(value.value)) return null;
      return String(value.value);
    case 'text':
      if (sharedStrings !== null) return String(sharedStrings.indexFor(value.value));
      return null;
    case 'boolean':
      return value.value ? '1' : '0';
    case 'error':
      return isCircularError(value) ? null : value.value;
    case 'blank':
      return null;
  }
}

function serializedFormulaValue(value: CellValue): string | null {
  switch (value.type) {
    case 'number':
      return Number.isFinite(value.value) ? String(value.value) : null;
    case 'text':
      return value.value;
    case 'boolean':
      return value.value ? '1' : '0';
    case 'error':
      // Formula semantics §6: #CIRC! is an in-app-only error. A formula cell
      // keeps its authored text and its previous cached <v>, never #CIRC!.
      return isCircularError(value) ? null : value.value;
    case 'blank':
      return null;
  }
}
