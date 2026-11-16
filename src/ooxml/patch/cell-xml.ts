import type { Cell, FormulaStorageKind } from '@model/types';
import { textValue, type CellValue } from '@model/value';
import { WorkbookSaveError } from '../errors';
import { escapeXmlAttribute, escapeXmlText } from '../xml';
import type { SharedStringsAppender } from './shared-strings-patch';

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
  cell: Cell | null,
  options: CellWriteOptions,
): string | null {
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
  const literalCell = formula === null && cell.formulaSource === undefined;
  const storedValue = literalCell ? storedLiteral(cell.value) : cell.value;
  const type = cellType(storedValue, cell.savedResult, formula, options.sharedStrings);
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
      formula !== null
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
