import type { ColumnDimension, FormulaStorageKind, RowDimension, Sheet } from '@model/types';
import { formatCellRef } from '@model/types';
import type { SharedStringsTable } from '../types';
import type { SharedStringsAppender } from './shared-strings-patch';
import { cellToXml } from './cell-xml';
import { AFTER_AUTO_FILTER, type AutoFilterSavePlan } from '../autofilter-xml';
import type { ConditionalFormatSavePlan } from '../conditional-format-xml';
import {
  findElement,
  findElements,
  getAttribute,
  localName,
  makeElement,
  parseXmlTree,
  qualifiedName,
  removeAttribute,
  serializeNode,
  serializeXmlDocument,
  setAttribute,
  elementPrefix,
  textContent,
  type XmlElement,
} from './xml-tree';
import type { DirtyCellInput, DirtyDimensions } from './worksheet-patch';

export interface TreePatchInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly sheet: Sheet;
  readonly dirtyCells: readonly DirtyCellInput[];
  readonly dirtyDimensions?: DirtyDimensions;
  readonly syncMerges?: boolean;
  readonly syncAutoFilter?: AutoFilterSavePlan;
  readonly syncConditionalFormats?: ConditionalFormatSavePlan;
  readonly sharedStringsTable: SharedStringsTable | null;
  readonly sharedStrings: SharedStringsAppender | null;
}

const AFTER_SHEET_DATA = new Set([
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

/** Direct children of `<worksheet>` that follow `conditionalFormatting` (CT_Worksheet). */
const AFTER_CONDITIONAL_FORMATTING: ReadonlySet<string> = new Set([
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

/**
 * Children of `<worksheet>` that follow `autoFilter` in the schema order
 * (CT_Worksheet). The first of them is where a new filter element goes.
 */
export function patchWorksheetAsTree(input: TreePatchInput): Uint8Array {
  const xml = new TextDecoder().decode(input.sourceBytes);
  const document = parseXmlTree(xml, input.partName);
  const rootPrefix = elementPrefix(document.root);
  let sheetData = findElement(document.root, 'sheetData');
  if (sheetData === undefined) {
    sheetData = makeElement(qualifiedName(rootPrefix, 'sheetData'));
    const insertion = document.root.children.findIndex(
      (child) => child.type === 'element' && AFTER_SHEET_DATA.has(localName(child)),
    );
    if (insertion < 0) document.root.children.push(sheetData);
    else document.root.children.splice(insertion, 0, sheetData);
  }

  const sheetDataPrefix = elementPrefix(sheetData);
  const pendingRows = new Map<number, DirtyCellInput[]>();
  for (const dirty of input.dirtyCells) {
    const row = findRow(sheetData, dirty.row);
    if (row === undefined) {
      const list = pendingRows.get(dirty.row) ?? [];
      list.push(dirty);
      pendingRows.set(dirty.row, list);
      continue;
    }
    updateCell(row, dirty, input, elementPrefix(row) || sheetDataPrefix);
  }

  for (const [rowNumber, cells] of [...pendingRows.entries()].sort((a, b) => a[0] - b[0])) {
    const row = makeElement(qualifiedName(sheetDataPrefix, 'row'), [
      { name: 'r', value: String(rowNumber + 1) },
    ]);
    insertRow(sheetData, row, rowNumber);
    for (const dirty of cells.sort((a, b) => a.col - b.col)) {
      updateCell(row, dirty, input, sheetDataPrefix);
    }
  }

  applyDimensions(document.root, sheetData, input.dirtyDimensions);
  applyMerges(document.root, sheetData, input);
  applyAutoFilter(document.root, input.syncAutoFilter);
  applyConditionalFormats(document.root, input.syncConditionalFormats);

  return new TextEncoder().encode(serializeXmlDocument(document));
}

/**
 * Replace, insert or remove the sheet's `autoFilter` element in the tree
 * fallback. The element itself is parsed from the text the model wrote, so an
 * untouched filter keeps the source bytes it was re-emitted from.
 */
function applyAutoFilter(root: XmlElement, plan: AutoFilterSavePlan | undefined): void {
  if (plan === undefined || !plan.changed) return;
  const existing = findElement(root, 'autoFilter');
  if (plan.xml === null) {
    if (existing === undefined) return;
    const index = root.children.indexOf(existing);
    if (index >= 0) root.children.splice(index, 1);
    return;
  }
  const element = parseXmlTree(plan.xml, 'autoFilter element').root;
  if (existing !== undefined) {
    const index = root.children.indexOf(existing);
    root.children[index] = element;
    return;
  }
  const insertion = root.children.findIndex(
    (child) => child.type === 'element' && AFTER_AUTO_FILTER.has(localName(child)),
  );
  if (insertion < 0) root.children.push(element);
  else root.children.splice(insertion, 0, element);
}

/**
 * Rewrite the base `conditionalFormatting` elements from the model's records.
 * The x14 extension block is never touched: an x14 rule's ranges come from the
 * base rule it duplicates, and this milestone does not author rules.
 */
function applyConditionalFormats(
  root: XmlElement,
  plan: ConditionalFormatSavePlan | undefined,
): void {
  if (plan === undefined || !plan.changed) return;
  const existing = findElements(root, 'conditionalFormatting');
  const shared = Math.min(existing.length, plan.elements.length);
  for (let index = 0; index < shared; index += 1) {
    const element = parseXmlTree(plan.elements[index]!, 'conditionalFormatting element').root;
    root.children[root.children.indexOf(existing[index]!)] = element;
  }
  for (let index = plan.elements.length; index < existing.length; index += 1) {
    const target = existing[index]!;
    const position = root.children.indexOf(target);
    if (position >= 0) root.children.splice(position, 1);
  }
  if (plan.elements.length <= existing.length) return;
  const additions = plan.elements.slice(existing.length).map((xml) => {
    return parseXmlTree(xml, 'conditionalFormatting element').root;
  });
  const insertion = root.children.findIndex(
    (child) => child.type === 'element' && AFTER_CONDITIONAL_FORMATTING.has(localName(child)),
  );
  if (insertion < 0) root.children.push(...additions);
  else root.children.splice(insertion, 0, ...additions);
}

function applyMerges(root: XmlElement, sheetData: XmlElement, input: TreePatchInput): void {
  if (input.syncMerges !== true) return;
  for (const existing of findElements(root, 'mergeCells')) {
    const index = root.children.indexOf(existing);
    if (index >= 0) root.children.splice(index, 1);
  }
  if (input.sheet.merges.length === 0) return;
  const prefix = elementPrefix(root);
  const mergeCells = makeElement(qualifiedName(prefix, 'mergeCells'), [
    { name: 'count', value: String(input.sheet.merges.length) },
  ]);
  for (const merge of input.sheet.merges) {
    mergeCells.children.push(
      makeElement(qualifiedName(prefix, 'mergeCell'), [
        {
          name: 'ref',
          value: `${formatCellRef(merge.range.startRow, merge.range.startCol)}:${formatCellRef(
            merge.range.endRow,
            merge.range.endCol,
          )}`,
        },
      ]),
    );
  }
  const position = root.children.indexOf(sheetData);
  root.children.splice(position + 1, 0, mergeCells);
}

function applyDimensions(
  root: XmlElement,
  sheetData: XmlElement,
  dimensions: DirtyDimensions | undefined,
): void {
  if (dimensions === undefined) return;
  const rootPrefix = elementPrefix(root);
  const rowPrefix = elementPrefix(sheetData) || rootPrefix;
  for (const [rowNumber, dimension] of dimensions.rows) {
    if (!Number.isInteger(rowNumber) || rowNumber < 0 || rowNumber >= 1_048_576) continue;
    const row = findRow(sheetData, rowNumber);
    if (row !== undefined) {
      applyRowDimension(row, dimension);
      continue;
    }
    if (dimension === null) continue;
    const newRow = makeElement(qualifiedName(rowPrefix, 'row'), [
      { name: 'r', value: String(rowNumber + 1) },
    ]);
    applyRowDimension(newRow, dimension);
    insertRow(sheetData, newRow, rowNumber);
  }

  if (dimensions.columns.size === 0) return;
  const dirtyColumns = [...dimensions.columns.entries()]
    .filter(([index]) => Number.isInteger(index) && index >= 0 && index < 16_384)
    .sort((a, b) => a[0] - b[0]);
  if (dirtyColumns.length === 0) return;

  let cols = findElement(root, 'cols');
  if (cols === undefined) {
    if (dirtyColumns.every(([, dimension]) => dimension === null)) return;
    cols = makeElement(qualifiedName(rootPrefix, 'cols'));
    const sheetDataIndex = root.children.indexOf(sheetData);
    if (sheetDataIndex < 0) root.children.push(cols);
    else root.children.splice(sheetDataIndex, 0, cols);
  }
  const colsPrefix = elementPrefix(cols) || rootPrefix;
  const covered = new Set<number>();

  for (const column of [...findElements(cols, 'col')]) {
    const min = readColumnIndex(column, 'min');
    const max = readColumnIndex(column, 'max');
    if (min === null || max === null || max < min) continue;
    const dirtyIndices = dirtyColumns.filter(([index]) => index >= min && index <= max);
    if (dirtyIndices.length === 0) continue;
    const replacements: XmlElement[] = [];
    let cursor = min;
    for (const [index, dimension] of dirtyIndices) {
      if (index > cursor)
        replacements.push(cloneColumn(column, colsPrefix, cursor, index - 1, undefined));
      replacements.push(cloneColumn(column, colsPrefix, index, index, dimension));
      covered.add(index);
      cursor = index + 1;
    }
    if (cursor <= max) replacements.push(cloneColumn(column, colsPrefix, cursor, max, undefined));
    const childIndex = cols.children.indexOf(column);
    cols.children.splice(childIndex, 1, ...replacements);
  }

  const uncovered = dirtyColumns.filter(
    ([index, dimension]) => dimension !== null && !covered.has(index),
  );
  for (const [index, dimension] of uncovered) {
    const element = cloneColumn(undefined, colsPrefix, index, index, dimension);
    const existing = findElements(cols, 'col');
    const next = existing.find((candidate) => {
      const candidateMin = readColumnIndex(candidate, 'min');
      return candidateMin !== null && candidateMin > index;
    });
    if (next !== undefined) {
      cols.children.splice(cols.children.indexOf(next), 0, element);
    } else {
      cols.children.push(element);
    }
  }
}

function applyRowDimension(element: XmlElement, dimension: RowDimension | null): void {
  setOrRemoveAttribute(
    element,
    'ht',
    dimension?.height === undefined ? null : dimensionNumber(dimension.height),
  );
  setOrRemoveAttribute(element, 'customHeight', dimension?.customHeight === true ? '1' : null);
  setOrRemoveAttribute(element, 'hidden', dimension?.hidden === true ? '1' : null);
  setOrRemoveAttribute(
    element,
    's',
    dimension?.styleId === undefined || dimension.styleId === 0 ? null : String(dimension.styleId),
  );
  setOrRemoveAttribute(
    element,
    'outlineLevel',
    dimension?.outlineLevel === undefined || dimension.outlineLevel <= 0
      ? null
      : String(dimension.outlineLevel),
  );
}

function applyColumnDimension(element: XmlElement, dimension: ColumnDimension | null): void {
  setOrRemoveAttribute(
    element,
    'width',
    dimension?.width === undefined ? null : dimensionNumber(dimension.width),
  );
  setOrRemoveAttribute(element, 'hidden', dimension?.hidden === true ? '1' : null);
  setOrRemoveAttribute(
    element,
    's',
    dimension?.styleId === undefined || dimension.styleId === 0 ? null : String(dimension.styleId),
  );
  setOrRemoveAttribute(element, 'customWidth', dimension?.customWidth === true ? '1' : null);
  setOrRemoveAttribute(element, 'bestFit', dimension?.bestFit === true ? '1' : null);
  setOrRemoveAttribute(
    element,
    'outlineLevel',
    dimension?.outlineLevel === undefined || dimension.outlineLevel <= 0
      ? null
      : String(dimension.outlineLevel),
  );
}

function cloneColumn(
  source: XmlElement | undefined,
  prefix: string,
  min: number,
  max: number,
  dimension: ColumnDimension | null | undefined,
): XmlElement {
  const element = makeElement(
    qualifiedName(source === undefined ? prefix : elementPrefix(source) || prefix, 'col'),
    source === undefined ? [] : source.attributes.map((attribute) => ({ ...attribute })),
  );
  setAttribute(element, 'min', String(min + 1));
  setAttribute(element, 'max', String(max + 1));
  if (dimension !== undefined) applyColumnDimension(element, dimension);
  return element;
}

function readColumnIndex(element: XmlElement, name: string): number | null {
  const raw = getAttribute(element, name);
  if (raw === undefined) return null;
  const value = Number(raw);
  return Number.isInteger(value) && value >= 1 ? value - 1 : null;
}

function setOrRemoveAttribute(element: XmlElement, name: string, value: string | null): void {
  if (value === null) removeAttribute(element, name);
  else setAttribute(element, name, value);
}

function dimensionNumber(value: number): string {
  return Number.isFinite(value) ? String(value) : '0';
}

function findRow(sheetData: XmlElement, row: number): XmlElement | undefined {
  let previous = -1;
  for (const element of findElements(sheetData, 'row')) {
    const raw = getAttribute(element, 'r');
    const numeric = raw === undefined ? NaN : Number(raw);
    const number = Number.isInteger(numeric) && numeric > 0 ? numeric - 1 : previous + 1;
    previous = number;
    if (number === row) return element;
  }
  return undefined;
}

function insertRow(sheetData: XmlElement, row: XmlElement, rowNumber: number): void {
  const rows = findElements(sheetData, 'row');
  let previous = -1;
  const next = rows.find((candidate) => {
    const raw = getAttribute(candidate, 'r');
    const numeric = raw === undefined ? NaN : Number(raw);
    const number = Number.isInteger(numeric) && numeric > 0 ? numeric - 1 : previous + 1;
    previous = number;
    return number > rowNumber;
  });
  if (next !== undefined) {
    const index = sheetData.children.indexOf(next);
    sheetData.children.splice(index, 0, row);
    return;
  }
  const lastRow = rows.at(-1);
  if (lastRow !== undefined) {
    const index = sheetData.children.indexOf(lastRow);
    sheetData.children.splice(index + 1, 0, row);
  } else {
    sheetData.children.push(row);
  }
}

function updateCell(
  row: XmlElement,
  dirty: DirtyCellInput,
  input: TreePatchInput,
  prefix: string,
): void {
  const existing = findCell(row, dirty.col);
  const original =
    existing === undefined
      ? null
      : {
          attributes: existing.attributes,
          unknownChildren: existing.children
            .filter(
              (child) => child.type === 'element' && !['f', 'v', 'is'].includes(localName(child)),
            )
            .map(serializeNode),
          ...formulaOriginal(existing),
        };
  const replacement = cellToXml(prefix, formatCellRef(dirty.row, dirty.col), dirty.cell, {
    sharedStrings: input.sharedStrings,
    original,
  });
  if (existing !== undefined) {
    if (replacement === null) {
      row.children = row.children.filter((child) => child !== existing);
      return;
    }
    const element = parseCellFragment(replacement, input.partName);
    const index = row.children.indexOf(existing);
    row.children[index] = element;
    return;
  }
  if (replacement === null) return;
  const element = parseCellFragment(replacement, input.partName);
  insertCell(row, element, dirty.col);
}

function formulaOriginal(cell: XmlElement): {
  formulaXml: string | null;
  formulaText: string | null;
  formulaKind: FormulaStorageKind | null;
} {
  const formula = findElement(cell, 'f');
  if (formula === undefined) {
    return { formulaXml: null, formulaText: null, formulaKind: null };
  }
  const text = textContent(formula);
  const type = getAttribute(formula, 't');
  let kind: FormulaStorageKind | null = null;
  if (type === 'array') kind = 'array';
  else if (type === 'shared') kind = text.length > 0 ? 'shared-master' : 'shared-follower';
  return {
    formulaXml: serializeNode(formula),
    formulaText: text.length > 0 ? text : null,
    formulaKind: kind,
  };
}

function findCell(row: XmlElement, col: number): XmlElement | undefined {
  return findElements(row, 'c').find((element) => {
    const ref = getAttribute(element, 'r');
    if (ref === undefined) return false;
    const parsed = /^([A-Z]+)(\d+)$/i.exec(ref);
    if (parsed === null) return false;
    return columnIndex(parsed[1]!) === col;
  });
}

function insertCell(row: XmlElement, cell: XmlElement, col: number): void {
  const cells = findElements(row, 'c');
  const next = cells.find((candidate) => {
    const ref = getAttribute(candidate, 'r');
    if (ref === undefined) return false;
    const parsed = /^([A-Z]+)(\d+)$/i.exec(ref);
    return parsed !== null && columnIndex(parsed[1]!) > col;
  });
  if (next !== undefined) {
    const index = row.children.indexOf(next);
    row.children.splice(index, 0, cell);
    return;
  }
  const lastCell = cells.at(-1);
  if (lastCell !== undefined) {
    const index = row.children.indexOf(lastCell);
    row.children.splice(index + 1, 0, cell);
    return;
  }
  const firstNonCell = row.children.findIndex(
    (child) => child.type === 'element' && localName(child) !== 'c',
  );
  if (firstNonCell < 0) row.children.push(cell);
  else row.children.splice(firstNonCell, 0, cell);
}

function parseCellFragment(xml: string, partName: string): XmlElement {
  const document = parseXmlTree(xml, `${partName} (cell)`);
  return document.root;
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}
