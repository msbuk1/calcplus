import type { FormulaStorageKind, Sheet } from '@model/types';
import { formatCellRef } from '@model/types';
import type { SharedStringsTable } from '../types';
import type { SharedStringsAppender } from './shared-strings-patch';
import { cellToXml } from './cell-xml';
import {
  findElement,
  findElements,
  getAttribute,
  localName,
  makeElement,
  parseXmlTree,
  qualifiedName,
  serializeNode,
  serializeXmlDocument,
  elementPrefix,
  textContent,
  type XmlElement,
} from './xml-tree';
import type { DirtyCellInput } from './worksheet-patch';

export interface TreePatchInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly sheet: Sheet;
  readonly dirtyCells: readonly DirtyCellInput[];
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

  return new TextEncoder().encode(serializeXmlDocument(document));
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
