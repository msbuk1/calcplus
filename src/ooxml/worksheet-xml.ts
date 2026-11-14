import type { Cell, FormulaSource, FormulaStorageKind, GridRange, Sheet } from '@model/types';
import {
  BLANK,
  booleanValue,
  errorValue,
  isCellErrorCode,
  numberValue,
  textValue,
  type CellValue,
} from '@model/value';
import { createSheet, createSheetId } from '@model/workbook';
import { WorkbookOpenError, WorkbookSaveError } from './errors';
import type { SharedStringEntry, SharedStringsTable } from './types';
import { SPREADSHEET_NS, STRICT_SPREADSHEET_NS, decodeXml, openTagStart, parseXml } from './xml';
import type { SaxesAttributeNS, SaxesTagNS } from 'saxes';

export interface RawAttribute {
  readonly name: string;
  readonly value: string;
}

export interface WorksheetCellPosition {
  readonly ref: string;
  readonly row: number;
  readonly col: number;
  readonly start: number;
  readonly end: number;
  readonly openEnd: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
  readonly attributes: readonly RawAttribute[];
  readonly styleId: number;
  readonly type: string | null;
  readonly unknownChildren: readonly string[];
  readonly formulaXml: string | null;
  readonly formulaText: string | null;
  readonly formulaKind: FormulaStorageKind | null;
}

export interface WorksheetRowPosition {
  readonly number: number;
  readonly start: number;
  readonly end: number;
  readonly openEnd: number;
  readonly contentEnd: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
  readonly cells: Map<number, WorksheetCellPosition>;
}

export interface SheetDataPosition {
  readonly start: number;
  readonly openEnd: number;
  readonly end: number;
  readonly contentEnd: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
}

export interface ParsedWorksheet {
  readonly sheet: Sheet;
  readonly xml: string;
  readonly sheetData: SheetDataPosition | null;
  readonly rows: Map<number, WorksheetRowPosition>;
  readonly unsafeReasons: readonly string[];
  readonly rootPrefix: string;
}

interface OpenElement {
  readonly local: string;
  readonly start: number;
  readonly openEnd: number;
  readonly prefix: string;
}

interface CellAccumulator {
  row: number;
  col: number;
  ref: string;
  start: number;
  openEnd: number;
  prefix: string;
  attributes: RawAttribute[];
  styleId: number;
  type: string | null;
  formulaText: string;
  formulaType: string | null;
  formulaStart: number;
  formulaEnd: number;
  formulaAttributes: RawAttribute[];
  valueText: string;
  hasFormula: boolean;
  hasValue: boolean;
  inlineText: string;
  inlineStart: number;
  inlineXml: string;
  inlineRich: boolean;
  inlineT: boolean;
  unknownChildren: string[];
}

export interface ParseWorksheetOptions {
  readonly sharedStrings: SharedStringsTable | null;
  readonly partPath: string;
  readonly name: string;
  readonly knownSheetNames?: ReadonlySet<string>;
}

export function parseWorksheetXml(
  bytes: Uint8Array,
  options: ParseWorksheetOptions,
): ParsedWorksheet {
  const xml = decodeXml(bytes);
  const sheet = createSheet(options.name, options.partPath, createSheetId(options.partPath));
  const unsafeReasons: string[] = [];
  const rows = new Map<number, WorksheetRowPosition>();
  const stack: OpenElement[] = [];
  const rootHolder: { root: OpenElement | null } = { root: null };
  const sheetDataHolder: { value: SheetDataPosition | null } = { value: null };
  let sheetDataCount = 0;
  let row: WorksheetRowPosition | null = null;
  let rowAccumulator: {
    number: number;
    start: number;
    openEnd: number;
    prefix: string;
    cells: Map<number, WorksheetCellPosition>;
    missingRef: boolean;
  } | null = null;
  let cell: CellAccumulator | null = null;
  let lastRowNumber = -1;
  let lastCellColumn = -1;
  let inValue = false;
  let inFormula = false;
  let inInlineString = false;
  let inInlineText = false;
  let inPhonetic = false;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        if (tag.uri === STRICT_SPREADSHEET_NS) {
          throw new WorkbookOpenError(
            'strict-ooxml',
            'This workbook uses an OOXML variant calcplus does not support yet.',
          );
        }
        const start = openTagStart(xml, position);
        stack.push({ local: tag.local, start, openEnd: position, prefix: tag.prefix });
        const parent = stack.length >= 2 ? stack[stack.length - 2]!.local : null;

        if (tag.local === 'worksheet' && rootHolder.root === null) {
          if (tag.uri !== SPREADSHEET_NS) {
            throw new WorkbookOpenError('corrupt', `${options.partPath} is not a worksheet part.`);
          }
          rootHolder.root = { local: tag.local, start, openEnd: position, prefix: tag.prefix };
          return;
        }

        if (tag.local === 'sheetData') {
          sheetDataCount += 1;
          if (sheetDataCount > 1) unsafeReasons.push('multiple sheetData elements');
          sheetDataHolder.value = {
            start,
            openEnd: position,
            end: position,
            contentEnd: position,
            selfClosing: false,
            prefix: tag.prefix,
          };
          return;
        }

        if (tag.local === 'row' && parent === 'sheetData') {
          if (rowAccumulator !== null) unsafeReasons.push('nested row elements');
          const refText = attributeValue(tag, 'r');
          const missingRef = refText === undefined || !/^\d+$/.test(refText);
          if (missingRef) unsafeReasons.push('row without a numeric r attribute');
          const number = missingRef ? lastRowNumber + 1 : Number(refText) - 1;
          if (!Number.isInteger(number) || number < 0) {
            unsafeReasons.push('row r attribute is outside the grid');
          }
          lastRowNumber = number;
          rowAccumulator = {
            number,
            start,
            openEnd: position,
            prefix: tag.prefix,
            cells: new Map(),
            missingRef,
          };
          row = {
            number,
            start,
            end: position,
            openEnd: position,
            contentEnd: position,
            selfClosing: false,
            prefix: tag.prefix,
            cells: rowAccumulator.cells,
          };
          applyRowAttributes(sheet, tag, number);
          const hidden = attributeValue(tag, 'hidden');
          if (hidden !== undefined) {
            const dimension = sheet.rowDimensions.get(number) ?? {};
            dimension.hidden = hidden === '1';
            sheet.rowDimensions.set(number, dimension);
          }
          return;
        }

        if (tag.local === 'c' && parent === 'row' && rowAccumulator !== null) {
          const refText = attributeValue(tag, 'r');
          const missingRef = refText === undefined || !/^[A-Z]+\d+$/i.test(refText);
          if (missingRef) unsafeReasons.push('cell without a valid r attribute');
          let col = lastCellColumn + 1;
          let ref = '';
          if (!missingRef) {
            const parsed = parseRef(refText!);
            col = parsed.col;
            ref = `${columnLetters(col)}${rowAccumulator.number + 1}`;
          } else {
            ref = `${columnLetters(col)}${rowAccumulator.number + 1}`;
          }
          if (col <= lastCellColumn) unsafeReasons.push('cells out of order');
          lastCellColumn = col;
          const styleText = attributeValue(tag, 's');
          const styleId = styleText === undefined ? 0 : Number(styleText) || 0;
          cell = {
            row: rowAccumulator.number,
            col,
            ref,
            start,
            openEnd: position,
            prefix: tag.prefix,
            attributes: attributesOf(tag),
            styleId,
            type: attributeValue(tag, 't') ?? null,
            formulaText: '',
            formulaType: null,
            formulaStart: -1,
            formulaEnd: -1,
            formulaAttributes: [],
            valueText: '',
            hasFormula: false,
            hasValue: false,
            inlineText: '',
            inlineStart: start,
            inlineXml: '',
            inlineRich: false,
            inlineT: false,
            unknownChildren: [],
          };
          return;
        }

        if (cell !== null && parent === 'c') {
          switch (tag.local) {
            case 'f':
              cell.hasFormula = true;
              cell.formulaType = attributeValue(tag, 't') ?? null;
              cell.formulaStart = start;
              cell.formulaAttributes = attributesOf(tag);
              inFormula = true;
              break;
            case 'v':
              cell.hasValue = true;
              inValue = true;
              break;
            case 'is':
              cell.inlineStart = start;
              inInlineString = true;
              break;
            case 'r':
              if (inInlineString) cell.inlineRich = true;
              break;
            case 'rPh':
              inPhonetic = true;
              break;
            case 't':
              if (inInlineString) {
                inInlineText = true;
                cell.inlineT = true;
              }
              break;
            default:
              break;
          }
          return;
        }

        if (cell !== null && inInlineString && parent === 'is') {
          if (tag.local === 't') {
            inInlineText = true;
            cell.inlineT = true;
          } else if (tag.local === 'r') {
            cell.inlineRich = true;
          } else if (tag.local === 'rPh') {
            inPhonetic = true;
          }
          return;
        }

        switch (tag.local) {
          case 'col': {
            applyColumnAttributes(sheet, tag);
            break;
          }
          case 'mergeCell': {
            const ref = attributeValue(tag, 'ref');
            if (ref !== undefined) {
              try {
                sheet.merges.push({ range: parseRange(ref) });
              } catch {
                unsafeReasons.push(`malformed merge range ${ref}`);
              }
            }
            break;
          }
          case 'sheetView': {
            const gridLines = attributeValue(tag, 'showGridLines');
            const headings = attributeValue(tag, 'showHeadings');
            const zoom = attributeValue(tag, 'zoomScale');
            if (gridLines !== undefined) sheet.view.showGridLines = gridLines !== '0';
            if (headings !== undefined) sheet.view.showHeadings = headings !== '0';
            if (zoom !== undefined) {
              const value = Number(zoom);
              if (Number.isFinite(value) && value > 0) sheet.view.zoom = value;
            }
            break;
          }
          case 'pane': {
            const state = attributeValue(tag, 'state');
            if (state === 'frozen' || state === 'frozenSplit') {
              const xSplit = Number(attributeValue(tag, 'xSplit') ?? '0');
              const ySplit = Number(attributeValue(tag, 'ySplit') ?? '0');
              if (Number.isFinite(xSplit) && xSplit > 0) sheet.view.freezeCols = xSplit;
              if (Number.isFinite(ySplit) && ySplit > 0) sheet.view.freezeRows = ySplit;
            }
            break;
          }
          case 'sheetProtection': {
            sheet.protection = { rawXml: xml.slice(start, position) };
            break;
          }
          default:
            break;
        }
      },
      onText(text) {
        if (cell === null) return;
        if (inValue) cell.valueText += text;
        else if (inFormula) cell.formulaText += text;
        else if (inInlineText && !inPhonetic) cell.inlineText += text;
      },
      onCdata(text) {
        if (cell === null) return;
        if (inValue) cell.valueText += text;
        else if (inFormula) cell.formulaText += text;
        else if (inInlineText && !inPhonetic) cell.inlineText += text;
      },
      onCloseTag(tag, start, end) {
        const open = stack.pop();
        const parent = stack.length > 0 ? stack[stack.length - 1]!.local : null;

        if (cell !== null && tag.local === 'c' && parent === 'row' && open !== undefined) {
          const positionRecord: WorksheetCellPosition = {
            ref: cell.ref,
            row: cell.row,
            col: cell.col,
            start: cell.start,
            end,
            openEnd: cell.openEnd,
            selfClosing: cell.openEnd === end,
            prefix: cell.prefix,
            attributes: cell.attributes,
            styleId: cell.styleId,
            type: cell.type,
            unknownChildren: cell.unknownChildren,
            formulaXml: formulaXmlOf(cell, xml),
            formulaText: cell.hasFormula && cell.formulaText.length > 0 ? cell.formulaText : null,
            formulaKind: formulaKindOf(cell),
          };
          if (rowAccumulator !== null) {
            if (rowAccumulator.cells.has(cell.col)) unsafeReasons.push('duplicate cell reference');
            rowAccumulator.cells.set(cell.col, positionRecord);
          }
          const modelCell = cellFromAccumulator(cell, options.sharedStrings, xml);
          if (modelCell !== null) {
            sheet.cells.set(`${cell.row}:${cell.col}`, modelCell);
          }
          cell = null;
          inValue = false;
          inFormula = false;
          inInlineString = false;
          inInlineText = false;
          inPhonetic = false;
          return;
        }

        if (tag.local === 'v') inValue = false;
        if (tag.local === 'f') {
          if (cell !== null) cell.formulaEnd = end;
          inFormula = false;
        }
        if (tag.local === 'is') {
          if (cell !== null && open !== undefined) cell.inlineXml = xml.slice(open.start, end);
          inInlineString = false;
          inInlineText = false;
          inPhonetic = false;
        }
        if (tag.local === 't' && inInlineString) inInlineText = false;
        if (tag.local === 'rPh') inPhonetic = false;

        if (
          cell !== null &&
          parent === 'c' &&
          tag.local !== 'f' &&
          tag.local !== 'v' &&
          tag.local !== 'is' &&
          open !== undefined
        ) {
          cell.unknownChildren.push(xml.slice(open.start, end));
        }

        if (tag.local === 'row' && open !== undefined && row !== null) {
          if (rowAccumulator !== null && rowAccumulator.cells.size === 0) {
            // An empty row is legal; it still carries dimensions and must be preserved.
          }
          const finalRow: WorksheetRowPosition = {
            ...row,
            end,
            contentEnd: start,
            selfClosing: row.openEnd === end,
          };
          if (rows.has(finalRow.number)) unsafeReasons.push('duplicate row number');
          rows.set(finalRow.number, finalRow);
          row = null;
          rowAccumulator = null;
          lastCellColumn = -1;
        }

        if (tag.local === 'sheetData' && sheetDataHolder.value !== null && open !== undefined) {
          sheetDataHolder.value = {
            ...sheetDataHolder.value,
            end,
            contentEnd: start,
            selfClosing: open.openEnd === end,
          };
        }
      },
    },
    options.partPath,
  );

  const root = rootHolder.root;
  const sheetData = sheetDataHolder.value;
  if (root === null) {
    throw new WorkbookOpenError('corrupt', `${options.partPath} is not a worksheet part.`);
  }
  if (sheetData === null) {
    unsafeReasons.push('missing sheetData element');
  } else if (sheetData.selfClosing) {
    // `<sheetData/>` is a valid empty sheet; the patch writer expands it.
  }

  if (sheet.view.freezeRows !== 0) sheet.view.freezeRows = Math.floor(sheet.view.freezeRows);
  if (sheet.view.freezeCols !== 0) sheet.view.freezeCols = Math.floor(sheet.view.freezeCols);

  return {
    sheet,
    xml,
    sheetData,
    rows,
    unsafeReasons,
    rootPrefix: root.prefix,
  };
}

function cellFromAccumulator(
  cell: CellAccumulator,
  sharedStrings: SharedStringsTable | null,
  xml: string,
): Cell | null {
  const type = cell.type ?? 'n';
  const parsed = parseValue(type, cell, sharedStrings);
  const value = parsed.value;
  const savedResult = cell.hasFormula && cell.hasValue ? value : cell.hasFormula ? null : null;
  const finalValue = cell.hasFormula ? (savedResult ?? BLANK) : value;
  const formulaText = cell.hasFormula && cell.formulaText.length > 0 ? cell.formulaText : null;
  const formulaDerived = cell.formulaType === 'shared' && formulaText === null;
  const formulaKind = formulaKindOf(cell);
  const formulaSource: FormulaSource | undefined =
    formulaKind === null
      ? undefined
      : {
          kind: formulaKind,
          ...(sharedIndexOf(cell) === undefined ? {} : { sharedIndex: sharedIndexOf(cell) }),
          ...(formulaRangeOf(cell) === undefined ? {} : { range: formulaRangeOf(cell) }),
          xml: formulaXmlOf(cell, xml) ?? '',
        };
  const hasAnything =
    finalValue.type !== 'blank' ||
    formulaText !== null ||
    formulaKind !== null ||
    parsed.richText !== null ||
    cell.styleId !== 0;
  if (!hasAnything) return null;
  return {
    value: finalValue,
    formula: formulaText,
    formulaDerived,
    savedResult,
    styleId: cell.styleId,
    richText: parsed.richText,
    ...(formulaSource === undefined ? {} : { formulaSource }),
  };
}

function formulaKindOf(cell: CellAccumulator): FormulaStorageKind | null {
  if (!cell.hasFormula) return null;
  if (cell.formulaType === 'array') return 'array';
  if (cell.formulaType !== 'shared') return null;
  return cell.formulaText.length > 0 ? 'shared-master' : 'shared-follower';
}

function formulaXmlOf(cell: CellAccumulator, xml: string): string | null {
  if (!cell.hasFormula || cell.formulaStart < 0 || cell.formulaEnd < cell.formulaStart) {
    return null;
  }
  return xml.slice(cell.formulaStart, cell.formulaEnd);
}

function formulaRangeOf(cell: CellAccumulator): string | undefined {
  return cell.formulaAttributes.find((attribute) => localAttributeName(attribute.name) === 'ref')
    ?.value;
}

function sharedIndexOf(cell: CellAccumulator): string | undefined {
  return cell.formulaAttributes.find((attribute) => localAttributeName(attribute.name) === 'si')
    ?.value;
}

function localAttributeName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

function parseValue(
  type: string,
  cell: CellAccumulator,
  sharedStrings: SharedStringsTable | null,
): { value: CellValue; richText: Cell['richText'] } {
  switch (type) {
    case 's': {
      const index = Number(cell.valueText);
      const entry = sharedStrings?.entries[index];
      if (entry === undefined) {
        return { value: errorValue('#REF!'), richText: null };
      }
      return {
        value: textValue(entry.text),
        richText: entry.richText ? { xml: entry.rawXml, plainText: entry.text } : null,
      };
    }
    case 'inlineStr': {
      return {
        value: textValue(cell.inlineText),
        richText: cell.inlineRich ? { xml: cell.inlineXml, plainText: cell.inlineText } : null,
      };
    }
    case 'str':
      return { value: textValue(cell.valueText), richText: null };
    case 'b':
      return { value: booleanValue(cell.valueText === '1'), richText: null };
    case 'e':
      return {
        value: isCellErrorCode(cell.valueText) ? errorValue(cell.valueText) : errorValue('#VALUE!'),
        richText: null,
      };
    case 'd':
      return { value: textValue(cell.valueText), richText: null };
    case 'n':
    default: {
      if (!cell.hasValue) return { value: BLANK, richText: null };
      const numeric = Number(cell.valueText);
      return Number.isFinite(numeric)
        ? { value: numberValue(numeric), richText: null }
        : { value: BLANK, richText: null };
    }
  }
}

function applyRowAttributes(sheet: Sheet, tag: SaxesTagNS, row: number): void {
  const heightText = attributeValue(tag, 'ht');
  const customHeight = attributeValue(tag, 'customHeight');
  const styleText = attributeValue(tag, 's');
  const dimension = sheet.rowDimensions.get(row) ?? {};
  if (heightText !== undefined) {
    const height = Number(heightText);
    if (Number.isFinite(height) && height >= 0) dimension.height = height;
  }
  if (customHeight !== undefined && customHeight !== '0') dimension.height = dimension.height ?? 15;
  if (styleText !== undefined) dimension.styleId = Number(styleText) || 0;
  if (Object.keys(dimension).length > 0) sheet.rowDimensions.set(row, dimension);
}

function applyColumnAttributes(sheet: Sheet, tag: SaxesTagNS): void {
  const min = Number(attributeValue(tag, 'min') ?? '0');
  const max = Number(attributeValue(tag, 'max') ?? '0');
  const widthText = attributeValue(tag, 'width');
  const hidden = attributeValue(tag, 'hidden');
  const styleText = attributeValue(tag, 's');
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min) return;
  const width = widthText === undefined ? undefined : Number(widthText);
  for (let column = min; column <= Math.min(max, 16_384); column += 1) {
    const index = column - 1;
    const dimension = sheet.columnDimensions.get(index) ?? {};
    if (width !== undefined && Number.isFinite(width) && width >= 0) dimension.width = width;
    if (hidden !== undefined && hidden !== '0') dimension.hidden = true;
    if (styleText !== undefined) dimension.styleId = Number(styleText) || 0;
    if (Object.keys(dimension).length > 0) sheet.columnDimensions.set(index, dimension);
  }
}

function attributesOf(tag: SaxesTagNS): RawAttribute[] {
  return (Object.values(tag.attributes) as SaxesAttributeNS[]).map((attribute) => ({
    name: attribute.name,
    value: attribute.value,
  }));
}

function attributeValue(tag: SaxesTagNS, local: string): string | undefined {
  for (const attribute of Object.values(tag.attributes) as SaxesAttributeNS[]) {
    if (attribute.local === local) return attribute.value;
  }
  return undefined;
}

function parseRef(ref: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/i.exec(ref);
  if (match === null) throw new Error(`Invalid ref ${ref}`);
  return { row: Number(match[2]) - 1, col: columnIndex(match[1]!) };
}

function parseRange(ref: string): GridRange {
  const [startText, endText = startText] = ref.split(':');
  const start = parseRef(startText!);
  const end = parseRef(endText!);
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function columnLetters(column: number): string {
  let value = column;
  let result = '';
  while (value >= 0) {
    const digit = value % 26;
    result = String.fromCharCode(65 + digit) + result;
    value = Math.floor(value / 26) - 1;
  }
  return result;
}

export function worksheetCellValueForSave(cell: Cell | null): CellValue {
  return cell?.value ?? BLANK;
}

export function assertWorksheetPartSize(partName: string, byteLength: number): void {
  if (byteLength >= 16 * 1024 * 1024) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${partName} could not be patched safely and is too large (${byteLength} bytes) for the tree fallback. The file was not written.`,
    );
  }
}

export function unknownChildFragment(cell: WorksheetCellPosition): string {
  return cell.unknownChildren.join('');
}

export function sharedStringEntry(
  table: SharedStringsTable,
  index: number,
): SharedStringEntry | undefined {
  return table.entries[index];
}
