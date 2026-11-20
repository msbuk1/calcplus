import type { Cell, ColumnDimension, RowDimension, Sheet } from '@model/types';
import { formatCellRef } from '@model/types';
import type { SharedStringsTable } from '../types';
import type { SharedStringsAppender } from './shared-strings-patch';
import { cellToXml } from './cell-xml';
import type {
  ParsedWorksheet,
  WorksheetElementPosition,
  WorksheetRowPosition,
} from '../worksheet-xml';
import { parseWorksheetXml } from '../worksheet-xml';
import { AFTER_AUTO_FILTER, type AutoFilterSavePlan } from '../autofilter-xml';
import type { ConditionalFormatSavePlan } from '../conditional-format-xml';
import type { PaneSavePlan } from '../pane-xml';
import { WorkbookSaveError } from '../errors';
import { escapeXmlAttribute, isWellFormedXml } from '../xml';
import { patchWorksheetAsTree } from './tree-patch';
import { MAX_TREE_PATCH_BYTES } from '../xml';

export interface DirtyCellInput {
  readonly row: number;
  readonly col: number;
  readonly cell: Cell | null;
}

/** Dimensions changed since the last save, with the exact record to write. */
export interface DirtyDimensions {
  readonly rows: ReadonlyMap<number, RowDimension | null>;
  readonly columns: ReadonlyMap<number, ColumnDimension | null>;
}

export interface WorksheetSaveInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly sheet: Sheet;
  readonly dirtyCells: readonly DirtyCellInput[];
  readonly dirtyDimensions?: DirtyDimensions;
  /** Replace the merge list from the model, even when no cell changed. */
  readonly syncMerges?: boolean;
  /** Write or remove the sheet's `autoFilter` element, when it changed. */
  readonly syncAutoFilter?: AutoFilterSavePlan;
  /** Rewrite the sheet's conditional-format elements, when they changed. */
  readonly syncConditionalFormats?: ConditionalFormatSavePlan;
  /** Replace the sheet view's frozen pane, when the split changed. */
  readonly syncPane?: PaneSavePlan;
  readonly sharedStringsTable: SharedStringsTable | null;
  readonly sharedStrings: SharedStringsAppender | null;
  readonly options?: WorksheetSaveOptions;
}

export interface WorksheetSaveOptions {
  readonly maxTreeFallbackBytes?: number;
  /** Test seam: prove the tree fallback without manufacturing an unsafe part. */
  readonly forceTreeFallback?: boolean;
}

export interface WorksheetSaveResult {
  readonly bytes: Uint8Array;
  readonly mode: 'none' | 'position' | 'tree';
}

interface Edit {
  readonly start: number;
  readonly end: number;
  readonly text: string;
  readonly kind: 'replace' | 'insert';
}

class PositionPatchUnsafe extends Error {}

export function patchWorksheetXml(input: WorksheetSaveInput): WorksheetSaveResult {
  if (
    input.dirtyCells.length === 0 &&
    !hasDirtyDimensions(input.dirtyDimensions) &&
    input.syncMerges !== true &&
    input.syncAutoFilter?.changed !== true &&
    input.syncConditionalFormats?.changed !== true &&
    input.syncPane?.changed !== true
  ) {
    return { bytes: input.sourceBytes, mode: 'none' };
  }
  // A changed freeze split is applied to the source text before the cell,
  // dimension and element patches run. It is one small element edit, so the
  // position path keeps working and the rest of the part stays untouched.
  const sourceBytes = input.syncPane?.changed === true ? input.syncPane.bytes : input.sourceBytes;
  const maxFallback = input.options?.maxTreeFallbackBytes ?? MAX_TREE_PATCH_BYTES;
  if (input.options?.forceTreeFallback !== true) {
    try {
      const parsed = parseWorksheetXml(sourceBytes, {
        sharedStrings: input.sharedStringsTable,
        partPath: input.partName,
        name: input.sheet.name,
      });
      if (parsed.unsafeReasons.length === 0 && parsed.sheetData !== null) {
        const patched = patchByPosition(parsed, input);
        const withMerges =
          input.syncMerges === true
            ? patchMergeCellsXml(patched, parsed.rootPrefix, input.sheet)
            : patched;
        if (isWellFormedXml(withMerges) && preservesCellOrder(withMerges, input)) {
          return { bytes: new TextEncoder().encode(withMerges), mode: 'position' };
        }
      }
    } catch (error) {
      if (!(error instanceof PositionPatchUnsafe)) {
        // A parse/position failure falls through to the tree fallback. The
        // tree path re-validates the result.
      }
    }
  }

  if (sourceBytes.byteLength >= maxFallback) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${input.partName} could not be patched safely and is too large (${sourceBytes.byteLength} bytes) for the tree fallback. The file was not written.`,
    );
  }
  const treeBytes = patchWorksheetAsTree({
    partName: input.partName,
    sourceBytes,
    sheet: input.sheet,
    dirtyCells: input.dirtyCells,
    dirtyDimensions: input.dirtyDimensions,
    syncMerges: input.syncMerges,
    syncAutoFilter: input.syncAutoFilter,
    syncConditionalFormats: input.syncConditionalFormats,
    sharedStringsTable: input.sharedStringsTable,
    sharedStrings: input.sharedStrings,
  });
  return { bytes: treeBytes, mode: 'tree' };
}

function patchByPosition(parsed: ParsedWorksheet, input: WorksheetSaveInput): string {
  const { xml, sheetData } = parsed;
  if (sheetData === null) throw new PositionPatchUnsafe('missing sheetData');
  const edits: Edit[] = [];
  const newRows = new Map<number, DirtyCellInput[]>();
  const insertedCells = new Map<number, DirtyCellInput[]>();
  const selfClosingRows = new Map<number, { row: WorksheetRowPosition; cells: DirtyCellInput[] }>();
  const dirtyRows = input.dirtyDimensions?.rows ?? new Map<number, RowDimension | null>();
  const dirtyColumns = input.dirtyDimensions?.columns ?? new Map<number, ColumnDimension | null>();

  for (const dirty of input.dirtyCells) {
    if (dirty.row < 0 || dirty.col < 0) throw new PositionPatchUnsafe('negative cell address');
    const rowRecord = parsed.rows.get(dirty.row);
    if (rowRecord === undefined) {
      const list = newRows.get(dirty.row) ?? [];
      list.push(dirty);
      newRows.set(dirty.row, list);
      continue;
    }
    if (rowRecord.selfClosing) {
      const entry = selfClosingRows.get(dirty.row) ?? { row: rowRecord, cells: [] };
      entry.cells.push(dirty);
      selfClosingRows.set(dirty.row, entry);
      continue;
    }

    const cellRecord = rowRecord.cells.get(dirty.col);
    if (cellRecord !== undefined) {
      const replacement = cellToXml(
        cellRecord.prefix,
        formatCellRef(dirty.row, dirty.col),
        dirty.cell,
        { sharedStrings: input.sharedStrings, original: cellRecord },
      );
      edits.push({
        start: cellRecord.start,
        end: cellRecord.end,
        text: replacement ?? '',
        kind: 'replace',
      });
      continue;
    }

    if (dirty.cell === null) continue;
    const list = insertedCells.get(dirty.row) ?? [];
    list.push(dirty);
    insertedCells.set(dirty.row, list);
  }

  for (const [rowNumber, dimension] of dirtyRows) {
    if (!Number.isInteger(rowNumber) || rowNumber < 0) {
      throw new PositionPatchUnsafe('invalid dirty row index');
    }
    const rowRecord = parsed.rows.get(rowNumber);
    if (rowRecord === undefined) continue;
    if (rowRecord.selfClosing && selfClosingRows.has(rowNumber)) continue;
    const openTag = xml.slice(rowRecord.start, rowRecord.openEnd);
    const updated = applyRowDimensionToOpenTag(openTag, dimension);
    if (updated === openTag) continue;
    edits.push({
      start: rowRecord.start,
      end: rowRecord.openEnd,
      text: updated,
      kind: 'replace',
    });
  }

  for (const [rowNumber, cells] of insertedCells) {
    const rowRecord = parsed.rows.get(rowNumber)!;
    const groups = new Map<number, string[]>();
    for (const dirty of [...cells].sort((a, b) => a.col - b.col)) {
      const text = cellToXml(rowRecord.prefix, formatCellRef(dirty.row, dirty.col), dirty.cell, {
        sharedStrings: input.sharedStrings,
        original: null,
      });
      if (text === null) continue;
      const position = insertionPosition(rowRecord, dirty.col);
      const group = groups.get(position) ?? [];
      group.push(text);
      groups.set(position, group);
    }
    for (const [position, text] of groups) {
      edits.push({ start: position, end: position, text: text.join(''), kind: 'insert' });
    }
  }

  for (const [rowNumber, { row, cells }] of selfClosingRows) {
    const cellXml = cells
      .sort((a, b) => a.col - b.col)
      .map((item) =>
        cellToXml(row.prefix, formatCellRef(item.row, item.col), item.cell, {
          sharedStrings: input.sharedStrings,
          original: null,
        }),
      )
      .filter((value): value is string => value !== null)
      .join('');
    const openTag = applyRowDimensionToOpenTag(
      xml.slice(row.start, row.openEnd),
      dirtyRows.get(rowNumber),
    );
    if (cellXml === '' && cells.every((cell) => cell.cell === null)) {
      if (openTag === xml.slice(row.start, row.openEnd)) continue;
      edits.push({ start: row.start, end: row.openEnd, text: openTag, kind: 'replace' });
      continue;
    }
    const open = openTag.replace(/\/\s*>$/, '>');
    const closing = row.prefix === '' ? '</row>' : `</${row.prefix}:row>`;
    edits.push({
      start: row.start,
      end: row.end,
      text: `${open}${cellXml}${closing}`,
      kind: 'replace',
    });
  }

  const newRowNumbers = new Set<number>([
    ...newRows.keys(),
    ...[...dirtyRows.entries()]
      .filter(([rowNumber, dimension]) => dimension !== null && !parsed.rows.has(rowNumber))
      .map(([rowNumber]) => rowNumber),
  ]);

  if (newRowNumbers.size > 0) {
    const rowPrefix = sheetData.prefix || parsed.rootPrefix;
    const rowParts = [...newRowNumbers]
      .sort((a, b) => a - b)
      .map((rowNumber) => {
        const cells = newRows.get(rowNumber) ?? [];
        const cellXml = [...cells]
          .sort((a, b) => a.col - b.col)
          .map((dirty) =>
            cellToXml(rowPrefix, formatCellRef(rowNumber, dirty.col), dirty.cell, {
              sharedStrings: input.sharedStrings,
              original: null,
            }),
          )
          .filter((value): value is string => value !== null)
          .join('');
        const openTag = applyRowDimensionToOpenTag(
          `<${qualified(rowPrefix, 'row')} r="${rowNumber + 1}"/>`,
          dirtyRows.get(rowNumber),
        );
        if (cellXml === '') {
          return openTag === `<${qualified(rowPrefix, 'row')} r="${rowNumber + 1}"/>`
            ? null
            : { rowNumber, text: openTag };
        }
        const open = openTag.replace(/\/\s*>$/, '>');
        return {
          rowNumber,
          text: `${open}${cellXml}</${qualified(rowPrefix, 'row')}>`,
        };
      })
      .filter((value): value is { rowNumber: number; text: string } => value !== null);

    if (sheetData.selfClosing) {
      const open = xml.slice(sheetData.start, sheetData.openEnd).replace(/\/>$/, '>');
      const closing = `</${qualified(sheetData.prefix || parsed.rootPrefix, 'sheetData')}>`;
      edits.push({
        start: sheetData.start,
        end: sheetData.end,
        text: `${open}${rowParts.map((part) => part.text).join('')}${closing}`,
        kind: 'replace',
      });
    } else {
      for (const part of rowParts) {
        const next = [...parsed.rows.values()]
          .filter((row) => row.number > part.rowNumber)
          .sort((a, b) => a.number - b.number)[0];
        const position = next === undefined ? sheetData.contentEnd : next.start;
        edits.push({ start: position, end: position, text: part.text, kind: 'insert' });
      }
    }
  }

  edits.push(...columnDimensionEdits(parsed, dirtyColumns));
  const autoFilterEdit = autoFilterPositionEdit(parsed, input.syncAutoFilter);
  if (autoFilterEdit !== null) edits.push(autoFilterEdit);
  if (input.syncConditionalFormats?.changed === true) {
    for (const edit of input.syncConditionalFormats.edits) {
      if (edit.start < 0 || edit.end < edit.start || edit.end > parsed.xml.length) {
        throw new PositionPatchUnsafe('conditional-format edit outside the part');
      }
      edits.push({ ...edit, kind: 'replace' });
    }
  }

  return applyEdits(xml, groupInsertions(edits));
}

/**
 * Children of `<worksheet>` that follow `autoFilter` in the schema order
 * (CT_Worksheet). Inserting before the first of them keeps the part valid.
 */
function autoFilterPositionEdit(
  parsed: ParsedWorksheet,
  plan: AutoFilterSavePlan | undefined,
): Edit | null {
  if (plan === undefined || !plan.changed) return null;
  const existing = parsed.rootChildren.find((child) => child.local === 'autoFilter') ?? null;
  if (plan.xml === null) {
    return existing === null
      ? null
      : { start: existing.start, end: existing.end, text: '', kind: 'replace' };
  }
  if (existing !== null) {
    if (parsed.xml.slice(existing.start, existing.end) === plan.xml) return null;
    return { start: existing.start, end: existing.end, text: plan.xml, kind: 'replace' };
  }
  const position = autoFilterInsertionPosition(parsed);
  return { start: position, end: position, text: plan.xml, kind: 'insert' };
}

function autoFilterInsertionPosition(parsed: ParsedWorksheet): number {
  const next = parsed.rootChildren.find((child: WorksheetElementPosition) =>
    AFTER_AUTO_FILTER.has(child.local),
  );
  if (next !== undefined) return next.start;
  const closing = parsed.xml.lastIndexOf('</');
  return closing < 0 ? parsed.xml.length : closing;
}

function columnDimensionEdits(
  parsed: ParsedWorksheet,
  dirtyColumns: ReadonlyMap<number, ColumnDimension | null>,
): Edit[] {
  if (dirtyColumns.size === 0) return [];
  const indices = [...dirtyColumns.keys()]
    .filter((index) => Number.isInteger(index) && index >= 0 && index < 16_384)
    .sort((a, b) => a - b);
  if (indices.length === 0) return [];
  const { xml, columns } = parsed;
  const edits: Edit[] = [];
  const covered = new Set<number>();

  for (const column of columns) {
    const dirtyIndices = indices.filter((index) => index >= column.min && index <= column.max);
    if (dirtyIndices.length === 0) continue;
    if (!column.selfClosing) throw new PositionPatchUnsafe('non-self-closing col element');
    const openTag = xml.slice(column.start, column.openEnd);
    const parts: string[] = [];
    let cursor = column.min;
    for (const index of dirtyIndices) {
      if (index > cursor) parts.push(columnRangeTag(openTag, cursor, index - 1));
      parts.push(columnDimensionTag(openTag, index, index, dirtyColumns.get(index) ?? null));
      covered.add(index);
      cursor = index + 1;
    }
    if (cursor <= column.max) parts.push(columnRangeTag(openTag, cursor, column.max));
    edits.push({
      start: column.start,
      end: column.end,
      text: parts.join(''),
      kind: 'replace',
    });
  }

  const uncovered = indices.filter(
    (index) => !covered.has(index) && dirtyColumns.get(index) !== null,
  );
  if (uncovered.length === 0) return edits;

  const cols = parsed.cols;
  if (cols === null) {
    if (parsed.sheetData === null) throw new PositionPatchUnsafe('missing sheetData');
    const prefix = parsed.rootPrefix;
    const colXml = uncovered
      .map((index) => newColumnTag(prefix, index, dirtyColumns.get(index) ?? null))
      .join('');
    const colsName = qualified(prefix, 'cols');
    edits.push({
      start: parsed.sheetData.start,
      end: parsed.sheetData.start,
      text: `<${colsName}>${colXml}</${colsName}>`,
      kind: 'insert',
    });
    return edits;
  }

  const colPrefix = columns[0]?.prefix ?? cols.prefix ?? parsed.rootPrefix;
  if (cols.selfClosing) {
    const colXml = uncovered
      .map((index) => newColumnTag(colPrefix, index, dirtyColumns.get(index) ?? null))
      .join('');
    const open = xml.slice(cols.start, cols.openEnd).replace(/\/\s*>$/, '>');
    const colsName = qualified(cols.prefix || parsed.rootPrefix, 'cols');
    edits.push({
      start: cols.start,
      end: cols.end,
      text: `${open}${colXml}</${colsName}>`,
      kind: 'replace',
    });
    return edits;
  }

  for (const index of uncovered) {
    const next = columns.find((column) => column.min > index);
    const position = next?.start ?? columns.at(-1)?.end ?? cols.contentEnd;
    edits.push({
      start: position,
      end: position,
      text: newColumnTag(colPrefix, index, dirtyColumns.get(index) ?? null),
      kind: 'insert',
    });
  }
  return edits;
}

function columnRangeTag(openTag: string, min: number, max: number): string {
  return updateOpenTagAttributes(
    openTag,
    new Map([
      ['min', String(min + 1)],
      ['max', String(max + 1)],
    ]),
  );
}

function columnDimensionTag(
  openTag: string,
  min: number,
  max: number,
  dimension: ColumnDimension | null,
): string {
  const updates = columnDimensionAttributes(dimension);
  updates.set('min', String(min + 1));
  updates.set('max', String(max + 1));
  return updateOpenTagAttributes(openTag, updates);
}

function newColumnTag(prefix: string, index: number, dimension: ColumnDimension | null): string {
  const attributes = new Map<string, string>([
    ['min', String(index + 1)],
    ['max', String(index + 1)],
  ]);
  for (const [name, value] of columnDimensionAttributes(dimension)) {
    if (value !== null) attributes.set(name, value);
  }
  return `<${qualified(prefix, 'col')}${[...attributes]
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(value)}"`)
    .join('')}/>`;
}

function applyRowDimensionToOpenTag(
  openTag: string,
  dimension: RowDimension | null | undefined,
): string {
  if (dimension === undefined) return openTag;
  return updateOpenTagAttributes(openTag, rowDimensionAttributes(dimension));
}

function rowDimensionAttributes(dimension: RowDimension | null): Map<string, string | null> {
  return new Map([
    ['ht', dimension?.height === undefined ? null : dimensionNumber(dimension.height)],
    ['customHeight', dimension?.customHeight === true ? '1' : null],
    ['hidden', dimension?.hidden === true ? '1' : null],
    [
      's',
      dimension?.styleId === undefined || dimension.styleId === 0
        ? null
        : String(dimension.styleId),
    ],
    [
      'outlineLevel',
      dimension?.outlineLevel === undefined || dimension.outlineLevel <= 0
        ? null
        : String(dimension.outlineLevel),
    ],
  ]);
}

function columnDimensionAttributes(dimension: ColumnDimension | null): Map<string, string | null> {
  return new Map([
    ['width', dimension?.width === undefined ? null : dimensionNumber(dimension.width)],
    ['hidden', dimension?.hidden === true ? '1' : null],
    [
      's',
      dimension?.styleId === undefined || dimension.styleId === 0
        ? null
        : String(dimension.styleId),
    ],
    ['customWidth', dimension?.customWidth === true ? '1' : null],
    ['bestFit', dimension?.bestFit === true ? '1' : null],
    [
      'outlineLevel',
      dimension?.outlineLevel === undefined || dimension.outlineLevel <= 0
        ? null
        : String(dimension.outlineLevel),
    ],
  ]);
}

function dimensionNumber(value: number): string {
  return Number.isFinite(value) ? String(value) : '0';
}

/**
 * Rewrite modelled attributes inside an open tag without touching any
 * unmodelled attribute, its order, the element prefix or the trailing slash.
 */
function updateOpenTagAttributes(
  openTag: string,
  updates: ReadonlyMap<string, string | null>,
): string {
  const pattern = /\s+([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  const applied = new Set<string>();
  let output = '';
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(openTag)) !== null) {
    const name = match[1]!;
    const local = tagLocalName(name);
    if (!updates.has(local)) continue;
    output += openTag.slice(lastIndex, match.index);
    const value = updates.get(local);
    if (value !== null && value !== undefined) {
      output += ` ${name}="${escapeXmlAttribute(value)}"`;
    }
    lastIndex = pattern.lastIndex;
    applied.add(local);
  }
  output += openTag.slice(lastIndex);
  const missing = [...updates.entries()].filter(
    ([name, value]) => value !== null && !applied.has(name),
  );
  if (missing.length === 0) return output;
  const close = output.lastIndexOf('>');
  if (close < 0) return output;
  const insertAt = output[close - 1] === '/' ? close - 1 : close;
  const added = missing.map(([name, value]) => ` ${name}="${escapeXmlAttribute(value!)}"`).join('');
  return output.slice(0, insertAt) + added + output.slice(insertAt);
}

function tagLocalName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

function hasDirtyDimensions(dimensions: DirtyDimensions | undefined): boolean {
  return dimensions !== undefined && (dimensions.rows.size > 0 || dimensions.columns.size > 0);
}

function preservesCellOrder(xml: string, input: WorksheetSaveInput): boolean {
  try {
    const reparsed = parseWorksheetXml(new TextEncoder().encode(xml), {
      sharedStrings: input.sharedStringsTable,
      partPath: input.partName,
      name: input.sheet.name,
    });
    return reparsed.unsafeReasons.length === 0;
  } catch {
    return false;
  }
}

/**
 * Replace the worksheet's `<mergeCells>` element with the model's merge list.
 * The paste is the only authoring path for merges in this milestone, and it
 * must survive a save exactly like a cell or a dimension.
 */
function patchMergeCellsXml(xml: string, prefix: string, sheet: Sheet): string {
  const mergeCellsTag = qualified(prefix, 'mergeCells');
  const mergeCellTag = qualified(prefix, 'mergeCell');
  const elementPattern = new RegExp(
    `<${escapeRegExp(mergeCellsTag)}\\b[^>]*?(?:/>|>[\\s\\S]*?</${escapeRegExp(mergeCellsTag)}\\s*>)`,
  );
  const replacement =
    sheet.merges.length === 0
      ? ''
      : `<${mergeCellsTag} count="${sheet.merges.length}">${sheet.merges
          .map((merge) => `<${mergeCellTag} ref="${formatRange(merge.range)}"/>`)
          .join('')}</${mergeCellsTag}>`;
  const existing = elementPattern.exec(xml);
  if (existing !== null) {
    return (
      xml.slice(0, existing.index) + replacement + xml.slice(existing.index + existing[0].length)
    );
  }
  if (replacement === '') return xml;
  const sheetDataTag = qualified(prefix, 'sheetData');
  const sheetDataPattern = new RegExp(
    `<${escapeRegExp(sheetDataTag)}\\b[^>]*?(?:/>|>[\\s\\S]*?</${escapeRegExp(sheetDataTag)}\\s*>)`,
  );
  const sheetData = sheetDataPattern.exec(xml);
  if (sheetData === null) return xml;
  const position = sheetData.index + sheetData[0].length;
  return xml.slice(0, position) + replacement + xml.slice(position);
}

function formatRange(range: Sheet['merges'][number]['range']): string {
  const start = formatCellRef(range.startRow, range.startCol);
  const end = formatCellRef(range.endRow, range.endCol);
  return `${start}:${end}`;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function insertionPosition(row: WorksheetRowPosition, col: number): number {
  let position = row.openEnd;
  if (row.cells.size > 0) {
    const sorted = [...row.cells.values()].sort((a, b) => a.col - b.col);
    const next = sorted.find((cell) => cell.col > col);
    position = next?.start ?? sorted.at(-1)!.end;
  }
  return position;
}

function groupInsertions(edits: Edit[]): Edit[] {
  const grouped = new Map<string, Edit[]>();
  for (const edit of edits) {
    const key = `${edit.start}:${edit.end}:${edit.kind}`;
    const list = grouped.get(key) ?? [];
    list.push(edit);
    grouped.set(key, list);
  }
  return [...grouped.values()].map((list) => {
    if (list.length === 1) return list[0]!;
    if (list[0]!.kind === 'insert') {
      return {
        ...list[0]!,
        text: list.map((edit) => edit.text).join(''),
      };
    }
    return list[0]!;
  });
}

function applyEdits(xml: string, edits: Edit[]): string {
  const sorted = [...edits].sort((a, b) => {
    if (b.start !== a.start) return b.start - a.start;
    if (a.kind !== b.kind) return a.kind === 'replace' ? -1 : 1;
    return 0;
  });
  let output = xml;
  for (const edit of sorted) {
    output = output.slice(0, edit.start) + edit.text + output.slice(edit.end);
  }
  return output;
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}
