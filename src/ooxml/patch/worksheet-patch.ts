import type { Cell, Sheet } from '@model/types';
import { formatCellRef } from '@model/types';
import type { SharedStringsTable } from '../types';
import type { SharedStringsAppender } from './shared-strings-patch';
import { cellToXml } from './cell-xml';
import type { ParsedWorksheet, WorksheetRowPosition } from '../worksheet-xml';
import { parseWorksheetXml } from '../worksheet-xml';
import { WorkbookSaveError } from '../errors';
import { isWellFormedXml } from '../xml';
import { patchWorksheetAsTree } from './tree-patch';
import { MAX_TREE_PATCH_BYTES } from '../xml';

export interface DirtyCellInput {
  readonly row: number;
  readonly col: number;
  readonly cell: Cell | null;
}

export interface WorksheetSaveInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly sheet: Sheet;
  readonly dirtyCells: readonly DirtyCellInput[];
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
  if (input.dirtyCells.length === 0) {
    return { bytes: input.sourceBytes, mode: 'none' };
  }
  const maxFallback = input.options?.maxTreeFallbackBytes ?? MAX_TREE_PATCH_BYTES;
  if (input.options?.forceTreeFallback !== true) {
    try {
      const parsed = parseWorksheetXml(input.sourceBytes, {
        sharedStrings: input.sharedStringsTable,
        partPath: input.partName,
        name: input.sheet.name,
      });
      if (parsed.unsafeReasons.length === 0 && parsed.sheetData !== null) {
        const patched = patchByPosition(parsed, input);
        if (isWellFormedXml(patched) && preservesCellOrder(patched, input)) {
          return { bytes: new TextEncoder().encode(patched), mode: 'position' };
        }
      }
    } catch (error) {
      if (!(error instanceof PositionPatchUnsafe)) {
        // A parse/position failure falls through to the tree fallback. The
        // tree path re-validates the result.
      }
    }
  }

  if (input.sourceBytes.byteLength >= maxFallback) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${input.partName} could not be patched safely and is too large (${input.sourceBytes.byteLength} bytes) for the tree fallback. The file was not written.`,
    );
  }
  const treeBytes = patchWorksheetAsTree({
    partName: input.partName,
    sourceBytes: input.sourceBytes,
    sheet: input.sheet,
    dirtyCells: input.dirtyCells,
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

  for (const { row, cells } of selfClosingRows.values()) {
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
    if (cellXml === '' && cells.every((cell) => cell.cell === null)) continue;
    const open = xml.slice(row.start, row.openEnd).replace(/\/>$/, '>');
    const closing = row.prefix === '' ? '</row>' : `</${row.prefix}:row>`;
    edits.push({
      start: row.start,
      end: row.end,
      text: `${open}${cellXml}${closing}`,
      kind: 'replace',
    });
  }

  if (newRows.size > 0) {
    const rowParts = [...newRows.entries()]
      .map(([rowNumber, cells]) => {
        const rowPrefix = sheetData.prefix || parsed.rootPrefix;
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
        if (cellXml === '') return null;
        return {
          rowNumber,
          text: `<${qualified(rowPrefix, 'row')} r="${rowNumber + 1}">${cellXml}</${qualified(
            rowPrefix,
            'row',
          )}>`,
        };
      })
      .filter((value): value is { rowNumber: number; text: string } => value !== null)
      .sort((a, b) => a.rowNumber - b.rowNumber);

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

  return applyEdits(xml, groupInsertions(edits));
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
