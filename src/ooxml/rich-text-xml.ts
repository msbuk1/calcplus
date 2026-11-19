import type { Sheet } from '@model/types';
import { cellKey, cellStateEquals } from '@model/types';
import type { DirtyCellInput } from './patch/worksheet-patch';
import { parseWorksheetXml } from './worksheet-xml';
import type { SharedStringsTable } from './types';

/**
 * Rich-text save policy (file-compat §9 F17, workbook-model §9).
 *
 * The model cannot write rich runs back; `cellToXml` serialises plain text. A
 * dirty cell is therefore only writable when it either was plain before or was
 * deliberately flattened by a command (which warns the user). Two cases need
 * special care:
 *
 * - **Undo.** Undoing a flatten restores the exact prior cell, rich payload
 *   included, but the dirty marker survives. Rewriting that cell would drop the
 *   runs silently, so a dirty cell whose model state equals the source cell is
 *   dropped from the save altogether: nothing changed, so nothing is written
 *   and the part's bytes stay as authored.
 * - **Anything else.** A cell that still carries its runs (a style-only edit),
 *   or a fill/paste that overwrote them without asking, has no
 *   `richTextFlattened` marker, so the save refuses it
 *   (`dirtyCellsChangeFormula` in `save.ts`) instead of flattening silently.
 *   `cell.set` (warned) and Clear contents (explicit) set the marker; a deleted
 *   cell needs none.
 *
 * The comparison is `cellStateEquals`, which ignores derived `stale` state and
 * compares the authored value, formula, formula storage, style and rich
 * payload.
 */
export interface UnchangedCellFilterInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly sheetName: string;
  readonly sharedStrings: SharedStringsTable | null;
  readonly sheet: Sheet;
  readonly dirtyCells: readonly DirtyCellInput[];
}

/** Drop dirty markers whose model cell already matches the source part. */
export function dropUnchangedCells(input: UnchangedCellFilterInput): DirtyCellInput[] {
  if (input.dirtyCells.length === 0) return [];
  let sourceCells: Sheet['cells'];
  try {
    sourceCells = parseWorksheetXml(input.sourceBytes, {
      sharedStrings: input.sharedStrings,
      partPath: input.partName,
      name: input.sheetName,
    }).sheet.cells;
  } catch {
    // A part the open path cannot read is left for the patch writer to report.
    return [...input.dirtyCells];
  }
  return input.dirtyCells.filter((dirty) => {
    const original = sourceCells.get(cellKey(dirty.row, dirty.col)) ?? null;
    return !cellStateEquals(original, dirty.cell);
  });
}
