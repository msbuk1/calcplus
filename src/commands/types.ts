import type { ClipboardSnapshot } from '@model/clipboard';
import type { FindMatch } from '@model/find';
import type {
  Cell,
  CellAddress,
  ColumnDimension,
  GridRange,
  RowDimension,
  Sheet,
  SheetId,
  StyleId,
  StyleRecord,
  Workbook,
} from '@model/types';

/** The live view flags a view or toggle command needs to report its state. */
export interface CommandViewState {
  readonly showGridLines: boolean;
  readonly showHeadings: boolean;
  readonly zoom: number;
  readonly freezeRows: number;
  readonly freezeCols: number;
}

export interface CommandContext {
  readonly workbook: Workbook;
  readonly activeSheetId: SheetId | null;
  readonly selection: GridRange | null;
  readonly activeCell: CellAddress | null;
  /** Transient view state, never document state. */
  readonly view?: CommandViewState;
  /** The file layer is writing or picking a file. */
  readonly busy?: boolean;
  readonly canUndo?: boolean;
  readonly canRedo?: boolean;
}

export type CommandCategory = 'edit' | 'file' | 'view' | 'data' | 'format' | 'formula' | 'help';

export interface CommandUiRequest {
  readonly type: 'toast' | 'dialog' | 'focus';
  readonly message?: string;
}

/**
 * The payload `find.search`, `find.replaceOne` and `find.replaceAll` report to
 * the Find & Replace dialog. `matches` is present for a search; `replaced`
 * counts cells changed by a replace command.
 */
export interface FindResultPayload {
  readonly matches?: readonly FindMatch[];
  readonly total?: number;
  readonly replaced?: number;
}

export interface CommandResult {
  readonly changes: readonly CommandChange[];
  readonly selectionAfter?: GridRange;
  /**
   * A view-state hint, never a document change: the sheet the caller should
   * activate after this command or undo. Undo of an insert names the fallback
   * tab; a sheet delete is terminal, so it names the neighbouring tab.
   */
  readonly activeSheetId?: SheetId | null;
  /**
   * The in-app clipboard snapshot captured by `edit.copy` or `edit.cut`. It is
   * view state at the clipboard edge, not a document change, so it never
   * enters the undo log.
   */
  readonly clipboard?: ClipboardSnapshot;
  /**
   * The style snapshot `painter.pickup` captured. Like the clipboard, it is
   * view state and never enters the undo log.
   */
  readonly painter?: PainterSnapshot;
  /** The matches `find.search` found, in document order. */
  readonly matches?: readonly FindMatch[];
  /** Cells changed by `find.replaceOne` or `find.replaceAll`. */
  readonly replaced?: number;
  /** The dialog payload for the find and replace commands. */
  readonly find?: FindResultPayload;
  readonly ui?: CommandUiRequest;
}

/** One range-based cell patch. `null` means the cell did not exist. */
export interface CellSetChange {
  readonly type: 'cell.set';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly before: readonly (readonly (Cell | null)[])[];
  readonly after: readonly (readonly (Cell | null)[])[];
}

/**
 * One range-based formatting patch. `null` means there was no cell record at
 * all, so undo of a format on an empty cell removes the record again.
 */
export interface StyleSetChange {
  readonly type: 'style.set';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly before: readonly (readonly (StyleId | null)[])[];
  readonly after: readonly (readonly (StyleId | null)[])[];
}

/**
 * A Format Painter pickup: the full style record at the anchor cell of the
 * source range. It is view state, not a document change, so it never enters the
 * undo log. Esc cancels the painter by clearing this snapshot in the UI; there
 * is no cancel command.
 */
export interface PainterSnapshot {
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly styleId: StyleId;
  readonly style: StyleRecord;
}

/** The workbook-level date-system flag. Undo restores the flag exactly. */
export interface WorkbookDateSystemChange {
  readonly type: 'workbook.dateSystem';
  readonly before: boolean;
  readonly after: boolean;
}

/**
 * A merged-range patch. Merge authoring has no per-cell before/after shape;
 * the sheet's merge list is the document state, so the change carries the
 * exact list on both sides. Undo restores the previous list in one step.
 */
export interface MergeSetChange {
  readonly type: 'merge.set';
  readonly sheetId: SheetId;
  readonly before: readonly GridRange[];
  readonly after: readonly GridRange[];
}

/** A row or column dimension snapshot; `null` means the record did not exist. */
export type DimensionRecord = Readonly<RowDimension & ColumnDimension>;

/** One row-height or column-width change, including hidden and best-fit flags. */
export interface DimensionSetChange {
  readonly type: 'dimension.set';
  readonly sheetId: SheetId;
  readonly axis: 'row' | 'column';
  readonly index: number;
  readonly before: DimensionRecord | null;
  readonly after: DimensionRecord | null;
}

/**
 * A sheet inserted into the tab order. The change carries the detached sheet
 * so undo restores the same stable id and the same tab position. The sheet is
 * owned by the change while it is outside the workbook; no other command can
 * mutate it there.
 */
export interface SheetAddChange {
  readonly type: 'sheet.add';
  readonly index: number;
  readonly before: null;
  readonly after: Sheet;
}

/**
 * A sheet removed from the tab order. The change applies the removal; the
 * delete command is terminal, so no inverse is recorded for it.
 */
export interface SheetDeleteChange {
  readonly type: 'sheet.delete';
  readonly index: number;
  readonly before: Sheet;
  readonly after: null;
}

/** A rename is a label change only; the stable sheet id does not move. */
export interface SheetRenameChange {
  readonly type: 'sheet.rename';
  readonly sheetId: SheetId;
  readonly before: string;
  readonly after: string;
}

/** A defined-name formula change, in workbook or sheet scope. */
export interface NameSetChange {
  readonly type: 'name.set';
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId: SheetId | null;
  readonly name: string;
  readonly before: string;
  readonly after: string;
}

export type CommandChange =
  | CellSetChange
  | StyleSetChange
  | WorkbookDateSystemChange
  | DimensionSetChange
  | MergeSetChange
  | SheetAddChange
  | SheetDeleteChange
  | SheetRenameChange
  | NameSetChange;

/**
 * The calculation layer observes committed changes. It never participates in
 * the command transaction; derived results are applied separately.
 */
export interface CalculationScheduler {
  afterChanges(changes: readonly CommandChange[]): void;
}

export interface CommandDefinition<Args = unknown> {
  readonly id: string;
  readonly label: string;
  readonly category: CommandCategory;
  readonly icon?: string;
  readonly description?: string;
  readonly keybindings: readonly string[];
  /**
   * `terminal` applies the changes but records no undo entry and clears the
   * undo and redo stacks. Sheet delete uses it: Excel makes a delete permanent
   * and offers save-as, not undo, as the recourse. The default is `record`.
   */
  readonly history?: 'record' | 'terminal';
  isEnabled(context: CommandContext): boolean;
  isChecked?(context: CommandContext): boolean;
  run(context: CommandContext, args: Args): CommandResult;
}

export function changeCells(change: CommandChange): CellAddress[] {
  if (change.type !== 'cell.set' && change.type !== 'style.set') return [];
  const cells: CellAddress[] = [];
  for (let row = change.range.startRow; row <= change.range.endRow; row += 1) {
    for (let col = change.range.startCol; col <= change.range.endCol; col += 1) {
      cells.push({ row, col });
    }
  }
  return cells;
}

export function inverseChange(change: CommandChange): CommandChange {
  switch (change.type) {
    case 'cell.set':
      return { ...change, before: change.after, after: change.before };
    case 'style.set':
      return { ...change, before: change.after, after: change.before };
    case 'workbook.dateSystem':
      return { ...change, before: change.after, after: change.before };
    case 'dimension.set':
      return { ...change, before: change.after, after: change.before };
    case 'merge.set':
      return { ...change, before: change.after, after: change.before };
    case 'sheet.add':
      return { type: 'sheet.delete', index: change.index, before: change.after, after: null };
    case 'sheet.delete':
      return { type: 'sheet.add', index: change.index, before: null, after: change.before };
    case 'sheet.rename':
      return { ...change, before: change.after, after: change.before };
    case 'name.set':
      return { ...change, before: change.after, after: change.before };
  }
}
