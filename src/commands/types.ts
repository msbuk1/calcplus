import type {
  Cell,
  CellAddress,
  ColumnDimension,
  GridRange,
  RowDimension,
  SheetId,
  StyleId,
  Workbook,
} from '@model/types';

export interface CommandContext {
  readonly workbook: Workbook;
  readonly activeSheetId: SheetId | null;
  readonly selection: GridRange | null;
  readonly activeCell: CellAddress | null;
}

export type CommandCategory = 'edit' | 'file' | 'view' | 'data' | 'format';

export interface CommandUiRequest {
  readonly type: 'toast' | 'dialog' | 'focus';
  readonly message?: string;
}

export interface CommandResult {
  readonly changes: readonly CommandChange[];
  readonly selectionAfter?: GridRange;
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

/** The workbook-level date-system flag. Undo restores the flag exactly. */
export interface WorkbookDateSystemChange {
  readonly type: 'workbook.dateSystem';
  readonly before: boolean;
  readonly after: boolean;
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

export type CommandChange =
  CellSetChange | StyleSetChange | WorkbookDateSystemChange | DimensionSetChange;

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
  isEnabled(context: CommandContext): boolean;
  isChecked?(context: CommandContext): boolean;
  run(context: CommandContext, args: Args): CommandResult;
}

export function changeCells(change: CommandChange): CellAddress[] {
  if (change.type === 'workbook.dateSystem' || change.type === 'dimension.set') return [];
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
  }
}
