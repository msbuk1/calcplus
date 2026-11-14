import type { Cell, CellAddress, GridRange, SheetId, Workbook } from '@model/types';

export interface CommandContext {
  readonly workbook: Workbook;
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

export interface CellSetChange {
  readonly type: 'cell.set';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly before: readonly (readonly (Cell | null)[])[];
  readonly after: readonly (readonly (Cell | null)[])[];
}

export type CommandChange = CellSetChange;

export interface CommandDefinition<Args = unknown> {
  readonly id: string;
  readonly label: string;
  readonly category: CommandCategory;
  readonly keybindings: readonly string[];
  isEnabled(context: CommandContext): boolean;
  isChecked?(context: CommandContext): boolean;
  run(context: CommandContext, args: Args): CommandResult;
}

export function changeCells(change: CommandChange): CellAddress[] {
  const cells: CellAddress[] = [];
  for (let row = change.range.startRow; row <= change.range.endRow; row += 1) {
    for (let col = change.range.startCol; col <= change.range.endCol; col += 1) {
      cells.push({ row, col });
    }
  }
  return cells;
}

export function inverseChange(change: CommandChange): CommandChange {
  return {
    ...change,
    before: change.after,
    after: change.before,
  };
}
