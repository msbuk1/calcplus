import type { ClipboardSnapshot } from '@model/clipboard';
import type { FindMatch } from '@model/find';
import type { StructuralEdit } from '@model/range-transform';
import type { SortBlockCode, SortStopReason, SortWarning } from '@model/sort';
import type {
  AutoFilter,
  Cell,
  CellAddress,
  ColumnDimension,
  CommentRecord,
  ConditionalFormatRecord,
  GridRange,
  HyperlinkRecord,
  RowDimension,
  Sheet,
  SheetId,
  StyleId,
  StyleRecord,
  ValidationRecord,
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

/** One row whose effective hidden flag changes when the AutoFilter changes. */
export interface FilterHiddenRowChange {
  readonly row: number;
  readonly before: boolean;
  readonly after: boolean;
}

/**
 * What an AutoFilter command reports back to its surface (sort-autofilter
 * §3.8-§3.9). `changed` is `false` for a no-op, which creates no transaction.
 * The criteria themselves live in the transaction's `filter.set` change.
 */
export interface FilterPayload {
  readonly state: 'applied' | 'changed' | 'cleared' | 'reapplied' | 'none';
  /** `false` when the command was a no-op and created no transaction (§3.9). */
  readonly changed: boolean;
  readonly sheetId: SheetId;
  /** The filter range after the change, or the prior one when the filter is removed. */
  readonly range: GridRange | null;
  readonly before: AutoFilter | null;
  readonly after: AutoFilter | null;
  /** Effective hidden-row changes, the same record the transaction stores (§3.9). */
  readonly hiddenChanges: readonly FilterHiddenRowChange[];
  /** `X of Y records` for the status bar (§3.8), computed after the change. */
  readonly rowsVisible: number;
  readonly rowsTotal: number;
  /** Columns whose imported criteria calcplus cannot edit (§3.7). */
  readonly unsupportedColumns: readonly number[];
}

/**
 * What `data.sort` reports back to its surface (sort-autofilter §2.1-§2.7):
 * the dialog prefill, the applied plan, or Excel's block message. The two
 * warnings — hidden rows crossed once (§2.6) and a sub-range sort (§2.1) —
 * travel here as data; they never block, and the UI owns how to show them.
 */
export interface SortPayload {
  readonly state: 'sorted' | 'unchanged' | 'blocked' | 'needs-keys';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  /** The guessed or overridden header state; absent when the sort was blocked. */
  readonly hasHeaders?: boolean;
  /** The labels the Sort dialog shows, in key order. */
  readonly keyLabels?: readonly string[];
  /** Why a sort was a no-op: no keys, too few rows, nothing to order (§2.1). */
  readonly reason?: SortStopReason;
  /** The blocking code and Excel's exact message (§2.5). */
  readonly code?: SortBlockCode;
  readonly message?: string;
  readonly warnings: readonly SortWarning[];
  readonly visibleRowCount?: number;
  readonly hiddenRowCount?: number;
}

/**
 * What a `structure.*` command reports back to its surface (feature matrix
 * "Insert/delete rows, columns and cells"). A blocked edit creates no
 * transaction; an applied edit is one transaction whose shape is
 * `StructureApplyChange`.
 */
export interface StructurePayload {
  readonly state: 'applied' | 'blocked' | 'none';
  readonly sheetId: SheetId;
  readonly edit: StructuralEdit | null;
  readonly message?: string;
  /** Cells moved, removed or with a rewritten formula. */
  readonly cellCount: number;
  readonly formulaRewrites: number;
  /** Row and column dimension records that moved or renumbered. */
  readonly dimensionCount: number;
  /** `true` when deleting the header row cleared the AutoFilter (§3.5). */
  readonly filterCleared: boolean;
}

export interface CommandResult {
  readonly changes: readonly CommandChange[];
  readonly selectionAfter?: GridRange;
  /** The structural-edit outcome: the applied edit or Excel's block message. */
  readonly structure?: StructurePayload; /**
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
  /** The matches `find.search` found, in scan order (document order unless direction is Up). */
  readonly matches?: readonly FindMatch[];
  /** Cells changed by `find.replaceOne` or `find.replaceAll`. */
  readonly replaced?: number;
  /** The dialog payload for the find and replace commands. */
  readonly find?: FindResultPayload;
  /** The sort outcome: dialog prefill, the applied plan, or Excel's block message. */
  readonly sort?: SortPayload;
  /** The AutoFilter outcome: criteria, hidden-row changes and the record count. */
  readonly filter?: FilterPayload;
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

/** A formula the moved row rewrote, kept in the log for replay (§2.7). */
export interface SortFormulaRewritePatch {
  readonly address: CellAddress;
  readonly before: string;
  readonly after: string;
}

/** One keyed comment or hyperlink entry, keyed `row:col` like `Sheet.comments`. */
export type SheetAttachmentEntry<T> = readonly [string, T];

/** The sheet's whole comment or hyperlink map on both sides of a sort. */
export interface SortAttachmentChange<T> {
  readonly before: readonly SheetAttachmentEntry<T>[];
  readonly after: readonly SheetAttachmentEntry<T>[];
}

/**
 * A sort as one transaction (sort-autofilter §2.7). The moved cells travel as
 * `cell.set` patches in the same transaction, so the calculation layer sees
 * them; this record stores what only a sort changes: the row permutation, the
 * formula rewrites for the log, the merged ranges and the comments and
 * hyperlinks that move with their rows.
 *
 * Comments and hyperlinks are applied the way `src/model/sort.ts` documents:
 * every attachment in a moved row (inside the sorted columns) is removed first
 * and the moved records are inserted afterwards, because an incoming cell can
 * leave its destination attachment empty.
 */
export interface SortApplyChange {
  readonly type: 'sort.apply';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  /** Destination offset inside the range -> source row; the header maps to itself. */
  readonly rowPermutation: readonly number[];
  readonly formulaRewrites: readonly SortFormulaRewritePatch[];
  /** The sheet's whole merged-range list on both sides of the sort. */
  readonly merges: { readonly before: readonly GridRange[]; readonly after: readonly GridRange[] };
  readonly comments: SortAttachmentChange<CommentRecord>;
  readonly hyperlinks: SortAttachmentChange<HyperlinkRecord>;
}

/**
 * One AutoFilter transaction (sort-autofilter §3.9): the criteria before and
 * after. Filter-hidden rows are derived from the criteria (§3.3), so applying
 * the change restores both visibility layers together; `hiddenChanges` records
 * them for the log and the status bar and is never a second place the hidden
 * flags live.
 */
export interface FilterSetChange {
  readonly type: 'filter.set';
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly before: AutoFilter | null;
  readonly after: AutoFilter | null;
  readonly hiddenChanges: readonly FilterHiddenRowChange[];
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

/**
 * One cell moved, removed or rewritten by a structural edit. `from` is `null`
 * only when a caller materialises a cell, `to` is `null` when the edit deletes
 * it. `before` and `after` are the exact cell states at each address.
 */
export interface StructureCellRecord {
  readonly sheetId: SheetId;
  readonly from: CellAddress | null;
  readonly to: CellAddress | null;
  readonly before: Cell | null;
  readonly after: Cell | null;
}

export interface StructureRowsChange {
  readonly before: readonly (readonly [number, RowDimension])[];
  readonly after: readonly (readonly [number, RowDimension])[];
}

export interface StructureColumnsChange {
  readonly before: readonly (readonly [number, ColumnDimension])[];
  readonly after: readonly (readonly [number, ColumnDimension])[];
}

export interface StructureListsChange<T> {
  readonly before: readonly T[];
  readonly after: readonly T[];
}

export interface StructureFilterChange {
  readonly before: AutoFilter | null;
  readonly after: AutoFilter | null;
}

/**
 * One structural edit as one transaction (feature matrix §1; command-system
 * §2 and §4; sort-autofilter §3.5; CF §7).
 *
 * The change carries every consequence of the edit: the moved, deleted and
 * formula-rewritten cells across every sheet, the renumbered row and column
 * dimensions, merged ranges, comments and hyperlinks, CF and validation
 * records, the AutoFilter, and the renumbered workbook-part baseline the next
 * save patches against. Undo restores every side exactly, so no reference
 * rewrite is ever recomputed.
 */
export interface StructureApplyChange {
  readonly type: 'structure.apply';
  readonly sheetId: SheetId;
  readonly partPath: string;
  readonly edit: StructuralEdit;
  readonly cells: readonly StructureCellRecord[];
  readonly rows: StructureRowsChange;
  readonly columns: StructureColumnsChange;
  readonly merges: StructureListsChange<GridRange>;
  readonly comments: SortAttachmentChange<CommentRecord>;
  readonly hyperlinks: SortAttachmentChange<HyperlinkRecord>;
  readonly validation: StructureListsChange<ValidationRecord>;
  readonly conditionalFormats: StructureListsChange<ConditionalFormatRecord>;
  readonly autoFilter: StructureFilterChange;
  /** The source part bytes before and after renumbering, or `null` when absent. */
  readonly sourceBefore: Uint8Array | null;
  readonly sourceAfter: Uint8Array | null;
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
  | NameSetChange
  | SortApplyChange
  | FilterSetChange
  | StructureApplyChange;

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
    case 'sort.apply':
      // The permutation and the formula rewrites describe the operation, not a
      // direction; redo reapplies them, and the attachments and merges swap.
      return {
        ...change,
        merges: { before: change.merges.after, after: change.merges.before },
        comments: { before: change.comments.after, after: change.comments.before },
        hyperlinks: { before: change.hyperlinks.after, after: change.hyperlinks.before },
      };
    case 'filter.set':
      return {
        ...change,
        before: change.after,
        after: change.before,
        hiddenChanges: change.hiddenChanges.map((row) => ({
          row: row.row,
          before: row.after,
          after: row.before,
        })),
      };
    case 'structure.apply':
      return {
        ...change,
        cells: change.cells.map((cell) => ({
          sheetId: cell.sheetId,
          from: cell.to,
          to: cell.from,
          before: cell.after,
          after: cell.before,
        })),
        rows: { before: change.rows.after, after: change.rows.before },
        columns: { before: change.columns.after, after: change.columns.before },
        merges: { before: change.merges.after, after: change.merges.before },
        comments: { before: change.comments.after, after: change.comments.before },
        hyperlinks: { before: change.hyperlinks.after, after: change.hyperlinks.before },
        validation: { before: change.validation.after, after: change.validation.before },
        conditionalFormats: {
          before: change.conditionalFormats.after,
          after: change.conditionalFormats.before,
        },
        autoFilter: { before: change.autoFilter.after, after: change.autoFilter.before },
        sourceBefore: change.sourceAfter,
        sourceAfter: change.sourceBefore,
      };
  }
}
