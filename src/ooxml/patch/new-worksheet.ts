import type { Cell, Sheet } from '@model/types';
import { columnToLetters, formatCellRef, parseCellKey, type RowDimension } from '@model/types';
import { filterHiddenRows, serialiseAutoFilterElement } from '../autofilter-xml';
import { WorkbookSaveError } from '../errors';
import { RELATIONSHIPS_NS, SPREADSHEET_NS, XML_DECLARATION_UTF8, escapeXmlAttribute } from '../xml';
import { cellToXml } from './cell-xml';
import type { SharedStringsAppender } from './shared-strings-patch';

/**
 * A full worksheet part for a sheet that has no source skeleton: a sheet
 * inserted by the user. The writer emits only modelled constructs, which is
 * all a calcplus-created sheet has. A read-only formula group or rich-text cell
 * has no source XML to preserve, so writing one would silently change its
 * meaning and is refused instead.
 */
export function createWorksheetBytes(
  sheet: Sheet,
  sharedStrings: SharedStringsAppender | null,
): Uint8Array {
  const cellsByRow = new Map<number, Map<number, Cell>>();
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = -1;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = -1;

  for (const [key, cell] of sheet.cells) {
    if (cell.formulaSource !== undefined) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        `The sheet ${sheet.name} holds a shared or array formula group that cannot be written into a new worksheet part. The file was not written.`,
      );
    }
    if (cell.richText !== null) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        `The sheet ${sheet.name} holds rich text that cannot be written into a new worksheet part without losing its runs. The file was not written.`,
      );
    }
    const { row, col } = parseCellKey(key);
    const rowCells = cellsByRow.get(row) ?? new Map<number, Cell>();
    rowCells.set(col, cell);
    cellsByRow.set(row, rowCells);
    if (row < minRow) minRow = row;
    if (row > maxRow) maxRow = row;
    if (col < minCol) minCol = col;
    if (col > maxCol) maxCol = col;
  }

  const dimension =
    maxRow < 0 || maxCol < 0
      ? 'A1'
      : `${columnToLetters(minCol)}${minRow + 1}:${columnToLetters(maxCol)}${maxRow + 1}`;
  const defaultRowHeight = sheet.defaultRowHeight ?? 15;
  const formatPr = [`defaultRowHeight="${number(defaultRowHeight)}"`];
  if (sheet.defaultColumnWidth !== null) {
    formatPr.push(`defaultColWidth="${number(sheet.defaultColumnWidth)}"`);
  }

  const rowNumbers = new Set<number>([...cellsByRow.keys(), ...sheet.rowDimensions.keys()]);
  // §3.3: the saved flag is the effective one, so a row the sheet's filter
  // hides is written hidden even when the model only holds the criteria.
  const filterHidden = filterHiddenRows(sheet);
  const rows = [...rowNumbers]
    .sort((left, right) => left - right)
    .map((row) =>
      rowXml(
        row,
        cellsByRow.get(row) ?? new Map<number, Cell>(),
        sheet.rowDimensions.get(row),
        sharedStrings,
        filterHidden.has(row),
      ),
    )
    .filter((xml) => xml !== '');

  const sheetView = sheetViewXml(sheet);
  const cols = colsXml(sheet);
  const merges = mergeXml(sheet);
  const protection = sheet.protection === null ? '' : sheet.protection.rawXml;
  const autoFilter = sheet.autoFilter === null ? '' : serialiseAutoFilterElement(sheet.autoFilter);

  const xml = `${XML_DECLARATION_UTF8}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIPS_NS}">
  <dimension ref="${dimension}"/>
  ${sheetView}
  <sheetFormatPr ${formatPr.join(' ')}/>
  ${cols}
  <sheetData>${rows.join('')}</sheetData>
  ${protection}
  ${autoFilter}
  ${merges}
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
  return new TextEncoder().encode(xml);
}

function sheetViewXml(sheet: Sheet): string {
  const attributes = ['workbookViewId="0"'];
  if (!sheet.view.showGridLines) attributes.push('showGridLines="0"');
  if (!sheet.view.showHeadings) attributes.push('showHeadings="0"');
  const zoom = Math.min(400, Math.max(10, Math.round(sheet.view.zoom)));
  if (zoom !== 100) attributes.push(`zoomScale="${zoom}"`);
  const pane = paneXml(sheet);
  if (pane === '') return `<sheetViews><sheetView ${attributes.join(' ')}/></sheetViews>`;
  return `<sheetViews><sheetView ${attributes.join(' ')}>${pane}</sheetView></sheetViews>`;
}

function paneXml(sheet: Sheet): string {
  const rows = Math.max(0, Math.floor(sheet.view.freezeRows));
  const cols = Math.max(0, Math.floor(sheet.view.freezeCols));
  if (rows === 0 && cols === 0) return '';
  const topLeftCell = `${columnToLetters(cols)}${rows + 1}`;
  const xSplit = cols > 0 ? ` xSplit="${cols}"` : '';
  const ySplit = rows > 0 ? ` ySplit="${rows}"` : '';
  const activePane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
  return `<pane${xSplit}${ySplit} topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`;
}

function colsXml(sheet: Sheet): string {
  if (sheet.columnDimensions.size === 0) return '';
  const entries = [...sheet.columnDimensions.entries()].sort(([left], [right]) => left - right);
  const columns = entries.map(([index, dimension]) => {
    const attributes = [`min="${index + 1}"`, `max="${index + 1}"`];
    if (dimension.width !== undefined) attributes.push(`width="${number(dimension.width)}"`);
    if (dimension.hidden === true) attributes.push('hidden="1"');
    if (dimension.styleId !== undefined && dimension.styleId !== 0) {
      attributes.push(`s="${dimension.styleId}"`);
    }
    if (dimension.customWidth === true) attributes.push('customWidth="1"');
    if (dimension.bestFit === true) attributes.push('bestFit="1"');
    if (dimension.outlineLevel !== undefined && dimension.outlineLevel > 0) {
      attributes.push(`outlineLevel="${dimension.outlineLevel}"`);
    }
    return `<col ${attributes.join(' ')}/>`;
  });
  return `<cols>${columns.join('')}</cols>`;
}

function rowXml(
  row: number,
  cells: Map<number, Cell>,
  dimension: RowDimension | undefined,
  sharedStrings: SharedStringsAppender | null,
  filterHidden = false,
): string {
  const cellXml = [...cells.entries()]
    .sort(([left], [right]) => left - right)
    .map(([col, cell]) =>
      cellToXml('', formatCellRef(row, col), cell, { sharedStrings, original: null }),
    )
    .filter((xml): xml is string => xml !== null)
    .join('');
  if (cellXml === '' && dimension === undefined) return '';
  const attributes = [`r="${row + 1}"`];
  if (dimension?.height !== undefined) attributes.push(`ht="${number(dimension.height)}"`);
  if (dimension?.customHeight === true) attributes.push('customHeight="1"');
  if (dimension?.hidden === true || filterHidden) attributes.push('hidden="1"');
  if (dimension?.styleId !== undefined && dimension.styleId !== 0) {
    attributes.push(`s="${dimension.styleId}"`);
  }
  if (dimension?.outlineLevel !== undefined && dimension.outlineLevel > 0) {
    attributes.push(`outlineLevel="${dimension.outlineLevel}"`);
  }
  return `<row ${attributes.join(' ')}>${cellXml}</row>`;
}

function mergeXml(sheet: Sheet): string {
  if (sheet.merges.length === 0) return '';
  const merges = sheet.merges
    .map((merge) => {
      const range = merge.range;
      const start = `${columnToLetters(range.startCol)}${range.startRow + 1}`;
      const end = `${columnToLetters(range.endCol)}${range.endRow + 1}`;
      return `<mergeCell ref="${escapeXmlAttribute(start === end ? start : `${start}:${end}`)}"/>`;
    })
    .join('');
  return `<mergeCells count="${sheet.merges.length}">${merges}</mergeCells>`;
}

function number(value: number): string {
  return Number.isFinite(value) ? String(value) : '0';
}
