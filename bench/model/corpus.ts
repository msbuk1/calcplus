import type { EngineCellInput, EngineWorkbookInput } from '@engine/port';
import { BLANK, numberValue } from '@model/value';
import { type CorpusShape, formulaCells, sheetCount, touchedCells } from '../corpus/generate';

/**
 * Turn a corpus shape (bench/corpus/generate.ts) into engine inputs.
 *
 * This is the deterministic Node-side interpretation the `bench/model`
 * microbenchmarks build their workbooks from. The browser harness already has
 * its own literal-only hook; the shapes that need formulas, several sheets,
 * comments or conditional formatting are measured here.
 */
export interface CorpusWorkbook {
  readonly shape: string;
  readonly workbook: EngineWorkbookInput;
  readonly cells: readonly EngineCellInput[];
  readonly touchedCells: number;
  readonly formulas: number;
  /** What the shape carries that the engine inputs cannot express. */
  readonly modelOnly: readonly string[];
}

export function buildCorpusWorkbook(shape: CorpusShape): CorpusWorkbook {
  const sheets = sheetCount(shape);
  const cells: EngineCellInput[] = [];
  const totalFormulas = formulaCells(shape);
  const modelOnly: string[] = [];

  if (shape.touchedEvery !== undefined) {
    for (let sheet = 0; sheet < sheets; sheet += 1) {
      for (let row = 0; row < shape.rows; row += shape.touchedEvery) {
        for (let col = 0; col < shape.cols; col += 1) {
          cells.push({
            sheet,
            row,
            col,
            formula: null,
            value: numberValue(literalValue(sheet, row, col)),
          });
        }
      }
    }
    return {
      shape: shape.name,
      workbook: {
        date1904: false,
        sheetNames: sheetNames(shape),
      },
      cells,
      touchedCells: touchedCells(shape),
      formulas: 0,
      modelOnly,
    };
  }

  let formulasLeft = totalFormulas;
  let formulaIndex = 0;
  for (let sheet = 0; sheet < sheets; sheet += 1) {
    const formulasThisSheet = Math.min(
      formulasLeft,
      Math.floor(totalFormulas / sheets) + (sheet < totalFormulas % sheets ? 1 : 0),
    );
    let remaining = formulasThisSheet;
    for (let row = 0; row < shape.rows; row += 1) {
      for (let col = 0; col < shape.cols; col += 1) {
        const isFormulaCell = col > 0 && remaining > 0;
        if (isFormulaCell) {
          remaining -= 1;
          formulasLeft -= 1;
          cells.push({
            sheet,
            row,
            col,
            formula: formulaText(shape, row, formulaIndex),
            value: BLANK,
          });
          formulaIndex += 1;
        } else {
          cells.push({
            sheet,
            row,
            col,
            formula: null,
            value: numberValue(literalValue(sheet, row, col)),
          });
        }
      }
    }
  }

  if ((shape.comments ?? 0) > 0) modelOnly.push('comments');
  if ((shape.hyperlinks ?? 0) > 0) modelOnly.push('hyperlinks');
  if ((shape.merges ?? 0) > 0) modelOnly.push('merged ranges');
  if ((shape.cfRules ?? 0) > 0) modelOnly.push(`${shape.cfRules} conditional-formatting rules`);
  if (shape.styleCount > 6) modelOnly.push(`${shape.styleCount} distinct styles`);

  return {
    shape: shape.name,
    workbook: {
      date1904: false,
      sheetNames: sheetNames(shape),
    },
    cells,
    touchedCells: touchedCells(shape),
    formulas: formulaIndex,
    modelOnly,
  };
}

function formulaText(shape: CorpusShape, row: number, formulaIndex: number): string {
  switch (shape.formulaPattern ?? 'row-local') {
    case 'fan-out-single':
      return '$A$1*2';
    case 'whole-column':
      return 'SUM($A:$A)';
    case 'mixed': {
      const slot = formulaIndex % 10;
      if (slot < 7) return `$A${row + 1}*2+A${row + 1}`;
      if (slot < 9) return '$A$1*2';
      return 'SUM($A:$A)';
    }
    case 'none':
      return `$A${row + 1}+1`;
    case 'row-local':
    default:
      return `$A${row + 1}*2`;
  }
}

function literalValue(sheet: number, row: number, col: number): number {
  return ((row * 31 + col * 17 + sheet * 7) % 1000) + 1;
}

function sheetNames(shape: CorpusShape): readonly string[] {
  const sheets = sheetCount(shape);
  return Array.from({ length: sheets }, (_, index) => `Sheet${index + 1}`);
}
