/** Typed cell values, per spec/workbook-model.md §1. */

export type CellErrorCode =
  '#DIV/0!' | '#VALUE!' | '#N/A' | '#REF!' | '#NAME?' | '#NUM!' | '#NULL!' | '#CIRC!';

export const CELL_ERRORS: readonly CellErrorCode[] = [
  '#DIV/0!',
  '#VALUE!',
  '#N/A',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#NULL!',
  '#CIRC!',
];

export type CellValue =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'text'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean }
  | { readonly type: 'error'; readonly value: CellErrorCode }
  | { readonly type: 'blank' };

export const BLANK: CellValue = Object.freeze({ type: 'blank' as const });

export function numberValue(value: number): CellValue {
  if (!Number.isFinite(value)) {
    throw new RangeError('Cell numbers must be finite');
  }
  return { type: 'number', value };
}

export function textValue(value: string): CellValue {
  return { type: 'text', value };
}

export function booleanValue(value: boolean): CellValue {
  return { type: 'boolean', value };
}

export function errorValue(value: CellErrorCode): CellValue {
  return { type: 'error', value };
}

export function isErrorValue(value: CellValue): value is { type: 'error'; value: CellErrorCode } {
  return value.type === 'error';
}

export function cloneCellValue(value: CellValue): CellValue {
  return value;
}

export function cellValueEquals(a: CellValue, b: CellValue): boolean {
  if (a.type !== b.type) return false;
  switch (a.type) {
    case 'number':
      return b.type === 'number' && Object.is(a.value, b.value);
    case 'text':
      return b.type === 'text' && a.value === b.value;
    case 'boolean':
      return b.type === 'boolean' && a.value === b.value;
    case 'error':
      return b.type === 'error' && a.value === b.value;
    case 'blank':
      return true;
  }
}

export function formatCellValue(value: CellValue): string {
  switch (value.type) {
    case 'number': {
      if (Object.is(value.value, -0)) return '0';
      return String(value.value);
    }
    case 'text':
      return value.value;
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return value.value;
    case 'blank':
      return '';
  }
}

export function isCellErrorCode(value: string): value is CellErrorCode {
  return (CELL_ERRORS as readonly string[]).includes(value);
}

/** Parse a user edit into a typed value. Formula entry is deliberately not part of M1. */
export function parseUserValue(raw: string): CellValue {
  if (raw === '') return BLANK;
  const trimmed = raw.trim();
  if (trimmed === '') return { type: 'text', value: raw };
  if (/^(true|false)$/i.test(trimmed)) {
    return booleanValue(trimmed.toLowerCase() === 'true');
  }
  if (isCellErrorCode(trimmed)) {
    return errorValue(trimmed);
  }
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    const numeric = Number(trimmed);
    if (Number.isFinite(numeric)) return numberValue(numeric);
  }
  return { type: 'text', value: raw };
}
