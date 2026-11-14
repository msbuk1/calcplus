import type { CellValue } from '@model/value';

export const RELATIVE_TOLERANCE = 1e-9;
export const ABSOLUTE_TOLERANCE = 1e-12;

export function numbersApproxEqual(actual: number, expected: number): boolean {
  const difference = Math.abs(actual - expected);
  return (
    difference <=
    Math.max(
      ABSOLUTE_TOLERANCE,
      RELATIVE_TOLERANCE * Math.max(Math.abs(actual), Math.abs(expected)),
    )
  );
}

export interface GoldenCell {
  readonly kind: 'text' | 'number' | 'boolean' | 'error' | 'blank';
  readonly value: string | number | boolean;
}

export function cellValueMatchesGolden(actual: CellValue, expected: GoldenCell): boolean {
  switch (expected.kind) {
    case 'text':
      return actual.type === 'text' && actual.value === expected.value;
    case 'number':
      return (
        actual.type === 'number' &&
        typeof expected.value === 'number' &&
        numbersApproxEqual(actual.value, expected.value)
      );
    case 'boolean':
      return actual.type === 'boolean' && actual.value === expected.value;
    case 'error':
      return actual.type === 'error' && actual.value === expected.value;
    case 'blank':
      return actual.type === 'blank';
  }
}

export function describeValue(value: CellValue): string {
  return `${value.type}:${'value' in value ? String(value.value) : ''}`;
}
