import { describe, expect, it } from 'vitest';
import {
  FUNCTION_CATALOG,
  FUNCTION_CATEGORY_OPTIONS,
  functionEntry,
  functionSkeleton,
  functionsInCategory,
  searchFunctions,
} from '@engine/function-catalog';
import { PUBLISHED_FUNCTIONS } from '@engine/function-registry';

const AUTHORED_CATEGORIES = FUNCTION_CATEGORY_OPTIONS.filter(
  (category) => category !== 'Recently used' && category !== 'All',
);

describe('Insert Function catalogue', () => {
  it('covers every published function exactly once', () => {
    expect(FUNCTION_CATALOG).toHaveLength(52);
    expect(new Set(FUNCTION_CATALOG.map((entry) => entry.name)).size).toBe(52);
    expect([...FUNCTION_CATALOG.map((entry) => entry.name)].sort()).toEqual(
      [...PUBLISHED_FUNCTIONS].sort(),
    );
  });

  it('gives every entry a valid category, signature and description', () => {
    for (const entry of FUNCTION_CATALOG) {
      expect(AUTHORED_CATEGORIES).toContain(entry.category);
      expect(entry.signature.startsWith(`${entry.name}(`)).toBe(true);
      expect(entry.signature.endsWith(')')).toBe(true);
      expect(entry.description.trim()).not.toBe('');
      expect(entry.description.endsWith('.')).toBe(true);
    }
  });

  it('filters each Excel category down to the supported functions', () => {
    const counts = Object.fromEntries(
      AUTHORED_CATEGORIES.map((category) => [category, functionsInCategory(category).length]),
    );
    expect(counts).toEqual({
      Financial: 0,
      'Date & Time': 6,
      'Math & Trig': 10,
      Statistical: 9,
      'Lookup & Reference': 8,
      Text: 8,
      Logical: 6,
      Information: 5,
    });
    for (const category of AUTHORED_CATEGORIES) {
      for (const entry of functionsInCategory(category)) {
        expect(entry.category).toBe(category);
      }
    }
  });

  it('lists All alphabetically and Recently used in most-recent-first order', () => {
    const all = functionsInCategory('All');
    expect(all).toHaveLength(52);
    expect(all.map((entry) => entry.name)).toEqual(
      [...all.map((entry) => entry.name)].sort((left, right) => left.localeCompare(right)),
    );

    expect(functionsInCategory('Recently used')).toEqual([]);
    expect(
      functionsInCategory('Recently used', ['AVERAGE', 'SUM', 'AVERAGE']).map(
        (entry) => entry.name,
      ),
    ).toEqual(['AVERAGE', 'SUM']);
    // An unknown name is ignored rather than breaking the dialog.
    expect(functionsInCategory('Recently used', ['NOT_A_FUNCTION'])).toEqual([]);
  });

  it('searches names case-insensitively and ranks prefix matches first', () => {
    expect(searchFunctions('')).toHaveLength(52);
    expect(searchFunctions('average').map((entry) => entry.name)).toEqual(['AVERAGE']);
    expect(searchFunctions('SUM').map((entry) => entry.name)).toEqual(['SUM', 'SUMIF']);
    // `IF` is a prefix of three names and a substring of two more.
    expect(searchFunctions('if').map((entry) => entry.name)).toEqual([
      'IF',
      'IFERROR',
      'IFNA',
      'COUNTIF',
      'SUMIF',
    ]);
    expect(searchFunctions('lookup').map((entry) => entry.name)).toEqual(['HLOOKUP', 'VLOOKUP']);
    expect(searchFunctions('  numbers  ')).toEqual([]);
  });

  it('exposes a lookup and the skeleton the dialog inserts', () => {
    expect(functionEntry('AVERAGE')?.category).toBe('Statistical');
    expect(functionEntry('average')).toBeUndefined();
    expect(functionEntry('NOT_A_FUNCTION')).toBeUndefined();
    expect(functionSkeleton('SUM')).toBe('=SUM(');
    expect(functionSkeleton('AVERAGE')).toBe('=AVERAGE(');
  });
});
