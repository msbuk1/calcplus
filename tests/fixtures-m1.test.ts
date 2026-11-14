import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import {
  cellValueMatchesGolden,
  describeValue,
  numbersApproxEqual,
  type GoldenCell,
} from '@testing/compare';

interface ExpectedFile {
  readonly fixture: string;
  readonly provenance: {
    readonly kind: string;
    readonly oracle: null;
    readonly note: string;
  };
  readonly cells: {
    readonly sheet: string;
    readonly ref: string;
    readonly kind: GoldenCell['kind'] | 'formula';
    readonly value: string | number | boolean;
    readonly cached?: number;
  }[];
}

const expected = JSON.parse(
  readFileSync(new URL('../fixtures/expected/m1-values.json', import.meta.url), 'utf8'),
) as ExpectedFile;

const oracleManifest = JSON.parse(
  readFileSync(new URL('../fixtures/oracle-manifest.json', import.meta.url), 'utf8'),
) as { status: string; goldens: unknown[] };

describe('M1 hand-authored fixture values', () => {
  it('declares hand-authored provenance, not Excel-oracle golden provenance', () => {
    expect(expected.provenance.kind).toBe('hand-authored');
    expect(expected.provenance.oracle).toBeNull();
    expect(expected.provenance.note).toMatch(/not Excel 365 oracle goldens/);
  });

  it('has no committed Excel-oracle goldens while the oracle is not configured', () => {
    expect(oracleManifest.status).toBe('not-configured');
    expect(oracleManifest.goldens).toEqual([]);
  });

  const opened = openWorkbook(
    new Uint8Array(
      readFileSync(new URL(`../fixtures/compat/${expected.fixture}`, import.meta.url)),
    ),
  );

  for (const golden of expected.cells) {
    it(`${golden.sheet}!${golden.ref} matches its golden within tolerance`, () => {
      const sheet = opened.workbook.sheets.find((candidate) => candidate.name === golden.sheet);
      expect(sheet).toBeDefined();
      const match = /^([A-Z]+)(\d+)$/.exec(golden.ref)!;
      let col = 0;
      for (const character of match[1]!) {
        col = col * 26 + (character.charCodeAt(0) - 64);
      }
      const cell = sheet!.cells.get(`${Number(match[2]) - 1}:${col - 1}`);
      expect(cell).toBeDefined();
      if (golden.kind === 'formula') {
        expect(cell!.formula).toBe(golden.value);
        expect(cell!.savedResult?.type).toBe('number');
        expect(
          cell!.savedResult?.type === 'number' &&
            numbersApproxEqual(cell!.savedResult.value, golden.cached ?? Number.NaN),
        ).toBe(true);
      } else {
        expect(
          cellValueMatchesGolden(cell!.value, {
            kind: golden.kind,
            value: golden.value,
          }),
        ).toBe(true);
        if (!cellValueMatchesGolden(cell!.value, { kind: golden.kind, value: golden.value })) {
          throw new Error(`${golden.ref} was ${describeValue(cell!.value)}`);
        }
      }
    });
  }
});
