import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { builtinNumberFormat } from '@model/style';
import { numberValue, textValue, type CellValue } from '@model/value';
import { formatValue } from '@view/number-format';

/**
 * The committed built-in number-format corpus (number-formats §7.1-7.2).
 *
 * `ssf` is a differential oracle only; it is not installed and this test never
 * reaches the network. See `fixtures/README.md` for the gate command.
 */

interface CorpusCase {
  readonly numFmtId: number;
  readonly code: string;
  readonly valueType: 'number' | 'text';
  readonly value: number | string;
  readonly expected: string;
}

interface Corpus {
  readonly schemaVersion: number;
  readonly provenance: {
    readonly kind: string;
    readonly engine: string;
    readonly oracle: string | null;
    readonly gate: string;
    readonly differentialOracle: string;
    readonly note: string;
  };
  readonly cases: readonly CorpusCase[];
}

const corpus = JSON.parse(
  readFileSync(new URL('../fixtures/expected/m4a-number-formats.json', import.meta.url), 'utf8'),
) as Corpus;

function corpusValue(testCase: CorpusCase): CellValue {
  return testCase.valueType === 'text'
    ? textValue(String(testCase.value))
    : numberValue(Number(testCase.value));
}

describe('M4A built-in number-format corpus', () => {
  it('pins the exact display text for every built-in id', () => {
    expect(corpus.schemaVersion).toBe(1);
    expect(corpus.provenance.kind).toBe('hand-authored');
    expect(corpus.provenance.engine).toBe('numfmt 4.1.3');
    expect(corpus.provenance.oracle).toBeNull();
    expect(corpus.provenance.gate).toBe('npm test -- tests/m4a-number-formats.test.ts');
    expect(corpus.provenance.differentialOracle).toContain('ssf');
    expect(corpus.cases).toHaveLength(74);

    for (const testCase of corpus.cases) {
      const label = `numFmtId ${testCase.numFmtId}`;
      expect(builtinNumberFormat(testCase.numFmtId), label).toBe(testCase.code);
      expect(formatValue(corpusValue(testCase), testCase.code), label).toBe(testCase.expected);
    }
  });
});
