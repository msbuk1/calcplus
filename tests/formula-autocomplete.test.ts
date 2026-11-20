import { describe, expect, it } from 'vitest';
import {
  applyFormulaCompletion,
  formulaCompletionAt,
  signatureSegments,
} from '@engine/formula-suggestions';

/**
 * Function autocomplete and signature hints (feature matrix §4, command
 * `formula.autocomplete`). The cell editor and the formula bar share this
 * module, so the tests cover the editor's hidden leading `=` as well as the
 * bar's authored text.
 */

describe('function name suggestions', () => {
  it('suggests the published functions that start with the typed token', () => {
    const completion = formulaCompletionAt('=SU', 3, { formula: true });
    expect(completion?.kind).toBe('names');
    if (completion?.kind !== 'names') throw new Error('expected a name list');
    expect(completion.query).toBe('SU');
    expect(completion.start).toBe(1);
    expect(completion.end).toBe(3);
    expect(completion.matches.map((entry) => entry.name)).toEqual(['SUM', 'SUMIF']);
  });

  it('treats the hidden leading equals of an in-cell formula as the start of the formula', () => {
    const completion = formulaCompletionAt('AVER', 4, { formula: true });
    expect(completion?.kind).toBe('names');
    if (completion?.kind !== 'names') throw new Error('expected a name list');
    expect(completion.start).toBe(0);
    expect(completion.matches.map((entry) => entry.name)).toEqual(['AVERAGE']);
  });

  it('suggests nothing for a cell reference or a plain value', () => {
    expect(formulaCompletionAt('A1', 2, { formula: true })).toBeNull();
    expect(formulaCompletionAt('hello', 5, { formula: false })).toBeNull();
  });

  it('ignores tokens inside a quoted text literal', () => {
    expect(formulaCompletionAt('="SU', 4, { formula: true })).toBeNull();
    // A doubled quote ends the literal and opens a new one in Excel's grammar.
    expect(formulaCompletionAt('="a""SU', 7, { formula: true })).toBeNull();
  });

  it('replaces the token with the skeleton and leaves the caret on the argument', () => {
    const completion = formulaCompletionAt('=SU', 3, { formula: true });
    if (completion?.kind !== 'names') throw new Error('expected a name list');
    const choice = applyFormulaCompletion('=SU', completion, 'SUM');
    expect(choice.text).toBe('=SUM(');
    expect(choice.caret).toBe(5);
  });
});

describe('signature hints', () => {
  it('shows the function signature and the current argument index', () => {
    const empty = formulaCompletionAt('=SUM(', 5, { formula: true });
    expect(empty).toEqual({
      kind: 'signature',
      name: 'SUM',
      signature: 'SUM(number1, [number2], ...)',
      argumentIndex: 0,
      arguments: ['number1', '[number2]', '...'],
    });

    const second = formulaCompletionAt('=SUM(A1,', 8, { formula: true });
    expect(second?.kind).toBe('signature');
    if (second?.kind !== 'signature') throw new Error('expected a signature');
    expect(second.argumentIndex).toBe(1);
  });

  it('finds the innermost call when the caret is in a nested argument', () => {
    const hint = formulaCompletionAt('=IF(A1,SUM(B1,', 14, { formula: true });
    expect(hint?.kind).toBe('signature');
    if (hint?.kind !== 'signature') throw new Error('expected a signature');
    expect(hint.name).toBe('SUM');
    expect(hint.argumentIndex).toBe(1);
  });

  it('emphasises only the argument being edited', () => {
    const hint = formulaCompletionAt('=SUM(A1,', 8, { formula: true });
    if (hint?.kind !== 'signature') throw new Error('expected a signature');
    expect(signatureSegments(hint).map((segment) => [segment.text, segment.active])).toEqual([
      ['SUM(', false],
      ['number1', false],
      [', ', false],
      ['[number2]', true],
      [', ', false],
      ['...', false],
      [')', false],
    ]);
  });

  it('offers no signature inside a quoted literal', () => {
    expect(formulaCompletionAt('="SUM(', 6, { formula: true })).toBeNull();
  });
});
