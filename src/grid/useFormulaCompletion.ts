import { useMemo, useState } from 'react';
import type { FunctionCatalogEntry } from '@engine/function-catalog';
import {
  applyFormulaCompletion,
  formulaCompletionAt,
  type CompletionChoice,
  type FormulaCompletion,
} from '@engine/formula-suggestions';

/**
 * The editing state shared by the formula bar and the in-cell editor: which
 * function names match the token at the caret, which row is highlighted, and
 * how a choice rewrites the text. Dismissal and the highlight are keyed to the
 * exact text and caret, so the next keystroke or caret move reopens the list.
 */
export interface FormulaCompletionController {
  /** The completion to render: a name list or a signature hint. */
  readonly completion: FormulaCompletion | null;
  /** A name list is visible and the caller should consume its navigation keys. */
  readonly listOpen: boolean;
  readonly activeIndex: number;
  readonly activeEntry: FunctionCatalogEntry | null;
  move(delta: number): void;
  accept(): CompletionChoice | null;
  choose(entry: FunctionCatalogEntry): CompletionChoice | null;
  dismiss(): void;
}

export function useFormulaCompletion(input: {
  readonly text: string;
  readonly caret: number | null;
  /** The edited text is formula entry. */
  readonly formula: boolean;
  /** The surface is focused and editing. */
  readonly enabled: boolean;
}): FormulaCompletionController {
  const { text, caret, formula, enabled } = input;
  const completion = useMemo(
    () => formulaCompletionAt(text, caret ?? text.length, { formula: enabled && formula }),
    [text, caret, formula, enabled],
  );

  // A dismissal hides the list until the text or caret changes again.
  const surfaceKey = `${text}\u0000${caret ?? text.length}`;
  const [dismissed, setDismissed] = useState<string | null>(null);
  const [selection, setSelection] = useState<{ key: string; index: number }>({
    key: '',
    index: 0,
  });

  const visible = dismissed === surfaceKey ? null : completion;
  const matches = visible?.kind === 'names' ? visible.matches : [];
  const listKey = visible?.kind === 'names' ? `${surfaceKey}\u0000${visible.query}` : '';
  const index =
    matches.length === 0 ? -1 : (selection.key === listKey ? selection.index : 0) % matches.length;
  const activeEntry = index < 0 ? null : (matches[index] ?? null);

  return {
    completion: visible,
    listOpen: visible?.kind === 'names' && matches.length > 0,
    activeIndex: index,
    activeEntry,
    move(delta) {
      if (matches.length === 0) return;
      const next =
        ((((index < 0 ? 0 : index) + delta) % matches.length) + matches.length) % matches.length;
      setSelection({ key: listKey, index: next });
    },
    accept() {
      if (visible?.kind !== 'names' || activeEntry === null) return null;
      return applyFormulaCompletion(text, visible, activeEntry.name);
    },
    choose(entry) {
      if (visible?.kind !== 'names') return null;
      return applyFormulaCompletion(text, visible, entry.name);
    },
    dismiss() {
      setDismissed(surfaceKey);
    },
  };
}
