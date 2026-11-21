import { describe, expect, it } from 'vitest';
import {
  EXTERNAL_EDIT_UNSUPPORTED_WARNING,
  TAB_CLOSE_WARNING,
  chooseExternalChange,
  confirmRecoveryRestore,
  externalChangePrompt,
  recoveryRestoreQuestion,
  shouldWarnBeforeUnload,
} from '@io/confirm';
import { stampChanged, stampOf } from '@io/file-system';

/**
 * M7 lane 5 — the conflict and close prompts (persistence §2, §4).
 *
 * The storage layer exposes the decision functions; the app injects an
 * in-page dialog asker that resolves with the chosen button. A native confirm
 * cannot carry the user gesture a file picker needs, so the prompt is one
 * dialog, and declining its every branch writes nothing.
 */

function answers(...values: boolean[]): {
  readonly confirm: (message: string) => boolean;
  readonly asked: string[];
} {
  const asked: string[] = [];
  return {
    confirm: (message) => {
      asked.push(message);
      return values[asked.length - 1] ?? false;
    },
    asked,
  };
}

describe('external edit detection', () => {
  it('compares the modification time and the size captured at open', () => {
    const recorded = { lastModified: 1_000, size: 4_096 };
    expect(stampChanged(recorded, { lastModified: 1_000, size: 4_096 })).toBe(false);
    expect(stampChanged(recorded, { lastModified: 2_000, size: 4_096 })).toBe(true);
    expect(stampChanged(recorded, { lastModified: 1_000, size: 8_192 })).toBe(true);
  });

  it('never claims a change when the browser exposes no stamp', () => {
    expect(stampChanged(null, { lastModified: 2_000, size: 1 })).toBe(false);
    expect(stampChanged({ lastModified: 1_000, size: 1 }, null)).toBe(false);
    expect(stampOf({ lastModified: 1_234, size: 8 })).toEqual({ lastModified: 1_234, size: 8 });
    expect(stampOf({})).toBeNull();
  });
});

describe('external change prompt', () => {
  it('offers Reload, Overwrite, Save As and Cancel in the contract order', () => {
    const prompt = externalChangePrompt('book.xlsx');
    expect(prompt.displayName).toBe('book.xlsx');
    expect(prompt.summary).toContain('book.xlsx');
    expect(prompt.summary).toContain('changed on disk');
    const order = [...prompt.outcomes.map((outcome) => outcome.choice), 'cancel'];
    expect(order).toEqual(['reload', 'overwrite', 'saveAs', 'cancel']);
    expect(prompt.outcomes.map((outcome) => outcome.label)).toEqual([
      'Reload',
      'Overwrite',
      'Save As',
    ]);
    expect(prompt.cancelLabel).toBe('Cancel');
  });

  it('states both destructive outcomes plainly', () => {
    const prompt = externalChangePrompt('book.xlsx');
    const [reload, overwrite, saveAs] = prompt.outcomes;
    // One of the two destructive choices loses work either way, so the prompt
    // names each loss and never overwrites the newer file without an explicit
    // Overwrite choice.
    expect(reload!.statement).toContain('discard your unsaved changes');
    expect(reload!.statement).toContain('unsaved changes are lost');
    expect(overwrite!.statement).toContain('overwrite the newer');
    expect(overwrite!.statement).toContain('external edit is lost');
    expect(saveAs!.statement).toContain('new file');
  });

  it('resumes the save flow with the chosen outcome', async () => {
    for (const choice of ['reload', 'overwrite', 'saveAs', 'cancel'] as const) {
      const asked: unknown[] = [];
      const picked = await chooseExternalChange('book.xlsx', async (prompt) => {
        asked.push(prompt);
        return choice;
      });
      expect(picked).toBe(choice);
      expect(asked).toHaveLength(1);
    }
  });

  it('writes nothing when no dialog is wired', async () => {
    // Declining every branch cancels and writes nothing (persistence §4).
    await expect(chooseExternalChange('book.xlsx')).resolves.toBe('cancel');
  });

  it('says external edits cannot be detected on a browser without handles', () => {
    expect(EXTERNAL_EDIT_UNSUPPORTED_WARNING).toContain('cannot be detected');
    expect(EXTERNAL_EDIT_UNSUPPORTED_WARNING).toContain('may be overwritten');
  });
});

describe('recovery prompt', () => {
  it('names the file and both outcomes', () => {
    const message = recoveryRestoreQuestion('book.xlsx');
    expect(message).toContain('book.xlsx');
    expect(message).toContain('Restore');
    expect(message).toContain('discard the recovery copy');
  });

  it('honours the injected answer for Restore and Discard', () => {
    const restore = answers(true);
    expect(confirmRecoveryRestore('book.xlsx', restore.confirm)).toBe(true);
    expect(restore.asked).toEqual([recoveryRestoreQuestion('book.xlsx')]);

    const discard = answers(false);
    expect(confirmRecoveryRestore('book.xlsx', discard.confirm)).toBe(false);
    expect(discard.asked).toEqual([recoveryRestoreQuestion('book.xlsx')]);
  });
});

describe('tab close prompt', () => {
  it('warns only when the document has unsaved changes', () => {
    expect(shouldWarnBeforeUnload(false)).toBe(false);
    expect(shouldWarnBeforeUnload(true)).toBe(true);
  });

  it('records the browser wording limit instead of promising custom text', () => {
    expect(TAB_CLOSE_WARNING).toContain('unsaved changes');
    expect(TAB_CLOSE_WARNING).toContain('browsers do not let a page choose the wording');
  });
});
