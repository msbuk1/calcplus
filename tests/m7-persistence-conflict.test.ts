import { describe, expect, it } from 'vitest';
import {
  EXTERNAL_CHANGE_SAVE_AS_QUESTION,
  EXTERNAL_EDIT_UNSUPPORTED_WARNING,
  TAB_CLOSE_WARNING,
  chooseExternalChange,
  confirmRecoveryRestore,
  externalChangeOverwriteQuestion,
  externalChangeReloadQuestion,
  recoveryRestoreQuestion,
  shouldWarnBeforeUnload,
} from '@io/confirm';
import { stampChanged, stampOf } from '@io/file-system';

/**
 * M7 lane 5 — the conflict and close prompts (persistence §2, §4).
 *
 * The storage layer exposes the decision functions; the app injects the
 * browser confirm exactly as the banked signature and macro prompts do. No
 * branch leaves the newer external file overwritten without an explicit yes.
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
  it('offers Reload first, then Overwrite, then Save As', () => {
    const reload = answers(true);
    expect(chooseExternalChange('book.xlsx', reload.confirm)).toBe('reload');
    expect(reload.asked).toEqual([externalChangeReloadQuestion('book.xlsx')]);
    expect(reload.asked[0]).toContain('changed on disk');
    expect(reload.asked[0]).toContain('discard your unsaved changes');

    const overwrite = answers(false, true);
    expect(chooseExternalChange('book.xlsx', overwrite.confirm)).toBe('overwrite');
    expect(overwrite.asked).toEqual([
      externalChangeReloadQuestion('book.xlsx'),
      externalChangeOverwriteQuestion('book.xlsx'),
    ]);
    expect(overwrite.asked[1]).toContain('overwrite the newer');

    const saveAs = answers(false, false, true);
    expect(chooseExternalChange('book.xlsx', saveAs.confirm)).toBe('saveAs');
    expect(saveAs.asked[2]).toBe(EXTERNAL_CHANGE_SAVE_AS_QUESTION);
    expect(saveAs.asked[2]).toContain('new file');
  });

  it('writes nothing when every branch is declined', () => {
    const cancelled = answers(false, false, false);
    expect(chooseExternalChange('book.xlsx', cancelled.confirm)).toBe('cancel');
    expect(cancelled.asked).toHaveLength(3);
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
