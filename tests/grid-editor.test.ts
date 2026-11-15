import { describe, expect, it } from 'vitest';
import {
  editorHasChanges,
  resolveEditorAction,
  startEdit,
  type EditorKeyEvent,
} from '@grid/editor';

const address = { row: 2, col: 1 };

function keyEvent(overrides: Partial<EditorKeyEvent> = {}): EditorKeyEvent {
  return {
    key: 'Enter',
    shiftKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    ...overrides,
  };
}

describe('in-cell editor', () => {
  it('commits only a changed F2 edit', () => {
    const session = startEdit(address, '42', true);
    expect(editorHasChanges(session)).toBe(false);
    expect(editorHasChanges({ ...session, text: '43' })).toBe(true);
  });

  it('always commits a typed edit', () => {
    const session = startEdit(address, '4', false);
    expect(editorHasChanges(session)).toBe(true);
  });

  it('commits with Enter and Tab, and moves in the Excel direction', () => {
    const session = startEdit(address, 'x', false);
    expect(resolveEditorAction(session, keyEvent({ key: 'Enter' }))).toEqual({
      type: 'commit',
      direction: 'down',
    });
    expect(resolveEditorAction(session, keyEvent({ key: 'Enter', shiftKey: true }))).toEqual({
      type: 'commit',
      direction: 'up',
    });
    expect(resolveEditorAction(session, keyEvent({ key: 'Tab' }))).toEqual({
      type: 'commit',
      direction: 'right',
    });
    expect(resolveEditorAction(session, keyEvent({ key: 'Tab', shiftKey: true }))).toEqual({
      type: 'commit',
      direction: 'left',
    });
    expect(resolveEditorAction(session, keyEvent({ key: 'Escape' }))).toEqual({ type: 'cancel' });
    expect(resolveEditorAction(session, keyEvent({ key: 'a' }))).toEqual({ type: 'ignore' });
  });

  it('never commits while an input method is composing', () => {
    const session = { ...startEdit(address, '', false), composing: true };
    expect(resolveEditorAction(session, keyEvent({ key: 'Enter' }))).toEqual({ type: 'ignore' });
    expect(resolveEditorAction(session, keyEvent({ key: 'Escape' }))).toEqual({ type: 'ignore' });
    const reported = { ...startEdit(address, '', false), composing: false };
    expect(resolveEditorAction(reported, keyEvent({ key: 'Enter', isComposing: true }))).toEqual({
      type: 'ignore',
    });
  });
});
