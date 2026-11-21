import { describe, expect, it } from 'vitest';
import { ensureHandlePermission, type FileHandleLike } from '@io/file-system';

/**
 * M7 lane 5 — re-requesting File System Access permission (persistence
 * §1/§3).
 *
 * A handle revived from a recovery copy carries no live permission after the
 * tab is gone, so save must ask again before touching the file. A handle with
 * no permission API (a plain stub) is treated as usable, and a browser that
 * refuses access is reported so the caller can fall back instead of failing
 * mid-write.
 */

interface StubOptions {
  /** The state `queryPermission` reports. */
  readonly query?: PermissionState;
  /** The state `requestPermission` resolves to. */
  readonly request?: PermissionState;
  /** A handle without any permission methods, like a plain stub. */
  readonly noApi?: boolean;
  /** A handle that answers `queryPermission` but cannot ask for more. */
  readonly queryOnly?: boolean;
  /** A handle whose permission methods throw, like a revoked handle. */
  readonly throws?: boolean;
}

function stubHandle(options: StubOptions = {}): {
  handle: FileHandleLike;
  queried: string[];
  requested: string[];
} {
  const queried: string[] = [];
  const requested: string[] = [];
  const writable = async (): Promise<{
    write(data: BufferSource): Promise<void>;
    close(): Promise<void>;
  }> => ({
    write: async () => undefined,
    close: async () => undefined,
  });
  const base: FileHandleLike = {
    name: 'book.xlsx',
    createWritable: writable,
  };
  if (options.noApi === true) return { handle: base, queried, requested };
  return {
    handle: {
      ...base,
      queryPermission: async (descriptor: { readonly mode: 'read' | 'readwrite' }) => {
        if (options.throws === true) throw new Error('permission state is gone');
        queried.push(descriptor.mode);
        return options.query ?? 'prompt';
      },
      ...(options.queryOnly === true
        ? {}
        : {
            requestPermission: async (descriptor: { readonly mode: 'read' | 'readwrite' }) => {
              if (options.throws === true) throw new Error('permission request is gone');
              requested.push(descriptor.mode);
              return options.request ?? 'granted';
            },
          }),
    },
    queried,
    requested,
  };
}

describe('ensureHandlePermission', () => {
  it('treats a stub with no permission API as usable', async () => {
    const { handle, queried, requested } = stubHandle({ noApi: true });
    expect(await ensureHandlePermission(handle)).toBe(true);
    expect(queried).toEqual([]);
    expect(requested).toEqual([]);
  });

  it('uses an already-granted handle without prompting', async () => {
    const { handle, queried, requested } = stubHandle({ query: 'granted' });
    expect(await ensureHandlePermission(handle)).toBe(true);
    expect(queried).toEqual(['readwrite']);
    expect(requested).toEqual([]);
  });

  it('requests permission on a prompted handle and reports the result', async () => {
    const granted = stubHandle({ query: 'prompt', request: 'granted' });
    expect(await ensureHandlePermission(granted.handle)).toBe(true);
    expect(granted.queried).toEqual(['readwrite']);
    expect(granted.requested).toEqual(['readwrite']);

    const refused = stubHandle({ query: 'prompt', request: 'denied' });
    expect(await ensureHandlePermission(refused.handle)).toBe(false);
    expect(refused.queried).toEqual(['readwrite']);
    expect(refused.requested).toEqual(['readwrite']);
  });

  it('passes the requested access mode through', async () => {
    const readOnly = stubHandle({ query: 'prompt', request: 'granted' });
    expect(await ensureHandlePermission(readOnly.handle, 'read')).toBe(true);
    expect(readOnly.queried).toEqual(['read']);
    expect(readOnly.requested).toEqual(['read']);
  });

  it('returns false when the handle is denied or unaskable', async () => {
    const denied = stubHandle({ query: 'denied', request: 'denied' });
    expect(await ensureHandlePermission(denied.handle)).toBe(false);
    expect(denied.requested).toEqual(['readwrite']);

    const cannotAsk = stubHandle({ query: 'prompt', queryOnly: true });
    expect(await ensureHandlePermission(cannotAsk.handle)).toBe(false);
    expect(cannotAsk.requested).toEqual([]);

    const broken = stubHandle({ throws: true });
    expect(await ensureHandlePermission(broken.handle)).toBe(false);
  });
});
