import { decodeXml, parseXml } from './xml';

/**
 * The `xl/threadedComments/*` part and the `xl/persons/*` part (file-compat
 * §6.6). v1 reads them for the comment model and the cell indicator; replies,
 * mentions and resolved state are read-only, and both parts are copied
 * verbatim on save because they are not modelled part-by-part.
 *
 * Nothing here executes, resolves or rewrites a comment: the part is opaque,
 * so a structural edit renumbers its references through the range-transform
 * module (`transformThreadedCommentRefs`), never through this parser.
 */

export interface ThreadedCommentEntry {
  /** The cell the thread entry is anchored to. */
  readonly ref: string;
  readonly text: string;
  readonly personId: string | null;
  readonly id: string | null;
  /** The entry this one replies to, or `null` for a root comment. */
  readonly parentId: string | null;
  /** The authored timestamp, ISO 8601 text, or `null`. */
  readonly date: string | null;
}

export interface PersonEntry {
  readonly id: string;
  readonly displayName: string;
  readonly userId: string | null;
  readonly providerId: string | null;
}

interface OpenThreadedComment {
  ref: string;
  personId: string | null;
  id: string | null;
  parentId: string | null;
  date: string | null;
  text: string;
  depth: number;
}

/** Parse one threaded-comments part. An unreadable part yields no entries. */
export function parseThreadedCommentsPart(
  bytes: Uint8Array,
  partPath: string,
): ThreadedCommentEntry[] {
  const xml = decodeXml(bytes);
  const entries: ThreadedCommentEntry[] = [];
  let current: OpenThreadedComment | null = null;
  let textDepth = -1;
  let depth = 0;

  parseXml(
    xml,
    {
      onOpenTag(tag) {
        depth += 1;
        if (tag.local === 'threadedComment' && current === null) {
          current = {
            ref: attribute(tag, 'ref') ?? '',
            personId: attribute(tag, 'personId'),
            id: attribute(tag, 'id'),
            parentId: attribute(tag, 'parentId'),
            date: attribute(tag, 'dT'),
            text: '',
            depth,
          };
          return;
        }
        if (tag.local === 'text' && current !== null) textDepth = depth;
      },
      onCloseTag(tag) {
        if (tag.local === 'text' && textDepth === depth) textDepth = -1;
        if (tag.local === 'threadedComment' && current !== null && current.depth === depth) {
          const entry = current;
          if (entry.ref !== '') {
            entries.push({
              ref: entry.ref,
              text: entry.text.trim(),
              personId: entry.personId,
              id: entry.id,
              parentId: entry.parentId,
              date: entry.date,
            });
          }
          current = null;
        }
        depth -= 1;
      },
      onText(text) {
        if (current !== null && textDepth >= 0) current.text += text;
      },
      onCdata(text) {
        if (current !== null && textDepth >= 0) current.text += text;
      },
    },
    partPath,
  );
  return entries;
}

/** Parse the `personList` part that names thread authors. */
export function parsePersonsPart(bytes: Uint8Array, partPath: string): PersonEntry[] {
  const xml = decodeXml(bytes);
  const persons: PersonEntry[] = [];
  parseXml(
    xml,
    {
      onOpenTag(tag) {
        if (tag.local !== 'person') return;
        const id = attribute(tag, 'id') ?? '';
        if (id === '') return;
        persons.push({
          id,
          displayName: attribute(tag, 'displayName') ?? '',
          userId: attribute(tag, 'userId'),
          providerId: attribute(tag, 'providerId'),
        });
      },
    },
    partPath,
  );
  return persons;
}

function attribute(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
): string | null {
  for (const candidate of Object.values(tag.attributes)) {
    if (candidate.local === local) return candidate.value;
  }
  return null;
}
