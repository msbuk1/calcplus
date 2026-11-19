import type { CommentRecord, Sheet } from '@model/types';
import { cellKey, parseCellRef } from '@model/types';
import { decodeXml, parseXml } from './xml';
import { findRelationship } from './relationships';
import type { Relationship } from './types';
import {
  parsePersonsPart,
  parseThreadedCommentsPart,
  type PersonEntry,
  type ThreadedCommentEntry,
} from './threaded-comments-xml';

/**
 * Legacy notes (`xl/comments*.xml`) and the comment model (file-compat §3, §9
 * F12). A note's cell, text and author are read into `Sheet.comments`; the
 * part itself, its VML shape and every threaded-comment part stay opaque and
 * byte-identical on save. v1 never authors, edits or resolves a comment
 * (feature matrix "Legacy notes and threaded comments: read-only").
 *
 * The cell indicator is painted from `Sheet.comments`; the grid receives the
 * key set through the sheet projection, so this module stays model-side.
 */

export interface LegacyCommentEntry {
  readonly ref: string;
  readonly text: string;
  readonly author: string | null;
}

/** Parse one `xl/comments*.xml` part. Unreadable content yields no notes. */
export function parseCommentsPart(bytes: Uint8Array, partPath: string): LegacyCommentEntry[] {
  const xml = decodeXml(bytes);
  const authors: string[] = [];
  const notes: LegacyCommentEntry[] = [];
  let depth = 0;
  let inAuthors = false;
  let inText = false;
  let phoneticDepth = -1;
  let authorText = '';
  let current: { ref: string; authorId: string | null; text: string; depth: number } | null = null;

  parseXml(
    xml,
    {
      onOpenTag(tag) {
        depth += 1;
        if (tag.local === 'authors') {
          inAuthors = true;
          return;
        }
        if (tag.local === 'author' && inAuthors) {
          authorText = '';
          return;
        }
        if (tag.local === 'comment') {
          if (current !== null) return;
          current = {
            ref: attribute(tag, 'ref') ?? '',
            authorId: attribute(tag, 'authorId'),
            text: '',
            depth,
          };
          return;
        }
        if (current !== null) {
          if (tag.local === 'text') inText = true;
          else if (tag.local === 'rPh') phoneticDepth = depth;
          else if (tag.local === 't' && inText && phoneticDepth < 0) phoneticDepth = -2;
        }
      },
      onCloseTag(tag) {
        if (tag.local === 't' && phoneticDepth === -2) phoneticDepth = -1;
        if (tag.local === 'rPh' && phoneticDepth === depth) phoneticDepth = -1;
        if (tag.local === 'text') inText = false;
        if (tag.local === 'author' && inAuthors) authors.push(authorText.trim());
        if (tag.local === 'authors') inAuthors = false;
        if (tag.local === 'comment' && current !== null && current.depth === depth) {
          const entry = current;
          if (entry.ref !== '') {
            const index = entry.authorId === null ? Number.NaN : Number(entry.authorId);
            notes.push({
              ref: entry.ref,
              text: entry.text.trim(),
              author: Number.isInteger(index) ? (authors[index] ?? null) : null,
            });
          }
          current = null;
        }
        depth -= 1;
      },
      onText(text) {
        if (inAuthors) authorText += text;
        else if (current !== null && inText && phoneticDepth < 0) current.text += text;
      },
      onCdata(text) {
        if (inAuthors) authorText += text;
        else if (current !== null && inText && phoneticDepth < 0) current.text += text;
      },
    },
    partPath,
  );
  return notes;
}

export interface SheetCommentParts {
  readonly entries: ReadonlyMap<string, Uint8Array>;
  /** The worksheet's own relationships (`xl/worksheets/_rels/sheetN.xml.rels`). */
  readonly relationships: readonly Relationship[];
  /** The workbook's relationships, which carry the persons part. */
  readonly workbookRelationships: readonly Relationship[];
}

/**
 * Populate `Sheet.comments` from the parts the worksheet's relationships name:
 * the legacy comments part, the threaded-comments part and (through the
 * workbook) the persons part. A threaded entry replaces a legacy note anchored
 * to the same cell, which is Excel's meaning for a converted comment.
 */
export function applySheetComments(sheet: Sheet, parts: SheetCommentParts): void {
  const comments = new Map<string, CommentRecord>();

  const commentsRelationship = findRelationship(parts.relationships, '/comments');
  if (commentsRelationship !== undefined) {
    const bytes = parts.entries.get(commentsRelationship.resolvedTarget);
    if (bytes !== undefined) {
      for (const note of parseCommentsPart(bytes, commentsRelationship.resolvedTarget)) {
        const key = keyOf(note.ref);
        if (key === null) continue;
        comments.set(key, {
          text: note.text,
          ...(note.author === null || note.author === '' ? {} : { author: note.author }),
          threaded: false,
        });
      }
    }
  }

  const threadedRelationship = findRelationship(parts.relationships, '/threadedComment');
  if (threadedRelationship !== undefined) {
    const bytes = parts.entries.get(threadedRelationship.resolvedTarget);
    if (bytes !== undefined) {
      const persons = readPersons(parts);
      for (const entry of parseThreadedCommentsPart(bytes, threadedRelationship.resolvedTarget)) {
        const key = keyOf(entry.ref);
        if (key === null) continue;
        const author = entry.personId === null ? undefined : persons.get(entry.personId);
        comments.set(key, {
          text: entry.text,
          ...(author === undefined || author === '' ? {} : { author }),
          threaded: true,
        });
      }
    }
  }

  sheet.comments.clear();
  for (const [key, comment] of comments) sheet.comments.set(key, comment);
}

/** The threaded entries of a sheet, read from the same parts `applySheetComments` uses. */
export function threadedCommentEntries(parts: SheetCommentParts): readonly ThreadedCommentEntry[] {
  const relationship = findRelationship(parts.relationships, '/threadedComment');
  if (relationship === undefined) return [];
  const bytes = parts.entries.get(relationship.resolvedTarget);
  return bytes === undefined ? [] : parseThreadedCommentsPart(bytes, relationship.resolvedTarget);
}

/** The persons a sheet's threaded comments name, for display and reporting. */
export function sheetPersons(parts: SheetCommentParts): readonly PersonEntry[] {
  return readPersonEntries(parts);
}

function readPersonEntries(parts: SheetCommentParts): PersonEntry[] {
  const relationship = findRelationship(parts.workbookRelationships, '/person');
  if (relationship === undefined) return [];
  const bytes = parts.entries.get(relationship.resolvedTarget);
  return bytes === undefined ? [] : parsePersonsPart(bytes, relationship.resolvedTarget);
}

function readPersons(parts: SheetCommentParts): ReadonlyMap<string, string> {
  const persons = new Map<string, string>();
  for (const person of readPersonEntries(parts)) persons.set(person.id, person.displayName);
  return persons;
}

/** `A1` -> `0:0`. A malformed ref is skipped rather than anchored at A1. */
function keyOf(ref: string): string | null {
  try {
    const address = parseCellRef(ref);
    return cellKey(address.row, address.col);
  } catch {
    return null;
  }
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
