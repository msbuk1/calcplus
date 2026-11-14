import type { Relationship } from './types';
import {
  PACKAGE_RELS_NS,
  RELATIONSHIPS_NS,
  decodeXml,
  parseXml,
  resolvePartPath,
  tagLocalName,
  attributeValue,
} from './xml';
import { WorkbookOpenError } from './errors';

const RELATIONSHIP_TYPES = {
  officeDocument: '/officeDocument',
  worksheet: '/worksheet',
  styles: '/styles',
  sharedStrings: '/sharedStrings',
  theme: '/theme',
  calcChain: '/calcChain',
  externalLink: '/externalLink',
} as const;

export function relationshipTypeEndsWith(type: string, suffix: string): boolean {
  return type.endsWith(suffix);
}

export function parseRelationships(
  bytes: Uint8Array,
  partName: string,
  basePartPath: string,
): Relationship[] {
  const xml = decodeXml(bytes);
  const relationships: Relationship[] = [];
  let baseDepth = 0;
  parseXml(
    xml,
    {
      onOpenTag(tag) {
        const local = tag.local;
        if (local === 'Relationships') {
          baseDepth = 1;
          return;
        }
        if (baseDepth === 0) return;
        if (local !== 'Relationship') return;
        const id = attributeValue(tag, 'Id') ?? '';
        const type = attributeValue(tag, 'Type') ?? '';
        const target = attributeValue(tag, 'Target') ?? '';
        const targetMode =
          attributeValue(tag, 'TargetMode') === 'External' ? 'External' : 'Internal';
        if (id === '' || type === '' || target === '') {
          throw new WorkbookOpenError('corrupt', `Malformed relationship in ${partName}`);
        }
        relationships.push({
          id,
          type,
          target,
          targetMode,
          resolvedTarget:
            targetMode === 'Internal' ? resolvePartPath(basePartPath, target) : target,
        } as Relationship);
      },
    },
    partName,
  );
  return relationships;
}

export function findRelationship(
  relationships: readonly Relationship[],
  suffix: string,
): Relationship | undefined {
  return relationships.find((relationship) => relationshipTypeEndsWith(relationship.type, suffix));
}

export function findRelationshipById(
  relationships: readonly Relationship[],
  id: string,
): Relationship | undefined {
  return relationships.find((relationship) => relationship.id === id);
}

export function isExternalLinkRelationship(relationship: Relationship): boolean {
  return relationshipTypeEndsWith(relationship.type, RELATIONSHIP_TYPES.externalLink);
}

export const RELATIONSHIP_SUFFIXES = RELATIONSHIP_TYPES;

export function packageRelationshipNamespace(uri: string): boolean {
  return uri === PACKAGE_RELS_NS || uri === RELATIONSHIPS_NS;
}

export function relationshipTagLocalName(name: string): string {
  return tagLocalName(name);
}
