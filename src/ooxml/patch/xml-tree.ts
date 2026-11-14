import { SaxesParser } from 'saxes';
import { assertSafeXml, escapeXmlAttribute, escapeXmlText } from '../xml';

export interface XmlAttribute {
  readonly name: string;
  readonly value: string;
}

export interface XmlElement {
  readonly type: 'element';
  readonly name: string;
  attributes: XmlAttribute[];
  children: XmlNode[];
}

export interface XmlText {
  readonly type: 'text';
  readonly text: string;
}

export interface XmlCdata {
  readonly type: 'cdata';
  readonly text: string;
}

export interface XmlComment {
  readonly type: 'comment';
  readonly text: string;
}

export interface XmlProcessingInstruction {
  readonly type: 'pi';
  readonly target: string;
  readonly data: string;
}

export type XmlNode = XmlText | XmlCdata | XmlComment | XmlProcessingInstruction | XmlElement;

export interface XmlDocument {
  readonly prolog: XmlNode[];
  readonly root: XmlElement;
  readonly epilog: XmlNode[];
}

export function parseXmlTree(xml: string, partName: string): XmlDocument {
  assertSafeXml(xml, partName);
  const parser = new SaxesParser({ xmlns: false, position: false });
  const prolog: XmlNode[] = [];
  const epilog: XmlNode[] = [];
  const stack: XmlElement[] = [];
  let root: XmlElement | null = null;

  const append = (node: XmlNode): void => {
    const parent = stack.at(-1);
    if (parent !== undefined) {
      parent.children.push(node);
    } else if (root === null) {
      prolog.push(node);
    } else {
      epilog.push(node);
    }
  };

  parser.on('opentag', (tag) => {
    const element: XmlElement = {
      type: 'element',
      name: tag.name,
      attributes: (Object.entries(tag.attributes) as [string, unknown][]).map(([name, value]) => ({
        name,
        value: String(value),
      })),
      children: [],
    };
    if (root === null) {
      root = element;
    } else {
      append(element);
    }
    stack.push(element);
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.on('text', (text) => append({ type: 'text', text }));
  parser.on('cdata', (text) => append({ type: 'cdata', text }));
  parser.on('comment', (text) => append({ type: 'comment', text }));
  parser.on('xmldecl', (declaration) => {
    const data = [
      declaration.version === undefined ? null : `version="${declaration.version}"`,
      declaration.encoding === undefined ? null : `encoding="${declaration.encoding}"`,
      declaration.standalone === undefined ? null : `standalone="${declaration.standalone}"`,
    ]
      .filter((part): part is string => part !== null)
      .join(' ');
    prolog.push({ type: 'pi', target: 'xml', data });
  });
  parser.on('processinginstruction', (instruction) =>
    append({ type: 'pi', target: instruction.target, data: instruction.body }),
  );
  parser.on('error', (error) => {
    throw new Error(`Could not parse ${partName}: ${error.message}`);
  });

  try {
    parser.write(xml).close();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Could not parse ${partName}: ${message}`);
  }
  if (root === null) throw new Error(`Could not parse ${partName}: no root element`);
  return { prolog, root, epilog };
}

export function serializeXmlDocument(document: XmlDocument): string {
  return [
    ...document.prolog.map(serializeNode),
    serializeNode(document.root),
    ...document.epilog.map(serializeNode),
  ].join('');
}

export function serializeNode(node: XmlNode): string {
  switch (node.type) {
    case 'text':
      return escapeXmlText(node.text);
    case 'cdata':
      return `<![CDATA[${node.text.replaceAll(']]>', ']]]]><![CDATA[>')}]]>`;
    case 'comment':
      return `<!--${node.text}-->`;
    case 'pi':
      return `<?${node.target}${node.data === '' ? '' : ` ${node.data}`}?>`;
    case 'element': {
      const attributes = node.attributes
        .map((attribute) => ` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`)
        .join('');
      if (node.children.length === 0) return `<${node.name}${attributes}/>`;
      return `<${node.name}${attributes}>${node.children
        .map(serializeNode)
        .join('')}</${node.name}>`;
    }
  }
}

export function localName(element: XmlElement): string {
  const separator = element.name.lastIndexOf(':');
  return separator < 0 ? element.name : element.name.slice(separator + 1);
}

export function elementPrefix(element: XmlElement): string {
  const separator = element.name.lastIndexOf(':');
  return separator < 0 ? '' : element.name.slice(0, separator);
}

export function findElement(element: XmlElement, local: string): XmlElement | undefined {
  return element.children.find(
    (child): child is XmlElement => child.type === 'element' && localName(child) === local,
  );
}

export function findElements(element: XmlElement, local: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && localName(child) === local,
  );
}

export function getAttribute(element: XmlElement, local: string): string | undefined {
  return element.attributes.find((attribute) => attributeName(attribute.name) === local)?.value;
}

export function setAttribute(element: XmlElement, name: string, value: string): void {
  const index = element.attributes.findIndex((attribute) => attribute.name === name);
  if (index >= 0) {
    element.attributes[index] = { name, value };
  } else {
    element.attributes.push({ name, value });
  }
}

export function removeAttribute(element: XmlElement, local: string): void {
  element.attributes = element.attributes.filter(
    (attribute) => attributeName(attribute.name) !== local,
  );
}

export function textContent(element: XmlElement): string {
  let result = '';
  for (const child of element.children) {
    if (child.type === 'text' || child.type === 'cdata') result += child.text;
    else if (child.type === 'element') result += textContent(child);
  }
  return result;
}

export function makeText(text: string): XmlText {
  return { type: 'text', text };
}

export function makeElement(name: string, attributes: XmlAttribute[] = []): XmlElement {
  return { type: 'element', name, attributes, children: [] };
}

export function qualifiedName(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function attributeName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}
