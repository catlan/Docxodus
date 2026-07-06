import { XElement, type XAttributeInfo, type XName } from './xelement.js';

export interface WriteXmlPartOptions {
  readonly includeDeclaration?: boolean;
}

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

export function writeXmlPart(root: XElement, options: WriteXmlPartOptions = {}): string {
  const includeDeclaration = options.includeDeclaration ?? true;
  const prolog = includeDeclaration ? root.documentProlog ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' : '';
  return `${prolog}${writeElement(root, new NamespaceScope(null))}`;
}

class NamespaceScope {
  private readonly bindings = new Map<string, string>();

  constructor(private readonly parent: NamespaceScope | null) {}

  child(): NamespaceScope {
    return new NamespaceScope(this);
  }

  declare(prefix: string, ns: string): void {
    this.bindings.set(prefix, ns);
  }

  lookupPrefix(prefix: string): string | null {
    if (this.bindings.has(prefix)) return this.bindings.get(prefix)!;
    return this.parent?.lookupPrefix(prefix) ?? null;
  }

  prefixFor(ns: string): string | null {
    // A candidate binding only counts if it still RESOLVES to the namespace from THIS scope —
    // an ancestor's default declaration can be shadowed by a nearer re-declaration
    // (<graphic xmlns="a"> … <pic xmlns="pic"> … a-ns child must NOT reuse the shadowed default).
    for (let scope: NamespaceScope | null = this; scope; scope = scope.parent) {
      for (const [prefix, uri] of scope.bindings) {
        if (uri === ns && this.lookupPrefix(prefix) === ns) return prefix;
      }
    }
    return null;
  }
}

/**
 * Serialize the way .NET's XDocument.Save(Stream) does with default SaveOptions —
 * XmlWriter Indent=true: two-space indent, "\n" newlines, a newline before the root
 * element (after the declaration), " />" empty-element close, and .NET's mixed-content
 * rule (once a text node is written inside an element, nothing else in THAT element is
 * indented; the flag is pushed/popped per element). The C# engine writes every imported
 * XML part through this shape (WmlComparer.MoveRelatedPartsToDestination).
 */
export function writeXmlPartFormatted(root: XElement, declaration: string): string {
  const out: string[] = [declaration];
  writeFormattedElement(root, new NamespaceScope(null), 0, { mixed: false }, out);
  return out.join('');
}

interface IndentState {
  mixed: boolean;
}

function writeFormattedElement(el: XElement, inherited: NamespaceScope, depth: number, parent: IndentState, out: string[]): void {
  const scope = inherited.child();
  const attrs = [...el.attributeInfos()];
  for (const attr of attrs) {
    if (attr.name.ns === XMLNS_NS) scope.declare(xmlnsPrefix(attr), attr.value);
  }
  if (!parent.mixed) out.push(`\n${'  '.repeat(depth)}`);
  const generated: string[] = [];
  const name = rawElementName(el, scope, generated);
  out.push(`<${name}`);
  for (const attr of attrs) out.push(` ${rawAttributeName(attr, scope)}="${escapeAttribute(attr.value)}"`);
  for (const decl of generated) out.push(decl);
  if (el.children.length === 0) {
    out.push(' />');
    return;
  }
  out.push('>');
  const state: IndentState = { mixed: false };
  let hadChildElement = false;
  for (const child of el.children) {
    if (typeof child === 'string') {
      out.push(escapeText(child));
      state.mixed = true;
    } else {
      writeFormattedElement(child, scope, depth + 1, state, out);
      hadChildElement = true;
    }
  }
  if (!state.mixed && hadChildElement) out.push(`\n${'  '.repeat(depth)}`);
  out.push(`</${name}>`);
}

function writeElement(el: XElement, inherited: NamespaceScope): string {
  const scope = inherited.child();
  const attrs = [...el.attributeInfos()];
  for (const attr of attrs) {
    if (attr.name.ns === XMLNS_NS) scope.declare(xmlnsPrefix(attr), attr.value);
  }

  const generated: string[] = [];
  const name = rawElementName(el, scope, generated);
  let result = `<${name}`;
  for (const attr of attrs) result += ` ${rawAttributeName(attr, scope)}="${escapeAttribute(attr.value)}"`;
  for (const decl of generated) result += decl;
  if (el.children.length === 0) return `${result}${el.emptyElementSpace ? ' />' : '/>'}`;

  result += '>';
  for (const child of el.children) {
    result += typeof child === 'string' ? escapeText(child) : writeElement(child, scope);
  }
  return `${result}</${name}>`;
}

function rawElementName(el: XElement, scope: NamespaceScope, generated?: string[]): string {
  if (el.name.ns === '') return el.rawName || el.name.local;
  if (el.rawName) {
    // Trust the lexical name only while its prefix still binds to the element's namespace in the
    // OUTPUT's scope — a right-sourced subtree can reference a prefix its new document never
    // declares. .NET then falls back to any bound prefix, or auto-declares a DEFAULT namespace on
    // the element itself (how an inserted <a:graphic> becomes <graphic xmlns="…"> in the oracle).
    const colon = el.rawName.indexOf(':');
    const rawPrefix = colon >= 0 ? el.rawName.slice(0, colon) : '';
    if (scope.lookupPrefix(rawPrefix) === el.name.ns) return el.rawName;
  }
  const declared = scope.prefixFor(el.name.ns);
  if (declared !== null && declared !== '') return `${declared}:${el.name.local}`;
  if (declared === '') return el.name.local;
  if (generated) {
    scope.declare('', el.name.ns);
    generated.push(` xmlns="${escapeAttribute(el.name.ns)}"`);
  }
  return el.name.local;
}

function rawAttributeName(attr: XAttributeInfo, scope: NamespaceScope): string {
  if (attr.rawName) return attr.rawName;
  if (attr.name.ns === '') return attr.name.local;
  if (attr.name.ns === XML_NS) return `xml:${attr.name.local}`;
  const declared = scope.prefixFor(attr.name.ns);
  return declared ? `${declared}:${attr.name.local}` : attr.name.local;
}

function xmlnsPrefix(attr: XAttributeInfo): string {
  return attr.rawName === 'xmlns' ? '' : attr.name.local;
}

function escapeText(value: string): string {
  let result = '';
  for (const ch of value) {
    if (ch === '&') result += '&amp;';
    else if (ch === '<') result += '&lt;';
    else if (ch === '>') result += '&gt;';
    else if (ch === '\r') result += '&#xD;';
    else result += ch;
  }
  return result;
}

function escapeAttribute(value: string): string {
  let result = '';
  for (const ch of value) {
    if (ch === '&') result += '&amp;';
    else if (ch === '<') result += '&lt;';
    else if (ch === '"') result += '&quot;';
    else if (ch === '\t') result += '&#x9;';
    else if (ch === '\n') result += '&#xA;';
    else if (ch === '\r') result += '&#xD;';
    else result += ch;
  }
  return result;
}
