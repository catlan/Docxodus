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
    for (const [prefix, uri] of this.bindings) if (uri === ns) return prefix;
    return this.parent?.prefixFor(ns) ?? null;
  }
}

function writeElement(el: XElement, inherited: NamespaceScope): string {
  const scope = inherited.child();
  const attrs = [...el.attributeInfos()];
  for (const attr of attrs) {
    if (attr.name.ns === XMLNS_NS) scope.declare(xmlnsPrefix(attr), attr.value);
  }

  const name = rawElementName(el, scope);
  let result = `<${name}`;
  for (const attr of attrs) result += ` ${rawAttributeName(attr, scope)}="${escapeAttribute(attr.value)}"`;
  if (el.children.length === 0) return `${result}${el.emptyElementSpace ? ' />' : '/>'}`;

  result += '>';
  for (const child of el.children) {
    result += typeof child === 'string' ? escapeText(child) : writeElement(child, scope);
  }
  return `${result}</${name}>`;
}

function rawElementName(el: XElement, scope: NamespaceScope): string {
  if (el.rawName) return el.rawName;
  if (el.name.ns === '') return el.name.local;
  const declared = scope.prefixFor(el.name.ns);
  if (declared !== null && declared !== '') return `${declared}:${el.name.local}`;
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
