import { XElement, nameEquals, type XAttributeInfo, type XName } from './xelement.js';

export type XChild = XElement | string;

export interface XElementCloneOptions {
  readonly name?: XName;
  readonly attributes?: readonly XAttributeInfo[];
  readonly children?: readonly XChild[];
  readonly documentProlog?: string | null;
  readonly rawName?: string;
  readonly prefix?: string;
  readonly emptyElementSpace?: boolean;
}

export function cloneElement(el: XElement, options: XElementCloneOptions = {}): XElement {
  const name = options.name ?? el.name;
  const attrs = options.attributes ?? [...el.attributeInfos()];
  const children = options.children ?? el.children.map((c) => (typeof c === 'string' ? c : cloneElement(c)));
  const ctorOptions: {
    attributes: readonly XAttributeInfo[];
    prefix: string;
    rawName?: string;
    isSelfClosing: boolean;
    emptyElementSpace: boolean;
    documentProlog: string | null;
  } = {
    attributes: attrs,
    prefix: options.prefix ?? (nameEquals(name, el.name) ? el.prefix : ''),
    isSelfClosing: children.length === 0,
    emptyElementSpace: options.emptyElementSpace ?? (children.length === 0 ? el.emptyElementSpace : false),
    documentProlog: options.documentProlog ?? el.documentProlog,
  };
  const rawName = options.rawName ?? (nameEquals(name, el.name) ? el.rawName : undefined);
  if (rawName !== undefined) ctorOptions.rawName = rawName;
  return new XElement(name, attrsToMap(attrs), [...children], ctorOptions);
}

export function element(name: XName, attrs: readonly XAttributeInfo[] = [], children: readonly XChild[] = []): XElement {
  return new XElement(name, attrsToMap(attrs), [...children], {
    attributes: attrs,
    prefix: prefixFor(name.ns),
    rawName: rawNameFor(name),
    isSelfClosing: children.length === 0,
  });
}

export function attr(name: XName, value: string | number): XAttributeInfo {
  const prefix = prefixFor(name.ns);
  return { name, value: String(value), prefix, rawName: prefix ? `${prefix}:${name.local}` : name.local };
}

export function attrsToMap(attrs: readonly XAttributeInfo[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const a of attrs) map.set(`${a.name.ns}\0${a.name.local}`, a.value);
  return map;
}

export function withoutAttributes(el: XElement, predicate: (attr: XAttributeInfo) => boolean): XElement {
  return cloneElement(el, { attributes: [...el.attributeInfos()].filter((a) => !predicate(a)) });
}

export function replaceName(el: XElement, name: XName): XElement {
  return cloneElement(el, { name, prefix: prefixFor(name.ns), rawName: rawNameFor(name) });
}

export function childrenElements(el: XElement, name?: XName): XElement[] {
  return [...el.elements(name)];
}

export function descendantsAndSelf(el: XElement): XElement[] {
  const out: XElement[] = [el];
  for (const child of el.children) if (child instanceof XElement) out.push(...descendantsAndSelf(child));
  return out;
}

function prefixFor(ns: string): string {
  switch (ns) {
    case 'http://schemas.openxmlformats.org/wordprocessingml/2006/main': return 'w';
    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships': return 'r';
    case 'http://www.w3.org/XML/1998/namespace': return 'xml';
    case 'http://powertools.codeplex.com/2011': return 'pt';
    default: return '';
  }
}

function rawNameFor(name: XName): string {
  const prefix = prefixFor(name.ns);
  return prefix ? `${prefix}:${name.local}` : name.local;
}
