import { XElement, nameEquals, type XName } from '../xml/xelement.js';
import { PT_NS, R_NS, W, WP } from './names.js';
import { irHashCompute, type IrHash } from './ir-hash.js';

const PT_INSERT_NS = 'http://powertools.codeplex.com/documentbuilder/2011/insert';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

export type RelResolver = (relId: string) => string;

interface CleanNode {
  readonly name: XName;
  readonly attrs: Array<readonly [XName, string]>;
  readonly children: Array<CleanNode | string>;
}

// TODO(parity): verify against C#-dumped canonical goldens.
export function canonicalizeXml(element: XElement, relResolver?: RelResolver): Uint8Array {
  return new TextEncoder().encode(serialize(clean(element, relResolver)));
}

export function canonicalHash(element: XElement, relResolver?: RelResolver): IrHash {
  return irHashCompute(canonicalizeXml(element, relResolver));
}

export function syntheticElement(
  local: string,
  children: ReadonlyArray<XElement> = [],
): XElement {
  return new XElement({ ns: '', local }, new Map(), [...children]);
}

function clean(element: XElement, relResolver?: RelResolver): CleanNode {
  const attrs = [...element.attributes()]
    .filter(([name]) => !shouldStripAttribute(name, element))
    .map(([name, value]) => [name, name.ns === R_NS && relResolver ? relResolver(value) : value] as const)
    // Ordinal comparison (C# StringComparer.Ordinal) — localeCompare is
    // locale-dependent and diverges from ordinal on non-ASCII.
    .sort(([a], [b]) =>
      a.ns < b.ns ? -1 : a.ns > b.ns ? 1 : a.local < b.local ? -1 : a.local > b.local ? 1 : 0,
    );

  const children: Array<CleanNode | string> = [];
  for (const child of element.children) {
    if (typeof child === 'string') {
      children.push(child);
    } else if (!nameEquals(child.name, W.proofErr) && !nameEquals(child.name, W.noProof)) {
      children.push(clean(child, relResolver));
    }
  }
  return { name: element.name, attrs, children };
}

function shouldStripAttribute(name: XName, owner: XElement): boolean {
  if (name.ns === XMLNS_NS || (name.ns === '' && name.local === 'xmlns')) return true;
  if (name.ns === PT_NS || name.ns === PT_INSERT_NS) return true;
  if (name.local.startsWith('rsid')) return true;
  if (name.local === 'id' && name.ns === '' && nameEquals(owner.name, WP.docPr)) return true;
  return false;
}

const escapeText = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeAttr = (text: string): string =>
  escapeText(text).replace(/"/g, '&quot;').replace(/\r/g, '&#xD;').replace(/\n/g, '&#xA;');

const qname = (name: XName): string => {
  if (name.ns === '') return name.local;
  if (name.ns === W.document.ns) return `w:${name.local}`;
  if (name.ns === R_NS) return `r:${name.local}`;
  if (name.ns === XML_NS) return `xml:${name.local}`;
  if (name.ns === PT_NS) return `pt14:${name.local}`;
  return `n${hashNamespace(name.ns)}:${name.local}`;
};

function serialize(node: CleanNode): string {
  let s = `<${qname(node.name)}`;
  for (const [name, value] of node.attrs) s += ` ${qname(name)}="${escapeAttr(value)}"`;
  if (node.children.length === 0) return `${s} />`;
  s += '>';
  for (const child of node.children) s += typeof child === 'string' ? escapeText(child) : serialize(child);
  return `${s}</${qname(node.name)}>`;
}

function hashNamespace(ns: string): string {
  let h = 0;
  for (let i = 0; i < ns.length; i++) h = (h * 31 + ns.charCodeAt(i)) >>> 0;
  return h.toString(36);
}
