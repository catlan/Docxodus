import { strFromU8, strToU8 } from 'fflate';
import { cloneElement, element, type XChild } from '../xml/xclone.js';
import { nameEquals, parseXml, XElement, type XName } from '../xml/xelement.js';
import { writeXmlPart } from '../xml/xwriter.js';
import { M, PT, W } from './names.js';

export type RevisionView = 'accept' | 'raw' | 'failIfPresent';

const REVISION_PART_RE =
  /^word\/(?:document|footnotes|endnotes|styles|header\d+|footer\d+)\.xml$/i;

const DROP_ACCEPT = [
  W.customXmlDelRangeStart,
  W.customXmlDelRangeEnd,
  W.customXmlInsRangeStart,
  W.customXmlInsRangeEnd,
  W.customXmlMoveFromRangeStart,
  W.customXmlMoveFromRangeEnd,
  W.customXmlMoveToRangeStart,
  W.customXmlMoveToRangeEnd,
  W.moveFromRangeStart,
  W.moveFromRangeEnd,
  W.moveToRangeStart,
  W.moveToRangeEnd,
  W.pPrChange,
  W.rPrChange,
  W.tblPrChange,
  W.tblGridChange,
  W.tcPrChange,
  W.trPrChange,
  W.tblPrExChange,
  W.sectPrChange,
  W.numberingChange,
  W.delInstrText,
  W.delText,
  W.cellIns,
] as const;

const RSID_ATTRS = new Set([
  key(W.rsid),
  key(W.rsidDel),
  key(W.rsidP),
  key(W.rsidR),
  key(W.rsidRDefault),
  key(W.rsidRPr),
  key(W.rsidSect),
  key(W.rsidTr),
]);

export function applyRevisionViewToParts(
  parts: ReadonlyMap<string, Uint8Array> | Record<string, Uint8Array>,
  view: RevisionView = 'accept',
): ReadonlyMap<string, Uint8Array> | Record<string, Uint8Array> {
  if (view === 'raw') return parts;
  const out = new Map<string, Uint8Array>();
  for (const [name, bytes] of partEntries(parts)) {
    if (REVISION_PART_RE.test(name)) {
      const root = parseXml(strFromU8(bytes));
      const rev = firstRevision(root);
      if (view === 'failIfPresent') {
        if (rev) throw new Error(`IrReader requires revision-free XML; found ${rev.rawName}.`);
        out.set(name, bytes);
      } else if (!rev) {
        out.set(name, bytes);
      } else {
        out.set(name, strToU8(writeXmlPart(cloneElement(acceptRevisionsForRoot(root), {
          documentProlog: '<?xml version="1.0" encoding="utf-8"?>',
        }))));
      }
    } else {
      out.set(name, bytes);
    }
  }
  return out;
}

export function acceptRevisionsForRoot(root: XElement): XElement {
  return stripPtArtifacts(
    pruneEmptyNumPr(
      addEmptyParagraphsToEmptyCells(
        acceptDeletedCells(
          acceptAllOtherRevisions(
            acceptMoveFromMoveTo(
              fixUpDeletedOrInsertedFieldCodes(removeRsid(root)),
            ),
          ),
        ),
      ),
    ),
  );
}

function acceptMoveFromMoveTo(node: XElement): XElement {
  return mapElement(node, (el) => {
    if (nameEquals(el.name, W.moveTo)) return acceptChildren(el, acceptMoveFromMoveTo);
    if (nameEquals(el.name, W.moveFrom)) return [];
    return null;
  });
}

function acceptAllOtherRevisions(node: XElement): XElement {
  return mapElement(node, (el) => {
    if (nameEquals(el.name, W.ins)) return acceptChildren(el, acceptAllOtherRevisions);
    if (DROP_ACCEPT.some((n) => nameEquals(el.name, n))) return [];
    if (nameEquals(el.name, M.f) && [...el.elements(M.fPr)].some((fPr) => [...fPr.elements(M.ctrlPr)].some((ctrl) => ctrl.element(W.del)))) return [];
    if (nameEquals(el.name, W.tr) && [...el.elements(W.trPr)].some((trPr) => trPr.element(W.del))) return [];
    if (nameEquals(el.name, W.tbl)) {
      const rows = [...el.elements(W.tr)];
      if (rows.length > 0 && rows.every((tr) => [...tr.elements(W.trPr)].some((trPr) => trPr.element(W.del)))) return [];
    }
    if (nameEquals(el.name, W.del)) return [];
    if (nameEquals(el.name, W.cellMerge) && el.parent && nameEquals(el.parent.name, W.tcPr)) {
      const vMerge = el.attribute(W.vMerge);
      if (vMerge === 'rest') return [element(W.vMerge, [{ name: W.val, value: 'restart', prefix: 'w', rawName: 'w:val' }])];
      if (vMerge === 'cont') return [element(W.vMerge, [{ name: W.val, value: 'continue', prefix: 'w', rawName: 'w:val' }])];
    }
    if (nameEquals(el.name, W.hyperlink)) {
      const transformed = cloneElement(el, { children: acceptChildren(el, acceptAllOtherRevisions) });
      return hasRunContent(transformed) ? [transformed] : [];
    }
    return null;
  });
}

function fixUpDeletedOrInsertedFieldCodes(node: XElement): XElement {
  return mapElement(node, (el) => {
    if (!nameEquals(el.name, W.p)) return null;
    const elementChildren = el.children.filter((c): c is XElement => c instanceof XElement);
    const groups: Array<{ key: number; items: XElement[] }> = [];
    for (const child of elementChildren) {
      const keyForChild =
        nameEquals(child.name, W.del) && [...child.elements(W.r)].some((r) => r.element(W.fldChar)) ? 2 :
        nameEquals(child.name, W.ins) && [...child.elements(W.r)].some((r) => r.element(W.fldChar)) ? 3 :
        nameEquals(child.name, W.r) && child.element(W.instrText) ? 4 :
        1;
      const last = groups.at(-1);
      if (last && last.key === keyForChild) last.items.push(child);
      else groups.push({ key: keyForChild, items: [child] });
    }
    if (groups.length === 0) return null;
    const children: XChild[] = [];
    for (let i = 0; i < groups.length; i++) {
      const g = groups[i]!;
      if (g.key !== 4 || i === 0 || i === groups.length - 1) {
        children.push(...g.items.map(fixUpDeletedOrInsertedFieldCodes));
      } else if (groups[i - 1]!.key === 2 && groups[i + 1]!.key === 2) {
        children.push(element(W.del, [], g.items.map(transformInstrTextToDelInstrText)));
      } else if (groups[i - 1]!.key === 3 && groups[i + 1]!.key === 3) {
        children.push(element(W.ins, [], g.items.map(fixUpDeletedOrInsertedFieldCodes)));
      } else {
        children.push(...g.items.map(fixUpDeletedOrInsertedFieldCodes));
      }
    }
    return [cloneElement(el, { children })];
  });
}

function transformInstrTextToDelInstrText(node: XElement): XElement {
  return mapElement(node, (el) => nameEquals(el.name, W.instrText)
    ? [cloneElement(el, { name: W.delInstrText, prefix: 'w', rawName: 'w:delInstrText' })]
    : null);
}

function acceptDeletedCells(root: XElement): XElement {
  return mapElement(root, (el) => {
    if (!nameEquals(el.name, W.tr)) return null;
    const children: XChild[] = [];
    const cells = [...el.elements(W.tc)];
    for (const child of el.children) {
      if (!(child instanceof XElement) || !nameEquals(child.name, W.tc)) {
        children.push(child);
        continue;
      }
      const idx = cells.indexOf(child);
      if (hasDescendant(child, W.cellDel)) continue;
      let span = gridSpan(child);
      let j = idx + 1;
      while (j < cells.length && hasDescendant(cells[j]!, W.cellDel)) {
        span += gridSpan(cells[j]!);
        j++;
      }
      children.push(span === gridSpan(child) ? acceptDeletedCells(child) : widenGridSpan(acceptDeletedCells(child), span));
    }
    return [cloneElement(el, { children })];
  });
}

function addEmptyParagraphsToEmptyCells(root: XElement): XElement {
  return mapElement(root, (el) => {
    if (!nameEquals(el.name, W.tc)) return null;
    const transformed = acceptChildren(el, addEmptyParagraphsToEmptyCells);
    if (transformed.some((c) => c instanceof XElement && !nameEquals(c.name, W.tcPr))) {
      return [cloneElement(el, { children: transformed })];
    }
    return [cloneElement(el, { children: [...transformed, element(W.p)] })];
  });
}

function removeRsid(root: XElement): XElement {
  return mapElement(root, (el) => {
    if (nameEquals(el.name, W.rsid)) return [];
    return [cloneElement(el, {
      attributes: [...el.attributeInfos()].filter((a) => !RSID_ATTRS.has(key(a.name))),
      children: acceptChildren(el, removeRsid),
    })];
  });
}

function stripPtArtifacts(root: XElement): XElement {
  return mapElement(root, (el) => [cloneElement(el, {
    attributes: [...el.attributeInfos()].filter((a) => !nameEquals(a.name, PT.unid) && a.name.local !== 'RunIds'),
    children: acceptChildren(el, stripPtArtifacts),
  })]);
}

function pruneEmptyNumPr(root: XElement): XElement {
  return mapElement(root, (el) => {
    if (nameEquals(el.name, W.numPr) && !el.elements().next().value) return [];
    return null;
  });
}

function mapElement(root: XElement, map: (el: XElement) => XChild[] | null): XElement {
  const first = mapNode(root, map).find((c): c is XElement => c instanceof XElement);
  return first ?? element(root.name);
}

function mapNode(node: XElement, map: (el: XElement) => XChild[] | null): XChild[] {
  const mapped = map(node);
  if (mapped !== null) return mapped;
  const children: XChild[] = [];
  for (const child of node.children) {
    if (typeof child === 'string') children.push(child);
    else children.push(...mapNode(child, map));
  }
  return [cloneElement(node, { children })];
}

function acceptChildren(el: XElement, fn: (el: XElement) => XElement | XChild[]): XChild[] {
  const out: XChild[] = [];
  for (const child of el.children) {
    if (typeof child === 'string') out.push(child);
    else {
      const transformed = fn(child);
      if (Array.isArray(transformed)) out.push(...transformed);
      else out.push(transformed);
    }
  }
  return out;
}

function firstRevision(root: XElement): XElement | null {
  for (const el of root.descendants()) {
    if (
      nameEquals(el.name, W.ins) || nameEquals(el.name, W.del) ||
      nameEquals(el.name, W.moveFrom) || nameEquals(el.name, W.moveTo) ||
      DROP_ACCEPT.some((n) => nameEquals(el.name, n))
    ) return el;
  }
  return null;
}

function hasRunContent(el: XElement): boolean {
  return [...el.elements()].some((e) =>
    nameEquals(e.name, W.r) || nameEquals(e.name, W.smartTag) || nameEquals(e.name, W.ins) ||
    nameEquals(e.name, W.del) || nameEquals(e.name, W.hyperlink) || nameEquals(e.name, W.fldSimple) ||
    nameEquals(e.name, W.sdt));
}

function hasDescendant(el: XElement, name: XName): boolean {
  if (nameEquals(el.name, name)) return true;
  for (const d of el.descendants(name)) return true;
  return false;
}

function gridSpan(tc: XElement): number {
  const raw = tc.element(W.tcPr)?.element(W.gridSpan)?.attribute(W.val);
  const value = raw ? Number.parseInt(raw, 10) : 1;
  return Number.isFinite(value) && value > 0 ? value : 1;
}

function widenGridSpan(tc: XElement, span: number): XElement {
  const tcPr = tc.element(W.tcPr) ?? element(W.tcPr);
  const newGrid = element(W.gridSpan, [{ name: W.val, value: String(span), prefix: 'w', rawName: 'w:val' }]);
  const newTcPr = cloneElement(tcPr, {
    children: [newGrid, ...tcPr.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, W.gridSpan)))],
  });
  return tc.element(W.tcPr)
    ? cloneElement(tc, { children: tc.children.map((c) => c === tcPr ? newTcPr : c) })
    : cloneElement(tc, { children: [newTcPr, ...tc.children] });
}

function key(name: { readonly ns: string; readonly local: string }): string {
  return `${name.ns}\0${name.local}`;
}

function partEntries(parts: ReadonlyMap<string, Uint8Array> | Record<string, Uint8Array>): Array<readonly [string, Uint8Array]> {
  return typeof (parts as ReadonlyMap<string, Uint8Array>).entries === 'function'
    ? [...(parts as ReadonlyMap<string, Uint8Array>).entries()]
    : Object.entries(parts as Record<string, Uint8Array>);
}
