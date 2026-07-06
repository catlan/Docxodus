import { nameEquals, type XElement } from '../xml/xelement.js';
import { W } from './names.js';
import type { IrAnchorKind } from './ir-anchor.js';

export function kindFor(pOrBlock: XElement, stylesRoot: XElement | null): IrAnchorKind | null {
  const n = pOrBlock.name;
  if (nameEquals(n, W.p)) return isHeading(pOrBlock) ? 'h' : isListItem(pOrBlock, stylesRoot) ? 'li' : 'p';
  if (nameEquals(n, W.tbl)) return 'tbl';
  if (nameEquals(n, W.tr)) return 'tr';
  if (nameEquals(n, W.tc)) return 'tc';
  if (nameEquals(n, W.sectPr)) return 'sec';
  if (nameEquals(n, W.footnote)) return 'fn';
  if (nameEquals(n, W.endnote)) return 'en';
  return null;
}

export function isHeading(p: XElement): boolean {
  const styleId = p.element(W.pPr)?.element(W.pStyle)?.attribute(W.val);
  if (!styleId) return false;
  return (
    styleId.toLowerCase().startsWith('heading') ||
    styleId.toLowerCase() === 'title' ||
    styleId.toLowerCase() === 'subtitle'
  );
}

export function isListItem(p: XElement, stylesRoot: XElement | null): boolean {
  if (p.element(W.pPr)?.element(W.numPr)) return true;
  const styleId = p.element(W.pPr)?.element(W.pStyle)?.attribute(W.val);
  if (!styleId || !stylesRoot) return false;

  const visited = new Set<string>();
  let current: string | null = styleId;
  for (let i = 0; i < 16 && current !== null; i++) {
    if (visited.has(current)) return false;
    visited.add(current);
    const style = [...stylesRoot.elements(W.style)].find((s) => s.attribute(W.styleId) === current);
    if (!style) return false;
    if (style.element(W.pPr)?.element(W.numPr)) return true;
    current = style.element(W.basedOn)?.attribute(W.val) ?? null;
  }
  return false;
}
