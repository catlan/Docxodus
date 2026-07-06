// OOXML names used by the IR port (C# W.* / PtOpenXml.* equivalents).

import { xname, type XName } from '../xml/xelement.js';

export const W_NS =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export const W = {
  document: xname(W_NS, 'document'),
  body: xname(W_NS, 'body'),
  p: xname(W_NS, 'p'),
  tbl: xname(W_NS, 'tbl'),
  tr: xname(W_NS, 'tr'),
  tc: xname(W_NS, 'tc'),
  t: xname(W_NS, 't'),
  pPr: xname(W_NS, 'pPr'),
  pStyle: xname(W_NS, 'pStyle'),
  numPr: xname(W_NS, 'numPr'),
  numId: xname(W_NS, 'numId'),
  val: xname(W_NS, 'val'),
  id: xname(W_NS, 'id'),
  footnote: xname(W_NS, 'footnote'),
  endnote: xname(W_NS, 'endnote'),
} as const satisfies Record<string, XName>;

/** PowerTools annotation namespace (C# PtOpenXml.pt). */
export const PT_NS = 'http://powertools.codeplex.com/2011';
export const PT = {
  unid: xname(PT_NS, 'Unid'),
} as const satisfies Record<string, XName>;
