// OOXML names used by the IR port (C# W.* / PtOpenXml.* equivalents).

import { xname, type XName } from '../xml/xelement.js';

export const W_NS =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export const W = {
  document: xname(W_NS, 'document'),
  body: xname(W_NS, 'body'),
  style: xname(W_NS, 'style'),
  basedOn: xname(W_NS, 'basedOn'),
  name: xname(W_NS, 'name'),
  docDefaults: xname(W_NS, 'docDefaults'),
  pPrDefault: xname(W_NS, 'pPrDefault'),
  rPrDefault: xname(W_NS, 'rPrDefault'),
  p: xname(W_NS, 'p'),
  tbl: xname(W_NS, 'tbl'),
  tblPr: xname(W_NS, 'tblPr'),
  tblGrid: xname(W_NS, 'tblGrid'),
  gridCol: xname(W_NS, 'gridCol'),
  tr: xname(W_NS, 'tr'),
  trPr: xname(W_NS, 'trPr'),
  tblPrEx: xname(W_NS, 'tblPrEx'),
  tc: xname(W_NS, 'tc'),
  tcPr: xname(W_NS, 'tcPr'),
  gridSpan: xname(W_NS, 'gridSpan'),
  vMerge: xname(W_NS, 'vMerge'),
  r: xname(W_NS, 'r'),
  rPr: xname(W_NS, 'rPr'),
  rStyle: xname(W_NS, 'rStyle'),
  b: xname(W_NS, 'b'),
  i: xname(W_NS, 'i'),
  strike: xname(W_NS, 'strike'),
  dstrike: xname(W_NS, 'dstrike'),
  caps: xname(W_NS, 'caps'),
  smallCaps: xname(W_NS, 'smallCaps'),
  vanish: xname(W_NS, 'vanish'),
  u: xname(W_NS, 'u'),
  vertAlign: xname(W_NS, 'vertAlign'),
  rFonts: xname(W_NS, 'rFonts'),
  sz: xname(W_NS, 'sz'),
  color: xname(W_NS, 'color'),
  highlight: xname(W_NS, 'highlight'),
  t: xname(W_NS, 't'),
  tab: xname(W_NS, 'tab'),
  br: xname(W_NS, 'br'),
  noBreakHyphen: xname(W_NS, 'noBreakHyphen'),
  softHyphen: xname(W_NS, 'softHyphen'),
  sym: xname(W_NS, 'sym'),
  drawing: xname(W_NS, 'drawing'),
  pict: xname(W_NS, 'pict'),
  txbxContent: xname(W_NS, 'txbxContent'),
  hyperlink: xname(W_NS, 'hyperlink'),
  fldSimple: xname(W_NS, 'fldSimple'),
  fldChar: xname(W_NS, 'fldChar'),
  instrText: xname(W_NS, 'instrText'),
  delInstrText: xname(W_NS, 'delInstrText'),
  instr: xname(W_NS, 'instr'),
  fldCharType: xname(W_NS, 'fldCharType'),
  footnoteReference: xname(W_NS, 'footnoteReference'),
  endnoteReference: xname(W_NS, 'endnoteReference'),
  lastRenderedPageBreak: xname(W_NS, 'lastRenderedPageBreak'),
  commentReference: xname(W_NS, 'commentReference'),
  bookmarkStart: xname(W_NS, 'bookmarkStart'),
  bookmarkEnd: xname(W_NS, 'bookmarkEnd'),
  commentRangeStart: xname(W_NS, 'commentRangeStart'),
  commentRangeEnd: xname(W_NS, 'commentRangeEnd'),
  proofErr: xname(W_NS, 'proofErr'),
  noProof: xname(W_NS, 'noProof'),
  sdt: xname(W_NS, 'sdt'),
  sdtContent: xname(W_NS, 'sdtContent'),
  smartTag: xname(W_NS, 'smartTag'),
  smartTagPr: xname(W_NS, 'smartTagPr'),
  pPr: xname(W_NS, 'pPr'),
  pStyle: xname(W_NS, 'pStyle'),
  jc: xname(W_NS, 'jc'),
  ind: xname(W_NS, 'ind'),
  spacing: xname(W_NS, 'spacing'),
  outlineLvl: xname(W_NS, 'outlineLvl'),
  keepNext: xname(W_NS, 'keepNext'),
  keepLines: xname(W_NS, 'keepLines'),
  pageBreakBefore: xname(W_NS, 'pageBreakBefore'),
  numPr: xname(W_NS, 'numPr'),
  numId: xname(W_NS, 'numId'),
  ilvl: xname(W_NS, 'ilvl'),
  val: xname(W_NS, 'val'),
  id: xname(W_NS, 'id'),
  styleId: xname(W_NS, 'styleId'),
  type: xname(W_NS, 'type'),
  default: xname(W_NS, 'default'),
  pgSz: xname(W_NS, 'pgSz'),
  pgMar: xname(W_NS, 'pgMar'),
  sectPr: xname(W_NS, 'sectPr'),
  ins: xname(W_NS, 'ins'),
  del: xname(W_NS, 'del'),
  delText: xname(W_NS, 'delText'),
  moveFrom: xname(W_NS, 'moveFrom'),
  moveTo: xname(W_NS, 'moveTo'),
  moveFromRangeStart: xname(W_NS, 'moveFromRangeStart'),
  moveFromRangeEnd: xname(W_NS, 'moveFromRangeEnd'),
  moveToRangeStart: xname(W_NS, 'moveToRangeStart'),
  moveToRangeEnd: xname(W_NS, 'moveToRangeEnd'),
  rPrChange: xname(W_NS, 'rPrChange'),
  pPrChange: xname(W_NS, 'pPrChange'),
  tblPrChange: xname(W_NS, 'tblPrChange'),
  tblGridChange: xname(W_NS, 'tblGridChange'),
  trPrChange: xname(W_NS, 'trPrChange'),
  tcPrChange: xname(W_NS, 'tcPrChange'),
  tblPrExChange: xname(W_NS, 'tblPrExChange'),
  sectPrChange: xname(W_NS, 'sectPrChange'),
  cellIns: xname(W_NS, 'cellIns'),
  cellDel: xname(W_NS, 'cellDel'),
  cellMerge: xname(W_NS, 'cellMerge'),
  footnote: xname(W_NS, 'footnote'),
  endnote: xname(W_NS, 'endnote'),
  headerReference: xname(W_NS, 'headerReference'),
  footerReference: xname(W_NS, 'footerReference'),
  comment: xname(W_NS, 'comment'),
  author: xname(W_NS, 'author'),
  initials: xname(W_NS, 'initials'),
  date: xname(W_NS, 'date'),
} as const satisfies Record<string, XName>;

export const R_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const R = {
  id: xname(R_NS, 'id'),
  embed: xname(R_NS, 'embed'),
} as const satisfies Record<string, XName>;

export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const REL = {
  Relationships: xname(REL_NS, 'Relationships'),
  Relationship: xname(REL_NS, 'Relationship'),
} as const satisfies Record<string, XName>;

export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const A = {
  blip: xname(A_NS, 'blip'),
} as const satisfies Record<string, XName>;

export const WP_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const WP = {
  extent: xname(WP_NS, 'extent'),
  docPr: xname(WP_NS, 'docPr'),
} as const satisfies Record<string, XName>;

/** PowerTools annotation namespace (C# PtOpenXml.pt). */
export const PT_NS = 'http://powertools.codeplex.com/2011';
export const PT = {
  unid: xname(PT_NS, 'Unid'),
} as const satisfies Record<string, XName>;
