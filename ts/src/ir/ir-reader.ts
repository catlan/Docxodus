import { unzipSync, strFromU8 } from 'fflate';
import { nameEquals, parseXml, type XElement, type XName } from '../xml/xelement.js';
import { anchorToString, irAnchor, type IrAnchor, type IrAnchorKind } from './ir-anchor.js';
import {
  IrContentHashBuilder,
  SENTINEL_COLUMN_BREAK,
  SENTINEL_ENDNOTE_REF,
  SENTINEL_FOOTNOTE_REF,
  SENTINEL_HYPERLINK,
  SENTINEL_HYPERLINK_TARGET_END,
  SENTINEL_IMAGE,
  SENTINEL_LINE_BREAK,
  SENTINEL_OPAQUE,
  SENTINEL_PAGE_BREAK,
  SENTINEL_TAB,
  SENTINEL_TEXTBOX,
  STRUCTURE_CELL,
  STRUCTURE_ROW,
} from './ir-content-hash-builder.js';
import { canonicalHash, syntheticElement } from './ir-canonicalize.js';
import type { IrBlock, IrCell, IrOpaqueBlock, IrParagraph, IrRow, IrSectionBreak, IrTable } from './ir-blocks.js';
import {
  type IrBreakKind,
  type IrJustification,
  type IrLineSpacing,
  type IrParaFormat,
  type IrRunFormat,
  type IrSectionFormat,
  type IrUnderline,
  type IrUnderlineKind,
  type IrVMerge,
  type IrListInfo,
} from './ir-formats.js';
import type { IrInline } from './ir-inlines.js';
import { irHashCompute, irHashToBytes, type IrHash } from './ir-hash.js';
import { emptyIrProvenance } from './ir-provenance.js';
import { kindFor, isListItem } from './kind-for.js';
import { A, R, REL, W, WP } from './names.js';
import { assignToAllElementsDeterministic, type UnidMap } from './unid-helper.js';

export interface IrScope {
  readonly name: string;
  readonly blocks: ReadonlyArray<IrBlock>;
  readonly partUri: string;
}

export interface IrDocument {
  readonly body: IrScope;
  readonly headers: ReadonlyArray<never>;
  readonly footers: ReadonlyArray<never>;
  readonly footnotes: ReadonlyMap<string, IrScope>;
  readonly endnotes: ReadonlyMap<string, IrScope>;
  readonly comments: ReadonlyArray<never>;
  readonly styles: IrStyleRegistry;
  readonly numbering: IrNumberingRegistry;
  readonly anchorIndex: ReadonlyMap<string, IrBlock>;
}

export interface IrReaderOptions {
  readonly retainSources?: boolean;
}

interface IrStyle {
  readonly id: string;
  readonly basedOn: string | null;
  readonly pPr: XElement | null;
  readonly rPr: XElement | null;
}

export interface IrStyleRegistry {
  readonly styles: ReadonlyMap<string, IrStyle>;
  readonly defaultParagraphStyleId: string | null;
  readonly docDefaultsPPr: XElement | null;
  readonly docDefaultsRPr: XElement | null;
}

export interface IrNumberingRegistry {
  readonly nums: ReadonlyMap<number, { readonly abstractNumId: number; readonly startOverrides: ReadonlyMap<number, number> }>;
  readonly abstractNums: ReadonlyMap<number, { readonly levels: ReadonlyMap<number, { readonly numberFormat: string }> }>;
}

interface ReadContext {
  readonly scope: string;
  readonly unids: UnidMap;
  readonly stylesRoot: XElement | null;
  readonly styles: IrStyleRegistry;
  readonly numbering: IrNumberingRegistry;
  readonly hyperlinks: ReadonlyMap<string, string>;
}

const EMPTY_HASH = irHashCompute(new Uint8Array());
const EMPTY_UNMODELED_DIGEST = canonicalHash(syntheticElement('unmodeled'));
const EMPTY_SHELL_DIGEST = canonicalHash(syntheticElement('shell'));
const EMPTY_PPR_PROPS_DIGEST = canonicalHash(syntheticElement('ppr-props'));

const nameKey = (name: XName): string => `${name.ns}\0${name.local}`;

const PPR_CONSUMED = new Set([
  W.pStyle,
  W.jc,
  W.ind,
  W.spacing,
  W.outlineLvl,
  W.keepNext,
  W.keepLines,
  W.pageBreakBefore,
  W.numPr,
].map(nameKey));

const RPR_CONSUMED = new Set([
  W.rStyle,
  W.b,
  W.i,
  W.strike,
  W.dstrike,
  W.caps,
  W.smallCaps,
  W.vanish,
  W.u,
  W.sz,
  W.color,
  W.highlight,
].map(nameKey));

const SECTPR_CONSUMED = new Set([W.pgSz, W.pgMar, W.type].map(nameKey));

export function readIrDocument(
  partsOrDocx: ReadonlyMap<string, Uint8Array> | Uint8Array,
  _options: IrReaderOptions = {},
): IrDocument {
  const parts = partsOrDocx instanceof Uint8Array ? unzipSync(partsOrDocx) : partsOrDocx;
  const documentXml = getPartText(parts, 'word/document.xml');
  if (documentXml === null) throw new Error('Document has no word/document.xml part.');
  const root = parseXml(documentXml);
  const body = root.element(W.body);
  if (!body) throw new Error('Document has no w:body element.');
  for (const el of body.descendants()) {
    if (nameEquals(el.name, W.ins) || nameEquals(el.name, W.del)) {
      throw new Error('Stage A IrReader requires revision-free body XML; found w:ins/w:del.');
    }
  }

  const stylesRoot = parseOptionalXml(getPartText(parts, 'word/styles.xml'));
  const relsRoot = parseOptionalXml(getPartText(parts, 'word/_rels/document.xml.rels'));
  const unids = assignToAllElementsDeterministic(root);
  const styles = buildStyleRegistry(stylesRoot);
  const numbering = buildEmptyNumberingRegistry();
  const ctx: ReadContext = {
    scope: 'body',
    unids,
    stylesRoot,
    styles,
    numbering,
    hyperlinks: buildHyperlinkMap(relsRoot),
  };

  const blocks: IrBlock[] = [];
  for (const child of body.elements()) appendBlocks(child, ctx, blocks);
  const anchorIndex = new Map<string, IrBlock>();
  for (const block of blocks) indexBlock(block, anchorIndex);

  return {
    body: { name: 'body', blocks, partUri: '/word/document.xml' },
    headers: [],
    footers: [],
    footnotes: new Map(),
    endnotes: new Map(),
    comments: [],
    styles,
    numbering,
    anchorIndex,
  };
}

function getPartText(parts: ReadonlyMap<string, Uint8Array> | Record<string, Uint8Array>, name: string): string | null {
  const bytes =
    typeof (parts as ReadonlyMap<string, Uint8Array>).get === 'function'
      ? (parts as ReadonlyMap<string, Uint8Array>).get(name)
      : (parts as Record<string, Uint8Array>)[name];
  return bytes ? strFromU8(bytes) : null;
}

function parseOptionalXml(xml: string | null): XElement | null {
  return xml === null ? null : parseXml(xml);
}

function appendBlocks(el: XElement, ctx: ReadContext, sink: IrBlock[]): void {
  if (isDroppedParagraphChild(el.name)) return;
  sink.push(buildBlock(el, ctx));
}

function buildBlock(el: XElement, ctx: ReadContext): IrBlock {
  if (nameEquals(el.name, W.p)) return buildParagraph(el, ctx);
  if (nameEquals(el.name, W.tbl)) return buildTable(el, ctx);
  if (nameEquals(el.name, W.sectPr)) return buildSectionBreak(el, ctx);
  return buildOpaqueBlock(el, ctx);
}

function unid(el: XElement, ctx: ReadContext): string {
  return ctx.unids.get(el) ?? el.attribute({ ns: 'http://powertools.codeplex.com/2011', local: 'Unid' }) ?? '';
}

function anchorFor(kind: IrAnchorKind, el: XElement, ctx: ReadContext): IrAnchor {
  return irAnchor(kind, ctx.scope, unid(el, ctx));
}

function buildParagraph(p: XElement, ctx: ReadContext): IrParagraph {
  const kind = kindFor(p, ctx.stylesRoot) ?? 'p';
  const pPr = p.element(W.pPr);
  const format = mapParaFormat(pPr);
  const processed = walkInlines([...p.elements()].filter((c) => !nameEquals(c.name, W.pPr)), ctx);
  const inlineSectPr = pPr?.element(W.sectPr) ?? null;
  const inlineSectionFormat = inlineSectPr ? mapSectionFormat(inlineSectPr) : null;
  return {
    kind: 'paragraph',
    anchor: anchorFor(kind, p, ctx),
    format,
    list: pPr ? resolveListInfo(pPr, ctx) : null,
    inlines: processed,
    inlineSectionBreakAnchor: inlineSectPr ? anchorFor('sec', inlineSectPr, ctx) : null,
    inlineSectionFormat,
    pPrDigest: pPrPropsDigest(pPr),
    resolvedListMarker: null,
    isListItemForLayout: isListItem(p, ctx.stylesRoot),
    contentHash: computeParagraphContentHash(processed),
    formatFingerprint: fingerprintBlock(format, runFormatsInOrder(processed), inlineSectionFormat),
    source: emptyIrProvenance,
  };
}

class InlineWalker {
  private readonly output: IrInline[] = [];
  private fieldDepth = 0;
  private inResult = false;
  private instruction = '';
  private readonly result: IrInline[] = [];
  private readonly captured: XElement[] = [];

  constructor(private readonly ctx: ReadContext) {}

  feed(child: XElement): void {
    if (nameEquals(child.name, W.r)) this.feedRun(child);
    else if (nameEquals(child.name, W.hyperlink)) this.emit(buildHyperlink(child, this.ctx), child);
    else if (nameEquals(child.name, W.fldSimple)) this.emit(buildFldSimple(child, this.ctx), child);
    else if (nameEquals(child.name, W.proofErr)) return;
    else if (isDroppedParagraphChild(child.name)) return;
    else this.emit({ kind: 'opaqueInline', elementName: child.name, canonicalHash: canonicalHash(child) }, child);
  }

  finish(): IrInline[] {
    if (this.fieldDepth > 0) {
      if (this.inResult) this.output.push({ kind: 'fieldRun', instruction: this.instruction, cachedResult: [...this.result], isSimpleField: false });
      else for (const el of this.captured) this.output.push({ kind: 'opaqueInline', elementName: el.name, canonicalHash: canonicalHash(el) });
    }
    return coalesceRuns(dropEmptyTextRuns(this.output));
  }

  private feedRun(r: XElement): void {
    const rPr = r.element(W.rPr);
    const runFormat = mapRunFormat(rPr);
    for (const child of r.elements()) {
      if (nameEquals(child.name, W.rPr)) continue;
      if (nameEquals(child.name, W.fldChar)) {
        this.handleFldChar(child);
        continue;
      }
      if (this.fieldDepth > 0) {
        this.captured.push(child);
        if (!this.inResult) {
          if (nameEquals(child.name, W.instrText) || nameEquals(child.name, W.delInstrText)) this.instruction += child.value;
        } else {
          emitRunChild(child, rPr, runFormat, this.ctx, this.result);
        }
        continue;
      }
      emitRunChild(child, rPr, runFormat, this.ctx, this.output);
    }
  }

  private handleFldChar(fldChar: XElement): void {
    const type = fldChar.attribute({ ns: W.val.ns, local: 'fldCharType' });
    if (type === 'begin') {
      this.fieldDepth++;
      if (this.fieldDepth === 1) {
        this.inResult = false;
        this.instruction = '';
        this.result.length = 0;
        this.captured.length = 0;
      }
    } else if (type === 'separate') {
      if (this.fieldDepth === 1) this.inResult = true;
    } else if (type === 'end' && this.fieldDepth > 0) {
      this.fieldDepth--;
      if (this.fieldDepth === 0) {
        this.output.push({ kind: 'fieldRun', instruction: this.instruction, cachedResult: [...this.result], isSimpleField: false });
        this.inResult = false;
        this.result.length = 0;
        this.captured.length = 0;
      }
    }
  }

  private emit(inline: IrInline, source: XElement | null = null): void {
    if (this.fieldDepth === 0) this.output.push(inline);
    else if (this.inResult) this.result.push(inline);
    else if (source) this.captured.push(source);
  }
}

function walkInlines(children: ReadonlyArray<XElement>, ctx: ReadContext): IrInline[] {
  const walker = new InlineWalker(ctx);
  for (const child of children) walker.feed(child);
  return walker.finish();
}

function emitRunChild(child: XElement, rPr: XElement | null, runFormat: IrRunFormat, ctx: ReadContext, sink: IrInline[]): void {
  if (nameEquals(child.name, W.t)) sink.push({ kind: 'textRun', text: child.value, format: runFormat, fromInlineSdt: false });
  else if (nameEquals(child.name, W.tab)) sink.push({ kind: 'tab', format: runFormat });
  else if (nameEquals(child.name, W.br)) sink.push({ kind: 'break', breakKind: breakKind(child) });
  else if (nameEquals(child.name, W.noBreakHyphen)) sink.push({ kind: 'textRun', text: '\u2011', format: runFormat, fromInlineSdt: false });
  else if (nameEquals(child.name, W.softHyphen)) sink.push({ kind: 'textRun', text: '\u00ad', format: runFormat, fromInlineSdt: false });
  else if (nameEquals(child.name, W.sym)) appendSym(child, rPr, runFormat, sink);
  else if (nameEquals(child.name, W.footnoteReference)) sink.push({ kind: 'noteRef', noteKind: 'Footnote', noteId: child.attribute(W.id) ?? '' });
  else if (nameEquals(child.name, W.endnoteReference)) sink.push({ kind: 'noteRef', noteKind: 'Endnote', noteId: child.attribute(W.id) ?? '' });
  else if (nameEquals(child.name, W.lastRenderedPageBreak) || nameEquals(child.name, W.commentReference) || isDroppedParagraphChild(child.name)) return;
  else if (nameEquals(child.name, W.drawing) || hasDescendantOrSelf(child, W.txbxContent)) {
    sink.push({ kind: 'opaqueInline', elementName: child.name, canonicalHash: canonicalHash(child) });
  } else {
    sink.push({ kind: 'opaqueInline', elementName: child.name, canonicalHash: canonicalHash(child) });
  }
}

function buildHyperlink(hyperlink: XElement, ctx: ReadContext): IrInline {
  const relId = hyperlink.attribute(R.id);
  const anchor = hyperlink.attribute({ ns: W.val.ns, local: 'anchor' });
  return {
    kind: 'hyperlink',
    target: relId ? ctx.hyperlinks.get(relId) ?? null : anchor ? `#${anchor}` : null,
    internalTarget: null,
    inlines: walkInlines([...hyperlink.elements()], ctx),
  };
}

function buildFldSimple(fldSimple: XElement, ctx: ReadContext): IrInline {
  return {
    kind: 'fieldRun',
    instruction: fldSimple.attribute({ ns: W.val.ns, local: 'instr' }) ?? '',
    cachedResult: walkInlines([...fldSimple.elements()], ctx),
    isSimpleField: true,
  };
}

function appendSym(sym: XElement, rPr: XElement | null, baseFormat: IrRunFormat, sink: IrInline[]): void {
  const raw = sym.attribute({ ns: W.val.ns, local: 'char' });
  const code = raw ? Number.parseInt(raw, 16) : Number.NaN;
  if (Number.isInteger(code) && code >= 0x20 && code <= 0xffff) {
    sink.push({ kind: 'textRun', text: String.fromCharCode(code), format: mapRunFormat(rPr, sym), fromInlineSdt: false });
  } else {
    sink.push({ kind: 'opaqueInline', elementName: sym.name, canonicalHash: canonicalHash(sym) });
  }
  void baseFormat;
}

function isDroppedParagraphChild(name: XName): boolean {
  return nameEquals(name, W.bookmarkStart) || nameEquals(name, W.bookmarkEnd) || nameEquals(name, W.commentRangeStart) || nameEquals(name, W.commentRangeEnd);
}

function breakKind(br: XElement): IrBreakKind {
  const type = br.attribute(W.type);
  return type === 'page' ? 'Page' : type === 'column' ? 'Column' : 'Line';
}

const dropEmptyTextRuns = (inlines: IrInline[]): IrInline[] => inlines.filter((i) => i.kind !== 'textRun' || i.text !== '');

function coalesceRuns(inlines: IrInline[]): IrInline[] {
  const result: IrInline[] = [];
  for (const inline of inlines) {
    const prev = result[result.length - 1];
    if (inline.kind === 'textRun' && prev?.kind === 'textRun' && sameJson(prev.format, inline.format) && prev.fromInlineSdt === inline.fromInlineSdt) {
      result[result.length - 1] = { ...prev, text: prev.text + inline.text };
    } else result.push(inline);
  }
  return result;
}

function* runFormatsInOrder(inlines: Iterable<IrInline>): IterableIterator<IrRunFormat> {
  for (const inline of inlines) {
    if (inline.kind === 'textRun' || inline.kind === 'tab') yield inline.format;
    else if (inline.kind === 'hyperlink') yield* runFormatsInOrder(inline.inlines);
    else if (inline.kind === 'fieldRun') yield* runFormatsInOrder(inline.cachedResult);
  }
}

function computeParagraphContentHash(inlines: Iterable<IrInline>): IrHash {
  const builder = new IrContentHashBuilder();
  appendInlinesToContentHash(inlines, builder);
  return builder.build();
}

function appendInlinesToContentHash(inlines: Iterable<IrInline>, builder: IrContentHashBuilder): void {
  for (const inline of inlines) {
    if (inline.kind === 'textRun') builder.appendText(inline.text);
    else if (inline.kind === 'tab') builder.appendSentinel(SENTINEL_TAB);
    else if (inline.kind === 'break') builder.appendSentinel(inline.breakKind === 'Page' ? SENTINEL_PAGE_BREAK : inline.breakKind === 'Column' ? SENTINEL_COLUMN_BREAK : SENTINEL_LINE_BREAK);
    else if (inline.kind === 'hyperlink') {
      builder.appendSentinel(SENTINEL_HYPERLINK);
      builder.appendText(inline.target ?? '');
      builder.appendSentinel(SENTINEL_HYPERLINK_TARGET_END);
      appendInlinesToContentHash(inline.inlines, builder);
    } else if (inline.kind === 'fieldRun') appendInlinesToContentHash(inline.cachedResult, builder);
    else if (inline.kind === 'noteRef') builder.appendSentinel(inline.noteKind === 'Footnote' ? SENTINEL_FOOTNOTE_REF : SENTINEL_ENDNOTE_REF);
    else if (inline.kind === 'inlineImage') {
      // Sentinel + image bytes hash (spec §6.1) — extent/alt do not hash.
      builder.appendSentinel(SENTINEL_IMAGE);
      builder.appendHash(inline.imageBytesHash);
    } else if (inline.kind === 'textbox') {
      // Sentinel + each inner block's contentHash (spec §6.1 M1.5).
      builder.appendSentinel(SENTINEL_TEXTBOX);
      for (const block of inline.blocks) builder.appendHash(block.contentHash);
    } else if (inline.kind === 'opaqueInline') {
      builder.appendSentinel(SENTINEL_OPAQUE);
      builder.appendHash(inline.canonicalHash);
    }
  }
}

function mapParaFormat(pPr: XElement | null): IrParaFormat {
  if (!pPr) return emptyParaFormat();
  const jcVal = attrVal(pPr.element(W.jc));
  const ind = pPr.element(W.ind);
  const spacing = pPr.element(W.spacing);
  const hanging = intAttr(ind, { ns: W.val.ns, local: 'hanging' });
  const line = intAttr(spacing, { ns: W.val.ns, local: 'line' });
  const lineSpacing: IrLineSpacing | null = line === null ? null : {
    valueTwips: line,
    rule: spacing?.attribute({ ns: W.val.ns, local: 'lineRule' }) === 'atLeast' ? 'AtLeast' : spacing?.attribute({ ns: W.val.ns, local: 'lineRule' }) === 'exact' ? 'Exact' : 'Auto',
  };
  const numPr = pPr.element(W.numPr);
  return {
    styleId: attrVal(pPr.element(W.pStyle)),
    justification: mapJustification(jcVal),
    indentLeftTwips: intAttr(ind, { ns: W.val.ns, local: 'left' }) ?? intAttr(ind, { ns: W.val.ns, local: 'start' }),
    indentRightTwips: intAttr(ind, { ns: W.val.ns, local: 'right' }) ?? intAttr(ind, { ns: W.val.ns, local: 'end' }),
    indentFirstLineTwips: hanging !== null ? -hanging : intAttr(ind, { ns: W.val.ns, local: 'firstLine' }),
    spacingBeforeTwips: intAttr(spacing, { ns: W.val.ns, local: 'before' }),
    spacingAfterTwips: intAttr(spacing, { ns: W.val.ns, local: 'after' }),
    lineSpacing,
    outlineLevel: intAttr(pPr.element(W.outlineLvl), W.val),
    keepNext: toggle(pPr.element(W.keepNext)),
    keepLines: toggle(pPr.element(W.keepLines)),
    pageBreakBefore: toggle(pPr.element(W.pageBreakBefore)),
    numId: intAttr(numPr?.element(W.numId) ?? null, W.val),
    ilvl: intAttr(numPr?.element(W.ilvl) ?? null, W.val),
    unmodeledDigest: unmodeledDigest(pPr, PPR_CONSUMED),
  };
}

function emptyParaFormat(): IrParaFormat {
  return {
    styleId: null, justification: null, indentLeftTwips: null, indentRightTwips: null,
    indentFirstLineTwips: null, spacingBeforeTwips: null, spacingAfterTwips: null,
    lineSpacing: null, outlineLevel: null, keepNext: null, keepLines: null, pageBreakBefore: null,
    numId: null, ilvl: null, unmodeledDigest: EMPTY_UNMODELED_DIGEST,
  };
}

function mapRunFormat(rPr: XElement | null, extraUnmodeled: XElement | null = null): IrRunFormat {
  if (!rPr && !extraUnmodeled) return emptyRunFormat();
  if (!rPr && extraUnmodeled) return { ...emptyRunFormat(), unmodeledDigest: extraUnmodeledDigest(extraUnmodeled) };
  const rp = rPr!;
  const vertVal = attrVal(rp.element(W.vertAlign));
  const vertAlign = vertVal === 'subscript' ? 'Subscript' : vertVal === 'superscript' ? 'Superscript' : null;
  return {
    styleId: attrVal(rp.element(W.rStyle)),
    bold: toggle(rp.element(W.b)),
    italic: toggle(rp.element(W.i)),
    underline: mapUnderline(rp.element(W.u)),
    strike: toggle(rp.element(W.strike)),
    doubleStrike: toggle(rp.element(W.dstrike)),
    vertAlign,
    fontAscii: rp.element(W.rFonts)?.attribute({ ns: W.val.ns, local: 'ascii' }) ?? null,
    sizeHalfPoints: intAttr(rp.element(W.sz), W.val),
    colorHex: attrVal(rp.element(W.color)),
    highlight: attrVal(rp.element(W.highlight)),
    caps: toggle(rp.element(W.caps)),
    smallCaps: toggle(rp.element(W.smallCaps)),
    vanish: toggle(rp.element(W.vanish)),
    unmodeledDigest: unmodeledDigest(rp, RPR_CONSUMED, vertAlign ? W.vertAlign : null, extraUnmodeled),
  };
}

function emptyRunFormat(): IrRunFormat {
  return {
    styleId: null, bold: null, italic: null, underline: null, strike: null, doubleStrike: null,
    vertAlign: null, fontAscii: null, sizeHalfPoints: null, colorHex: null, highlight: null,
    caps: null, smallCaps: null, vanish: null, unmodeledDigest: EMPTY_UNMODELED_DIGEST,
  };
}

function mapUnderline(u: XElement | null): IrUnderline | null {
  if (!u) return null;
  const val = u.attribute(W.val);
  const kind: IrUnderlineKind = val === 'single' ? 'Single' : val === 'double' ? 'Double' : val === 'thick' ? 'Thick' : val === 'dotted' ? 'Dotted' : val === 'dash' || val === 'dashed' ? 'Dashed' : val === 'wave' ? 'Wave' : val === 'words' ? 'Words' : val === 'none' ? 'None' : 'Other';
  return { kind, colorHex: u.attribute({ ns: W.val.ns, local: 'color' }) };
}

function mapJustification(value: string | null): IrJustification | null {
  if (value === null) return null;
  return value === 'left' || value === 'start' ? 'Left' : value === 'center' ? 'Center' : value === 'right' || value === 'end' ? 'Right' : value === 'both' ? 'Both' : value === 'distribute' ? 'Distribute' : 'Other';
}

function buildTable(tbl: XElement, ctx: ReadContext): IrTable {
  const rows: IrRow[] = [];
  const cellFingerprints: IrHash[] = [];
  const contentBuilder = new IrContentHashBuilder();
  for (const tr of tbl.elements(W.tr)) {
    const { row, fingerprints } = buildRow(tr, ctx);
    rows.push(row);
    contentBuilder.appendHash(row.contentHash);
    cellFingerprints.push(...fingerprints);
  }
  const tblGrid = tbl.element(W.tblGrid);
  const tblPrDigest = shellChildrenDigest([...tbl.elements()].filter((e) => !nameEquals(e.name, W.tr) && !nameEquals(e.name, W.tblGrid)));
  const tblGridDigest = shellChildrenDigest(tblGrid ? [tblGrid] : []);
  const fp = new IrContentHashBuilder();
  fp.appendHash(tblPrDigest);
  fp.appendHash(tblGridDigest);
  for (const row of rows) fp.appendHash(row.trPrDigest);
  for (const hash of cellFingerprints) fp.appendHash(hash);
  return {
    kind: 'table',
    anchor: anchorFor('tbl', tbl, ctx),
    rows,
    tblPrDigest,
    tblGridDigest,
    contentHash: contentBuilder.build(),
    formatFingerprint: fp.build(),
    source: emptyIrProvenance,
  };
}

function buildRow(tr: XElement, ctx: ReadContext): { row: IrRow; fingerprints: IrHash[] } {
  const cells: IrCell[] = [];
  const fingerprints: IrHash[] = [];
  const builder = new IrContentHashBuilder();
  builder.appendStructure(STRUCTURE_ROW);
  for (const tc of tr.elements(W.tc)) {
    const built = buildCell(tc, ctx);
    cells.push(built.cell);
    fingerprints.push(...built.fingerprints);
    builder.appendHash(built.cell.contentHash);
  }
  const trPr = tr.element(W.trPr);
  const row: IrRow = {
    kind: 'row',
    anchor: anchorFor('tr', tr, ctx),
    cells,
    contentHash: builder.build(),
    source: emptyIrProvenance,
    trPrDigest: shellChildrenDigest([...tr.elements()].filter((e) => !nameEquals(e.name, W.tc))),
    trPrShellDigest: shellChildrenDigest(trPr ? [trPr] : []),
    trPrExDigest: shellChildrenDigest([...tr.elements(W.tblPrEx)]),
    fromTableSdt: false,
  };
  return { row, fingerprints };
}

function buildCell(tc: XElement, ctx: ReadContext): { cell: IrCell; fingerprints: IrHash[] } {
  const tcPr = tc.element(W.tcPr);
  const blocks: IrBlock[] = [];
  const fingerprints: IrHash[] = [];
  const builder = new IrContentHashBuilder();
  builder.appendStructure(STRUCTURE_CELL);
  const shellDigest = tcPr ? canonicalHash(tcPr) : EMPTY_HASH;
  if (tcPr) builder.appendHash(shellDigest);
  for (const child of tc.elements()) {
    if (nameEquals(child.name, W.tcPr)) continue;
    const before = blocks.length;
    appendBlocks(child, ctx, blocks);
    for (let i = before; i < blocks.length; i++) {
      builder.appendHash(blocks[i]!.contentHash);
      fingerprints.push(blocks[i]!.formatFingerprint);
    }
  }
  return {
    cell: {
      kind: 'cell',
      anchor: anchorFor('tc', tc, ctx),
      blocks,
      gridSpan: intAttr(tcPr?.element(W.gridSpan) ?? null, W.val) ?? 1,
      vMerge: mapVMerge(tcPr?.element(W.vMerge) ?? null),
      contentHash: builder.build(),
      source: emptyIrProvenance,
      shellDigest,
      tcPrShellDigest: shellChildrenDigest(tcPr ? [tcPr] : []),
      fromRowSdt: false,
    },
    fingerprints,
  };
}

function mapVMerge(vMerge: XElement | null): IrVMerge {
  return !vMerge ? 'None' : vMerge.attribute(W.val) === 'restart' ? 'Restart' : 'Continue';
}

function buildSectionBreak(sectPr: XElement, ctx: ReadContext): IrSectionBreak {
  const format = mapSectionFormat(sectPr);
  const builder = new IrContentHashBuilder();
  builder.appendHash(canonicalHash(sectPr));
  return {
    kind: 'sectionBreak',
    anchor: anchorFor('sec', sectPr, ctx),
    format,
    contentHash: builder.build(),
    formatFingerprint: fingerprintSectionFormat(format),
    source: emptyIrProvenance,
  };
}

function mapSectionFormat(sectPr: XElement): IrSectionFormat {
  const pgSz = sectPr.element(W.pgSz);
  const pgMar = sectPr.element(W.pgMar);
  const orient = pgSz?.attribute({ ns: W.val.ns, local: 'orient' }) ?? null;
  return {
    pageWidthTwips: intAttr(pgSz, { ns: W.val.ns, local: 'w' }),
    pageHeightTwips: intAttr(pgSz, { ns: W.val.ns, local: 'h' }),
    landscape: orient === null ? null : orient === 'landscape',
    marginTopTwips: intAttr(pgMar, { ns: W.val.ns, local: 'top' }),
    marginBottomTwips: intAttr(pgMar, { ns: W.val.ns, local: 'bottom' }),
    marginLeftTwips: intAttr(pgMar, { ns: W.val.ns, local: 'left' }),
    marginRightTwips: intAttr(pgMar, { ns: W.val.ns, local: 'right' }),
    sectionType: attrVal(sectPr.element(W.type)),
    unmodeledDigest: unmodeledDigest(sectPr, SECTPR_CONSUMED),
  };
}

function buildOpaqueBlock(el: XElement, ctx: ReadContext): IrOpaqueBlock {
  return {
    kind: 'opaqueBlock',
    anchor: anchorFor('unk', el, ctx),
    elementName: el.name,
    contentHash: canonicalHash(el),
    formatFingerprint: EMPTY_UNMODELED_DIGEST,
    source: emptyIrProvenance,
  };
}

function shellChildrenDigest(wrappers: ReadonlyArray<XElement>): IrHash {
  if (wrappers.length === 0) return EMPTY_SHELL_DIGEST;
  const children: XElement[] = [];
  for (const wrapper of wrappers) children.push(...wrapper.elements());
  return children.length === 0 ? EMPTY_SHELL_DIGEST : canonicalHash(syntheticElement('shell', children));
}

function pPrPropsDigest(pPr: XElement | null): IrHash {
  if (!pPr) return EMPTY_PPR_PROPS_DIGEST;
  const children = [...pPr.elements()].filter((e) => !nameEquals(e.name, W.rPr) && !nameEquals(e.name, W.sectPr) && e.name.local !== 'pPrChange');
  return children.length === 0 ? EMPTY_PPR_PROPS_DIGEST : canonicalHash(syntheticElement('ppr-props', children));
}

function unmodeledDigest(props: XElement, consumed: ReadonlySet<string>, alsoConsumed: XName | null = null, extra: XElement | null = null): IrHash {
  const leftovers = [...props.elements()].filter((e) => !consumed.has(nameKey(e.name)) && (!alsoConsumed || !nameEquals(e.name, alsoConsumed)));
  if (leftovers.length === 0 && !extra) return EMPTY_UNMODELED_DIGEST;
  return canonicalHash(syntheticElement('unmodeled', extra ? [...leftovers, extra] : leftovers));
}

function extraUnmodeledDigest(extra: XElement): IrHash {
  return canonicalHash(syntheticElement('unmodeled', [extra]));
}

function resolveListInfo(pPr: XElement, ctx: ReadContext): IrListInfo | null {
  let numPr = pPr.element(W.numPr);
  let fromStyle = false;
  if (!numPr) {
    numPr = resolveStyleNumPr(attrVal(pPr.element(W.pStyle)), ctx.styles);
    fromStyle = !!numPr;
  }
  if (!numPr) return null;
  const numId = intAttr(numPr.element(W.numId), W.val);
  if (numId === null || numId === 0) return null;
  const ilvl = intAttr(numPr.element(W.ilvl), W.val) ?? 0;
  const num = ctx.numbering.nums.get(numId);
  return { numId, abstractNumId: num?.abstractNumId ?? null, ilvl, numberFormat: '', startOverride: num?.startOverrides.get(ilvl) ?? null, fromStyle };
}

function resolveStyleNumPr(styleId: string | null, styles: IrStyleRegistry): XElement | null {
  let current = styleId;
  const visited = new Set<string>();
  for (let i = 0; i < 16 && current; i++) {
    if (visited.has(current)) return null;
    visited.add(current);
    const style = styles.styles.get(current);
    if (!style) return null;
    const numPr = style.pPr?.element(W.numPr) ?? null;
    if (numPr) return numPr;
    current = style.basedOn;
  }
  return null;
}

function buildStyleRegistry(root: XElement | null): IrStyleRegistry {
  if (!root) return { styles: new Map(), defaultParagraphStyleId: null, docDefaultsPPr: null, docDefaultsRPr: null };
  const styles = new Map<string, IrStyle>();
  let defaultParagraphStyleId: string | null = null;
  for (const styleEl of root.elements(W.style)) {
    const id = styleEl.attribute(W.styleId);
    if (!id || styles.has(id)) continue;
    const type = styleEl.attribute(W.type) ?? 'paragraph';
    const isDefault = ['1', 'true', 'on'].includes(styleEl.attribute(W.default) ?? '');
    styles.set(id, { id, basedOn: attrVal(styleEl.element(W.basedOn)), pPr: styleEl.element(W.pPr), rPr: styleEl.element(W.rPr) });
    if (!defaultParagraphStyleId && isDefault && type === 'paragraph') defaultParagraphStyleId = id;
  }
  const docDefaults = root.element(W.docDefaults);
  return {
    styles,
    defaultParagraphStyleId,
    docDefaultsPPr: docDefaults?.element(W.pPrDefault)?.element(W.pPr) ?? null,
    docDefaultsRPr: docDefaults?.element(W.rPrDefault)?.element(W.rPr) ?? null,
  };
}

function buildEmptyNumberingRegistry(): IrNumberingRegistry {
  return { nums: new Map(), abstractNums: new Map() };
}

function buildHyperlinkMap(root: XElement | null): ReadonlyMap<string, string> {
  const map = new Map<string, string>();
  if (!root) return map;
  for (const rel of root.elements(REL.Relationship)) {
    const id = rel.attribute({ ns: '', local: 'Id' });
    const target = rel.attribute({ ns: '', local: 'Target' });
    const type = rel.attribute({ ns: '', local: 'Type' });
    if (id && target && type?.endsWith('/hyperlink')) map.set(id, target);
  }
  return map;
}

function indexBlock(block: IrBlock, index: Map<string, IrBlock>): void {
  const key = anchorToString(block.anchor);
  if (index.has(key)) throw new Error(`Duplicate IR anchor '${key}' (invariant violation).`);
  index.set(key, block);
  if (block.kind === 'table') {
    for (const row of block.rows) for (const cell of row.cells) for (const child of cell.blocks) indexBlock(child, index);
  }
}

function fingerprintBlock(paraFormat: IrParaFormat, runFormats: Iterable<IrRunFormat>, inlineSection: IrSectionFormat | null): IrHash {
  const bytes: Uint8Array[] = [irHashToBytes(fingerprintParaFormat(paraFormat))];
  for (const rf of runFormats) bytes.push(irHashToBytes(fingerprintRunFormat(rf)));
  if (inlineSection) bytes.push(irHashToBytes(fingerprintSectionFormat(inlineSection)));
  return irHashCompute(concat(bytes));
}

function fingerprintRunFormat(f: IrRunFormat): IrHash {
  const fields: string[] = [];
  appendField(fields, 'StyleId', f.styleId);
  appendField(fields, 'Bold', boolString(f.bold));
  appendField(fields, 'Italic', boolString(f.italic));
  appendField(fields, 'Underline', f.underline ? (f.underline.colorHex ? `${f.underline.kind}|${f.underline.colorHex}` : f.underline.kind) : null);
  appendField(fields, 'Strike', boolString(f.strike));
  appendField(fields, 'DoubleStrike', boolString(f.doubleStrike));
  appendField(fields, 'VertAlign', f.vertAlign);
  appendField(fields, 'FontAscii', f.fontAscii);
  appendField(fields, 'SizeHalfPoints', numString(f.sizeHalfPoints));
  appendField(fields, 'ColorHex', f.colorHex);
  appendField(fields, 'Highlight', f.highlight);
  appendField(fields, 'Caps', boolString(f.caps));
  appendField(fields, 'SmallCaps', boolString(f.smallCaps));
  appendField(fields, 'Vanish', boolString(f.vanish));
  return hashFields(fields.join(''), f.unmodeledDigest);
}

function fingerprintParaFormat(f: IrParaFormat): IrHash {
  const fields: string[] = [];
  appendField(fields, 'StyleId', f.styleId);
  appendField(fields, 'Justification', f.justification);
  appendField(fields, 'IndentLeftTwips', numString(f.indentLeftTwips));
  appendField(fields, 'IndentRightTwips', numString(f.indentRightTwips));
  appendField(fields, 'IndentFirstLineTwips', numString(f.indentFirstLineTwips));
  appendField(fields, 'SpacingBeforeTwips', numString(f.spacingBeforeTwips));
  appendField(fields, 'SpacingAfterTwips', numString(f.spacingAfterTwips));
  appendField(fields, 'LineSpacing', f.lineSpacing ? `${f.lineSpacing.valueTwips}|${f.lineSpacing.rule}` : null);
  appendField(fields, 'OutlineLevel', numString(f.outlineLevel));
  appendField(fields, 'KeepNext', boolString(f.keepNext));
  appendField(fields, 'KeepLines', boolString(f.keepLines));
  appendField(fields, 'PageBreakBefore', boolString(f.pageBreakBefore));
  appendField(fields, 'NumId', numString(f.numId));
  appendField(fields, 'Ilvl', numString(f.ilvl));
  return hashFields(fields.join(''), f.unmodeledDigest);
}

function fingerprintSectionFormat(f: IrSectionFormat): IrHash {
  const fields: string[] = [];
  appendField(fields, 'PageWidthTwips', numString(f.pageWidthTwips));
  appendField(fields, 'PageHeightTwips', numString(f.pageHeightTwips));
  appendField(fields, 'Landscape', boolString(f.landscape));
  appendField(fields, 'MarginTopTwips', numString(f.marginTopTwips));
  appendField(fields, 'MarginBottomTwips', numString(f.marginBottomTwips));
  appendField(fields, 'MarginLeftTwips', numString(f.marginLeftTwips));
  appendField(fields, 'MarginRightTwips', numString(f.marginRightTwips));
  appendField(fields, 'SectionType', f.sectionType);
  return hashFields(fields.join(''), f.unmodeledDigest);
}

function appendField(fields: string[], name: string, value: string | null): void {
  if (value !== null) fields.push(`${name}=${value.length}:${value};`);
}

function hashFields(fields: string, digest: IrHash): IrHash {
  return irHashCompute(concat([new TextEncoder().encode(fields), irHashToBytes(digest)]));
}

function concat(chunks: ReadonlyArray<Uint8Array>): Uint8Array {
  const length = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

const boolString = (v: boolean | null): string | null => (v === null ? null : v ? 'true' : 'false');
const numString = (v: number | null): string | null => (v === null ? null : String(v));
const sameJson = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);
const attrVal = (el: XElement | null | undefined): string | null => el?.attribute(W.val) ?? null;

function intAttr(el: XElement | null | undefined, name: XName): number | null {
  const raw = el?.attribute(name);
  if (raw === undefined || raw === null || raw === '') return null;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) ? n : null;
}

function toggle(el: XElement | null | undefined): boolean | null {
  if (!el) return null;
  const val = el.attribute(W.val);
  return val === '0' || val === 'false' || val === 'off' ? false : true;
}

function hasDescendantOrSelf(el: XElement, name: XName): boolean {
  if (nameEquals(el.name, name)) return true;
  for (const d of el.descendants(name)) return !!d;
  return false;
}

void A;
void WP;
