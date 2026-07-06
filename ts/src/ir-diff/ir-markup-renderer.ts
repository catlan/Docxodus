import { unzipSync, strFromU8, strToU8 } from 'fflate';
import { anchorToString } from '../ir/ir-anchor.js';
import { W, W_NS, PT_NS, R, R_NS, REL } from '../ir/names.js';
import type { IrBlock, IrParagraph, IrTable } from '../ir/ir-blocks.js';
import { readIrDocument, type IrDocument } from '../ir/ir-reader.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import { normalizeIrDiffSettings, type IrDiffSettings, type IrDiffSettingsOptions } from './ir-diff-settings.js';
import type { IrCellOp, IrEditOp, IrEditScript, IrHeaderFooterDiff, IrNoteDiff, IrRowOp, IrTableDiff } from './ir-edit-script.js';
import type { IrDiffToken } from './ir-diff-token.js';
import type { IrTokenDiff, IrTokenOp } from './ir-token-diff.js';
import { writeXmlPart, writeXmlPartFormatted } from '../xml/xwriter.js';
import { XElement, nameEquals, parseXml, xname, type XAttributeInfo, type XName } from '../xml/xelement.js';
import { attr, childrenElements, cloneElement, descendantsAndSelf, element, replaceName, type XChild } from '../xml/xclone.js';

type RevKind = 'Ins' | 'Del' | 'MoveFrom' | 'MoveTo';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XML_SPACE = xname(XML_NS, 'space');
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const SOURCE_LINK_ID = xname(PT_NS, 'SourceLinkId');
const REL_ID = xname('', 'Id');
const REL_TYPE = xname('', 'Type');
const REL_TARGET = xname('', 'Target');
const REL_TARGET_MODE = xname('', 'TargetMode');
const CONTENT_TYPES = '[Content_Types].xml';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CT_TYPES = xname(CT_NS, 'Types');
const CT_OVERRIDE = xname(CT_NS, 'Override');
const CT_DEFAULT = xname(CT_NS, 'Default');
const CT_PART_NAME = xname('', 'PartName');
const CT_CONTENT_TYPE = xname('', 'ContentType');
const CT_EXTENSION = xname('', 'Extension');
const W_FOOTNOTES = xname(W_NS, 'footnotes');
const W_ENDNOTES = xname(W_NS, 'endnotes');
const WORD_REL_PREFIX = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT_FOOTNOTES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';
const CT_ENDNOTES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml';

interface RenderState {
  readonly left: IrDocument;
  readonly right: IrDocument;
  readonly settings: IrDiffSettings;
  nextId: number;
  nextMoveName: number;
  readonly moveNames: Map<number, string>;
  readonly rightSourcedClones: XElement[];
}

export function renderIrMarkup(
  leftDocx: Uint8Array,
  rightDocx: Uint8Array,
  script: IrEditScript,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): Map<string, Uint8Array> {
  const leftParts = unzipSync(leftDocx);
  const rightParts = unzipSync(rightDocx);
  const settings = normalizeIrDiffSettings(options);
  const left = readIrDocument(leftDocx, { retainSources: true });
  const right = readIrDocument(rightDocx, { retainSources: true });
  const state: RenderState = { left, right, settings, nextId: 1, nextMoveName: 1, moveNames: new Map(), rightSourcedClones: [] };
  const bodyBlocks: XElement[] = [];
  for (const op of script.ops) renderBlockOp(op, state, bodyBlocks);

  const root = left.parsedPartRoots?.get('/word/document.xml');
  if (!root) throw new Error('renderIrMarkup requires retained /word/document.xml source.');
  const body = root.element(W.body);
  if (!body) throw new Error('LEFT document has no w:body.');

  const trailingSectPr = childrenElements(body, W.sectPr).at(-1) ?? null;
  const newBodyChildren: XChild[] = [];
  for (const child of body.children) {
    if (!(child instanceof XElement)) {
      newBodyChildren.push(child);
      continue;
    }
    if (nameEquals(child.name, W.sectPr)) continue;
  }
  newBodyChildren.length = 0;
  newBodyChildren.push(...bodyBlocks);
  if (trailingSectPr) {
    const rightSect = right.parsedPartRoots?.get('/word/document.xml')?.element(W.body)?.elements(W.sectPr).next().value as XElement | undefined;
    newBodyChildren.push(stripUnids(cloneElement(trailingSectPr)));
    const outSect = newBodyChildren[newBodyChildren.length - 1] as XElement;
    if (settings.trackBlockFormatChanges && rightSect && sectPrPropsString(trailingSectPr) !== sectPrPropsString(rightSect)) {
      applySectPrChange(outSect, trailingSectPr, rightSect, state);
    }
  }

  const newBody = cloneElement(body, { children: newBodyChildren });
  const result = new Map<string, Uint8Array>();
  for (const [name, bytes] of Object.entries(leftParts)) result.set(name, bytes);

  // Import media referenced by RIGHT-sourced clones (registry-based, per attribute occurrence,
  // no dedup — the C# MoveRelatedPartsToDestination semantics). Mutates the registered clones'
  // relationship ids IN PLACE, so it must run BEFORE the functional stripUnids deep-clone below.
  importRightSourcedMedia(state, result, rightParts);
  const imported = stripUnids(replaceChild(root, body, newBody));
  renderNoteScopes(script, state, result, leftParts, rightParts);
  renderHeaderFooterScopes(script, state, result, rightParts, imported);
  const footnoteRemap = renumberNoteIds(imported, result, 'word/footnotes.xml', W.footnoteReference, W.footnote, W_FOOTNOTES);
  const endnoteRemap = renumberNoteIds(imported, result, 'word/endnotes.xml', W.endnoteReference, W.endnote, W_ENDNOTES);
  remapNestedNoteReferences(result, footnoteRemap, endnoteRemap);
  const normalized = normalizeBookmarks(imported, bodyBookmarkNames(left), bodyBookmarkNames(right));
  const remapped = remapHyperlinkRelationshipCollisions(normalized, result, rightParts);
  const outputShape = linqOutputShape(remapped.root);
  const newRoot = cloneElement(outputShape, {
    documentProlog: hasRevisionMarkup(rightParts['word/document.xml']) && !hasRenderedRevisionMarkup(outputShape)
      ? '<?xml version="1.0" encoding="utf-8"?>'
      : '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n',
  });

  result.set('word/document.xml', strToU8(writeXmlPart(newRoot)));
  if (remapped.relationshipsXml) result.set('word/_rels/document.xml.rels', strToU8(remapped.relationshipsXml));
  return result;
}

function hasRevisionMarkup(bytes: Uint8Array | undefined): boolean {
  if (!bytes) return false;
  return /\bw:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|delText|delInstrText)\b/.test(strFromU8(bytes));
}

function hasRenderedRevisionMarkup(root: XElement): boolean {
  for (const el of root.descendants()) {
    if (
      nameEquals(el.name, W.ins) || nameEquals(el.name, W.del) ||
      nameEquals(el.name, W.moveFrom) || nameEquals(el.name, W.moveTo) ||
      nameEquals(el.name, W.rPrChange) || nameEquals(el.name, W.pPrChange) ||
      nameEquals(el.name, W.tblPrChange) || nameEquals(el.name, W.trPrChange) ||
      nameEquals(el.name, W.tcPrChange) || nameEquals(el.name, W.sectPrChange)
    ) return true;
  }
  return false;
}

function renderBlockOp(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  if (isSectionBreakOp(op, state)) return;
  switch (op.kind) {
    case 'EqualBlock':
      emitVerbatimRight(op.rightAnchor, state, sink);
      return;
    case 'FormatOnlyBlock':
      if (resolveBlock(op.rightAnchor, state.right)?.kind === 'table') emitFormatOnlyTable(op, state, sink);
      else emitFormatOnlyParagraph(op, state, sink);
      return;
    case 'InsertBlock':
      emitWholeBlock(op.rightAnchor, state.right, state, sink, 'Ins');
      return;
    case 'DeleteBlock':
      emitWholeBlock(op.leftAnchor, state.left, state, sink, 'Del');
      return;
    case 'ModifyBlock':
      renderModifyBlock(op, state, sink);
      return;
    case 'MoveBlock':
    case 'MoveModifyBlock':
      if (!state.settings.renderMoves) {
        if (op.isMoveSource) emitWholeBlock(op.leftAnchor, state.left, state, sink, 'Del');
        else emitWholeBlock(op.rightAnchor, state.right, state, sink, 'Ins');
      } else if (op.isMoveSource) emitMoveSource(op, state, sink);
      else emitMoveDestination(op, state, sink);
      return;
    case 'SplitBlock':
      renderSplitBlock(op, state, sink);
      return;
    case 'MergeBlock':
      renderMergeBlock(op, state, sink);
      return;
  }
}

function renderModifyBlock(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  if (op.tokenDiff && !op.textboxDiffs && resolveBlock(op.leftAnchor, state.left)?.kind === 'paragraph' && resolveBlock(op.rightAnchor, state.right)?.kind === 'paragraph') {
    renderModifiedParagraph(op, op.tokenDiff as IrTokenDiff, state, sink);
    return;
  }
  if (op.tableDiff && resolveBlock(op.leftAnchor, state.left)?.kind === 'table' && resolveBlock(op.rightAnchor, state.right)?.kind === 'table') {
    if (renderModifiedTable(op, op.tableDiff as IrTableDiff, state, sink)) return;
  }
  if (op.leftAnchor) emitWholeBlock(op.leftAnchor, state.left, state, sink, 'Del');
  if (op.rightAnchor) emitWholeBlock(op.rightAnchor, state.right, state, sink, 'Ins');
}

function emitVerbatim(anchor: string | null, doc: IrDocument, sink: XElement[]): void {
  const src = sourceElement(anchor, doc);
  if (src) sink.push(stripUnids(cloneElement(src)));
}

function emitVerbatimRight(anchor: string | null, state: RenderState, sink: XElement[]): void {
  const src = sourceElement(anchor, state.right);
  if (!src) return;
  const clone = stripUnids(cloneElement(src));
  registerRightSourcedClone(clone, state);
  sink.push(clone);
}

function emitWholeBlock(anchor: string | null, doc: IrDocument, state: RenderState, sink: XElement[], kind: RevKind): void {
  const src = sourceElement(anchor, doc);
  if (!src) return;
  const clone = stripUnids(cloneElement(src));
  if (doc === state.right) registerRightSourcedClone(clone, state);
  if (nameEquals(clone.name, W.p)) {
    markWholeParagraph(clone, kind, state);
    sink.push(clone);
  } else if (nameEquals(clone.name, W.tbl)) {
    markWholeTable(clone, kind, state);
    sink.push(clone);
  } else {
    sink.push(clone);
  }
}

function markWholeTable(tbl: XElement, kind: RevKind, state: RenderState): void {
  const rows = childrenElements(tbl, W.tr).map((tr) => markWholeRow(tr, kind, state));
  replaceChildrenInPlace(tbl, (child) => child instanceof XElement && nameEquals(child.name, W.tr) ? rows.shift()! : child);
}

function markWholeRow(tr: XElement, kind: RevKind, state: RenderState): XElement {
  let out = tr;
  let trPr = out.element(W.trPr);
  if (!trPr) {
    trPr = element(W.trPr);
    out = cloneElement(out, { children: [trPr, ...out.children] });
  }
  const newTrPr = cloneElement(trPr, {
    children: [...trPr.children.filter((c) => !(c instanceof XElement && (nameEquals(c.name, W.ins) || nameEquals(c.name, W.del)))), element(isDeleteGrade(kind) ? W.del : W.ins, revisionAttrs(state))],
  });
  out = replaceChild(out, trPr, newTrPr);
  return transformDescendants(out, (el) => nameEquals(el.name, W.p) ? markWholeParagraphClone(el, kind, state) : null);
}

function markWholeParagraphClone(para: XElement, kind: RevKind, state: RenderState): XElement {
  const p = cloneElement(para);
  markWholeParagraph(p, kind, state);
  return p;
}

function markWholeParagraph(para: XElement, kind: RevKind, state: RenderState): void {
  const pPr = para.element(W.pPr);
  const runChildren = childrenElements(para).filter((e) => !nameEquals(e.name, W.pPr));
  const wrapped = runChildren.flatMap((c) => wrapFieldAware(c, kind, state));
  para.children.splice(0, para.children.length, ...(pPr ? [pPr, ...wrapped] : wrapped));
  for (const child of para.children) if (child instanceof XElement) child.parent = para;
  markParagraphMark(para, kind, state);
}

function wrapFieldAware(runLevel: XElement, kind: RevKind, state: RenderState): XElement[] {
  return expandFieldForRevision(runLevel).map((part) => wrapRunLevel(part, kind, state));
}

function wrapRunLevel(runLevel: XElement, kind: RevKind, state: RenderState): XElement {
  if (nameEquals(runLevel.name, W.hyperlink) || nameEquals(runLevel.name, W.sdt) || nameEquals(runLevel.name, W.smartTag)) {
    return cloneElement(runLevel, { children: childrenElements(runLevel).map((c) => wrapContainerChild(c, kind, state)) });
  }
  const clone = cloneElement(runLevel);
  const child = isDeleteGrade(kind) ? convertTextToDelText(clone) : clone;
  if (!isDeleteGrade(kind)) registerRightSourcedClone(child, state);
  return element(revElementName(kind), revisionAttrs(state), [child]);
}

function wrapContainerChild(child: XElement, kind: RevKind, state: RenderState): XElement {
  if (nameEquals(child.name, W.r) || nameEquals(child.name, W.hyperlink) || nameEquals(child.name, W.smartTag) || nameEquals(child.name, W.sdt)) {
    return wrapRunLevel(child, kind, state);
  }
  if (nameEquals(child.name, W.sdtContent)) {
    return cloneElement(child, { children: childrenElements(child).map((inner) => wrapContainerChild(inner, kind, state)) });
  }
  return cloneElement(child);
}

function expandFieldForRevision(runLevel: XElement): XElement[] {
  if (!nameEquals(runLevel.name, W.fldSimple)) return [runLevel];
  const instr = runLevel.attribute(W.instr) ?? '';
  return [
    element(W.r, [], [element(W.fldChar, [attr(W.fldCharType, 'begin')])]),
    element(W.r, [], [element(W.instrText, [attr(XML_SPACE, 'preserve')], [instr])]),
    element(W.r, [], [element(W.fldChar, [attr(W.fldCharType, 'separate')])]),
    ...childrenElements(runLevel).map((c) => cloneElement(c)),
    element(W.r, [], [element(W.fldChar, [attr(W.fldCharType, 'end')])]),
  ];
}

function markParagraphMark(para: XElement, kind: RevKind, state: RenderState): void {
  let pPr = para.element(W.pPr);
  if (!pPr) {
    pPr = element(W.pPr);
    para.children.unshift(pPr);
    pPr.parent = para;
  }
  let rPr = pPr.element(W.rPr);
  if (!rPr) {
    rPr = element(W.rPr);
    const idx = pPr.children.findIndex((c) => c instanceof XElement && (nameEquals(c.name, W.sectPr) || nameEquals(c.name, W.pPrChange)));
    pPr.children.splice(idx >= 0 ? idx : pPr.children.length, 0, rPr);
    rPr.parent = pPr;
  }
  rPr.children.splice(0, rPr.children.length, element(isDeleteGrade(kind) ? W.del : W.ins, revisionAttrs(state)), ...rPr.children.filter((c) => !(c instanceof XElement && (nameEquals(c.name, W.ins) || nameEquals(c.name, W.del)))));
  for (const child of rPr.children) if (child instanceof XElement) child.parent = rPr;
}

function renderModifiedParagraph(op: IrEditOp, tokenDiff: IrTokenDiff, state: RenderState, sink: XElement[]): void {
  const leftPara = sourceElement(op.leftAnchor, state.left);
  const rightPara = sourceElement(op.rightAnchor, state.right);
  if (!leftPara || !rightPara) return;
  const newPara = element(W.p);
  const rightPPr = rightPara.element(W.pPr);
  if (rightPPr) newPara.children.push(stripUnids(cloneElement(rightPPr)));
  applyBlockFormatChanges(newPara, leftPara, rightPara, state);
  const leftRuns = new SourceRunModel(leftPara);
  const rightRuns = new SourceRunModel(rightPara);
  const leftTokens = paragraphTokens(op.leftAnchor, state.left, state.settings);
  const rightTokens = paragraphTokens(op.rightAnchor, state.right, state.settings);
  newPara.children.push(...buildTokenOpContent(tokenDiff, leftTokens, rightTokens, leftRuns, rightRuns, state));
  sink.push(newPara);
}

function buildTokenOpContent(
  diff: IrTokenDiff,
  leftTokens: ReadonlyArray<IrDiffToken>,
  rightTokens: ReadonlyArray<IrDiffToken>,
  leftRuns: SourceRunModel,
  rightRuns: SourceRunModel,
  state: RenderState,
): XElement[] {
  const content: XElement[] = [];
  for (const op of diff.ops) {
    switch (op.kind) {
      case 'Equal': {
        const [rs, re] = rightSpanChars(rightTokens, op);
        const [ls, le] = leftSpanChars(leftTokens, op);
        const [rzs, rze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        const [lzs, lze] = zeroWidthBoundaries(leftTokens, op.leftStart, op.leftEnd);
        if (rawSpanText(leftTokens, op.leftStart, op.leftEnd) !== rawSpanText(rightTokens, op.rightStart, op.rightEnd)) {
          for (const r of leftRuns.slice(ls, le, lzs, lze)) content.push(...wrapFieldAware(r, 'Del', state));
          for (const r of rightRuns.slice(rs, re, rzs, rze)) content.push(...wrapFieldAware(r, 'Ins', state));
        } else {
          for (const r of rightRuns.slice(rs, re, rzs, rze)) {
            registerRightSourcedClone(r, state);
            content.push(r);
          }
        }
        break;
      }
      case 'FormatChanged': {
        const [rs, re] = rightSpanChars(rightTokens, op);
        const [ls] = leftSpanChars(leftTokens, op);
        const [zs, ze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        // Text-equal span: the left char at offset k matches the right char at offset k, so each
        // emitted right run's OLD rPr comes from the left run at the ALIGNED char — a cursor per
        // stamped run, not the span start (runs inside the span can carry different left formats).
        let cursor = rs;
        for (const r of rightRuns.slice(rs, re, zs, ze)) {
          registerRightSourcedClone(r, state);
          for (const run of runsForFormatStamp(r)) {
            applyRPrChange(run, leftRuns.rPrAtChar(ls + (cursor - rs)), state);
            cursor += runTextLength(run);
          }
          content.push(r);
        }
        break;
      }
      case 'Insert': {
        const [s, e] = rightSpanChars(rightTokens, op);
        const [zs, ze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        for (const r of rightRuns.slice(s, e, zs, ze)) content.push(...wrapFieldAware(r, 'Ins', state));
        break;
      }
      case 'Delete': {
        const [s, e] = leftSpanChars(leftTokens, op);
        const [zs, ze] = zeroWidthBoundaries(leftTokens, op.leftStart, op.leftEnd);
        for (const r of leftRuns.slice(s, e, zs, ze)) content.push(...wrapFieldAware(r, 'Del', state));
        break;
      }
    }
  }
  return coalesceAdjacentHyperlinks(content);
}

function emitFormatOnlyParagraph(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const leftPara = sourceElement(op.leftAnchor, state.left);
  const rightPara = sourceElement(op.rightAnchor, state.right);
  if (!leftPara || !rightPara || !nameEquals(leftPara.name, W.p) || !nameEquals(rightPara.name, W.p)) {
    emitVerbatim(op.rightAnchor, state.right, sink);
    return;
  }
  const leftRuns = new SourceRunModel(leftPara);
  const out = stripUnids(cloneElement(rightPara));
  applyBlockFormatChanges(out, leftPara, rightPara, state);
  for (const child of childrenElements(out).filter((e) => !nameEquals(e.name, W.pPr))) registerRightSourcedClone(child, state);
  let cursor = 0;
  for (const child of childrenElements(out).filter((e) => !nameEquals(e.name, W.pPr))) {
    for (const run of runsForFormatStamp(child)) {
      applyRPrChange(run, leftRuns.rPrAtChar(cursor), state);
      cursor += runTextLength(run);
    }
  }
  sink.push(out);
}

function emitFormatOnlyTable(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const leftTbl = sourceElement(op.leftAnchor, state.left);
  const rightTbl = sourceElement(op.rightAnchor, state.right);
  if (!leftTbl || !rightTbl) return emitVerbatim(op.rightAnchor, state.right, sink);
  const out = stripUnids(cloneElement(rightTbl));
  registerRightSourcedClone(out, state);
  applyTableLevelShellChanges(out, leftTbl, state);
  sink.push(out);
}

function renderModifiedTable(op: IrEditOp, tableDiff: IrTableDiff, state: RenderState, sink: XElement[]): boolean {
  const leftTblSrc = sourceElement(op.leftAnchor, state.left);
  const rightTblSrc = sourceElement(op.rightAnchor, state.right);
  if (!leftTblSrc || !rightTblSrc) return false;
  const leftRows = indexRows(resolveBlock(op.leftAnchor, state.left) as IrTable | undefined);
  const rightRows = indexRows(resolveBlock(op.rightAnchor, state.right) as IrTable | undefined);
  const newTbl = element(W.tbl);
  newTbl.children.push(...childrenElements(rightTblSrc).filter((e) => !nameEquals(e.name, W.tr)).map((e) => stripUnids(cloneElement(e))));
  applyTableLevelShellChanges(newTbl, leftTblSrc, state);
  for (const rowOp of tableDiff.rowOps) {
    if (!renderRowOp(rowOp, leftRows, rightRows, state, newTbl)) return false;
  }
  sink.push(newTbl);
  return true;
}

function renderRowOp(rowOp: IrRowOp, leftRows: Map<string, XElement>, rightRows: Map<string, XElement>, state: RenderState, newTbl: XElement): boolean {
  if (rowOp.kind === 'EqualRow') {
    const src = rightRows.get(rowOp.rightRowAnchor ?? '');
    if (!src) return false;
    let row = stripUnids(cloneElement(src));
    registerRightSourcedClone(row, state);
    const left = rowOp.leftRowAnchor ? leftRows.get(rowOp.leftRowAnchor) : undefined;
    if (left) row = applyRowAndCellShellChanges(row, left, state);
    newTbl.children.push(row);
    return true;
  }
  if (rowOp.kind === 'InsertRow' || (rowOp.kind === 'MovedRow' && !rowOp.isMoveSource)) {
    const src = rightRows.get(rowOp.rightRowAnchor ?? '');
    if (!src) return false;
    const row = markWholeRow(stripUnids(cloneElement(src)), 'Ins', state);
    registerRightSourcedClone(row, state);
    newTbl.children.push(row);
    return true;
  }
  if (rowOp.kind === 'DeleteRow' || (rowOp.kind === 'MovedRow' && rowOp.isMoveSource)) {
    const src = leftRows.get(rowOp.leftRowAnchor ?? '');
    if (!src) return false;
    newTbl.children.push(markWholeRow(stripUnids(cloneElement(src)), 'Del', state));
    return true;
  }
  if (rowOp.kind === 'ModifyRow') return renderModifyRow(rowOp, leftRows, rightRows, state, newTbl);
  return false;
}

function renderModifyRow(rowOp: IrRowOp, leftRows: Map<string, XElement>, rightRows: Map<string, XElement>, state: RenderState, newTbl: XElement): boolean {
  const rightSrc = rightRows.get(rowOp.rightRowAnchor ?? '');
  if (!rightSrc) return false;
  const leftSrc = rowOp.leftRowAnchor ? leftRows.get(rowOp.leftRowAnchor) : undefined;
  if (!rowOp.cellOps) {
    let row = stripUnids(cloneElement(rightSrc));
    registerRightSourcedClone(row, state);
    if (leftSrc) row = applyRowAndCellShellChanges(row, leftSrc, state);
    newTbl.children.push(row);
    return true;
  }
  if (rowOp.cellOps.some((c) => !c.leftCellAnchor || !c.rightCellAnchor)) return false;
  const newRow = element(W.tr);
  newRow.children.push(...childrenElements(rightSrc).filter((e) => !nameEquals(e.name, W.tc)).map((e) => stripUnids(cloneElement(e))));
  const rightCells = childrenElements(rightSrc, W.tc);
  let ci = 0;
  for (const cellOp of rowOp.cellOps) {
    const cellSrc = rightCells[ci++];
    if (!cellSrc) break;
    newRow.children.push(renderCell(cellOp, cellSrc, state));
  }
  for (; ci < rightCells.length; ci++) newRow.children.push(stripUnids(cloneElement(rightCells[ci]!)));
  registerRightSourcedClone(newRow, state);
  const out = leftSrc ? applyRowAndCellShellChanges(newRow, leftSrc, state) : newRow;
  newTbl.children.push(out);
  return true;
}

function renderCell(cellOp: IrCellOp, cellSrc: XElement, state: RenderState): XElement {
  const newCell = element(W.tc);
  newCell.children.push(...childrenElements(cellSrc).filter((e) => !nameEquals(e.name, W.p) && !nameEquals(e.name, W.tbl)).map((e) => stripUnids(cloneElement(e))));
  if (cellOp.blockOps) {
    const cellSink: XElement[] = [];
    for (const bop of cellOp.blockOps) renderBlockOp(bop, state, cellSink);
    if (cellSink.length > 0) newCell.children.push(...cellSink);
  }
  if (newCell.children.every((c) => !(c instanceof XElement && (nameEquals(c.name, W.p) || nameEquals(c.name, W.tbl))))) {
    newCell.children.push(...childrenElements(cellSrc).filter((e) => nameEquals(e.name, W.p) || nameEquals(e.name, W.tbl)).map((e) => stripUnids(cloneElement(e))));
  }
  return newCell;
}

function renderSplitBlock(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const leftPara = sourceElement(op.leftAnchor, state.left);
  const anchors = op.splitMergeAnchors;
  const diffs = op.segmentDiffs as ReadonlyArray<IrTokenDiff> | null | undefined;
  if (!leftPara || !anchors || !diffs || anchors.length !== diffs.length) {
    emitWholeBlock(op.leftAnchor, state.left, state, sink, 'Del');
    for (const a of anchors ?? []) emitWholeBlock(a, state.right, state, sink, 'Ins');
    return;
  }
  const leftRuns = new SourceRunModel(leftPara);
  const leftTokens = paragraphTokens(op.leftAnchor, state.left, state.settings);
  let offset = 0;
  for (let s = 0; s < anchors.length; s++) {
    const diff = diffs[s]!;
    const sliceLen = segmentSliceLength(diff, true);
    const slice = leftTokens.slice(offset, offset + sliceLen);
    offset += sliceLen;
    const memberPara = sourceElement(anchors[s]!, state.right);
    if (!memberPara) { emitWholeBlock(anchors[s]!, state.right, state, sink, 'Ins'); continue; }
    const newPara = element(W.p);
    const pPr = memberPara.element(W.pPr);
    if (pPr) newPara.children.push(stripUnids(cloneElement(pPr)));
    newPara.children.push(...buildTokenOpContent(diff, slice, paragraphTokens(anchors[s]!, state.right, state.settings), leftRuns, new SourceRunModel(memberPara), state));
    if (s < anchors.length - 1) markParagraphMark(newPara, 'Ins', state);
    sink.push(newPara);
  }
}

function renderMergeBlock(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const rightPara = sourceElement(op.rightAnchor, state.right);
  const anchors = op.splitMergeAnchors;
  const diffs = op.segmentDiffs as ReadonlyArray<IrTokenDiff> | null | undefined;
  if (!rightPara || !anchors || !diffs || anchors.length !== diffs.length) {
    for (const a of anchors ?? []) emitWholeBlock(a, state.left, state, sink, 'Del');
    emitWholeBlock(op.rightAnchor, state.right, state, sink, 'Ins');
    return;
  }
  const rightTokens = paragraphTokens(op.rightAnchor, state.right, state.settings);
  let offset = 0;
  for (let m = 0; m < anchors.length; m++) {
    const diff = diffs[m]!;
    const sliceLen = segmentSliceLength(diff, false);
    const slice = rightTokens.slice(offset, offset + sliceLen);
    offset += sliceLen;
    const memberPara = sourceElement(anchors[m]!, state.left);
    if (!memberPara) { emitWholeBlock(anchors[m]!, state.left, state, sink, 'Del'); continue; }
    const newPara = element(W.p);
    const pPr = (m === anchors.length - 1 ? rightPara : memberPara).element(W.pPr);
    if (pPr) newPara.children.push(stripUnids(cloneElement(pPr)));
    newPara.children.push(...buildTokenOpContent(diff, paragraphTokens(anchors[m]!, state.left, state.settings), slice, new SourceRunModel(memberPara), new SourceRunModel(rightPara), state));
    if (m < anchors.length - 1) markParagraphMark(newPara, 'Del', state);
    sink.push(newPara);
  }
}

function emitMoveSource(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const src = sourceElement(op.leftAnchor, state.left);
  if (!src || !op.moveGroupId || !nameEquals(src.name, W.p)) return emitWholeBlock(op.leftAnchor, state.left, state, sink, 'Del');
  const para = stripUnids(cloneElement(src));
  markWholeParagraph(para, 'MoveFrom', state);
  bracketParagraphWithMoveRange(para, true, moveName(op.moveGroupId, state), state);
  sink.push(para);
}

function emitMoveDestination(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  const src = sourceElement(op.rightAnchor, state.right);
  if (!src || !op.moveGroupId || !nameEquals(src.name, W.p)) return emitWholeBlock(op.rightAnchor, state.right, state, sink, 'Ins');
  if (op.kind === 'MoveModifyBlock' && op.tokenDiff && !op.textboxDiffs) {
    const para = buildMoveModifyDestination(op, op.tokenDiff as IrTokenDiff, state);
    if (para) {
      markParagraphMark(para, 'MoveTo', state);
      bracketParagraphWithMoveRange(para, false, moveName(op.moveGroupId, state), state);
      sink.push(para);
      return;
    }
  }
  const dest = stripUnids(cloneElement(src));
  registerRightSourcedClone(dest, state);
  const left = sourceElement(op.leftAnchor, state.left);
  if (left) applyBlockFormatChanges(dest, left, src, state);
  markWholeParagraph(dest, 'MoveTo', state);
  bracketParagraphWithMoveRange(dest, false, moveName(op.moveGroupId, state), state);
  sink.push(dest);
}

function buildMoveModifyDestination(op: IrEditOp, tokenDiff: IrTokenDiff, state: RenderState): XElement | null {
  const leftPara = sourceElement(op.leftAnchor, state.left);
  const rightPara = sourceElement(op.rightAnchor, state.right);
  if (!leftPara || !rightPara) return null;
  const newPara = element(W.p);
  const pPr = rightPara.element(W.pPr);
  if (pPr) newPara.children.push(stripUnids(cloneElement(pPr)));
  applyBlockFormatChanges(newPara, leftPara, rightPara, state);
  const leftRuns = new SourceRunModel(leftPara);
  const rightRuns = new SourceRunModel(rightPara);
  const leftTokens = paragraphTokens(op.leftAnchor, state.left, state.settings);
  const rightTokens = paragraphTokens(op.rightAnchor, state.right, state.settings);
  for (const tokenOp of tokenDiff.ops) {
    if (tokenOp.kind === 'Equal' || tokenOp.kind === 'FormatChanged') {
      const [s, e] = rightSpanChars(rightTokens, tokenOp);
      const [zs, ze] = zeroWidthBoundaries(rightTokens, tokenOp.rightStart, tokenOp.rightEnd);
      for (const r of rightRuns.slice(s, e, zs, ze)) newPara.children.push(...wrapFieldAware(r, 'MoveTo', state));
    } else if (tokenOp.kind === 'Insert') {
      const [s, e] = rightSpanChars(rightTokens, tokenOp);
      const [zs, ze] = zeroWidthBoundaries(rightTokens, tokenOp.rightStart, tokenOp.rightEnd);
      for (const r of rightRuns.slice(s, e, zs, ze)) newPara.children.push(...wrapFieldAware(r, 'Ins', state));
    } else {
      const [s, e] = leftSpanChars(leftTokens, tokenOp);
      const [zs, ze] = zeroWidthBoundaries(leftTokens, tokenOp.leftStart, tokenOp.leftEnd);
      for (const r of leftRuns.slice(s, e, zs, ze)) newPara.children.push(...wrapFieldAware(r, 'Del', state));
    }
  }
  return newPara;
}

function bracketParagraphWithMoveRange(para: XElement, isFrom: boolean, name: string, state: RenderState): void {
  const id = nextId(state);
  const start = element(isFrom ? W.moveFromRangeStart : W.moveToRangeStart, [
    attr(W.id, id),
    attr(W.name, name),
    attr(W.author, state.settings.authorForRevisions),
    attr(W.date, state.settings.dateTimeForRevisions),
  ]);
  const end = element(isFrom ? W.moveFromRangeEnd : W.moveToRangeEnd, [attr(W.id, id)]);
  const pPrIndex = para.children.findIndex((c) => c instanceof XElement && nameEquals(c.name, W.pPr));
  para.children.splice(pPrIndex >= 0 ? pPrIndex + 1 : 0, 0, start);
  para.children.push(end);
}

class SourceRunModel {
  private readonly segments: Segment[] = [];
  private readonly claimed = new Set<XElement>();
  private readonly hyperlinkOrdinal = new Map<XElement, number>();
  private nextHyperlinkOrdinal = 0;

  constructor(para: XElement) {
    const offset = { value: 0 };
    for (const child of childrenElements(para).filter((e) => !nameEquals(e.name, W.pPr))) this.walkRunLevel(child, offset, []);
  }

  private walkRunLevel(runLevel: XElement, offset: { value: number }, chain: XElement[]): void {
    if (nameEquals(runLevel.name, W.r)) this.walkRun(runLevel, offset, chain);
    else if (nameEquals(runLevel.name, W.hyperlink)) {
      this.hyperlinkOrdinal.set(runLevel, this.nextHyperlinkOrdinal++);
      const childChain = [...chain, runLevel];
      const kids = childrenElements(runLevel);
      if (kids.length === 0) this.segments.push({ element: runLevel, start: offset.value, end: offset.value, kind: 'zero', chain: childChain, alwaysKeep: false });
      for (const child of kids) this.walkRunLevel(child, offset, childChain);
    } else if (nameEquals(runLevel.name, W.ins) || nameEquals(runLevel.name, W.del) || nameEquals(runLevel.name, W.sdt) || nameEquals(runLevel.name, W.smartTag)) {
      const start = offset.value;
      for (const t of runLevel.descendants(W.t)) offset.value += t.value.length;
      this.segments.push({ element: runLevel, start, end: offset.value, kind: 'container', chain, alwaysKeep: false });
    } else {
      this.segments.push({ element: runLevel, start: offset.value, end: offset.value, kind: 'zero', chain, alwaysKeep: isAlwaysKeepMarker(runLevel.name) });
    }
  }

  private walkRun(run: XElement, offset: { value: number }, chain: XElement[]): void {
    const kids = childrenElements(run).filter((e) => !nameEquals(e.name, W.rPr));
    if (kids.length === 0) this.segments.push({ element: run, start: offset.value, end: offset.value, kind: 'other', chain, alwaysKeep: false });
    for (const child of kids) {
      if (nameEquals(child.name, W.t)) {
        const start = offset.value;
        offset.value += child.value.length;
        this.segments.push({ element: run, start, end: offset.value, kind: 'text', textChild: child, chain, alwaysKeep: false });
      } else if (nameEquals(child.name, W.fldSimple) || isContainer(child.name)) {
        const start = offset.value;
        for (const t of child.descendants(W.t)) offset.value += t.value.length;
        this.segments.push({ element: run, start, end: offset.value, kind: 'other', otherChild: child, chain, alwaysKeep: false });
      } else if (nameEquals(child.name, W.noBreakHyphen) || nameEquals(child.name, W.softHyphen) || nameEquals(child.name, W.sym)) {
        const start = offset.value;
        offset.value += 1;
        this.segments.push({ element: run, start, end: offset.value, kind: 'other', otherChild: child, chain, alwaysKeep: false });
      } else {
        this.segments.push({ element: run, start: offset.value, end: offset.value, kind: 'other', otherChild: child, chain, alwaysKeep: fieldPlumbingKeep(child.name) || nameEquals(child.name, W.commentReference) });
      }
    }
  }

  slice(start: number, end: number, includeStartZeroWidth = true, includeEndZeroWidth = false): XElement[] {
    const result: XElement[] = [];
    let groupChain: XElement[] = [];
    let groupChildren: XElement[] = [];
    let currentRun: XElement | null = null;
    let rebuilt: XElement | null = null;

    const flushRun = () => {
      if (rebuilt && childrenElements(rebuilt).some((e) => !nameEquals(e.name, W.rPr))) groupChildren.push(rebuilt);
      rebuilt = null; currentRun = null;
    };
    const sameChain = (a: XElement[], b: XElement[]) => a.length === b.length && a.every((x, i) => x === b[i]);
    const flushGroup = () => {
      flushRun();
      if (groupChildren.length > 0) {
        let content = groupChildren;
        for (let i = groupChain.length - 1; i >= 0; i--) {
          const shell = groupChain[i]!;
          const attrs = [...shell.attributeInfos()];
          if (nameEquals(shell.name, W.hyperlink) && this.hyperlinkOrdinal.has(shell)) attrs.push(attr(SOURCE_LINK_ID, this.hyperlinkOrdinal.get(shell)!));
          content = [cloneElement(shell, { attributes: attrs, children: content })];
        }
        result.push(...content);
      }
      groupChildren = []; groupChain = [];
    };
    const startGroup = (chain: XElement[]) => {
      if ((groupChildren.length > 0 || currentRun) && !sameChain(groupChain, chain)) flushGroup();
      groupChain = chain;
    };

    for (const seg of this.segments) {
      let overlaps = false;
      if (start === end) overlaps = seg.start === start && seg.start === seg.end;
      else if (seg.start === seg.end && seg.alwaysKeep) overlaps = seg.start >= start && seg.start <= end;
      else if (seg.start === seg.end) overlaps = (seg.start > start && seg.start < end) || (seg.start === start && includeStartZeroWidth) || (seg.start === end && includeEndZeroWidth);
      else overlaps = seg.start < end && seg.end > start;
      if (overlaps && seg.start === seg.end && seg.kind !== 'container') {
        const key = seg.otherChild ?? seg.element;
        if (this.claimed.has(key)) overlaps = false;
        else this.claimed.add(key);
      }
      if (overlaps && seg.kind === 'container') {
        if (this.claimed.has(seg.element)) overlaps = false;
        else this.claimed.add(seg.element);
      }
      if (!overlaps) {
        if (seg.kind === 'container' || seg.kind === 'zero') flushRun();
        continue;
      }
      if (seg.kind === 'zero' || seg.kind === 'container') {
        startGroup(seg.chain); flushRun(); groupChildren.push(cloneElement(seg.element)); continue;
      }
      startGroup(seg.chain);
      if (currentRun !== seg.element) {
        flushRun(); currentRun = seg.element; rebuilt = element(W.r);
        const rPr = seg.element.element(W.rPr);
        if (rPr) rebuilt.children.push(cloneElement(rPr));
      }
      if (seg.kind === 'text' && seg.textChild) {
        const s = Math.max(start, seg.start), e = Math.min(end, seg.end);
        const piece = seg.textChild.value.slice(s - seg.start, e - seg.start);
        const attrs = preserveSpace(piece) ? [attr(XML_SPACE, 'preserve')] : [];
        rebuilt!.children.push(element(W.t, attrs, [piece]));
      } else if (seg.otherChild) rebuilt!.children.push(cloneElement(seg.otherChild));
    }
    flushGroup();
    return result;
  }

  rPrAtChar(at: number): XElement | null {
    const hit = this.segments.find((s) => s.kind === 'text' && s.start <= at && at < s.end) ?? this.segments.find((s) => s.start <= at && (at < s.end || (s.start === s.end && s.start === at)));
    const rPr = hit?.element.element(W.rPr);
    return rPr ? cloneElement(rPr) : null;
  }
}

interface Segment {
  readonly element: XElement;
  readonly start: number;
  readonly end: number;
  readonly kind: 'text' | 'other' | 'zero' | 'container';
  readonly textChild?: XElement;
  readonly otherChild?: XElement;
  readonly chain: XElement[];
  readonly alwaysKeep: boolean;
}

function convertTextToDelText(el: XElement): XElement {
  return transformDescendants(el, (node) => {
    if (nameEquals(node.name, W.t)) return replaceName(node, W.delText);
    if (nameEquals(node.name, W.instrText)) return replaceName(node, W.delInstrText);
    return null;
  });
}

function applyRPrChange(run: XElement, oldRPr: XElement | null, state: RenderState): void {
  let rPr = run.element(W.rPr);
  if (!rPr) {
    rPr = element(W.rPr);
    run.children.unshift(rPr);
    rPr.parent = run;
  }
  rPr.children.splice(0, rPr.children.length, ...rPr.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, W.rPrChange))), element(W.rPrChange, revisionAttrs(state), [oldRPr ? cloneElement(oldRPr) : element(W.rPr)]));
}

function applyBlockFormatChanges(newPara: XElement, leftPara: XElement, rightPara: XElement, state: RenderState): void {
  if (!state.settings.trackParagraphFormatChanges) return;
  const leftPPr = leftPara.element(W.pPr);
  const rightPPr = rightPara.element(W.pPr);
  if (compareXml(pPrForCompare(leftPPr)) === compareXml(pPrForCompare(rightPPr)) && compareXml(markRPrForCompare(leftPPr)) === compareXml(markRPrForCompare(rightPPr))) return;
  let pPr = newPara.element(W.pPr);
  if (!pPr) {
    pPr = element(W.pPr);
    newPara.children.unshift(pPr);
  }
  if (compareXml(markRPrForCompare(leftPPr)) !== compareXml(markRPrForCompare(rightPPr))) {
    let markRPr = pPr.element(W.rPr);
    if (!markRPr) {
      markRPr = element(W.rPr);
      pPr.children.push(markRPr);
    }
    markRPr.children.push(element(W.rPrChange, revisionAttrs(state), [markRPrForCompare(leftPPr)]));
  }
  if (compareXml(pPrForCompare(leftPPr)) !== compareXml(pPrForCompare(rightPPr))) {
    pPr.children.splice(0, pPr.children.length, ...pPr.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, W.pPrChange))));
    pPr.children.push(element(W.pPrChange, revisionAttrs(state), [pPrForCompare(leftPPr)]));
  }
}

function applyTableLevelShellChanges(newTbl: XElement, leftTbl: XElement, state: RenderState): void {
  if (!state.settings.trackBlockFormatChanges) return;
  applyShellChange(newTbl, W.tblPr, W.tblPrChange, leftTbl.element(W.tblPr), state, false, [W.tblPrChange]);
  applyShellChange(newTbl, W.tblGrid, W.tblGridChange, leftTbl.element(W.tblGrid), state, true, [W.tblGridChange]);
}

function applyRowAndCellShellChanges(newRow: XElement, leftRow: XElement, state: RenderState): XElement {
  if (!state.settings.trackBlockFormatChanges) return newRow;
  applyShellChange(newRow, W.trPr, W.trPrChange, leftRow.element(W.trPr), state, false, [W.ins, W.del, W.trPrChange]);
  applyShellChange(newRow, W.tblPrEx, W.tblPrExChange, leftRow.element(W.tblPrEx), state, false, [W.tblPrExChange]);
  const leftCells = childrenElements(leftRow, W.tc);
  const newCells = childrenElements(newRow, W.tc);
  for (let i = 0; i < leftCells.length && i < newCells.length; i++) applyShellChange(newCells[i]!, W.tcPr, W.tcPrChange, leftCells[i]!.element(W.tcPr), state, false, [W.cellIns, W.cellDel, W.cellMerge, W.tcPrChange]);
  return newRow;
}

function applyShellChange(host: XElement, shellName: XName, changeName: XName, leftShell: XElement | null, state: RenderState, idOnly: boolean, exclude: readonly XName[]): void {
  let rightShell = host.element(shellName);
  if (compareXml(cleanShell(leftShell, changeName, exclude)) === compareXml(cleanShell(rightShell, changeName, exclude))) return;
  if (!rightShell) {
    rightShell = element(shellName);
    insertShellInSchemaOrder(host, rightShell, shellName);
  }
  rightShell.children.splice(0, rightShell.children.length, ...rightShell.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, changeName))));
  const inner = leftShell ? cloneElement(leftShell, { children: childrenElements(leftShell).filter((e) => !nameEquals(e.name, changeName) && !exclude.some((x) => nameEquals(x, e.name))).map((e) => stripUnids(cloneElement(e))) }) : element(shellName);
  rightShell.children.push(element(changeName, idOnly ? [attr(W.id, nextId(state))] : revisionAttrs(state), [inner]));
}

function insertShellInSchemaOrder(host: XElement, shell: XElement, shellName: XName): void {
  if (nameEquals(shellName, W.tblGrid)) {
    const tblPrIndex = host.children.findIndex((c) => c instanceof XElement && nameEquals(c.name, W.tblPr));
    if (tblPrIndex >= 0) {
      host.children.splice(tblPrIndex + 1, 0, shell);
      return;
    }
  } else if (nameEquals(shellName, W.trPr)) {
    const tblPrExIndex = host.children.findIndex((c) => c instanceof XElement && nameEquals(c.name, W.tblPrEx));
    if (tblPrExIndex >= 0) {
      host.children.splice(tblPrExIndex + 1, 0, shell);
      return;
    }
  }
  host.children.unshift(shell);
}

function applySectPrChange(output: XElement, oldSect: XElement, rightSect: XElement, state: RenderState): void {
  const inner = element(W.sectPr, [], childrenElements(oldSect).filter(isSectPrProp).map((e) => stripUnids(cloneElement(e))));
  const refs = output.children.filter((c) => !(c instanceof XElement) || !isSectPrProp(c));
  const rightProps = childrenElements(rightSect).filter(isSectPrProp).map((e) => stripUnids(cloneElement(e)));
  output.children.splice(0, output.children.length, ...refs, ...rightProps, element(W.sectPrChange, revisionAttrs(state), [inner]));
}

function stripUnids<T extends XElement>(el: T): T {
  for (const node of descendantsAndSelf(el)) {
    const attrs = [...node.attributeInfos()].filter((a) => a.name.ns !== PT_NS);
    if (attrs.length !== [...node.attributeInfos()].length) {
      const replacement = cloneElement(node, { attributes: attrs, children: node.children });
      node.children.splice(0, node.children.length, ...replacement.children);
      (node as unknown as { attrInfos?: readonly XAttributeInfo[] });
    }
  }
  return transformDescendants(el, (node) => {
    const attrs = [...node.attributeInfos()].filter((a) => a.name.ns !== PT_NS);
    return attrs.length === [...node.attributeInfos()].length ? null : cloneElement(node, { attributes: attrs, children: node.children });
  }) as T;
}

function replaceChild(root: XElement, oldChild: XElement, newChild: XElement): XElement {
  return cloneElement(root, { children: root.children.map((c) => c === oldChild ? newChild : c instanceof XElement ? replaceChild(c, oldChild, newChild) : c) });
}

function transformDescendants(el: XElement, transform: (el: XElement) => XElement | null): XElement {
  const transformedChildren = el.children.map((c) => c instanceof XElement ? transformDescendants(c, transform) : c);
  const cloned = cloneElement(el, { children: transformedChildren });
  return transform(cloned) ?? cloned;
}

function linqOutputShape(el: XElement): XElement {
  return transformDescendants(el, (node) => cloneElement(node, {
    children: node.children,
    emptyElementSpace: node.children.length === 0,
  }));
}

function remapHyperlinkRelationshipCollisions(
  root: XElement,
  outParts: ReadonlyMap<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
): { readonly root: XElement; readonly relationshipsXml: string | null } {
  const leftRelsBytes = outParts.get('word/_rels/document.xml.rels');
  const leftRels = leftRelsBytes ? parseXml(strFromU8(leftRelsBytes)) : null;
  const rightRels = relationshipRoot(rightParts);
  if (!leftRels || !rightRels) return { root, relationshipsXml: null };

  const leftById = hyperlinkRelationships(leftRels);
  const rightById = hyperlinkRelationships(rightRels);
  const used = new Set([...relationshipsById(leftRels).keys()]);
  const remap = new Map<string, string>();
  const newRelationships: XElement[] = [];
  for (const [id, rightRel] of rightById) {
    const leftRel = leftById.get(id);
    if (!leftRel) continue;
    if (leftRel.attribute(REL_TARGET) === rightRel.attribute(REL_TARGET)) continue;
    const fresh = freshRelationshipId(used);
    used.add(fresh);
    remap.set(id, fresh);
    newRelationships.push(cloneElement(rightRel, {
      attributes: [...rightRel.attributeInfos()].map((a) => nameEquals(a.name, REL_ID) ? { ...a, value: fresh } : a),
    }));
  }
  if (remap.size === 0) return { root, relationshipsXml: null };

  const rewrittenRoot = rewriteInsertedHyperlinkIds(root, remap);
  const rewrittenRels = cloneElement(leftRels, { children: [...leftRels.children, ...newRelationships] });
  return { root: rewrittenRoot, relationshipsXml: writeRelationshipsPart(rewrittenRels) };
}

// Relationship attribute names the media import remaps (C# ComparisonUnitWord.s_RelationshipAttributeNames).
const MEDIA_REL_LOCALS = new Set(['embed', 'link', 'id', 'cs', 'dm', 'lo', 'qs', 'href', 'pict']);
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const C_EXTERNAL_DATA = xname(C_NS, 'externalData');

function isMediaRelationshipAttribute(name: XName): boolean {
  return name.ns === R_NS && MEDIA_REL_LOCALS.has(name.local);
}

/**
 * Import media parts referenced by the registered RIGHT-sourced clones into the output package —
 * the C# ImportRightSourcedMedia / WmlComparer.MoveRelatedPartsToDestination semantics: walk each
 * registered clone's DESCENDANTS (self excluded) in registry order; every relationship attribute
 * that resolves in the RIGHT part's relationships imports a FRESH copy of its target part (per
 * occurrence, no dedup) under a generated P-name, appends a fresh relationship, and remaps the
 * attribute IN PLACE. w:headerReference/w:footerReference elements are skipped (left parts are
 * authoritative), as are c:externalData and External-mode relationships. An attribute already
 * remapped by an earlier (containing) clone no longer resolves in the right rels — skipped, which
 * is exactly the C# dangling-relationship skip.
 */
function importRightSourcedMedia(
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
): void {
  if (state.rightSourcedClones.length === 0) return;
  const rightRelsRoot = relationshipRootForPart(rightParts, 'word/document.xml');
  if (!rightRelsRoot) return;
  const rightRels = relationshipsById(rightRelsRoot);
  const destRelsName = relsPartName('word/document.xml');
  const destRelsRoot = ensureRelationshipsRoot(outParts, destRelsName);
  const used = new Set(relationshipsById(destRelsRoot).keys());
  let changed = false;
  for (const clone of state.rightSourcedClones) {
    if (importCloneMedia(clone, rightRels, 'word/document.xml', destRelsRoot, used, outParts, rightParts, true)) changed = true;
  }
  if (changed) outParts.set(destRelsName, strToU8(writeRelationshipsPart(destRelsRoot)));
}

function importCloneMedia(
  clone: XElement,
  sourceRels: ReadonlyMap<string, XElement>,
  sourcePartName: string,
  destRelsRoot: XElement,
  used: Set<string>,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  skipHeaderFooterReferences: boolean,
): boolean {
  let changed = false;
  for (const el of clone.descendants()) {
    if (nameEquals(el.name, C_EXTERNAL_DATA)) continue;
    if (skipHeaderFooterReferences && (nameEquals(el.name, W.headerReference) || nameEquals(el.name, W.footerReference))) continue;
    for (const a of [...el.attributeInfos()]) {
      if (!isMediaRelationshipAttribute(a.name)) continue;
      const rel = sourceRels.get(a.value);
      if (!rel || rel.attribute(REL_TARGET_MODE) === 'External') continue;
      const newId = importPartOccurrence(outParts, rightParts, sourcePartName, rel, destRelsRoot, used);
      if (!newId) continue;
      el.setAttributeValue(a.name, newId);
      changed = true;
    }
  }
  return changed;
}

/** Copy one relationship's target part from the RIGHT package as a FRESH part (P-name, original
 * extension), register its content type, append a fresh relationship to the destination rels root,
 * and — for XML parts — recursively import ITS related parts, then re-save it in the formatted
 * XDocument.Save shape (the oracle round-trips every imported XML part that way). */
function importPartOccurrence(
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  sourcePartName: string,
  rightRel: XElement,
  destRelsRoot: XElement,
  used: Set<string>,
): string | null {
  const target = rightRel.attribute(REL_TARGET);
  const type = rightRel.attribute(REL_TYPE);
  if (!target || !type) return null;
  const sourcePart = resolveTargetPart(sourcePartName, target);
  const bytes = rightParts[sourcePart];
  if (!bytes) return null;
  const contentType = contentTypeForPart(rightParts, sourcePart);
  const importedPart = freshImportedPartName(outParts, sourcePart);
  outParts.set(importedPart, bytes);
  if (contentType) addContentTypeForImportedPart(outParts, importedPart, contentType);

  const newId = freshImportedRelationshipId(used);
  used.add(newId);
  const newRelationship = packageRelationship(type, `/${importedPart}`, newId);
  destRelsRoot.children.push(newRelationship);
  newRelationship.parent = destRelsRoot;

  if (importedPart.toLowerCase().endsWith('.xml')) {
    const importedRoot = parseXml(strFromU8(bytes));
    importPartTreeMedia(importedRoot, outParts, rightParts, sourcePart, importedPart);
    outParts.set(importedPart, strToU8(writeXmlPartFormatted(importedRoot, xdocumentDeclaration(importedRoot))));
  }
  return newId;
}

/** The nested-import step of the C# recursion: walk the WHOLE imported part tree (no clone
 * registry at this depth), importing every resolvable media relationship per occurrence into the
 * imported part's OWN rels file. */
function importPartTreeMedia(
  partRoot: XElement,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  rightPartName: string,
  destPartName: string,
): void {
  const rightRelsRoot = relationshipRootForPart(rightParts, rightPartName);
  if (!rightRelsRoot) return;
  const rightRels = relationshipsById(rightRelsRoot);
  const destRelsName = relsPartName(destPartName);
  const destRelsRoot = ensureRelationshipsRoot(outParts, destRelsName);
  const used = new Set(relationshipsById(destRelsRoot).keys());
  if (importCloneMedia(partRoot, rightRels, rightPartName, destRelsRoot, used, outParts, rightParts, false)) {
    outParts.set(destRelsName, strToU8(writeRelationshipsPart(destRelsRoot)));
  }
}

function renderNoteScopes(
  script: IrEditScript,
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  leftParts: Record<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
): void {
  if (!script.noteOps || script.noteOps.length === 0) return;
  ensureNotePart(outParts, rightParts, 'word/footnotes.xml', W.footnote, W_FOOTNOTES);
  ensureNotePart(outParts, rightParts, 'word/endnotes.xml', W.endnote, W_ENDNOTES);
  applyNoteDiffs(script.noteOps.filter((n) => n.kind === 'Footnote'), state, outParts, leftParts, rightParts, 'word/footnotes.xml', W.footnote, W_FOOTNOTES);
  applyNoteDiffs(script.noteOps.filter((n) => n.kind === 'Endnote'), state, outParts, leftParts, rightParts, 'word/endnotes.xml', W.endnote, W_ENDNOTES);
}

function ensureNotePart(
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  partName: string,
  noteName: XName,
  rootName: XName,
): void {
  if (outParts.has(partName)) return;
  const rightRoot = rightParts[partName] ? parseXml(strFromU8(rightParts[partName]!)) : null;
  if (!rightRoot) return;
  ensureMainRelationship(outParts, partName, noteName);
  addContentTypeOverride(outParts, partName, nameEquals(noteName, W.footnote) ? CT_FOOTNOTES : CT_ENDNOTES);
  const root = element(rootName, [...rightRoot.attributeInfos()]);
  for (const n of rightRoot.elements(noteName)) {
    const id = n.attribute(W.id);
    if (id !== null && Number.isFinite(Number(id)) && Number(id) <= 0) root.children.push(cloneElement(n));
  }
  for (const child of root.children) if (child instanceof XElement) child.parent = root;
  outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(root))));
}

function applyNoteDiffs(
  diffs: ReadonlyArray<IrNoteDiff>,
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  leftParts: Record<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  partName: string,
  noteName: XName,
  rootName: XName,
): void {
  if (diffs.length === 0) return;
  const rightRoot = rightParts[partName] ? parseXml(strFromU8(rightParts[partName]!)) : null;
  let root = outParts.get(partName) ? parseXml(strFromU8(outParts.get(partName)!)) : null;
  if (!root) {
    if (!rightRoot) return;
    root = element(rootName, [...rightRoot.attributeInfos()]);
    for (const n of rightRoot.elements(noteName)) {
      const id = n.attribute(W.id);
      if (id !== null && Number.isFinite(Number(id)) && Number(id) <= 0) root.children.push(cloneElement(n));
    }
    for (const child of root.children) if (child instanceof XElement) child.parent = root;
  }

  let changed = false;
  for (const diff of diffs) {
    let noteEl = diff.leftNoteId ? childrenElements(root, noteName).find((e) => e.attribute(W.id) === diff.leftNoteId) ?? null : null;
    if (!noteEl) {
      const rightNote = rightRoot ? childrenElements(rightRoot, noteName).find((e) => e.attribute(W.id) === diff.noteId) ?? null : null;
      if (!rightNote) continue;
      noteEl = element(noteName, [...rightNote.attributeInfos()], childrenElements(rightNote).filter((e) => !nameEquals(e.name, W.p) && !nameEquals(e.name, W.tbl)).map((e) => stripUnids(cloneElement(e))));
      root.children.push(noteEl);
      noteEl.parent = root;
    }
    const blocks: XElement[] = [];
    for (const op of diff.ops) renderBlockOp(op, state, blocks);
    noteEl.children.splice(0, noteEl.children.length, ...noteEl.children.filter((c) => !(c instanceof XElement && (nameEquals(c.name, W.p) || nameEquals(c.name, W.tbl)))), ...blocks.map((b) => stripUnids(b)));
    for (const child of noteEl.children) if (child instanceof XElement) child.parent = noteEl;
    if (diff.leftNoteId !== null && diff.noteId !== diff.leftNoteId) {
      const rewritten = replaceElementAttributes(noteEl, W.id, diff.noteId);
      replaceElementInParent(noteEl, rewritten);
      noteEl = rewritten;
    }
    changed = true;
  }
  if (changed) outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(stripUnids(root)))));
}

// ----------------------------------------------------------------- header/footer story markup

/**
 * Apply header/footer story edit ops inside the output package's header/footer parts (the C#
 * RenderHeaderFooterScopes port). A MATCHED story (left part URI set) rebuilds that part's
 * block-level children from its ops with native markup; a RIGHT-only story creates a fresh part,
 * attaches a typed w:headerReference/w:footerReference to the target section's sectPr, and ensures
 * the story's visibility flag (w:titlePg / w:evenAndOddHeaders). Right-sourced clones inside a
 * story import their media relationships into THAT part (relationship ids are part-scoped).
 */
function renderHeaderFooterScopes(
  script: IrEditScript,
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  documentRoot: XElement,
): void {
  if (!script.headerFooterOps || script.headerFooterOps.length === 0) return;
  for (const diff of script.headerFooterOps) {
    if (diff.leftPartUri) applyHeaderFooterDiffToPart(diff, state, outParts, rightParts);
    else insertHeaderFooterStory(diff, state, outParts, rightParts, documentRoot);
  }
}

function applyHeaderFooterDiffToPart(
  diff: IrHeaderFooterDiff,
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
): void {
  const partName = diff.leftPartUri!.replace(/^\//, '');
  const bytes = outParts.get(partName);
  if (!bytes) return; // left part vanished (malformed input) — keep the carry-over
  const root = parseXml(strFromU8(bytes));

  // Slice the clone registry around this story's render so ONLY its clones import into this part.
  const clonesBefore = state.rightSourcedClones.length;
  const blocks: XElement[] = [];
  for (const op of diff.ops) renderBlockOp(op, state, blocks);

  const kept = root.children.filter((c) => !(c instanceof XElement && (nameEquals(c.name, W.p) || nameEquals(c.name, W.tbl))));
  root.children.splice(0, root.children.length, ...kept, ...blocks);
  for (const child of root.children) if (child instanceof XElement) child.parent = root;

  importStorySourcedRelationships(diff, state, clonesBefore, partName, outParts, rightParts);
  outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(stripUnids(root)))));
}

function insertHeaderFooterStory(
  diff: IrHeaderFooterDiff,
  state: RenderState,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
  documentRoot: XElement,
): void {
  if (!diff.rightPartUri) return;
  const rightPartName = diff.rightPartUri.replace(/^\//, '');
  const rightBytes = rightParts[rightPartName];
  if (!rightBytes) return;
  const rightRoot = parseXml(strFromU8(rightBytes));

  // Locate the target sectPr FIRST — a story that cannot attach must not create an orphan part.
  const body = documentRoot.element(W.body);
  if (!body) return;
  const sectPrs = [...body.descendants(W.sectPr)];
  if (diff.sectionIndex >= sectPrs.length) return;
  const sectPr = sectPrs[diff.sectionIndex]!;

  const partName = freshHeaderFooterPartName(outParts, diff.isHeader);
  const newRoot = element(rightRoot.name, [...rightRoot.attributeInfos()]);
  const clonesBefore = state.rightSourcedClones.length;
  const blocks: XElement[] = [];
  for (const op of diff.ops) renderBlockOp(op, state, blocks);
  newRoot.children.push(...blocks);
  for (const child of newRoot.children) if (child instanceof XElement) child.parent = newRoot;

  addContentTypeOverride(outParts, partName, diff.isHeader
    ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
  const relsName = relsPartName('word/document.xml');
  const relsRoot = ensureRelationshipsRoot(outParts, relsName);
  const used = new Set(relationshipsById(relsRoot).keys());
  const relId = freshImportedRelationshipId(used);
  const rel = packageRelationship(`${WORD_REL_PREFIX}${diff.isHeader ? 'header' : 'footer'}`, `/${partName}`, relId);
  relsRoot.children.push(rel);
  rel.parent = relsRoot;
  outParts.set(relsName, strToU8(writeRelationshipsPart(relsRoot)));

  importStorySourcedRelationships(diff, state, clonesBefore, partName, outParts, rightParts);
  outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(stripUnids(newRoot)))));

  // Attach the reference (header/footer references lead the CT_SectPr sequence, so prepending is
  // always schema-ordered) and ensure the story's visibility flag.
  const refName = diff.isHeader ? W.headerReference : W.footerReference;
  const typeValue = diff.kind === 'First' ? 'first' : diff.kind === 'Even' ? 'even' : 'default';
  const ref = element(refName, [attr(W.type, typeValue), attr(R.id, relId)]);
  sectPr.children.unshift(ref);
  ref.parent = sectPr;
  if (diff.kind === 'First' && !sectPr.element(xname(W_NS, 'titlePg'))) {
    insertIntoSectPr(sectPr, element(xname(W_NS, 'titlePg')));
  }
  if (diff.kind === 'Even') ensureEvenAndOddHeaders(outParts);
}

/** Elements that FOLLOW w:titlePg in the CT_SectPr sequence — an insertion lands before the first
 * of these (or at the end), keeping the sectPr schema-ordered. */
const SECT_PR_AFTER_TITLE_PG = ['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];

function insertIntoSectPr(sectPr: XElement, el: XElement): void {
  const idx = sectPr.children.findIndex((c) => c instanceof XElement && c.name.ns === W_NS && SECT_PR_AFTER_TITLE_PG.includes(c.name.local));
  sectPr.children.splice(idx >= 0 ? idx : sectPr.children.length, 0, el);
  el.parent = sectPr;
}

/** Ensure the settings part carries w:evenAndOddHeaders (required for an Even story to render). */
function ensureEvenAndOddHeaders(outParts: Map<string, Uint8Array>): void {
  const settingsName = 'word/settings.xml';
  const bytes = outParts.get(settingsName);
  const root = bytes ? parseXml(strFromU8(bytes)) : element(xname(W_NS, 'settings'), [attr(xname(XMLNS_NS, 'w'), W_NS)]);
  const flagName = xname(W_NS, 'evenAndOddHeaders');
  if (root.element(flagName)) return;
  const flag = element(flagName);
  root.children.push(flag);
  flag.parent = root;
  outParts.set(settingsName, strToU8(writeNoteXmlPart(linqOutputShape(root))));
}

function freshHeaderFooterPartName(outParts: ReadonlyMap<string, Uint8Array>, isHeader: boolean): string {
  const stem = isHeader ? 'header' : 'footer';
  let n = 1;
  while (outParts.has(`word/${stem}${n}.xml`)) n++;
  return `word/${stem}${n}.xml`;
}

/**
 * Import media parts referenced by the clones registered during ONE story's render (the
 * watermark marks the registry position before the story rendered) from the RIGHT story part into
 * the output story part. Relationship ids are part-scoped in OOXML, which is why the body's
 * main-part import cannot serve header/footer content.
 */
function importStorySourcedRelationships(
  diff: IrHeaderFooterDiff,
  state: RenderState,
  clonesBefore: number,
  outputPartName: string,
  outParts: Map<string, Uint8Array>,
  rightParts: Record<string, Uint8Array>,
): void {
  if (!diff.rightPartUri) return;
  const storyClones = state.rightSourcedClones.slice(clonesBefore);
  if (storyClones.length === 0) return;
  const rightPartName = diff.rightPartUri.replace(/^\//, '');
  const rightRelsRoot = relationshipRootForPart(rightParts, rightPartName);
  if (!rightRelsRoot) return;
  const rightRels = relationshipsById(rightRelsRoot);
  const destRelsName = relsPartName(outputPartName);
  const destRelsRoot = ensureRelationshipsRoot(outParts, destRelsName);
  const used = new Set(relationshipsById(destRelsRoot).keys());
  let changed = false;
  for (const clone of storyClones) {
    if (importCloneMedia(clone, rightRels, rightPartName, destRelsRoot, used, outParts, rightParts, true)) changed = true;
  }
  if (changed) outParts.set(destRelsName, strToU8(writeRelationshipsPart(destRelsRoot)));
}

function renumberNoteIds(
  documentRoot: XElement,
  outParts: Map<string, Uint8Array>,
  partName: string,
  refName: XName,
  noteName: XName,
  rootName: XName,
): Map<string, string> {
  const bytes = outParts.get(partName);
  if (!bytes) return new Map();
  const noteRoot = parseXml(strFromU8(bytes));
  const body = documentRoot.element(W.body);
  if (!body) return new Map();
  const refs = [...body.descendants(refName)];
  if (refs.length === 0) return new Map();

  const notes = childrenElements(noteRoot, noteName);
  const reserved = notes.filter(isReservedNote);
  const real = notes.filter((n) => !isReservedNote(n));
  const delDefs = real.filter(isDeletedOnlyNote);
  const liveById = new Map<string, XElement>();
  for (const n of real.filter((x) => !isDeletedOnlyNote(x))) {
    const id = n.attribute(W.id);
    if (id) liveById.set(id, n);
  }
  let delIndex = 0;
  let next = Math.max(0, ...reserved.map((n) => Number(n.attribute(W.id) ?? 0)).filter((n) => n > 0)) + 1;
  const assigned = new Map<XElement, string>();
  const remap = new Map<string, string>();
  const ordered: XElement[] = [];
  for (const ref of refs) {
    const oldId = ref.attribute(W.id);
    if (!oldId) continue;
    const isDel = hasAncestor(ref, W.del);
    const def = isDel ? (delDefs[delIndex++] ?? liveById.get(oldId) ?? null) : (liveById.get(oldId) ?? null);
    const existing = def ? assigned.get(def) : undefined;
    if (existing) {
      replaceElementInParent(ref, replaceElementAttributes(ref, W.id, existing));
      continue;
    }
    const newId = String(next++);
    replaceElementInParent(ref, replaceElementAttributes(ref, W.id, newId));
    if (def) {
      const defOldId = def.attribute(W.id);
      const rewritten = replaceElementAttributes(def, W.id, newId);
      replaceElementInParent(def, rewritten);
      assigned.set(def, newId);
      if (defOldId && defOldId !== newId) remap.set(defOldId, newId);
      ordered.push(rewritten);
    }
  }
  for (const n of real) if (![...assigned.keys()].includes(n) && !ordered.includes(n)) ordered.push(n);
  const nonNoteChildren = noteRoot.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, noteName)));
  noteRoot.children.splice(0, noteRoot.children.length, ...nonNoteChildren, ...reserved, ...ordered);
  for (const child of noteRoot.children) if (child instanceof XElement) child.parent = noteRoot;
  outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(noteRoot))));
  return remap;
}

function remapNestedNoteReferences(outParts: Map<string, Uint8Array>, footnoteRemap: ReadonlyMap<string, string>, endnoteRemap: ReadonlyMap<string, string>): void {
  if (footnoteRemap.size === 0 && endnoteRemap.size === 0) return;
  for (const partName of ['word/footnotes.xml', 'word/endnotes.xml']) {
    const bytes = outParts.get(partName);
    if (!bytes) continue;
    const root = parseXml(strFromU8(bytes));
    const rewritten = transformDescendants(root, (node) => {
      const remap = nameEquals(node.name, W.footnoteReference) ? footnoteRemap : nameEquals(node.name, W.endnoteReference) ? endnoteRemap : null;
      if (!remap) return null;
      const id = node.attribute(W.id);
      const mapped = id ? remap.get(id) : undefined;
      return mapped ? replaceElementAttributes(node, W.id, mapped) : null;
    });
    outParts.set(partName, strToU8(writeNoteXmlPart(linqOutputShape(rewritten))));
  }
}

function isReservedNote(note: XElement): boolean {
  const id = Number(note.attribute(W.id) ?? Number.NaN);
  return note.attribute(W.type) !== null || (Number.isFinite(id) && id <= 0);
}

function isDeletedOnlyNote(note: XElement): boolean {
  const hasDelText = [...note.descendants(W.delText)].length > 0;
  const hasLiveText = [...note.descendants(W.t)].some((t) => !hasAncestor(t, W.del));
  return hasDelText && !hasLiveText;
}

function hasAncestor(el: XElement, name: XName): boolean {
  for (let p = el.parent; p; p = p.parent) if (nameEquals(p.name, name)) return true;
  return false;
}

function replaceElementAttributes(el: XElement, name: XName, value: string): XElement {
  return withAttribute(el, name, value);
}

function xdocumentDeclaration(root: XElement): string {
  const standalone = root.documentProlog?.includes('standalone="yes"') ?? false;
  return standalone
    ? '<?xml version="1.0" encoding="utf-8" standalone="yes"?>'
    : '<?xml version="1.0" encoding="utf-8"?>';
}

function ensureRelationshipsRoot(parts: Map<string, Uint8Array>, name: string): XElement {
  const existing = parts.get(name);
  if (existing) return parseXml(strFromU8(existing));
  return relationshipRootElement();
}

function ensureMainRelationship(parts: Map<string, Uint8Array>, partName: string, noteName: XName): void {
  const relsName = relsPartName('word/document.xml');
  const root = ensureRelationshipsRoot(parts, relsName);
  const target = `/${partName}`;
  const type = `${WORD_REL_PREFIX}${nameEquals(noteName, W.footnote) ? 'footnotes' : 'endnotes'}`;
  for (const rel of childrenElements(root, REL.Relationship)) {
    if (rel.attribute(REL_TYPE) === type && rel.attribute(REL_TARGET) === target) return;
  }
  const used = new Set(relationshipsById(root).keys());
  const id = freshImportedRelationshipId(used);
  const rel = packageRelationship(type, target, id);
  root.children.push(rel);
  rel.parent = root;
  parts.set(relsName, strToU8(writeRelationshipsPart(root)));
}

function relationshipRootElement(): XElement {
  return element(REL.Relationships, [attr(xname(XMLNS_NS, 'xmlns'), REL_NS)]);
}

function contentTypesRootElement(): XElement {
  return element(CT_TYPES, [attr(xname(XMLNS_NS, 'xmlns'), CT_NS)]);
}

function packageRelationship(type: string, target: string, id: string, targetMode: string | null = null): XElement {
  return cloneElement(element(REL.Relationship, [
    attr(REL_TYPE, type),
    attr(REL_TARGET, target),
    ...(targetMode ? [attr(REL_TARGET_MODE, targetMode)] : []),
    attr(REL_ID, id),
  ]), { emptyElementSpace: true });
}

function writeMainXmlPart(root: XElement): string {
  return writeXmlPart(cloneElement(root, { documentProlog: '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' }));
}

// A note part CREATED by this render (absent in the left package) serializes with a bare
// declaration — the oracle writes fresh parts via XmlWriter defaults (no newline), while parts
// that existed in the left package round-trip through its XDocument save shape (newline).
// Freshly created parts only ever carry the bare prolog (we wrote it), so preserving a bare
// parsed prolog and normalizing everything else reproduces the distinction without threading
// freshness through every write site.
const BARE_NOTE_PROLOG = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>';

function writeNoteXmlPart(root: XElement): string {
  const fresh = root.documentProlog === null || root.documentProlog === BARE_NOTE_PROLOG;
  const prolog = fresh ? BARE_NOTE_PROLOG : `${BARE_NOTE_PROLOG}\n`;
  return writeXmlPart(cloneElement(root, { documentProlog: prolog }));
}

function writeRelationshipsPart(root: XElement): string {
  return writeXmlPart(relationshipOutputShape(root));
}

function writeContentTypesPart(root: XElement): string {
  return writeXmlPart(cloneElement(linqOutputShape(root), { documentProlog: '<?xml version="1.0" encoding="utf-8"?>' }));
}

function relationshipOutputShape(root: XElement): XElement {
  const children = root.children.map((child) => {
    if (!(child instanceof XElement) || !nameEquals(child.name, REL.Relationship)) return child;
    const type = child.attribute(REL_TYPE);
    const target = child.attribute(REL_TARGET);
    const targetMode = child.attribute(REL_TARGET_MODE);
    const id = child.attribute(REL_ID);
    if (!type || !target || !id) return cloneElement(child, { emptyElementSpace: true });
    return packageRelationship(type, target, id, targetMode);
  });
  return cloneElement(root, {
    children,
    documentProlog: '<?xml version="1.0" encoding="utf-8"?>',
  });
}

function relationshipRootForPart(parts: Record<string, Uint8Array>, partName: string): XElement | null {
  const bytes = parts[relsPartName(partName)];
  return bytes ? parseXml(strFromU8(bytes)) : null;
}

function relsPartName(partName: string): string {
  const slash = partName.lastIndexOf('/');
  const dir = slash < 0 ? '' : partName.slice(0, slash + 1);
  const file = slash < 0 ? partName : partName.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}




function isRelationshipAttribute(name: XName): boolean {
  return name.ns === R_NS;
}


function freshImportedRelationshipId(used: ReadonlySet<string>): string {
  let n = 1;
  while (used.has(`R${n.toString(16).padStart(32, '0')}`)) n++;
  return `R${n.toString(16).padStart(32, '0')}`;
}

function freshImportedPartName(parts: ReadonlyMap<string, Uint8Array>, sourcePart: string): string {
  const slash = sourcePart.lastIndexOf('/');
  const dir = slash < 0 ? '' : sourcePart.slice(0, slash + 1);
  const file = slash < 0 ? sourcePart : sourcePart.slice(slash + 1);
  const dot = file.lastIndexOf('.');
  const ext = dot >= 0 ? file.slice(dot) : '';
  let n = 1;
  let candidate = `${dir}P${n.toString(16).padStart(32, '0')}${ext}`;
  while (parts.has(candidate)) {
    n++;
    candidate = `${dir}P${n.toString(16).padStart(32, '0')}${ext}`;
  }
  return candidate;
}


function resolveTargetPart(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const slash = sourcePart.lastIndexOf('/');
  const base = slash < 0 ? '' : sourcePart.slice(0, slash + 1);
  const stack: string[] = [];
  for (const segment of `${base}${target}`.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') stack.pop();
    else stack.push(segment);
  }
  return stack.join('/');
}

function contentTypeForPart(parts: Record<string, Uint8Array>, partName: string): string | null {
  const root = parts[CONTENT_TYPES] ? parseXml(strFromU8(parts[CONTENT_TYPES]!)) : null;
  if (!root) return null;
  const partPath = `/${partName}`;
  for (const over of root.elements(CT_OVERRIDE)) {
    if (over.attribute(CT_PART_NAME) === partPath) return over.attribute(CT_CONTENT_TYPE);
  }
  const ext = extensionOf(partName);
  if (ext) {
    for (const def of root.elements(CT_DEFAULT)) {
      if ((def.attribute(CT_EXTENSION) ?? '').toLowerCase() === ext.toLowerCase()) return def.attribute(CT_CONTENT_TYPE);
    }
  }
  return null;
}

function addContentTypeForImportedPart(parts: Map<string, Uint8Array>, partName: string, contentType: string): void {
  if (partName.toLowerCase().endsWith('.xml')) addContentTypeOverride(parts, partName, contentType);
  else addContentTypeDefault(parts, extensionOf(partName), contentType);
}

function addContentTypeOverride(parts: Map<string, Uint8Array>, partName: string, contentType: string): void {
  const root = parts.get(CONTENT_TYPES) ? parseXml(strFromU8(parts.get(CONTENT_TYPES)!)) : contentTypesRootElement();
  const partPath = `/${partName}`;
  for (const over of root.elements(CT_OVERRIDE)) {
    if (over.attribute(CT_PART_NAME) === partPath) return;
  }
  const over = element(CT_OVERRIDE, [attr(CT_PART_NAME, partPath), attr(CT_CONTENT_TYPE, contentType)]);
  root.children.push(over);
  over.parent = root;
  parts.set(CONTENT_TYPES, strToU8(writeContentTypesPart(root)));
}

function addContentTypeDefault(parts: Map<string, Uint8Array>, extension: string | null, contentType: string): void {
  if (!extension) return;
  const root = parts.get(CONTENT_TYPES) ? parseXml(strFromU8(parts.get(CONTENT_TYPES)!)) : contentTypesRootElement();
  for (const def of root.elements(CT_DEFAULT)) {
    if ((def.attribute(CT_EXTENSION) ?? '').toLowerCase() === extension.toLowerCase()) return;
  }
  const def = element(CT_DEFAULT, [attr(CT_EXTENSION, extension), attr(CT_CONTENT_TYPE, contentType)]);
  // System.IO.Packaging keeps the part grouped: Defaults lead, Overrides follow — a new Default
  // lands after the last existing Default, not at the end of the part.
  const lastDefault = root.children.reduce((acc, c, i) => c instanceof XElement && nameEquals(c.name, CT_DEFAULT) ? i : acc, -1);
  root.children.splice(lastDefault >= 0 ? lastDefault + 1 : 0, 0, def);
  def.parent = root;
  parts.set(CONTENT_TYPES, strToU8(writeContentTypesPart(root)));
}

function extensionOf(partName: string): string | null {
  const file = partName.slice(partName.lastIndexOf('/') + 1);
  const dot = file.lastIndexOf('.');
  return dot >= 0 && dot + 1 < file.length ? file.slice(dot + 1) : null;
}

function relationshipRoot(parts: Record<string, Uint8Array>): XElement | null {
  const bytes = parts['word/_rels/document.xml.rels'];
  return bytes ? parseXml(strFromU8(bytes)) : null;
}

function relationshipsById(root: XElement): Map<string, XElement> {
  const map = new Map<string, XElement>();
  for (const rel of childrenElements(root, xname(REL_NS, 'Relationship'))) {
    const id = rel.attribute(REL_ID);
    if (id) map.set(id, rel);
  }
  return map;
}

function hyperlinkRelationships(root: XElement): Map<string, XElement> {
  const map = new Map<string, XElement>();
  for (const [id, rel] of relationshipsById(root)) {
    if (rel.attribute(REL_TYPE) === REL_HYPERLINK) map.set(id, rel);
  }
  return map;
}

function freshRelationshipId(used: ReadonlySet<string>): string {
  let n = 1;
  while (used.has(`rIdRemap${n}`)) n++;
  return `rIdRemap${n}`;
}

function rewriteInsertedHyperlinkIds(root: XElement, remap: ReadonlyMap<string, string>): XElement {
  return transformDescendants(root, (node) => {
    if (!nameEquals(node.name, W.hyperlink) || !isInsertedOnlyHyperlink(node)) return null;
    const rewritten = [...node.attributeInfos()].map((a) => {
      const replacement = nameEquals(a.name, R.id) ? remap.get(a.value) : undefined;
      return replacement ? { ...a, value: replacement } : a;
    });
    return cloneElement(node, { attributes: rewritten, children: node.children });
  });
}

function isInsertedOnlyHyperlink(link: XElement): boolean {
  return [...link.descendants(W.ins)].length > 0 && [...link.descendants(W.del)].length === 0;
}

function normalizeBookmarks(root: XElement, leftNames: ReadonlySet<string>, rightNames: ReadonlySet<string>): XElement {
  const body = root.element(W.body);
  if (!body) return root;
  let starts = runLevelBookmarks(body, W.bookmarkStart);
  let ends = runLevelBookmarks(body, W.bookmarkEnd);
  if (starts.length === 0 && ends.length === 0) return root;

  for (const [name, nameStarts] of groupBy(starts.filter((s) => (s.attribute(W.name) ?? '') !== ''), (s) => s.attribute(W.name)!)) {
    const ids = new Set(nameStarts.map(idOf).filter((id): id is string => id !== null));
    const nameEnds = ends.filter((e) => {
      const id = idOf(e);
      return id !== null && ids.has(id);
    });
    if (!leftNames.has(name) || !rightNames.has(name)) continue;
    if ([...nameStarts, ...nameEnds].some(isInWholeBlockRevisedParagraph)) continue;

    const keepStart = nameStarts[0]!;
    const keepEnd = nameEnds[0] ?? null;
    for (const s of nameStarts) if (s !== keepStart) removeBookmarkMarker(s);
    for (const e of nameEnds) if (e !== keepEnd) removeBookmarkMarker(e);
    liftBookmarkBare(keepStart);
    if (keepEnd) liftBookmarkBare(keepEnd);
  }

  starts = runLevelBookmarks(body, W.bookmarkStart);
  ends = runLevelBookmarks(body, W.bookmarkEnd);
  const liveStarts = groupBy(starts, (s) => idOf(s) ?? '');
  const liveEnds = groupBy(ends, (e) => idOf(e) ?? '');
  const dupIds = new Set<string>();
  for (const [id, list] of liveStarts) if (list.length > 1) dupIds.add(id);
  for (const [id, list] of liveEnds) if (list.length > 1) dupIds.add(id);
  if (dupIds.size > 0) {
    let next = globalMaxBookmarkId(root) + 1;
    for (const id of dupIds) {
      const ss = liveStarts.get(id) ?? [];
      const es = liveEnds.get(id) ?? [];
      const copies = Math.max(ss.length, es.length);
      for (let k = 1; k < copies; k++) {
        const fresh = String(next++);
        if (k < ss.length) replaceElementInParent(ss[k]!, withAttribute(ss[k]!, W.id, fresh));
        if (k < es.length) replaceElementInParent(es[k]!, withAttribute(es[k]!, W.id, fresh));
      }
    }
  }

  starts = runLevelBookmarks(body, W.bookmarkStart);
  ends = runLevelBookmarks(body, W.bookmarkEnd);
  const startById = groupBy(starts, (s) => idOf(s) ?? '');
  const endById = groupBy(ends, (e) => idOf(e) ?? '');
  for (const [id, sl] of startById) {
    const have = endById.get(id)?.length ?? 0;
    for (let k = have; k < sl.length; k++) insertAfter(sl[k]!, element(W.bookmarkEnd, [attr(W.id, id)]));
  }
  for (const [id, el] of endById) {
    const have = startById.get(id)?.length ?? 0;
    for (let k = have; k < el.length; k++) removeBookmarkMarker(el[k]!);
  }
  return root;
}

function bodyBookmarkNames(doc: IrDocument): Set<string> {
  const names = new Set<string>();
  for (const block of doc.body.blocks) {
    const el = block.source.element;
    if (!el) continue;
    for (const bk of descendantsAndSelf(el).filter((e) => nameEquals(e.name, W.bookmarkStart))) {
      const name = bk.attribute(W.name);
      if (name) names.add(name);
    }
  }
  return names;
}

function runLevelBookmarks(root: XElement, name: XName): XElement[] {
  return [...root.descendants(name)].filter(isRunLevelBookmark);
}

function isRunLevelBookmark(marker: XElement): boolean {
  for (let a = marker.parent; a; a = a.parent) {
    if (nameEquals(a.name, W.p) || nameEquals(a.name, W.body) || nameEquals(a.name, W.tc)) return true;
    if (
      !nameEquals(a.name, W.ins) &&
      !nameEquals(a.name, W.del) &&
      !nameEquals(a.name, W.hyperlink) &&
      !nameEquals(a.name, W.sdt) &&
      !nameEquals(a.name, W.sdtContent) &&
      !nameEquals(a.name, W.smartTag) &&
      !nameEquals(a.name, W.fldSimple)
    ) return false;
  }
  return false;
}

function isInWholeBlockRevisedParagraph(marker: XElement): boolean {
  for (let a = marker.parent; a; a = a.parent) {
    if (!nameEquals(a.name, W.p)) continue;
    const mark = a.element(W.pPr)?.element(W.rPr);
    return !!mark?.element(W.del) || !!mark?.element(W.ins);
  }
  return false;
}

function removeBookmarkMarker(marker: XElement): void {
  const parent = marker.parent;
  if (!parent) return;
  removeChild(parent, marker);
  if ((nameEquals(parent.name, W.ins) || nameEquals(parent.name, W.del)) && parent.parent && childrenElements(parent).length === 0) {
    removeChild(parent.parent, parent);
  }
}

function liftBookmarkBare(marker: XElement): boolean {
  const parent = marker.parent;
  const grand = parent?.parent ?? null;
  if (!parent || !grand || (!nameEquals(parent.name, W.ins) && !nameEquals(parent.name, W.del))) return false;
  removeChild(parent, marker);
  if (nameEquals(marker.name, W.bookmarkEnd)) insertAfter(parent, marker);
  else insertBefore(parent, marker);
  if (childrenElements(parent).length === 0 && parent.parent) removeChild(parent.parent, parent);
  return true;
}

function idOf(el: XElement): string | null {
  return el.attribute(W.id);
}

function globalMaxBookmarkId(root: XElement): number {
  let max = 0;
  for (const el of root.descendants()) {
    if (!nameEquals(el.name, W.bookmarkStart) && !nameEquals(el.name, W.bookmarkEnd)) continue;
    const id = el.attribute(W.id);
    if (id && /^\d+$/.test(id)) max = Math.max(max, Number(id));
  }
  return max;
}

function withAttribute(el: XElement, name: XName, value: string | number): XElement {
  const attrs = [...el.attributeInfos()];
  const replacement = attr(name, value);
  const idx = attrs.findIndex((a) => nameEquals(a.name, name));
  if (idx >= 0) attrs.splice(idx, 1, replacement);
  else attrs.push(replacement);
  return cloneElement(el, { attributes: attrs });
}

function replaceElementInParent(oldElement: XElement, replacement: XElement): void {
  const parent = oldElement.parent;
  if (!parent) return;
  const idx = parent.children.indexOf(oldElement);
  if (idx < 0) return;
  parent.children[idx] = replacement;
  replacement.parent = parent;
  oldElement.parent = null;
}

function removeChild(parent: XElement, child: XElement): void {
  const idx = parent.children.indexOf(child);
  if (idx < 0) return;
  parent.children.splice(idx, 1);
  child.parent = null;
}

function insertBefore(ref: XElement, child: XElement): void {
  const parent = ref.parent;
  if (!parent) return;
  const idx = parent.children.indexOf(ref);
  if (idx < 0) return;
  parent.children.splice(idx, 0, child);
  child.parent = parent;
}

function insertAfter(ref: XElement, child: XElement): void {
  const parent = ref.parent;
  if (!parent) return;
  const idx = parent.children.indexOf(ref);
  if (idx < 0) return;
  parent.children.splice(idx + 1, 0, child);
  child.parent = parent;
}

function groupBy<T, K>(items: Iterable<T>, key: (item: T) => K): Map<K, T[]> {
  const map = new Map<K, T[]>();
  for (const item of items) {
    const k = key(item);
    const list = map.get(k);
    if (list) list.push(item);
    else map.set(k, [item]);
  }
  return map;
}

function replaceChildrenInPlace(el: XElement, replace: (child: XChild) => XChild): void {
  el.children.splice(0, el.children.length, ...el.children.map(replace));
  for (const child of el.children) if (child instanceof XElement) child.parent = el;
}

function revisionAttrs(state: RenderState): XAttributeInfo[] {
  return [attr(W.author, state.settings.authorForRevisions), attr(W.id, nextId(state)), attr(W.date, state.settings.dateTimeForRevisions)];
}

function nextId(state: RenderState): number { return state.nextId++; }
function moveName(id: number, state: RenderState): string {
  let name = state.moveNames.get(id);
  if (!name) { name = `move${state.nextMoveName++}`; state.moveNames.set(id, name); }
  return name;
}
function registerRightSourcedClone(clone: XElement, state: RenderState): void {
  if (descendantsAndSelf(clone).some((el) => [...el.attributeInfos()].some((a) => isRelationshipAttribute(a.name)))) {
    state.rightSourcedClones.push(clone);
  }
}

function resolveBlock(anchor: string | null, doc: IrDocument): IrBlock | undefined { return anchor ? doc.anchorIndex.get(anchor) : undefined; }
function sourceElement(anchor: string | null, doc: IrDocument): XElement | null { return resolveBlock(anchor, doc)?.source.element ?? null; }
function paragraphTokens(anchor: string | null, doc: IrDocument, settings: IrDiffSettings): ReadonlyArray<IrDiffToken> {
  const b = resolveBlock(anchor, doc);
  return b?.kind === 'paragraph' ? tokenizeIrParagraph(b as IrParagraph, settings) : [];
}
function isSectionBreakOp(op: IrEditOp, state: RenderState): boolean {
  return op.rightAnchor?.startsWith('sec:') === true || op.leftAnchor?.startsWith('sec:') === true || resolveBlock(op.rightAnchor, state.right)?.kind === 'sectionBreak' || resolveBlock(op.leftAnchor, state.left)?.kind === 'sectionBreak';
}
function revElementName(kind: RevKind): XName { return kind === 'Ins' ? W.ins : kind === 'Del' ? W.del : kind === 'MoveFrom' ? W.moveFrom : W.moveTo; }
function isDeleteGrade(kind: RevKind): boolean { return kind === 'Del' || kind === 'MoveFrom'; }
function isAlwaysKeepMarker(name: XName): boolean { return nameEquals(name, W.bookmarkStart) || nameEquals(name, W.bookmarkEnd) || nameEquals(name, W.commentRangeStart) || nameEquals(name, W.commentRangeEnd); }
function fieldPlumbingKeep(name: XName): boolean { return nameEquals(name, W.fldChar) || nameEquals(name, W.instrText) || nameEquals(name, W.delInstrText); }
function isContainer(name: XName): boolean { return nameEquals(name, W.hyperlink) || nameEquals(name, W.ins) || nameEquals(name, W.del) || nameEquals(name, W.sdt) || nameEquals(name, W.smartTag); }
function preserveSpace(s: string): boolean { return s.length > 0 && (/\s/.test(s[0]!) || /\s/.test(s[s.length - 1]!)); }
function zeroWidthBoundaries(tokens: ReadonlyArray<IrDiffToken>, start: number, end: number): [boolean, boolean] { return end <= start ? [false, false] : [tokens[start]?.startChar === tokens[start]?.endChar, tokens[end - 1]?.startChar === tokens[end - 1]?.endChar]; }
function rightSpanChars(tokens: ReadonlyArray<IrDiffToken>, op: IrTokenOp): [number, number] {
  if (op.rightStart >= op.rightEnd) { const at = op.rightStart < tokens.length ? tokens[op.rightStart]!.startChar : tokens.at(-1)?.endChar ?? 0; return [at, at]; }
  return [tokens[op.rightStart]!.startChar, tokens[op.rightEnd - 1]!.endChar];
}
function leftSpanChars(tokens: ReadonlyArray<IrDiffToken>, op: IrTokenOp): [number, number] {
  if (op.leftStart >= op.leftEnd) { const at = op.leftStart < tokens.length ? tokens[op.leftStart]!.startChar : tokens.at(-1)?.endChar ?? 0; return [at, at]; }
  return [tokens[op.leftStart]!.startChar, tokens[op.leftEnd - 1]!.endChar];
}
function segmentSliceLength(diff: IrTokenDiff, leftSide: boolean): number { return diff.ops.reduce((n, o) => n + (leftSide ? (o.kind === 'Insert' ? 0 : o.leftEnd - o.leftStart) : (o.kind === 'Delete' ? 0 : o.rightEnd - o.rightStart)), 0); }
function rawSpanText(tokens: ReadonlyArray<IrDiffToken>, start: number, end: number): string { return tokens.slice(start, end).map((t) => t.text).join(''); }
function runsForFormatStamp(el: XElement): XElement[] { return nameEquals(el.name, W.r) ? [el] : [...el.descendants(W.r)]; }
function runTextLength(run: XElement): number { return childrenElements(run, W.t).reduce((n, t) => n + t.value.length, 0); }
function compareXml(el: XElement): string { return writeXmlPart(stripUnids(cloneElement(el)), { includeDeclaration: false }); }
function pPrForCompare(pPr: XElement | null): XElement { return pPr ? element(W.pPr, [...pPr.attributeInfos()], childrenElements(pPr).filter((e) => !nameEquals(e.name, W.rPr) && !nameEquals(e.name, W.sectPr) && !nameEquals(e.name, W.pPrChange)).map((e) => cloneElement(e))) : element(W.pPr); }
function markRPrForCompare(pPr: XElement | null): XElement { const rPr = pPr?.element(W.rPr); return rPr ? element(W.rPr, [...rPr.attributeInfos()], childrenElements(rPr).filter((e) => !nameEquals(e.name, W.rPrChange)).map((e) => cloneElement(e))) : element(W.rPr); }
function cleanShell(shell: XElement | null, changeName: XName, exclude: readonly XName[]): XElement { return element(xname('', 'shell'), shell ? [...shell.attributeInfos()] : [], shell ? childrenElements(shell).filter((e) => !nameEquals(e.name, changeName) && !exclude.some((x) => nameEquals(x, e.name))).map((e) => cloneElement(e)) : []); }
function isSectPrProp(e: XElement): boolean { return !nameEquals(e.name, W.headerReference) && !nameEquals(e.name, W.footerReference) && !nameEquals(e.name, W.sectPrChange); }
function sectPrPropsString(e: XElement): string { return writeXmlPart(element(xname('', 'sect'), [], childrenElements(e).filter(isSectPrProp).map((c) => stripUnids(cloneElement(c)))), { includeDeclaration: false }); }
function indexRows(table: IrTable | undefined): Map<string, XElement> {
  const map = new Map<string, XElement>();
  for (const row of table?.rows ?? []) if (row.source.element) map.set(anchorToString(row.anchor), row.source.element);
  return map;
}
function coalesceAdjacentHyperlinks(content: XElement[]): XElement[] {
  const merged: XElement[] = [];
  for (let i = 0; i < content.length;) {
    const el = content[i]!;
    if (!nameEquals(el.name, W.hyperlink)) { merged.push(el); i++; continue; }
    let j = i + 1;
    while (j < content.length && nameEquals(content[j]!.name, W.hyperlink) && sameSourceLink(el, content[j]!)) j++;
    const group = content.slice(i, j);
    if (group.length > 1 && group.some((g) => !childrenElements(g).every((c) => nameEquals(c.name, W.del) || nameEquals(c.name, W.ins)))) {
      merged.push(cloneElement(group[0]!, { attributes: [...group[0]!.attributeInfos()].filter((a) => !nameEquals(a.name, SOURCE_LINK_ID)), children: group.flatMap((g) => childrenElements(g)) }));
    } else {
      merged.push(...group.map((g) => cloneElement(g, { attributes: [...g.attributeInfos()].filter((a) => !nameEquals(a.name, SOURCE_LINK_ID)) })));
    }
    i = j;
  }
  return merged;
}
function sameSourceLink(a: XElement, b: XElement): boolean {
  const ao = a.attribute(SOURCE_LINK_ID), bo = b.attribute(SOURCE_LINK_ID);
  if (ao !== null || bo !== null) return ao === bo;
  const aa = [...a.attributeInfos()].filter((x) => x.name.ns !== PT_NS);
  const bb = [...b.attributeInfos()].filter((x) => x.name.ns !== PT_NS);
  return aa.length === bb.length && aa.every((x, i) => nameEquals(x.name, bb[i]!.name) && x.value === bb[i]!.value);
}
