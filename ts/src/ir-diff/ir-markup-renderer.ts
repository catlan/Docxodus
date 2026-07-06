import { unzipSync, strToU8 } from 'fflate';
import { W, PT_NS, R } from '../ir/names.js';
import type { IrBlock, IrParagraph, IrTable } from '../ir/ir-blocks.js';
import { readIrDocument, type IrDocument } from '../ir/ir-reader.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import { normalizeIrDiffSettings, type IrDiffSettings, type IrDiffSettingsOptions } from './ir-diff-settings.js';
import type { IrCellOp, IrEditOp, IrEditScript, IrRowOp, IrTableDiff } from './ir-edit-script.js';
import type { IrDiffToken } from './ir-diff-token.js';
import type { IrTokenDiff, IrTokenOp } from './ir-token-diff.js';
import { writeXmlPart } from '../xml/xwriter.js';
import { XElement, nameEquals, xname, type XAttributeInfo, type XName } from '../xml/xelement.js';
import { attr, childrenElements, cloneElement, descendantsAndSelf, element, replaceName, type XChild } from '../xml/xclone.js';

type RevKind = 'Ins' | 'Del' | 'MoveFrom' | 'MoveTo';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XML_SPACE = xname(XML_NS, 'space');
const SOURCE_LINK_ID = xname(PT_NS, 'SourceLinkId');

interface RenderState {
  readonly left: IrDocument;
  readonly right: IrDocument;
  readonly settings: IrDiffSettings;
  nextId: number;
  nextMoveName: number;
  readonly moveNames: Map<number, string>;
}

export function renderIrMarkup(
  leftDocx: Uint8Array,
  rightDocx: Uint8Array,
  script: IrEditScript,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): Map<string, Uint8Array> {
  const leftParts = unzipSync(leftDocx);
  const settings = normalizeIrDiffSettings(options);
  const left = readIrDocument(leftDocx, { retainSources: true });
  const right = readIrDocument(rightDocx, { retainSources: true });
  const state: RenderState = { left, right, settings, nextId: 1, nextMoveName: 1, moveNames: new Map() };
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
  const newRoot = cloneElement(linqOutputShape(normalizeBookmarkMarkers(stripUnids(replaceChild(root, body, newBody)))), {
    documentProlog: '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n',
  });

  const result = new Map<string, Uint8Array>();
  for (const [name, bytes] of Object.entries(leftParts)) result.set(name, bytes);
  result.set('word/document.xml', strToU8(writeXmlPart(newRoot)));
  return result;
}

function renderBlockOp(op: IrEditOp, state: RenderState, sink: XElement[]): void {
  if (isSectionBreakOp(op, state)) return;
  switch (op.kind) {
    case 'EqualBlock':
      emitVerbatim(op.rightAnchor, state.right, sink);
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

function emitWholeBlock(anchor: string | null, doc: IrDocument, state: RenderState, sink: XElement[], kind: RevKind): void {
  const src = sourceElement(anchor, doc);
  if (!src) return;
  const clone = stripUnids(cloneElement(src));
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
        const [s, e] = rightSpanChars(rightTokens, op);
        const [zs, ze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        content.push(...rightRuns.slice(s, e, zs, ze));
        break;
      }
      case 'FormatChanged': {
        const [rs, re] = rightSpanChars(rightTokens, op);
        const [zs, ze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        for (const r of rightRuns.slice(rs, re, zs, ze)) {
          const leftRPr = leftRuns.rPrAtChar(leftSpanChars(leftTokens, op)[0]);
          for (const run of runsForFormatStamp(r)) applyRPrChange(run, leftRPr, state);
          content.push(r);
        }
        break;
      }
      case 'Insert': {
        const [s, e] = rightSpanChars(rightTokens, op);
        const [zs, ze] = zeroWidthBoundaries(rightTokens, op.rightStart, op.rightEnd);
        for (const r of rightRuns.slice(s, e, zs, ze)) {
          if (isBookmarkMarker(r.name)) revisionAttrs(state);
          else content.push(...wrapFieldAware(r, 'Ins', state));
        }
        break;
      }
      case 'Delete': {
        const [s, e] = leftSpanChars(leftTokens, op);
        const [zs, ze] = zeroWidthBoundaries(leftTokens, op.leftStart, op.leftEnd);
        for (const r of leftRuns.slice(s, e, zs, ze)) {
          if (isBookmarkMarker(r.name)) {
            revisionAttrs(state); // C# emits then normalizes these wrappers away; ids are still consumed.
            content.push(r);
          }
          else content.push(...wrapFieldAware(r, 'Del', state));
        }
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
    const left = rowOp.leftRowAnchor ? leftRows.get(rowOp.leftRowAnchor) : undefined;
    if (left) row = applyRowAndCellShellChanges(row, left, state);
    newTbl.children.push(row);
    return true;
  }
  if (rowOp.kind === 'InsertRow' || (rowOp.kind === 'MovedRow' && !rowOp.isMoveSource)) {
    const src = rightRows.get(rowOp.rightRowAnchor ?? '');
    if (!src) return false;
    newTbl.children.push(markWholeRow(stripUnids(cloneElement(src)), 'Ins', state));
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
    let offset = 0;
    for (const child of childrenElements(para).filter((e) => !nameEquals(e.name, W.pPr))) this.walkRunLevel(child, { value: offset }, []);
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
    host.children.unshift(rightShell);
  }
  rightShell.children.splice(0, rightShell.children.length, ...rightShell.children.filter((c) => !(c instanceof XElement && nameEquals(c.name, changeName))));
  const inner = leftShell ? cloneElement(leftShell, { children: childrenElements(leftShell).filter((e) => !nameEquals(e.name, changeName) && !exclude.some((x) => nameEquals(x, e.name))).map((e) => stripUnids(cloneElement(e))) }) : element(shellName);
  rightShell.children.push(element(changeName, idOnly ? [attr(W.id, nextId(state))] : revisionAttrs(state), [inner]));
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

function normalizeBookmarkMarkers(root: XElement): XElement {
  return transformDescendants(root, (node) => {
    if (!nameEquals(node.name, W.p)) return null;
    const seenStarts = new Set<string>();
    const seenEnds = new Set<string>();
    const children = node.children.filter((child) => {
      if (!(child instanceof XElement)) return true;
      if (nameEquals(child.name, W.bookmarkStart)) {
        const key = `${child.attribute(W.id) ?? ''}\0${child.attribute(W.name) ?? ''}`;
        if (seenStarts.has(key)) return false;
        seenStarts.add(key);
      } else if (nameEquals(child.name, W.bookmarkEnd)) {
        const key = child.attribute(W.id) ?? '';
        if (seenEnds.has(key)) return false;
        seenEnds.add(key);
      }
      return true;
    });
    return cloneElement(node, { children });
  });
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
function isBookmarkMarker(name: XName): boolean { return nameEquals(name, W.bookmarkStart) || nameEquals(name, W.bookmarkEnd); }
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
  for (const row of table?.rows ?? []) if (row.source.element) map.set(String(row.anchor), row.source.element);
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
