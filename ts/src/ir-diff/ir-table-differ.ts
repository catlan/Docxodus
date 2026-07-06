// Port of Docxodus/Ir/Diff/IrTableDiffer.cs

import { anchorToString } from '../ir/ir-anchor.js';
import type { IrBlock, IrCell, IrRow, IrTable } from '../ir/ir-blocks.js';
import { alignIrBlocks } from './ir-block-aligner.js';
import type { IrAlignedBlock } from './ir-block-alignment.js';
import {
  normalizeIrDiffSettings,
  type IrDiffSettings,
  type IrDiffSettingsOptions,
} from './ir-diff-settings.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import type { IrCellOp, IrEditOp, IrRowOp, IrRowOpKind, IrTableDiff } from './ir-edit-script.js';
import { projectIrAlignment } from './ir-edit-script-builder.js';
import { diffIrTokens } from './ir-token-differ.js';

interface RowCand {
  readonly left: number;
  readonly right: number;
}

export function diffIrTables(
  left: IrTable,
  right: IrTable,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrTableDiff {
  const settings = normalizeIrDiffSettings(options);
  const leftRows = left.rows;
  const rightRows = right.rows;
  const leftKind: Array<IrRowOpKind | null> = Array(leftRows.length).fill(null);
  const rightKind: Array<IrRowOpKind | null> = Array(rightRows.length).fill(null);
  const leftMatch = Array<number>(leftRows.length).fill(-1);
  const rightMatch = Array<number>(rightRows.length).fill(-1);

  const candidates = collectRowAnchors(leftRows, rightRows, leftMatch, rightMatch);
  candidates.sort((a, b) => a.left - b.left);
  const onSpine = lis(candidates);
  for (let c = 0; c < candidates.length; c++) {
    const { left: li, right: rj } = candidates[c]!;
    const kind: IrRowOpKind = onSpine.has(c) ? 'EqualRow' : 'MovedRow';
    leftKind[li] = kind;
    rightKind[rj] = kind;
  }

  const spinePairs = [...onSpine]
    .map((c) => ({ left: candidates[c]!.left, right: candidates[c]!.right }))
    .sort((a, b) => a.left - b.left);
  fillRowGaps(leftRows, rightRows, spinePairs, leftKind, rightKind, leftMatch, rightMatch);
  return { rowOps: emitRowOps(leftRows, rightRows, leftKind, rightKind, leftMatch, rightMatch, settings) };
}

export const IrTableDiffer = {
  diff: diffIrTables,
};

function collectRowAnchors(
  leftRows: ReadonlyArray<IrRow>,
  rightRows: ReadonlyArray<IrRow>,
  leftMatch: number[],
  rightMatch: number[],
): RowCand[] {
  const leftByHash = uniqueRowsByHash(leftRows);
  const rightByHash = uniqueRowsByHash(rightRows);
  const candidates: RowCand[] = [];
  for (let i = 0; i < leftRows.length; i++) {
    const h = leftRows[i]!.contentHash;
    if (leftByHash.get(h) !== i) continue;
    const rj = rightByHash.get(h);
    if (rj === undefined) continue;
    leftMatch[i] = rj;
    rightMatch[rj] = i;
    candidates.push({ left: i, right: rj });
  }
  return candidates;
}

function uniqueRowsByHash(rows: ReadonlyArray<IrRow>): Map<string, number> {
  const counts = new Map<string, number>();
  const first = new Map<string, number>();
  for (let i = 0; i < rows.length; i++) {
    const h = rows[i]!.contentHash;
    counts.set(h, (counts.get(h) ?? 0) + 1);
    if (!first.has(h)) first.set(h, i);
  }
  const unique = new Map<string, number>();
  for (const [h, i] of first) if (counts.get(h) === 1) unique.set(h, i);
  return unique;
}

function lis(candidates: ReadonlyArray<RowCand>): Set<number> {
  const result = new Set<number>();
  if (candidates.length === 0) return result;
  const tails: number[] = [];
  const prev = Array<number>(candidates.length).fill(-1);
  for (let i = 0; i < candidates.length; i++) {
    const right = candidates[i]!.right;
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (candidates[tails[mid]!]!.right < right) lo = mid + 1;
      else hi = mid;
    }
    if (lo > 0) prev[i] = tails[lo - 1]!;
    if (lo === tails.length) tails.push(i);
    else tails[lo] = i;
  }
  for (let i = tails[tails.length - 1]!; i !== -1; i = prev[i]!) result.add(i);
  return result;
}

function fillRowGaps(
  leftRows: ReadonlyArray<IrRow>,
  rightRows: ReadonlyArray<IrRow>,
  spinePairs: ReadonlyArray<{ readonly left: number; readonly right: number }>,
  leftKind: Array<IrRowOpKind | null>,
  rightKind: Array<IrRowOpKind | null>,
  leftMatch: number[],
  rightMatch: number[],
): void {
  let prevLeft = -1;
  let prevRight = -1;
  for (const { left, right } of spinePairs) {
    fillOneRowGap(prevLeft + 1, left, prevRight + 1, right, leftKind, rightKind, leftMatch, rightMatch);
    prevLeft = left;
    prevRight = right;
  }
  fillOneRowGap(prevLeft + 1, leftRows.length, prevRight + 1, rightRows.length, leftKind, rightKind, leftMatch, rightMatch);
}

function fillOneRowGap(
  leftFrom: number,
  leftTo: number,
  rightFrom: number,
  rightTo: number,
  leftKind: Array<IrRowOpKind | null>,
  rightKind: Array<IrRowOpKind | null>,
  leftMatch: number[],
  rightMatch: number[],
): void {
  const freeLeft: number[] = [];
  for (let i = leftFrom; i < leftTo; i++) if (leftMatch[i] === -1) freeLeft.push(i);
  const freeRight: number[] = [];
  for (let j = rightFrom; j < rightTo; j++) if (rightMatch[j] === -1) freeRight.push(j);
  const paired = Math.min(freeLeft.length, freeRight.length);
  for (let k = 0; k < paired; k++) {
    const li = freeLeft[k]!;
    const rj = freeRight[k]!;
    leftKind[li] = 'ModifyRow';
    rightKind[rj] = 'ModifyRow';
    leftMatch[li] = rj;
    rightMatch[rj] = li;
  }
  for (let k = paired; k < freeLeft.length; k++) leftKind[freeLeft[k]!] = 'DeleteRow';
  for (let k = paired; k < freeRight.length; k++) rightKind[freeRight[k]!] = 'InsertRow';
}

function emitRowOps(
  leftRows: ReadonlyArray<IrRow>,
  rightRows: ReadonlyArray<IrRow>,
  leftKind: ReadonlyArray<IrRowOpKind | null>,
  rightKind: ReadonlyArray<IrRowOpKind | null>,
  leftMatch: ReadonlyArray<number>,
  rightMatch: ReadonlyArray<number>,
  settings: IrDiffSettings,
): IrRowOp[] {
  const moveGroup = new Map<number, number>();
  let nextGroup = 1;
  for (let j = 0; j < rightRows.length; j++) if (rightKind[j] === 'MovedRow') moveGroup.set(rightMatch[j]!, nextGroup++);

  const sourcesAfterLeft = new Map<number, number[]>();
  let lastPaired = -1;
  for (let i = 0; i < leftRows.length; i++) {
    if (leftKind[i] === 'DeleteRow' || leftKind[i] === 'MovedRow') {
      const list = sourcesAfterLeft.get(lastPaired) ?? [];
      list.push(i);
      sourcesAfterLeft.set(lastPaired, list);
    } else if (leftMatch[i] !== -1) {
      lastPaired = i;
    }
  }

  const ops: IrRowOp[] = [];
  emitRowSources(sourcesAfterLeft, -1, leftRows, leftKind, moveGroup, ops);
  for (let j = 0; j < rightRows.length; j++) {
    const kind = rightKind[j] ?? 'InsertRow';
    const li = rightMatch[j]!;
    if (kind === 'EqualRow') {
      ops.push(rowOp('EqualRow', anchorToString(leftRows[li]!.anchor), anchorToString(rightRows[j]!.anchor), null));
    } else if (kind === 'ModifyRow') {
      ops.push(rowOp('ModifyRow', anchorToString(leftRows[li]!.anchor), anchorToString(rightRows[j]!.anchor), diffCells(leftRows[li]!, rightRows[j]!, settings)));
    } else if (kind === 'MovedRow') {
      ops.push({ ...rowOp('MovedRow', null, anchorToString(rightRows[j]!.anchor), null), moveGroupId: moveGroup.get(li)!, isMoveSource: false });
    } else if (kind === 'InsertRow') {
      ops.push(rowOp('InsertRow', null, anchorToString(rightRows[j]!.anchor), null));
    }
    if (li !== -1 && (kind === 'EqualRow' || kind === 'ModifyRow')) {
      emitRowSources(sourcesAfterLeft, li, leftRows, leftKind, moveGroup, ops);
    }
  }
  return ops;
}

function rowOp(kind: IrRowOpKind, leftRowAnchor: string | null, rightRowAnchor: string | null, cellOps: ReadonlyArray<IrCellOp> | null): IrRowOp {
  return { kind, leftRowAnchor, rightRowAnchor, cellOps, moveGroupId: null, isMoveSource: null };
}

function emitRowSources(
  sourcesAfterLeft: ReadonlyMap<number, ReadonlyArray<number>>,
  anchorLeft: number,
  leftRows: ReadonlyArray<IrRow>,
  leftKind: ReadonlyArray<IrRowOpKind | null>,
  moveGroup: ReadonlyMap<number, number>,
  ops: IrRowOp[],
): void {
  for (const li of sourcesAfterLeft.get(anchorLeft) ?? []) {
    if (leftKind[li] === 'MovedRow') {
      ops.push({ ...rowOp('MovedRow', anchorToString(leftRows[li]!.anchor), null, null), moveGroupId: moveGroup.get(li)!, isMoveSource: true });
    } else {
      ops.push(rowOp('DeleteRow', anchorToString(leftRows[li]!.anchor), null, null));
    }
  }
}

function diffCells(left: IrRow, right: IrRow, settings: IrDiffSettings): IrCellOp[] {
  const paired = Math.min(left.cells.length, right.cells.length);
  const cellOps: IrCellOp[] = [];
  for (let k = 0; k < paired; k++) {
    const lc = left.cells[k]!;
    const rc = right.cells[k]!;
    const blockOps = lc.contentHash === rc.contentHash ? null : diffCellBlocks(lc, rc, settings);
    cellOps.push({ leftCellAnchor: anchorToString(lc.anchor), rightCellAnchor: anchorToString(rc.anchor), blockOps });
  }
  for (let k = paired; k < left.cells.length; k++) {
    cellOps.push({ leftCellAnchor: anchorToString(left.cells[k]!.anchor), rightCellAnchor: null, blockOps: null });
  }
  for (let k = paired; k < right.cells.length; k++) {
    cellOps.push({ leftCellAnchor: null, rightCellAnchor: anchorToString(right.cells[k]!.anchor), blockOps: null });
  }
  return cellOps;
}

function diffCellBlocks(left: IrCell, right: IrCell, settings: IrDiffSettings): IrEditOp[] {
  return projectIrAlignment(left.blocks, alignIrBlocks(left.blocks, right.blocks, settings), settings);
}

function projectAlignment(entries: ReadonlyArray<IrAlignedBlock>, settings: IrDiffSettings): IrEditOp[] {
  return entries.map((e) => projectEntry(e, settings));
}

function projectEntry(e: IrAlignedBlock, settings: IrDiffSettings): IrEditOp {
  const base = (kind: IrEditOp['kind'], leftBlock: IrBlock | null, rightBlock: IrBlock | null): IrEditOp => ({
    kind,
    leftAnchor: leftBlock ? anchorToString(leftBlock.anchor) : null,
    rightAnchor: rightBlock ? anchorToString(rightBlock.anchor) : null,
    tokenDiff: null,
    moveGroupId: null,
    isMoveSource: null,
    tableDiff: null,
    textboxDiffs: null,
    splitMergeAnchors: null,
    segmentDiffs: null,
  });
  switch (e.kind) {
    case 'Unchanged':
      return base('EqualBlock', e.left, e.right);
    case 'FormatOnly':
      return base('FormatOnlyBlock', e.left, e.right);
    case 'Inserted':
      return base('InsertBlock', null, e.right);
    case 'Deleted':
      return base('DeleteBlock', e.left, null);
    case 'Moved':
      return base('MoveBlock', e.left, e.right);
    case 'MovedModified':
      return withNestedDiff(base('MoveModifyBlock', e.left, e.right), e.left, e.right, settings);
    case 'Modified':
      return withNestedDiff(base('ModifyBlock', e.left, e.right), e.left, e.right, settings);
    case 'Split':
      return { ...base('SplitBlock', e.left, null), splitMergeAnchors: e.multiBlocks?.map((b) => anchorToString(b.anchor)) ?? null };
    case 'Merge':
      return { ...base('MergeBlock', null, e.right), splitMergeAnchors: e.multiBlocks?.map((b) => anchorToString(b.anchor)) ?? null };
  }
}

function withNestedDiff(op: IrEditOp, left: IrBlock | null, right: IrBlock | null, settings: IrDiffSettings): IrEditOp {
  if (left?.kind === 'paragraph' && right?.kind === 'paragraph') {
    return { ...op, tokenDiff: diffIrTokens(tokenizeIrParagraph(left, settings), tokenizeIrParagraph(right, settings), settings) };
  }
  if (left?.kind === 'table' && right?.kind === 'table') {
    return { ...op, tableDiff: diffIrTables(left, right, settings) };
  }
  return op;
}
