// Port of Docxodus/Ir/Diff/IrBlockAligner.cs

import type { IrBlock } from '../ir/ir-blocks.js';
import type { IrDocument } from '../ir/ir-reader.js';
import { blockSignature } from './ir-modeled-format.js';
import type { IrAlignedBlock, IrAlignmentKind, IrBlockAlignment } from './ir-block-alignment.js';
import { IrBlockSimilarity } from './ir-block-similarity.js';
import { normalizeIrDiffSettings, type IrDiffSettings, type IrDiffSettingsOptions } from './ir-diff-settings.js';

interface Candidate {
  readonly leftIndex: number;
  readonly rightIndex: number;
  readonly anchorKind: IrAlignmentKind;
}

export function alignIrBlocks(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrBlockAlignment {
  const settings = normalizeIrDiffSettings(options);
  const similarity = new IrBlockSimilarity(settings);
  const leftKind: Array<IrAlignmentKind | null> = Array(leftBlocks.length).fill(null);
  const rightKind: Array<IrAlignmentKind | null> = Array(rightBlocks.length).fill(null);
  const leftMatch = Array(leftBlocks.length).fill(-1) as number[];
  const rightMatch = Array(rightBlocks.length).fill(-1) as number[];
  const candidates: Candidate[] = [];

  collectAnchors(leftBlocks, rightBlocks, keyAB, 'Unchanged', leftMatch, rightMatch, candidates, settings);
  collectAnchors(leftBlocks, rightBlocks, keyContentOnly, 'FormatOnly', leftMatch, rightMatch, candidates, settings);

  candidates.sort((a, b) => a.leftIndex - b.leftIndex);
  let onSpine = longestIncreasingSubsequence(candidates);
  if (anyOffSpineStructuralBlock(candidates, onSpine, leftBlocks)) {
    onSpine = longestIncreasingSubsequencePreferringStructuralAnchors(candidates, leftBlocks, rightBlocks.length);
  }

  for (let c = 0; c < candidates.length; c++) {
    const cand = candidates[c]!;
    const kind = onSpine.has(c) ? cand.anchorKind : 'Moved';
    leftKind[cand.leftIndex] = kind;
    rightKind[cand.rightIndex] = kind;
  }

  const spinePairs = [...onSpine]
    .map((c) => ({ left: candidates[c]!.leftIndex, right: candidates[c]!.rightIndex }))
    .sort((a, b) => a.left - b.left);

  fillGaps(leftBlocks, rightBlocks, spinePairs, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);
  detectCrossGapMoves(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);

  return { entries: emitEntries(leftBlocks, rightBlocks, leftKind, rightKind, rightMatch, leftMatch) };
}

export function alignIrDocuments(
  left: IrDocument,
  right: IrDocument,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrBlockAlignment {
  return alignIrBlocks(left.body.blocks, right.body.blocks, options);
}

export const IrBlockAligner = {
  align: alignIrDocuments,
  alignBlocks: alignIrBlocks,
};

const keyAB = (b: IrBlock): string => `${b.contentHash}\0${b.formatFingerprint}`;
const keyContentOnly = (b: IrBlock): string => `${b.contentHash}\0`;

function collectAnchors(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  key: (block: IrBlock) => string,
  anchorKind: IrAlignmentKind,
  leftMatch: number[],
  rightMatch: number[],
  candidates: Candidate[],
  settings: IrDiffSettings,
): void {
  const leftByKey = buildUniqueIndex(leftBlocks, leftMatch, key);
  const rightByKey = buildUniqueIndex(rightBlocks, rightMatch, key);
  for (let i = 0; i < leftBlocks.length; i++) {
    if (leftMatch[i] !== -1) continue;
    const k = key(leftBlocks[i]!);
    if (leftByKey.get(k) !== i) continue;
    const rj = rightByKey.get(k);
    if (rj === undefined || rightMatch[rj] !== -1) continue;
    const resolvedKind =
      anchorKind === 'FormatOnly'
        ? formatEqual(leftBlocks[i]!, rightBlocks[rj]!, settings)
          ? 'Unchanged'
          : 'FormatOnly'
        : anchorKind;
    leftMatch[i] = rj;
    rightMatch[rj] = i;
    candidates.push({ leftIndex: i, rightIndex: rj, anchorKind: resolvedKind });
  }
}

function buildUniqueIndex(
  blocks: ReadonlyArray<IrBlock>,
  matched: number[],
  key: (block: IrBlock) => string,
): Map<string, number> {
  const counts = new Map<string, number>();
  const first = new Map<string, number>();
  for (let i = 0; i < blocks.length; i++) {
    if (matched[i] !== -1) continue;
    const k = key(blocks[i]!);
    counts.set(k, (counts.get(k) ?? 0) + 1);
    if (!first.has(k)) first.set(k, i);
  }
  const unique = new Map<string, number>();
  for (const [k, i] of first) if (counts.get(k) === 1) unique.set(k, i);
  return unique;
}

export function formatEqual(left: IrBlock, right: IrBlock, settings: IrDiffSettings): boolean {
  if (settings.formatComparison === 'ModeledOnly' && left.kind === 'paragraph' && right.kind === 'paragraph') {
    return blockSignature(left, settings) === blockSignature(right, settings);
  }
  return left.formatFingerprint === right.formatFingerprint;
}

function longestIncreasingSubsequence(candidates: ReadonlyArray<Candidate>): Set<number> {
  const result = new Set<number>();
  if (candidates.length === 0) return result;
  const tails: number[] = [];
  const prev = Array(candidates.length).fill(-1) as number[];
  for (let i = 0; i < candidates.length; i++) {
    const right = candidates[i]!.rightIndex;
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (candidates[tails[mid]!]!.rightIndex < right) lo = mid + 1;
      else hi = mid;
    }
    if (lo > 0) prev[i] = tails[lo - 1]!;
    if (lo === tails.length) tails.push(i);
    else tails[lo] = i;
  }
  for (let i = tails[tails.length - 1]!; i !== -1; i = prev[i]!) result.add(i);
  return result;
}

function anyOffSpineStructuralBlock(
  candidates: ReadonlyArray<Candidate>,
  onSpine: ReadonlySet<number>,
  leftBlocks: ReadonlyArray<IrBlock>,
): boolean {
  for (let c = 0; c < candidates.length; c++) {
    if (!onSpine.has(c) && leftBlocks[candidates[c]!.leftIndex]!.kind !== 'paragraph') return true;
  }
  return false;
}

function longestIncreasingSubsequencePreferringStructuralAnchors(
  candidates: ReadonlyArray<Candidate>,
  leftBlocks: ReadonlyArray<IrBlock>,
  nRight: number,
): Set<number> {
  const result = new Set<number>();
  if (candidates.length === 0) return result;
  const big = BigInt(candidates.length + 1);
  const treeWeight = Array<bigint>(nRight + 1).fill(0n);
  const treeIndex = Array<number>(nRight + 1).fill(-1);
  const update = (pos0: number, weight: bigint, candIndex: number) => {
    for (let pos = pos0; pos <= nRight; pos += pos & -pos) {
      if (weight > treeWeight[pos]!) {
        treeWeight[pos] = weight;
        treeIndex[pos] = candIndex;
      }
    }
  };
  const query = (pos0: number): [bigint, number] => {
    let bestWeight = 0n;
    let bestIndex = -1;
    for (let pos = pos0; pos > 0; pos -= pos & -pos) {
      if (treeWeight[pos]! > bestWeight) {
        bestWeight = treeWeight[pos]!;
        bestIndex = treeIndex[pos]!;
      }
    }
    return [bestWeight, bestIndex];
  };
  const dp = Array<bigint>(candidates.length).fill(0n);
  const parent = Array<number>(candidates.length).fill(-1);
  for (let c = 0; c < candidates.length; c++) {
    const r = candidates[c]!.rightIndex;
    const weight = big + (leftBlocks[candidates[c]!.leftIndex]!.kind === 'paragraph' ? 0n : 1n);
    const [predWeight, predIndex] = query(r);
    dp[c] = weight + predWeight;
    parent[c] = predIndex;
    update(r + 1, dp[c]!, c);
  }
  let globalBest = -1n;
  let globalEnd = -1;
  for (let c = 0; c < candidates.length; c++) {
    if (dp[c]! > globalBest) {
      globalBest = dp[c]!;
      globalEnd = c;
    }
  }
  for (let c = globalEnd; c !== -1; c = parent[c]!) result.add(c);
  return result;
}

function fillGaps(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  spinePairs: ReadonlyArray<{ readonly left: number; readonly right: number }>,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
  similarity: IrBlockSimilarity,
  settings: IrDiffSettings,
): void {
  let prevLeft = -1;
  let prevRight = -1;
  for (const { left, right } of spinePairs) {
    fillOneGap(leftBlocks, rightBlocks, prevLeft + 1, left, prevRight + 1, right, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);
    prevLeft = left;
    prevRight = right;
  }
  fillOneGap(leftBlocks, rightBlocks, prevLeft + 1, leftBlocks.length, prevRight + 1, rightBlocks.length, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);
}

function fillOneGap(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  leftFrom: number,
  leftTo: number,
  rightFrom: number,
  rightTo: number,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
  similarity: IrBlockSimilarity,
  settings: IrDiffSettings,
): void {
  const freeLeft: number[] = [];
  for (let i = leftFrom; i < leftTo; i++) if (leftMatch[i] === -1) freeLeft.push(i);
  const freeRight: number[] = [];
  for (let j = rightFrom; j < rightTo; j++) if (rightMatch[j] === -1) freeRight.push(j);

  inOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch, true, 'Unchanged', settings);
  inOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch, false, 'FormatOnly', settings);

  similarityPair(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch, similarity, settings.blockSimilarityThreshold);

  const leftoverLeft = freeLeft.filter((i) => leftMatch[i] === -1);
  const leftoverRight = freeRight.filter((j) => rightMatch[j] === -1);

  const tableLeft = leftoverLeft.filter((i) => leftBlocks[i]!.kind === 'table');
  const tableRight = leftoverRight.filter((j) => rightBlocks[j]!.kind === 'table');
  if (tableLeft.length === 1 && tableRight.length === 1) {
    pairAs('Modified', tableLeft[0]!, tableRight[0]!, leftKind, rightKind, leftMatch, rightMatch);
    removeValue(leftoverLeft, tableLeft[0]!);
    removeValue(leftoverRight, tableRight[0]!);
  }

  if (leftoverLeft.length === 1 && leftoverRight.length === 1) {
    pairAs('Modified', leftoverLeft[0]!, leftoverRight[0]!, leftKind, rightKind, leftMatch, rightMatch);
    return;
  }

  for (const li of leftoverLeft) if (leftMatch[li] === -1) leftKind[li] = 'Deleted';
  for (const rj of leftoverRight) if (rightMatch[rj] === -1) rightKind[rj] = 'Inserted';
}

function inOrderRefine(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  freeLeft: ReadonlyArray<number>,
  freeRight: ReadonlyArray<number>,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
  requireFormatEqual: boolean,
  kind: IrAlignmentKind,
  settings: IrDiffSettings,
): void {
  for (const rj of freeRight) {
    if (rightMatch[rj] !== -1) continue;
    for (const li of freeLeft) {
      if (leftMatch[li] !== -1) continue;
      if (leftBlocks[li]!.anchor.unid !== rightBlocks[rj]!.anchor.unid) continue;
      if (leftBlocks[li]!.contentHash !== rightBlocks[rj]!.contentHash) continue;
      if (requireFormatEqual !== formatEqual(leftBlocks[li]!, rightBlocks[rj]!, settings)) continue;
      pairAs(kind, li, rj, leftKind, rightKind, leftMatch, rightMatch);
      break;
    }
  }
  for (const rj of freeRight) {
    if (rightMatch[rj] !== -1) continue;
    for (const li of freeLeft) {
      if (leftMatch[li] !== -1) continue;
      if (leftBlocks[li]!.contentHash !== rightBlocks[rj]!.contentHash) continue;
      if (requireFormatEqual !== formatEqual(leftBlocks[li]!, rightBlocks[rj]!, settings)) continue;
      pairAs(kind, li, rj, leftKind, rightKind, leftMatch, rightMatch);
      break;
    }
  }
}

function similarityPair(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  freeLeft: ReadonlyArray<number>,
  freeRight: ReadonlyArray<number>,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
  similarity: IrBlockSimilarity,
  threshold: number,
): void {
  while (true) {
    let bestScore = threshold;
    let bestLeft = -1;
    let bestRight = -1;
    let found = false;
    for (const li of freeLeft) {
      if (leftMatch[li] !== -1) continue;
      for (const rj of freeRight) {
        if (rightMatch[rj] !== -1) continue;
        const score = similarity.score(leftBlocks[li]!, rightBlocks[rj]!);
        if (score > bestScore || (!found && score >= threshold)) {
          bestScore = score;
          bestLeft = li;
          bestRight = rj;
          found = true;
        }
      }
    }
    if (!found) return;
    pairAs('Modified', bestLeft, bestRight, leftKind, rightKind, leftMatch, rightMatch);
  }
}

function detectCrossGapMoves(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
  similarity: IrBlockSimilarity,
  settings: IrDiffSettings,
): void {
  const deleted: number[] = [];
  for (let i = 0; i < leftBlocks.length; i++) if (leftKind[i] === 'Deleted') deleted.push(i);
  const inserted: number[] = [];
  for (let j = 0; j < rightBlocks.length; j++) if (rightKind[j] === 'Inserted') inserted.push(j);
  while (true) {
    let bestScore = settings.moveSimilarityThreshold;
    let bestLeft = -1;
    let bestRight = -1;
    let found = false;
    for (const li of deleted) {
      if (leftMatch[li] !== -1 || similarity.wordCount(leftBlocks[li]!) < settings.moveMinimumTokenCount) continue;
      for (const rj of inserted) {
        if (rightMatch[rj] !== -1 || similarity.wordCount(rightBlocks[rj]!) < settings.moveMinimumTokenCount) continue;
        const score = similarity.score(leftBlocks[li]!, rightBlocks[rj]!);
        if (score > bestScore || (!found && score >= settings.moveSimilarityThreshold)) {
          bestScore = score;
          bestLeft = li;
          bestRight = rj;
          found = true;
        }
      }
    }
    if (!found) return;
    const kind = bestScore >= 1 && leftBlocks[bestLeft]!.contentHash === rightBlocks[bestRight]!.contentHash ? 'Moved' : 'MovedModified';
    pairAs(kind, bestLeft, bestRight, leftKind, rightKind, leftMatch, rightMatch);
  }
}

function pairAs(
  kind: IrAlignmentKind,
  li: number,
  rj: number,
  leftKind: Array<IrAlignmentKind | null>,
  rightKind: Array<IrAlignmentKind | null>,
  leftMatch: number[],
  rightMatch: number[],
): void {
  leftKind[li] = kind;
  rightKind[rj] = kind;
  leftMatch[li] = rj;
  rightMatch[rj] = li;
}

function emitEntries(
  leftBlocks: ReadonlyArray<IrBlock>,
  rightBlocks: ReadonlyArray<IrBlock>,
  leftKind: ReadonlyArray<IrAlignmentKind | null>,
  rightKind: ReadonlyArray<IrAlignmentKind | null>,
  rightMatch: ReadonlyArray<number>,
  leftMatch: ReadonlyArray<number>,
): IrAlignedBlock[] {
  const deletionsAfterLeft = new Map<number, number[]>();
  let lastPairedLeft = -1;
  for (let i = 0; i < leftBlocks.length; i++) {
    if (leftKind[i] === 'Deleted') {
      const bucket = deletionsAfterLeft.get(lastPairedLeft) ?? [];
      bucket.push(i);
      deletionsAfterLeft.set(lastPairedLeft, bucket);
    } else if (leftMatch[i] !== -1 && leftKind[i] !== 'Moved' && leftKind[i] !== 'MovedModified') {
      lastPairedLeft = i;
    }
  }

  const entries: IrAlignedBlock[] = [];
  emitDeletions(deletionsAfterLeft, -1, leftBlocks, entries);
  for (let j = 0; j < rightBlocks.length; j++) {
    const kind = rightKind[j] ?? 'Inserted';
    const li = rightMatch[j]!;
    entries.push({
      kind,
      left: li !== -1 ? leftBlocks[li]! : null,
      right: rightBlocks[j]!,
      multiBlocks: null,
    });
    if (li !== -1) emitDeletions(deletionsAfterLeft, li, leftBlocks, entries);
  }
  return entries;
}

function emitDeletions(
  deletionsAfterLeft: ReadonlyMap<number, ReadonlyArray<number>>,
  anchorLeftIndex: number,
  leftBlocks: ReadonlyArray<IrBlock>,
  entries: IrAlignedBlock[],
): void {
  for (const li of deletionsAfterLeft.get(anchorLeftIndex) ?? []) {
    entries.push({ kind: 'Deleted', left: leftBlocks[li]!, right: null, multiBlocks: null });
  }
}

function removeValue(values: number[], value: number): void {
  const i = values.indexOf(value);
  if (i >= 0) values.splice(i, 1);
}
