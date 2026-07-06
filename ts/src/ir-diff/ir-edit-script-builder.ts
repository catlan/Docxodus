import { anchorToString } from '../ir/ir-anchor.js';
import type { IrBlock, IrParagraph } from '../ir/ir-blocks.js';
import type { IrNoteKind } from '../ir/ir-formats.js';
import type { IrHeaderFooter, IrHeaderFooterKind, IrNoteStore } from '../ir/ir-notes.js';
import type { IrDocument } from '../ir/ir-reader.js';
import type { IrInline, IrTextbox } from '../ir/ir-inlines.js';
import { alignIrBlocks, alignIrDocuments } from './ir-block-aligner.js';
import type { IrAlignedBlock, IrBlockAlignment, IrAlignmentKind } from './ir-block-alignment.js';
import {
  normalizeIrDiffSettings,
  type IrDiffSettings,
  type IrDiffSettingsOptions,
} from './ir-diff-settings.js';
import type { IrDiffToken } from './ir-diff-token.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import type { IrEditOp, IrEditOpKind, IrEditScript, IrHeaderFooterDiff, IrNoteDiff, IrTextboxDiff } from './ir-edit-script.js';
import { IrSplitSegmenter } from './ir-split-segmenter.js';
import { diffIrTables } from './ir-table-differ.js';
import { diffIrTokens } from './ir-token-differ.js';
import type { IrTokenDiff } from './ir-token-diff.js';

interface MoveInfo {
  readonly groupId: number;
  readonly leftBlock: IrBlock;
  readonly opKind: IrEditOpKind;
}

export function buildIrEditScript(
  left: IrDocument,
  right: IrDocument,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrEditScript {
  const settings = normalizeIrDiffSettings(options);
  const alignment = alignIrDocuments(left, right, settings);
  const ops = projectIrAlignment(left.body.blocks, alignment, settings);

  const noteOps = buildNoteOps(left, right, settings);
  const headerFooterOps = buildHeaderFooterOps(left, right, settings);
  return {
    ops,
    ...(noteOps.length > 0 ? { noteOps } : {}),
    ...(headerFooterOps.length > 0 ? { headerFooterOps } : {}),
  };
}

export function projectIrAlignment(
  leftBlocks: ReadonlyArray<IrBlock>,
  alignment: IrBlockAlignment,
  settings: IrDiffSettings,
): IrEditOp[] {
  const leftIndex = buildLeftIndexMap(leftBlocks);
  const moves = new Map<number, MoveInfo>();
  let nextGroup = 1;
  for (const entry of alignment.entries) {
    if (entry.kind === 'Moved' || entry.kind === 'MovedModified') {
      const li = leftIndex.get(entry.left!);
      if (li === undefined) throw new Error('Moved entry left block is not in left block list.');
      moves.set(li, {
        groupId: nextGroup++,
        leftBlock: entry.left!,
        opKind: entry.kind === 'MovedModified' ? 'MoveModifyBlock' : 'MoveBlock',
      });
    }
  }

  const sourcesAfterLeft = buildSourceInterleave(leftBlocks, alignment, leftIndex, moves);
  const ops: IrEditOp[] = [];
  const pendingSources: number[] = [];
  stageSources(sourcesAfterLeft, -1, pendingSources);

  for (const entry of alignment.entries) {
    if (pendingSources.length > 0) {
      const limit = entry.kind === 'Deleted' ? leftIndex.get(entry.left!)! : Number.MAX_SAFE_INTEGER;
      flushPendingSources(pendingSources, limit, moves, ops);
    }

    switch (entry.kind) {
      case 'Unchanged':
        ops.push(editOp('EqualBlock', entry.left!, entry.right!));
        break;
      case 'FormatOnly':
        ops.push(editOp('FormatOnlyBlock', entry.left!, entry.right!));
        break;
      case 'Modified':
        ops.push(makeModifyOp(entry.left!, entry.right!, settings));
        break;
      case 'Inserted':
        ops.push(editOp('InsertBlock', null, entry.right!));
        break;
      case 'Deleted':
        ops.push(editOp('DeleteBlock', entry.left!, null));
        break;
      case 'Split': {
        const lp = entry.left as IrParagraph;
        const members = (entry.multiBlocks ?? []) as IrParagraph[];
        ops.push({
          ...editOp('SplitBlock', lp, null),
          splitMergeAnchors: members.map((m) => anchorToString(m.anchor)),
          segmentDiffs: IrSplitSegmenter.computeSegmentDiffs(lp, members, settings),
        });
        break;
      }
      case 'Merge': {
        const rp = entry.right as IrParagraph;
        const members = (entry.multiBlocks ?? []) as IrParagraph[];
        const sliced = IrSplitSegmenter.computeSegmentDiffs(rp, members, settings);
        ops.push({
          ...editOp('MergeBlock', null, rp),
          splitMergeAnchors: members.map((m) => anchorToString(m.anchor)),
          segmentDiffs: sliced.map((d) => IrSplitSegmenter.mirrorDiff(d)),
        });
        break;
      }
      case 'Moved':
      case 'MovedModified': {
        const li = leftIndex.get(entry.left!)!;
        const move = moves.get(li)!;
        ops.push({
          ...editOp(move.opKind, null, entry.right!),
          tokenDiff: move.opKind === 'MoveModifyBlock' ? tokenDiffFor(entry.left!, entry.right!, settings) : null,
          moveGroupId: move.groupId,
          isMoveSource: false,
        });
        break;
      }
    }

    if (entry.left !== null && isPairedInPlace(entry.kind)) {
      stageSources(sourcesAfterLeft, leftIndex.get(entry.left)!, pendingSources);
    }
    if (entry.kind === 'Merge' && entry.multiBlocks !== null) {
      for (const lb of entry.multiBlocks) stageSources(sourcesAfterLeft, leftIndex.get(lb)!, pendingSources);
    }
  }

  flushPendingSources(pendingSources, Number.MAX_SAFE_INTEGER, moves, ops);
  return ops;
}

function editOp(kind: IrEditOpKind, left: IrBlock | null, right: IrBlock | null): IrEditOp {
  return {
    kind,
    leftAnchor: left === null ? null : anchorToString(left.anchor),
    rightAnchor: right === null ? null : anchorToString(right.anchor),
    tokenDiff: null,
    moveGroupId: null,
    isMoveSource: null,
    tableDiff: null,
    textboxDiffs: null,
    splitMergeAnchors: null,
    segmentDiffs: null,
  };
}

function stageSources(sourcesAfterLeft: ReadonlyMap<number, ReadonlyArray<number>>, anchorLeftIndex: number, pendingSources: number[]): void {
  pendingSources.push(...(sourcesAfterLeft.get(anchorLeftIndex) ?? []));
}

function flushPendingSources(
  pendingSources: number[],
  limit: number,
  moves: ReadonlyMap<number, MoveInfo>,
  ops: IrEditOp[],
): void {
  let n = 0;
  while (n < pendingSources.length && pendingSources[n]! < limit) {
    const move = moves.get(pendingSources[n]!)!;
    ops.push({
      ...editOp(move.opKind, move.leftBlock, null),
      moveGroupId: move.groupId,
      isMoveSource: true,
    });
    n++;
  }
  pendingSources.splice(0, n);
}

function makeModifyOp(left: IrBlock, right: IrBlock, settings: IrDiffSettings): IrEditOp {
  if (left.kind === 'table' && right.kind === 'table') {
    return { ...editOp('ModifyBlock', left, right), tableDiff: diffIrTables(left, right, settings) };
  }
  if (left.kind === 'paragraph' && right.kind === 'paragraph') {
    return makeParagraphModifyOp(left, right, settings);
  }
  return { ...editOp('ModifyBlock', left, right), tokenDiff: tokenDiffFor(left, right, settings) };
}

function makeParagraphModifyOp(left: IrParagraph, right: IrParagraph, settings: IrDiffSettings): IrEditOp {
  const leftBoxes = collectTextboxes(left.inlines);
  const rightBoxes = collectTextboxes(right.inlines);
  const textboxDiffs = buildTextboxDiffs(leftBoxes, rightBoxes, settings);
  const nest = textboxDiffs !== null;
  const leftTokens = tokenizeIrParagraph(left, settings);
  const rightTokens = tokenizeIrParagraph(right, settings);
  const tokenDiff = diffIrTokens(nest ? maskTextboxKeys(leftTokens) : leftTokens, nest ? maskTextboxKeys(rightTokens) : rightTokens, settings);
  return {
    ...editOp('ModifyBlock', left, right),
    tokenDiff,
    textboxDiffs: nest ? textboxDiffs : null,
  };
}

function collectTextboxes(inlines: ReadonlyArray<IrInline>): IrTextbox[] {
  const boxes: IrTextbox[] = [];
  walkForTextboxes(inlines, boxes);
  return boxes;
}

function walkForTextboxes(inlines: ReadonlyArray<IrInline>, sink: IrTextbox[]): void {
  for (const inline of inlines) {
    if (inline.kind === 'textbox') sink.push(inline);
    else if (inline.kind === 'fieldRun') walkForTextboxes(inline.cachedResult, sink);
    else if (inline.kind === 'hyperlink') walkForTextboxes(inline.inlines, sink);
  }
}

function buildTextboxDiffs(
  leftBoxes: ReadonlyArray<IrTextbox>,
  rightBoxes: ReadonlyArray<IrTextbox>,
  settings: IrDiffSettings,
): IrTextboxDiff[] | null {
  if (leftBoxes.length === 0 && rightBoxes.length === 0) return null;
  const paired = Math.min(leftBoxes.length, rightBoxes.length);
  const diffs: IrTextboxDiff[] = [];
  let anyChange = false;
  for (let i = 0; i < paired; i++) {
    const alignment = alignIrBlocks(leftBoxes[i]!.blocks, rightBoxes[i]!.blocks, settings);
    const ops = projectIrAlignment(leftBoxes[i]!.blocks, alignment, settings);
    diffs.push({ ops });
    if (ops.some((o) => o.kind !== 'EqualBlock')) anyChange = true;
  }
  for (let i = paired; i < leftBoxes.length; i++) {
    const ops = leftBoxes[i]!.blocks.map((b) => editOp('DeleteBlock', b, null));
    diffs.push({ ops });
    if (ops.length > 0) anyChange = true;
  }
  for (let i = paired; i < rightBoxes.length; i++) {
    const ops = rightBoxes[i]!.blocks.map((b) => editOp('InsertBlock', null, b));
    diffs.push({ ops });
    if (ops.length > 0) anyChange = true;
  }
  return anyChange ? diffs : null;
}

const MASKED_TEXTBOX_KEY = '\u0001tbx';

function maskTextboxKeys(tokens: ReadonlyArray<IrDiffToken>): IrDiffToken[] {
  return tokens.map((t) => (t.kind === 'Textbox' ? { ...t, matchKey: MASKED_TEXTBOX_KEY } : t));
}

function tokenDiffFor(left: IrBlock, right: IrBlock, settings: IrDiffSettings): IrTokenDiff | null {
  if (left.kind === 'paragraph' && right.kind === 'paragraph') {
    return diffIrTokens(tokenizeIrParagraph(left, settings), tokenizeIrParagraph(right, settings), settings);
  }
  return null;
}

function buildLeftIndexMap(blocks: ReadonlyArray<IrBlock>): Map<IrBlock, number> {
  const map = new Map<IrBlock, number>();
  for (let i = 0; i < blocks.length; i++) map.set(blocks[i]!, i);
  return map;
}

function buildSourceInterleave(
  blocks: ReadonlyArray<IrBlock>,
  alignment: IrBlockAlignment,
  leftIndex: ReadonlyMap<IrBlock, number>,
  moves: ReadonlyMap<number, MoveInfo>,
): Map<number, number[]> {
  const pairedInPlace = new Set<number>();
  for (const entry of alignment.entries) {
    if (entry.left !== null && isPairedInPlace(entry.kind)) pairedInPlace.add(leftIndex.get(entry.left)!);
    if (entry.kind === 'Merge' && entry.multiBlocks !== null) {
      for (const lb of entry.multiBlocks) pairedInPlace.add(leftIndex.get(lb)!);
    }
  }

  const result = new Map<number, number[]>();
  let lastPairedLeft = -1;
  for (let i = 0; i < blocks.length; i++) {
    if (moves.has(i)) {
      const list = result.get(lastPairedLeft) ?? [];
      list.push(i);
      result.set(lastPairedLeft, list);
    } else if (pairedInPlace.has(i)) {
      lastPairedLeft = i;
    }
  }
  return result;
}

function isPairedInPlace(kind: IrAlignmentKind): boolean {
  return kind === 'Unchanged' || kind === 'FormatOnly' || kind === 'Modified' || kind === 'Split';
}

function buildNoteOps(left: IrDocument, right: IrDocument, settings: IrDiffSettings) {
  return [
    ...buildOneNoteStore(left, right, 'Footnote', settings),
    ...buildOneNoteStore(left, right, 'Endnote', settings),
  ];
}

function buildHeaderFooterOps(left: IrDocument, right: IrDocument, settings: IrDiffSettings) {
  if (!settings.compareHeadersFooters) return [];
  const diffs: IrHeaderFooterDiff[] = [];
  buildOneHeaderFooterScope(left.headers, right.headers, true, settings, diffs);
  buildOneHeaderFooterScope(left.footers, right.footers, false, settings, diffs);
  return diffs;
}

const HEADER_FOOTER_KINDS: IrHeaderFooterKind[] = ['Default', 'First', 'Even'];

function buildOneHeaderFooterScope(
  leftParts: ReadonlyArray<IrHeaderFooter>,
  rightParts: ReadonlyArray<IrHeaderFooter>,
  isHeader: boolean,
  settings: IrDiffSettings,
  diffs: IrHeaderFooterDiff[],
): void {
  let sectionCount = 0;
  for (const hf of [...leftParts, ...rightParts]) {
    for (const ref of hf.references) sectionCount = Math.max(sectionCount, ref.sectionIndex + 1);
  }
  const leftGrid = effectiveStoryGrid(leftParts, sectionCount);
  const rightGrid = effectiveStoryGrid(rightParts, sectionCount);
  const seenPairs = new Set<string>();
  const usedLeftParts = new Set<string>();
  for (let section = 0; section < sectionCount; section++) {
    for (const kind of HEADER_FOOTER_KINDS) {
      const l = leftGrid.get(hfKey(section, kind)) ?? null;
      const r = rightGrid.get(hfKey(section, kind)) ?? null;
      if (!l && !r) continue;
      const pairKey = `${l?.scope.partUri ?? ''}\0${r?.scope.partUri ?? ''}`;
      if (seenPairs.has(pairKey)) continue;
      seenPairs.add(pairKey);
      if (l && usedLeftParts.has(l.scope.partUri)) continue;
      if (l) usedLeftParts.add(l.scope.partUri);

      let ops: IrEditOp[];
      if (l && r) {
        const alignment = alignIrBlocks(l.scope.blocks, r.scope.blocks, settings);
        ops = projectIrAlignment(l.scope.blocks, alignment, settings);
        if (!ops.some((o) => o.kind !== 'EqualBlock')) continue;
      } else if (r) {
        ops = r.scope.blocks.map((b) => editOp('InsertBlock', null, b));
      } else {
        ops = l!.scope.blocks.map((b) => editOp('DeleteBlock', b, null));
      }
      if (ops.length === 0) continue;
      diffs.push({
        isHeader,
        kind,
        sectionIndex: section,
        scopeName: (r ?? l)!.scopeName,
        leftScopeName: l?.scopeName ?? null,
        leftPartUri: l?.scope.partUri ?? null,
        rightPartUri: r?.scope.partUri ?? null,
        ops,
      });
    }
  }
}

function effectiveStoryGrid(parts: ReadonlyArray<IrHeaderFooter>, sectionCount: number): Map<string, IrHeaderFooter> {
  const explicit = new Map<string, IrHeaderFooter>();
  for (const hf of parts) {
    for (const ref of hf.references) {
      const key = hfKey(ref.sectionIndex, ref.kind);
      if (!explicit.has(key)) explicit.set(key, hf);
    }
  }
  const grid = new Map<string, IrHeaderFooter>();
  for (const kind of HEADER_FOOTER_KINDS) {
    let current: IrHeaderFooter | null = null;
    for (let section = 0; section < sectionCount; section++) {
      current = explicit.get(hfKey(section, kind)) ?? current;
      if (current) grid.set(hfKey(section, kind), current);
    }
  }
  return grid;
}

function hfKey(section: number, kind: IrHeaderFooterKind): string {
  return `${section}\0${kind}`;
}

function buildOneNoteStore(left: IrDocument, right: IrDocument, kind: IrNoteKind, settings: IrDiffSettings): IrNoteDiff[] {
  const leftStore = kind === 'Footnote' ? left.footnotes : left.endnotes;
  const rightStore = kind === 'Footnote' ? right.footnotes : right.endnotes;
  const refsLeft = distinctInOrder(collectNoteReferenceOrder(left, kind), leftStore);
  const refsRight = distinctInOrder(collectNoteReferenceOrder(right, kind), rightStore);
  const correspondence = alignNoteReferences(refsLeft, refsRight, leftStore, rightStore, settings);
  appendUnreferencedNotes(correspondence, refsLeft, refsRight, leftStore, rightStore);
  const diffs: IrNoteDiff[] = [];
  for (const [leftId, rightId] of correspondence) {
    const hasLeft = leftId !== null && leftStore.notes.has(leftId);
    const hasRight = rightId !== null && rightStore.notes.has(rightId);
    const noteId = rightId ?? leftId!;
    let ops: IrEditOp[];
    if (hasLeft && hasRight) {
      const leftScope = leftStore.notes.get(leftId!)!;
      const rightScope = rightStore.notes.get(rightId!)!;
      ops = projectIrAlignment(leftScope.blocks, alignIrBlocks(leftScope.blocks, rightScope.blocks, settings), settings);
    } else if (hasRight) {
      ops = rightStore.notes.get(rightId!)!.blocks.map((b) => editOp('InsertBlock', null, b));
    } else {
      ops = leftStore.notes.get(leftId!)!.blocks.map((b) => editOp('DeleteBlock', b, null));
    }
    const hasRealChange = ops.some((o) => o.kind !== 'EqualBlock');
    const idShifted = hasLeft && hasRight && leftId !== rightId;
    if (hasRealChange || idShifted) diffs.push({ kind, noteId, ops, leftNoteId: hasLeft ? leftId : null });
  }
  return diffs;
}

function collectNoteReferenceOrder(doc: IrDocument, kind: IrNoteKind): string[] {
  const ids: string[] = [];
  walkBlocksForNoteRefs(doc.body.blocks, kind, ids);
  return ids;
}

function walkBlocksForNoteRefs(blocks: ReadonlyArray<IrBlock>, kind: IrNoteKind, sink: string[]): void {
  for (const block of blocks) {
    if (block.kind === 'paragraph') walkInlinesForNoteRefs(block.inlines, kind, sink);
    else if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) walkBlocksForNoteRefs(cell.blocks, kind, sink);
  }
}

function walkInlinesForNoteRefs(inlines: ReadonlyArray<IrInline>, kind: IrNoteKind, sink: string[]): void {
  for (const inline of inlines) {
    if (inline.kind === 'noteRef' && inline.noteKind === kind) sink.push(inline.noteId);
    else if (inline.kind === 'fieldRun') walkInlinesForNoteRefs(inline.cachedResult, kind, sink);
    else if (inline.kind === 'hyperlink') walkInlinesForNoteRefs(inline.inlines, kind, sink);
    else if (inline.kind === 'textbox') walkBlocksForNoteRefs(inline.blocks, kind, sink);
  }
}

function distinctInOrder(ids: string[], store: IrNoteStore): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const id of ids) if (store.notes.has(id) && !seen.has(id)) {
    seen.add(id);
    result.push(id);
  }
  return result;
}

function alignNoteReferences(
  refsLeft: string[],
  refsRight: string[],
  leftStore: IrNoteStore,
  rightStore: IrNoteStore,
  settings: IrDiffSettings,
): Array<[string | null, string | null]> {
  const leftPartner = new Array(refsLeft.length).fill(-1) as number[];
  const rightPartner = new Array(refsRight.length).fill(-1) as number[];
  const spine = longestCommonSubsequence(refsLeft.length, refsRight.length, (li, rj) =>
    noteContentEqual(leftStore.notes.get(refsLeft[li]!)!, rightStore.notes.get(refsRight[rj]!)!),
  );
  for (const [li, rj] of spine) {
    leftPartner[li] = rj;
    rightPartner[rj] = li;
  }
  const bagCache = new Map<string, Map<string, number>>();
  greedyResiduePair(refsLeft.length, refsRight.length, leftPartner, rightPartner, (li, rj) =>
    bagJaccard(
      noteTokenBag(bagCache, `L:${refsLeft[li]!}`, leftStore.notes.get(refsLeft[li]!)!, settings),
      noteTokenBag(bagCache, `R:${refsRight[rj]!}`, rightStore.notes.get(refsRight[rj]!)!, settings),
    ),
  );
  const result: Array<[string | null, string | null]> = [];
  let nextLeft = 0;
  for (let rj = 0; rj < refsRight.length; rj++) {
    const li = rightPartner[rj]!;
    if (li >= 0) {
      while (nextLeft < li) {
        if (leftPartner[nextLeft] === -1) result.push([refsLeft[nextLeft]!, null]);
        nextLeft++;
      }
      nextLeft = li + 1;
      result.push([refsLeft[li]!, refsRight[rj]!]);
    } else {
      result.push([null, refsRight[rj]!]);
    }
  }
  for (; nextLeft < refsLeft.length; nextLeft++) if (leftPartner[nextLeft] === -1) result.push([refsLeft[nextLeft]!, null]);
  return result;
}

function longestCommonSubsequence(nLeft: number, nRight: number, match: (li: number, rj: number) => boolean): Array<[number, number]> {
  const dp = Array.from({ length: nLeft + 1 }, () => new Array(nRight + 1).fill(0) as number[]);
  for (let i = nLeft - 1; i >= 0; i--) {
    for (let j = nRight - 1; j >= 0; j--) dp[i]![j] = match(i, j) ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
  }
  const pairs: Array<[number, number]> = [];
  for (let i = 0, j = 0; i < nLeft && j < nRight;) {
    if (match(i, j)) {
      pairs.push([i, j]);
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) i++;
    else j++;
  }
  return pairs;
}

function greedyResiduePair(nLeft: number, nRight: number, leftPartner: number[], rightPartner: number[], score: (li: number, rj: number) => number): void {
  for (;;) {
    let best = Number.NEGATIVE_INFINITY;
    let bestLi = -1;
    let bestRj = -1;
    for (let li = 0; li < nLeft; li++) {
      if (leftPartner[li]! >= 0) continue;
      for (let rj = 0; rj < nRight; rj++) {
        if (rightPartner[rj]! >= 0 || crossesExistingPair(li, rj, leftPartner)) continue;
        const s = score(li, rj);
        if (s > best) {
          best = s;
          bestLi = li;
          bestRj = rj;
        }
      }
    }
    if (bestLi < 0) return;
    leftPartner[bestLi] = bestRj;
    rightPartner[bestRj] = bestLi;
  }
}

function crossesExistingPair(li: number, rj: number, leftPartner: readonly number[]): boolean {
  for (let k = 0; k < leftPartner.length; k++) {
    const p = leftPartner[k]!;
    if (p < 0) continue;
    if ((k < li && p > rj) || (k > li && p < rj)) return true;
  }
  return false;
}

function noteContentEqual(a: { readonly blocks: ReadonlyArray<IrBlock> }, b: { readonly blocks: ReadonlyArray<IrBlock> }): boolean {
  return a.blocks.length === b.blocks.length && a.blocks.every((block, i) => block.contentHash === b.blocks[i]!.contentHash);
}

function noteTokenBag(cache: Map<string, Map<string, number>>, key: string, scope: { readonly blocks: ReadonlyArray<IrBlock> }, settings: IrDiffSettings): Map<string, number> {
  const cached = cache.get(key);
  if (cached) return cached;
  const bag = new Map<string, number>();
  for (const block of scope.blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const token of tokenizeIrParagraph(block, settings)) {
      if (token.kind === 'Word') bag.set(token.matchKey, (bag.get(token.matchKey) ?? 0) + 1);
    }
  }
  cache.set(key, bag);
  return bag;
}

function bagJaccard(a: ReadonlyMap<string, number>, b: ReadonlyMap<string, number>): number {
  let totalA = 0;
  let totalB = 0;
  for (const v of a.values()) totalA += v;
  for (const v of b.values()) totalB += v;
  if (totalA === 0 && totalB === 0) return 1;
  if (totalA === 0 || totalB === 0) return 0;
  let intersection = 0;
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  for (const [key, value] of small) intersection += Math.min(value, large.get(key) ?? 0);
  return intersection === 0 ? 0 : intersection / (totalA + totalB - intersection);
}

function appendUnreferencedNotes(
  correspondence: Array<[string | null, string | null]>,
  refsLeft: string[],
  refsRight: string[],
  leftStore: IrNoteStore,
  rightStore: IrNoteStore,
): void {
  const referencedLeft = new Set(refsLeft);
  const referencedRight = new Set(refsRight);
  const ids = [...new Set([...leftStore.notes.keys(), ...rightStore.notes.keys()])]
    .filter((id) => !referencedLeft.has(id) || !referencedRight.has(id))
    .sort(compareNoteIds);
  for (const id of ids) {
    const l = leftStore.notes.has(id) && !referencedLeft.has(id);
    const r = rightStore.notes.has(id) && !referencedRight.has(id);
    if (r && !l) {
      const idx = correspondence.findIndex(([left, right]) => left === id && right === null);
      if (idx >= 0) {
        correspondence[idx] = [id, id];
        continue;
      }
    } else if (l && !r) {
      const idx = correspondence.findIndex(([left, right]) => right === id && left === null);
      if (idx >= 0) {
        correspondence[idx] = [id, id];
        continue;
      }
    }
    if (l && r) correspondence.push([id, id]);
    else if (r) correspondence.push([null, id]);
    else if (l) correspondence.push([id, null]);
  }
}

function compareNoteIds(a: string, b: string): number {
  const an = /^-?\d+$/.test(a);
  const bn = /^-?\d+$/.test(b);
  if (an && bn) return Number(a) - Number(b);
  if (an) return -1;
  if (bn) return 1;
  return a < b ? -1 : a > b ? 1 : 0;
}

export const IrEditScriptBuilder = {
  build: buildIrEditScript,
  projectAlignment: projectIrAlignment,
};
