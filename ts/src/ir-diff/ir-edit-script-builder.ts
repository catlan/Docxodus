import { anchorToString } from '../ir/ir-anchor.js';
import type { IrBlock, IrParagraph } from '../ir/ir-blocks.js';
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
import type { IrEditOp, IrEditOpKind, IrEditScript, IrTextboxDiff } from './ir-edit-script.js';
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
  if (left.footnotes.size === 0 && right.footnotes.size === 0 && left.endnotes.size === 0 && right.endnotes.size === 0) {
    return [];
  }
  void settings;
  return [];
}

function buildHeaderFooterOps(left: IrDocument, right: IrDocument, settings: IrDiffSettings) {
  if (!settings.compareHeadersFooters || (left.headers.length === 0 && right.headers.length === 0 && left.footers.length === 0 && right.footers.length === 0)) {
    return [];
  }
  return [];
}

export const IrEditScriptBuilder = {
  build: buildIrEditScript,
  projectAlignment: projectIrAlignment,
};
