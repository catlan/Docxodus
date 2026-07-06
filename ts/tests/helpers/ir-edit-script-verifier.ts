import { expect } from 'vitest';
import {
  anchorToString,
  IrDiffTokenizer,
  IrSplitSegmenter,
  normalizeIrDiffSettings,
  type IrBlock,
  type IrDiffSettings,
  type IrDiffToken,
  type IrDocument,
  type IrEditOp,
  type IrEditScript,
  type IrParagraph,
  type IrTable,
  type IrTableDiff,
  type IrTokenDiff,
} from '../../src/index.js';
import { assertTokenDiffInvariants } from './ir-token-diff-asserts.js';

export function verifyIrEditScript(
  left: IrDocument,
  right: IrDocument,
  script: IrEditScript,
  options: Partial<IrDiffSettings> = {},
): void {
  const settings = normalizeIrDiffSettings(options);
  assertAnchorsResolve(left, right, script);
  assertMovePairing(script);
  assertSplitMergePairing(script);

  const moveSourceBlock = new Map<number, IrBlock>();
  for (const op of script.ops) {
    if (op.isMoveSource === true) moveSourceBlock.set(op.moveGroupId!, resolve(left.anchorIndex, op.leftAnchor!));
  }

  const reconstructed: Array<{ readonly rightBlock: IrBlock; readonly tokens: ReadonlyArray<string> | null; readonly sourceBlock: IrBlock }> = [];
  for (const op of script.ops) {
    switch (op.kind) {
      case 'EqualBlock':
      case 'FormatOnlyBlock': {
        const lb = resolve(left.anchorIndex, op.leftAnchor!);
        const rb = resolve(right.anchorIndex, op.rightAnchor!);
        reconstructed.push({ rightBlock: rb, tokens: tokensOrNull(lb, settings), sourceBlock: lb });
        break;
      }
      case 'ModifyBlock': {
        const lb = resolve(left.anchorIndex, op.leftAnchor!);
        const rb = resolve(right.anchorIndex, op.rightAnchor!);
        if (lb.kind === 'table' && rb.kind === 'table' && op.tableDiff) verifyTableDiff(lb, rb, op.tableDiff as IrTableDiff, settings);
        reconstructed.push({ rightBlock: rb, tokens: applyModify(lb, rb, op.tokenDiff as IrTokenDiff | null, settings), sourceBlock: rb });
        break;
      }
      case 'InsertBlock': {
        const rb = resolve(right.anchorIndex, op.rightAnchor!);
        reconstructed.push({ rightBlock: rb, tokens: tokensOrNull(rb, settings), sourceBlock: rb });
        break;
      }
      case 'DeleteBlock':
        break;
      case 'MoveBlock':
      case 'MoveModifyBlock': {
        if (op.isMoveSource === true) break;
        const source = moveSourceBlock.get(op.moveGroupId!)!;
        const rb = resolve(right.anchorIndex, op.rightAnchor!);
        reconstructed.push({
          rightBlock: rb,
          tokens: op.kind === 'MoveModifyBlock' ? applyModify(source, rb, op.tokenDiff as IrTokenDiff, settings) : tokensOrNull(source, settings),
          sourceBlock: op.kind === 'MoveModifyBlock' ? rb : source,
        });
        break;
      }
      case 'SplitBlock': {
        const lp = resolve(left.anchorIndex, op.leftAnchor!) as IrParagraph;
        const members = op.splitMergeAnchors!.map((a) => resolve(right.anchorIndex, a) as IrParagraph);
        const segments = applySplitOp(lp, members, op.segmentDiffs as ReadonlyArray<IrTokenDiff>, settings);
        for (let i = 0; i < members.length; i++) {
          reconstructed.push({ rightBlock: members[i]!, tokens: segments[i]!, sourceBlock: members[i]! });
        }
        break;
      }
      case 'MergeBlock': {
        const rp = resolve(right.anchorIndex, op.rightAnchor!) as IrParagraph;
        const members = op.splitMergeAnchors!.map((a) => resolve(left.anchorIndex, a) as IrParagraph);
        reconstructed.push({ rightBlock: rp, tokens: applyMergeOp(members, rp, op.segmentDiffs as ReadonlyArray<IrTokenDiff>, settings), sourceBlock: rp });
        break;
      }
    }
  }

  expect(reconstructed).toHaveLength(right.body.blocks.length);
  for (let i = 0; i < right.body.blocks.length; i++) {
    const actual = right.body.blocks[i]!;
    const got = reconstructed[i]!;
    expect(got.rightBlock).toBe(actual);
    if (got.tokens !== null) expect(got.tokens.join('')).toBe(tokens(actual as IrParagraph, settings).join(''));
    else expect(got.sourceBlock.contentHash).toBe(actual.contentHash);
  }
}

function verifyTableDiff(left: IrTable, right: IrTable, diff: IrTableDiff, settings: IrDiffSettings): void {
  const leftRows = new Map(left.rows.map((r) => [anchorToString(r.anchor), r]));
  const rightRows = new Map(right.rows.map((r) => [anchorToString(r.anchor), r]));
  const moveSources = new Map<number, string>();
  for (const op of diff.rowOps) {
    if (op.isMoveSource === true) moveSources.set(op.moveGroupId!, leftRows.get(op.leftRowAnchor!)!.contentHash);
  }
  const rowHashes: string[] = [];
  for (const op of diff.rowOps) {
    if (op.kind === 'EqualRow') {
      const lr = leftRows.get(op.leftRowAnchor!)!;
      const rr = rightRows.get(op.rightRowAnchor!)!;
      expect(lr.contentHash).toBe(rr.contentHash);
      rowHashes.push(lr.contentHash);
    } else if (op.kind === 'ModifyRow') {
      const lr = leftRows.get(op.leftRowAnchor!)!;
      const rr = rightRows.get(op.rightRowAnchor!)!;
      expect(op.cellOps).not.toBeNull();
      expect(op.cellOps).toHaveLength(Math.max(lr.cells.length, rr.cells.length));
      rowHashes.push(rr.contentHash);
      void settings;
    } else if (op.kind === 'InsertRow') {
      rowHashes.push(rightRows.get(op.rightRowAnchor!)!.contentHash);
    } else if (op.kind === 'DeleteRow') {
      expect(leftRows.has(op.leftRowAnchor!)).toBe(true);
    } else if (op.kind === 'MovedRow' && op.isMoveSource === false) {
      const rr = rightRows.get(op.rightRowAnchor!)!;
      expect(moveSources.get(op.moveGroupId!)).toBe(rr.contentHash);
      rowHashes.push(rr.contentHash);
    }
  }
  expect(rowHashes).toEqual(right.rows.map((r) => r.contentHash));
}

function assertAnchorsResolve(left: IrDocument, right: IrDocument, script: IrEditScript): void {
  for (const op of allOps(script)) {
    if (op.leftAnchor !== null) expect(left.anchorIndex.has(op.leftAnchor)).toBe(true);
    if (op.rightAnchor !== null) expect(right.anchorIndex.has(op.rightAnchor)).toBe(true);
    if (op.splitMergeAnchors) {
      const index = op.kind === 'SplitBlock' ? right.anchorIndex : left.anchorIndex;
      for (const a of op.splitMergeAnchors) expect(index.has(a)).toBe(true);
    }
  }
}

function assertMovePairing(script: IrEditScript): void {
  const sources = new Map<number, IrEditOp>();
  const destinations = new Map<number, IrEditOp>();
  const destinationOrder: number[] = [];
  for (const op of script.ops) {
    if (op.kind !== 'MoveBlock' && op.kind !== 'MoveModifyBlock') {
      expect(op.moveGroupId).toBeNull();
      expect(op.isMoveSource).toBeNull();
      continue;
    }
    expect(op.moveGroupId).not.toBeNull();
    expect(op.isMoveSource).not.toBeNull();
    if (op.isMoveSource) sources.set(op.moveGroupId!, op);
    else {
      destinations.set(op.moveGroupId!, op);
      destinationOrder.push(op.moveGroupId!);
    }
  }
  expect([...sources.keys()].sort()).toEqual([...destinations.keys()].sort());
  expect(destinationOrder).toEqual(destinationOrder.map((_, i) => i + 1));
}

export function assertSplitMergePairing(script: IrEditScript): void {
  const seen = new Set<string>();
  for (const op of allOps(script)) {
    if (op.kind !== 'SplitBlock' && op.kind !== 'MergeBlock') {
      expect(op.splitMergeAnchors ?? null).toBeNull();
      expect(op.segmentDiffs ?? null).toBeNull();
      continue;
    }
    expect(op.moveGroupId).toBeNull();
    expect(op.isMoveSource).toBeNull();
    expect(op.tokenDiff).toBeNull();
    expect(op.tableDiff ?? null).toBeNull();
    expect(op.splitMergeAnchors!.length).toBeGreaterThanOrEqual(2);
    expect(op.segmentDiffs).toHaveLength(op.splitMergeAnchors!.length);
    if (op.kind === 'SplitBlock') {
      expect(op.leftAnchor).not.toBeNull();
      expect(op.rightAnchor).toBeNull();
    } else {
      expect(op.leftAnchor).toBeNull();
      expect(op.rightAnchor).not.toBeNull();
    }
    for (const a of op.splitMergeAnchors!) {
      expect(seen.has(a)).toBe(false);
      seen.add(a);
    }
  }
}

function* allOps(script: IrEditScript): Iterable<IrEditOp> {
  function* expand(op: IrEditOp): Iterable<IrEditOp> {
    yield op;
    for (const tbx of op.textboxDiffs ?? []) for (const inner of tbx.ops) yield* expand(inner);
    const td = op.tableDiff as IrTableDiff | null | undefined;
    for (const row of td?.rowOps ?? []) {
      for (const cell of row.cellOps ?? []) for (const inner of cell.blockOps ?? []) yield* expand(inner);
    }
  }
  for (const op of script.ops) yield* expand(op);
  for (const note of script.noteOps ?? []) for (const op of note.ops) yield* expand(op);
}

function resolve(index: ReadonlyMap<string, IrBlock>, anchor: string): IrBlock {
  const block = index.get(anchor);
  expect(block, `anchor ${anchor} missing`).toBeTruthy();
  return block!;
}

function tokensOrNull(block: IrBlock, settings: IrDiffSettings): ReadonlyArray<string> | null {
  return block.kind === 'paragraph' ? tokens(block, settings) : null;
}

function tokens(p: IrParagraph, settings: IrDiffSettings): ReadonlyArray<string> {
  return maskedTokenize(p, settings).map((t) => t.matchKey);
}

function maskedTokenize(p: IrParagraph, settings: IrDiffSettings): ReadonlyArray<IrDiffToken> {
  return IrDiffTokenizer.tokenize(p, settings).map((t) => (t.kind === 'Textbox' ? { ...t, matchKey: 'tbx' } : t));
}

function applyModify(left: IrBlock, right: IrBlock, diff: IrTokenDiff | null, settings: IrDiffSettings): ReadonlyArray<string> | null {
  if (left.kind !== 'paragraph' || right.kind !== 'paragraph') return null;
  expect(diff).not.toBeNull();
  const leftTokens = maskedTokenize(left, settings);
  const rightTokens = maskedTokenize(right, settings);
  assertTokenDiffInvariants(leftTokens, rightTokens, diff!, settings);
  return applyDiff(leftTokens, rightTokens, diff!);
}

function applySplitOp(left: IrParagraph, members: ReadonlyArray<IrParagraph>, diffs: ReadonlyArray<IrTokenDiff>, settings: IrDiffSettings): ReadonlyArray<ReadonlyArray<string>> {
  const leftTokens = maskedTokenize(left, settings);
  const result: ReadonlyArray<string>[] = [];
  let offset = 0;
  for (let i = 0; i < members.length; i++) {
    const memberTokens = maskedTokenize(members[i]!, settings);
    const sliceLen = diffs[i]!.ops.filter((o) => o.kind !== 'Insert').reduce((sum, o) => sum + o.leftEnd - o.leftStart, 0);
    const slice = leftTokens.slice(offset, offset + sliceLen);
    assertTokenDiffInvariants(slice, memberTokens, diffs[i]!, settings);
    result.push(applyDiff(slice, memberTokens, diffs[i]!));
    offset += sliceLen;
  }
  expect(offset).toBe(leftTokens.length);
  return result;
}

function applyMergeOp(members: ReadonlyArray<IrParagraph>, right: IrParagraph, diffs: ReadonlyArray<IrTokenDiff>, settings: IrDiffSettings): ReadonlyArray<string> {
  const rightTokens = maskedTokenize(right, settings);
  const combined: string[] = [];
  let offset = 0;
  for (let i = 0; i < members.length; i++) {
    const memberTokens = maskedTokenize(members[i]!, settings);
    const sliceLen = diffs[i]!.ops.filter((o) => o.kind !== 'Delete').reduce((sum, o) => sum + o.rightEnd - o.rightStart, 0);
    const slice = rightTokens.slice(offset, offset + sliceLen);
    assertTokenDiffInvariants(memberTokens, slice, diffs[i]!, settings);
    combined.push(...applyDiff(memberTokens, slice, diffs[i]!));
    offset += sliceLen;
  }
  expect(offset).toBe(rightTokens.length);
  return combined;
}

function applyDiff(leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, diff: IrTokenDiff): string[] {
  const result: string[] = [];
  for (const op of diff.ops) {
    if (op.kind === 'Equal' || op.kind === 'FormatChanged') {
      for (let i = op.leftStart; i < op.leftEnd; i++) result.push(leftTokens[i]!.matchKey);
    } else if (op.kind === 'Insert') {
      for (let i = op.rightStart; i < op.rightEnd; i++) result.push(rightTokens[i]!.matchKey);
    }
  }
  return result;
}
