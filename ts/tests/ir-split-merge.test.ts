import { describe, expect, test } from 'vitest';
import {
  IrBlockAligner,
  IrDiffTokenizer,
  IrSplitSegmenter,
  type IrDiffSettings,
  type IrDocument,
  type IrParagraph,
  type IrTokenDiff,
  type IrTokenOp,
} from '../src/index.js';
import { assertInvariants, count } from './helpers/ir-alignment-asserts.js';
import { fromBodyXml, p } from './helpers/ir-test-documents.js';

const S: Partial<IrDiffSettings> = { detectSplitMerge: true };

const diff = (...ops: IrTokenOp[]): IrTokenDiff => ({ ops });

function splitOp() {
  return {
    kind: 'SplitBlock',
    leftAnchor: 'p:body:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
    rightAnchor: null,
    tokenDiff: null,
    moveGroupId: null,
    isMoveSource: null,
    splitMergeAnchors: ['p:body:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb', 'p:body:cccccccccccccccccccccccccccccccc'],
    segmentDiffs: [
      diff({ kind: 'Equal', leftStart: 0, leftEnd: 3, rightStart: 0, rightEnd: 3 }),
      diff({ kind: 'Equal', leftStart: 0, leftEnd: 2, rightStart: 0, rightEnd: 2 }, { kind: 'Insert', leftStart: 2, leftEnd: 2, rightStart: 2, rightEnd: 4 }),
    ],
  };
}

function mergeOp() {
  return {
    kind: 'MergeBlock',
    leftAnchor: null,
    rightAnchor: 'p:body:99999999999999999999999999999999',
    tokenDiff: null,
    moveGroupId: null,
    isMoveSource: null,
    splitMergeAnchors: ['p:body:11111111111111111111111111111111', 'p:body:22222222222222222222222222222222'],
    segmentDiffs: [
      diff({ kind: 'Equal', leftStart: 0, leftEnd: 2, rightStart: 0, rightEnd: 2 }),
      diff({ kind: 'Equal', leftStart: 0, leftEnd: 2, rightStart: 2, rightEnd: 4 }),
    ],
  };
}

function readParas(...texts: string[]): { doc: IrDocument; paras: IrParagraph[] } {
  const doc = fromBodyXml(texts.map(p).join(''));
  return { doc, paras: doc.body.blocks.filter((b): b is IrParagraph => b.kind === 'paragraph') };
}

const align = (l: IrDocument, r: IrDocument, settings: Partial<IrDiffSettings> = S) => IrBlockAligner.align(l, r, settings);

describe('IrSplitMerge', () => {
  test('split op model carries split fields deterministically', () => {
    expect(splitOp()).toEqual(splitOp());
    expect(splitOp().rightAnchor).toBeNull();
    expect(splitOp().splitMergeAnchors).toHaveLength(2);
    expect(splitOp().segmentDiffs).toHaveLength(2);
  });

  test('merge op model carries merge fields deterministically', () => {
    expect(mergeOp()).toEqual(mergeOp());
    expect(mergeOp().leftAnchor).toBeNull();
    expect(mergeOp().splitMergeAnchors).toHaveLength(2);
    expect(mergeOp().segmentDiffs).toHaveLength(2);
  });

  test('scripts without splits omit split fields when absent', () => {
    const op = { kind: 'InsertBlock', leftAnchor: null, rightAnchor: 'p:body:dddddddddddddddddddddddddddddddd', tokenDiff: null };
    expect('splitMergeAnchors' in op).toBe(false);
    expect('segmentDiffs' in op).toBe(false);
  });

  test('segmenter scores a clean split at full coverage zero slack', () => {
    const { paras: lp } = readParas('alpha bravo charlie. delta echo foxtrot.');
    const { paras: rp } = readParas('alpha bravo charlie. ', 'delta echo foxtrot.');
    const score = IrSplitSegmenter.score(lp[0]!, [rp[0]!, rp[1]!], S);
    expect(score.coverage).toBeGreaterThanOrEqual(0.99);
    expect(score.foreignSlack).toBeLessThanOrEqual(0.01);
  });

  test('segmenter scores keyword coincidence below threshold', () => {
    const { paras: lp } = readParas('the contract terminates on delivery of the goods.');
    const { paras: rp } = readParas('the parties agree on many things.', 'delivery of pizza is unrelated to the goods here.');
    const score = IrSplitSegmenter.score(lp[0]!, [rp[0]!, rp[1]!], S);
    expect(score.coverage < 0.9 || score.foreignSlack > 0.34).toBe(true);
  });

  test('segmenter segment diffs tile the left token stream exactly', () => {
    const { paras: lp } = readParas('alpha bravo charlie. delta echo foxtrot.');
    const { paras: rp } = readParas('alpha bravo charlie. ', 'NEW WORDS HERE', 'delta echo foxtrot.');
    const rights = [rp[0]!, rp[1]!, rp[2]!];
    const diffs = IrSplitSegmenter.computeSegmentDiffs(lp[0]!, rights, S);
    expect(diffs).toHaveLength(3);
    const leftTotal = IrDiffTokenizer.tokenize(lp[0]!, S).length;
    expect(diffs.reduce((sum, d) => sum + d.ops.filter((o) => o.kind !== 'Insert').reduce((s, o) => s + o.leftEnd - o.leftStart, 0), 0)).toBe(leftTotal);
    for (let i = 0; i < rights.length; i++) {
      const rightCount = IrDiffTokenizer.tokenize(rights[i]!, S).length;
      expect(diffs[i]!.ops.filter((o) => o.kind !== 'Delete').reduce((s, o) => s + o.rightEnd - o.rightStart, 0)).toBe(rightCount);
    }
  });

  test('mirrorDiff swaps sides and flips insert delete', () => {
    const original = diff({ kind: 'Delete', leftStart: 0, leftEnd: 2, rightStart: 0, rightEnd: 0 }, { kind: 'Equal', leftStart: 2, leftEnd: 5, rightStart: 0, rightEnd: 3 });
    const mirrored = IrSplitSegmenter.mirrorDiff(original);
    expect(mirrored.ops[0]).toEqual({ kind: 'Insert', leftStart: 0, leftEnd: 0, rightStart: 0, rightEnd: 2 });
    expect(mirrored.ops[1]).toEqual({ kind: 'Equal', leftStart: 0, leftEnd: 3, rightStart: 2, rightEnd: 5 });
    expect(IrSplitSegmenter.mirrorDiff(mirrored)).toEqual(original);
  });

  test('detection fires for a clean two way split', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd. eee fff ggg hhh.', 'unrelated anchor paragraph one two three.');
    const { doc: r } = readParas('aaa bbb ccc ddd. ', 'eee fff ggg hhh.', 'unrelated anchor paragraph one two three.');
    const a = align(l, r);
    assertInvariants(l, r, a, S);
    const split = a.entries.find((e) => e.kind === 'Split')!;
    expect(split.multiBlocks).toHaveLength(2);
    expect(count(a, 'Inserted')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
  });

  test('detection fires for a fully free three way split', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd. eee fff ggg hhh. iii jjj kkk lll.', 'anchor one two three four five.');
    const { doc: r } = readParas('aaa bbb ccc ddd. ', 'eee fff ggg hhh. ', 'iii jjj kkk lll.', 'anchor one two three four five.');
    const a = align(l, r);
    assertInvariants(l, r, a, S);
    expect(a.entries.find((e) => e.kind === 'Split')!.multiBlocks).toHaveLength(3);
    expect(count(a, 'Inserted')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
  });

  test('detection absorbs an interior net new block', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd eee fff. ggg hhh iii jjj kkk lll.');
    const { doc: r } = readParas('aaa bbb ccc ddd eee fff. ', 'zzz', 'ggg hhh iii jjj kkk lll.');
    const a = align(l, r);
    expect(a.entries.find((e) => e.kind === 'Split')!.multiBlocks).toHaveLength(3);
    expect(count(a, 'Inserted')).toBe(0);
  });

  test('detection promotes a similarity paired prefix with trailing tail inserts', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd eee fff ggg hhh iii jjj. kkk lll.');
    const { doc: r } = readParas('aaa bbb ccc ddd eee fff ggg hhh iii jjj. ', 'kkk lll.');
    const a = align(l, r);
    expect(a.entries.filter((e) => e.kind === 'Split')).toHaveLength(1);
    expect(count(a, 'Modified')).toBe(0);
  });

  test('detection does not fire on keyword coincidence', () => {
    const { doc: l } = readParas('the contract terminates on delivery of the goods.');
    const { doc: r } = readParas('the parties agree on many things today.', 'delivery of pizza is unrelated to goods.');
    expect(align(l, r).entries.filter((e) => e.kind === 'Split' || e.kind === 'Merge')).toHaveLength(0);
  });

  test('detection excludes an unrelated edge insert from the run', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd. eee fff ggg hhh.', 'anchor one two three four five.');
    const { doc: r } = readParas('aaa bbb ccc ddd. ', 'eee fff ggg hhh.', 'totally unrelated new paragraph words.', 'anchor one two three four five.');
    const a = align(l, r);
    expect(a.entries.find((e) => e.kind === 'Split')!.multiBlocks).toHaveLength(2);
    expect(count(a, 'Inserted')).toBe(1);
  });

  test('detection never promotes an identity reserved unchanged pair', () => {
    const { doc: l } = readParas('same text here one two three.');
    const { doc: r } = readParas('same text here one two three.', 'a new paragraph appended after.');
    const a = align(l, r);
    expect(a.entries.filter((e) => e.kind === 'Split')).toHaveLength(0);
    expect(count(a, 'Unchanged')).toBe(1);
    expect(count(a, 'Inserted')).toBe(1);
  });

  test('detection merge mirror fires for a clean merge', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd. ', 'eee fff ggg hhh.', 'anchor one two three four.');
    const { doc: r } = readParas('aaa bbb ccc ddd. eee fff ggg hhh.', 'anchor one two three four.');
    expect(align(l, r).entries.find((e) => e.kind === 'Merge')!.multiBlocks).toHaveLength(2);
  });

  test('detection two adjacent splits never share a right block', () => {
    const { doc: l } = readParas('aaa bbb ccc. ddd eee fff.', 'ggg hhh iii. jjj kkk lll.');
    const { doc: r } = readParas('aaa bbb ccc. ', 'ddd eee fff.', 'ggg hhh iii. ', 'jjj kkk lll.');
    const splits = align(l, r).entries.filter((e) => e.kind === 'Split');
    expect(splits).toHaveLength(2);
    const members = splits.flatMap((e) => e.multiBlocks ?? []);
    expect(new Set(members).size).toBe(members.length);
  });

  test('detection disabled changes nothing', () => {
    const { doc: l } = readParas('aaa bbb ccc ddd. eee fff ggg hhh.');
    const { doc: r } = readParas('aaa bbb ccc ddd. ', 'eee fff ggg hhh.');
    expect(align(l, r, { detectSplitMerge: false }).entries.filter((e) => e.kind === 'Split' || e.kind === 'Merge')).toHaveLength(0);
  });
});
