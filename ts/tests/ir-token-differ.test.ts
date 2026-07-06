import { describe, expect, test } from 'vitest';
import { IrDiffTokenizer, IrTokenDiffer, irHashCompute, type IrDiffToken, type IrRunFormat, type IrTokenDiff } from '../src/index.js';
import { fromBodyXml } from './helpers/ir-test-documents.js';
import { assertTokenDiffInvariants } from './helpers/ir-token-diff-asserts.js';

const plain: IrRunFormat = {
  styleId: null,
  bold: false,
  italic: null,
  underline: null,
  strike: null,
  doubleStrike: null,
  vertAlign: null,
  fontAscii: null,
  sizeHalfPoints: null,
  colorHex: null,
  highlight: null,
  caps: null,
  smallCaps: null,
  vanish: null,
  unmodeledDigest: irHashCompute(''),
};
const bold: IrRunFormat = { ...plain, bold: true };

const w = (word: string, format: IrRunFormat = plain): IrDiffToken => ({
  kind: 'Word',
  text: word,
  matchKey: word,
  startChar: 0,
  endChar: word.length,
  format,
});

const sep = (text = ' ', format: IrRunFormat = plain): IrDiffToken => ({
  kind: 'Separator',
  text,
  matchKey: text,
  startChar: 0,
  endChar: text.length,
  format,
});

function words(...items: string[]): IrDiffToken[] {
  const tokens: IrDiffToken[] = [];
  for (let i = 0; i < items.length; i++) {
    if (i > 0) tokens.push(sep());
    tokens.push(w(items[i]!));
  }
  return tokens;
}

function diff(left: ReadonlyArray<IrDiffToken>, right: ReadonlyArray<IrDiffToken>): IrTokenDiff {
  const d = IrTokenDiffer.diff(left, right);
  assertTokenDiffInvariants(left, right, d);
  return d;
}

const sig = (d: IrTokenDiff): string =>
  d.ops.map((o) => `${o.kind}(${o.leftStart},${o.leftEnd}|${o.rightStart},${o.rightEnd})`).join(' ');

describe('IrTokenDiffer', () => {
  test('single word change in the middle', () => {
    expect(sig(diff(words('the', 'quick', 'fox'), words('the', 'slow', 'fox')))).toBe(
      'Equal(0,2|0,2) Delete(2,3|2,2) Insert(3,3|2,3) Equal(3,5|3,5)',
    );
  });

  test('prefix edit only', () => {
    expect(sig(diff(words('alpha', 'beta', 'gamma'), words('zeta', 'beta', 'gamma')))).toBe(
      'Delete(0,1|0,0) Insert(1,1|0,1) Equal(1,5|1,5)',
    );
  });

  test('suffix edit only', () => {
    expect(sig(diff(words('alpha', 'beta', 'gamma'), words('alpha', 'beta', 'omega')))).toBe(
      'Equal(0,4|0,4) Delete(4,5|4,4) Insert(5,5|4,5)',
    );
  });

  test('pure insertion at end', () => {
    expect(sig(diff(words('a', 'b'), words('a', 'b', 'c')))).toBe('Equal(0,3|0,3) Insert(3,3|3,5)');
  });

  test('pure deletion at end', () => {
    expect(sig(diff(words('a', 'b', 'c'), words('a', 'b')))).toBe('Equal(0,3|0,3) Delete(3,5|3,3)');
  });

  test('all changed no common tokens still covers both sides', () => {
    const d = diff(words('one', 'two'), words('three', 'four'));
    expect(d.ops.some((o) => o.kind === 'Delete')).toBe(true);
    expect(d.ops.some((o) => o.kind === 'Insert')).toBe(true);
  });

  test('all changed truly disjoint tokens', () => {
    expect(sig(diff([w('xxx')], [w('yyy')]))).toBe('Delete(0,1|0,0) Insert(1,1|0,1)');
  });

  test('all equal identical lists', () => {
    const d = diff(words('same', 'text', 'here'), words('same', 'text', 'here'));
    expect(sig(d)).toBe('Equal(0,5|0,5)');
    expect(d.ops).toHaveLength(1);
  });

  test('separator only change', () => {
    expect(sig(diff([w('a'), sep('-'), w('b')], [w('a'), sep(' '), w('b')]))).toBe(
      'Equal(0,1|0,1) Delete(1,2|1,1) Insert(2,2|1,2) Equal(2,3|2,3)',
    );
  });

  test('bold word becomes FormatChanged span exactly over that word', () => {
    const d = diff([w('plain'), sep(), w('word'), sep(), w('tail')], [w('plain'), sep(), w('word', bold), sep(), w('tail')]);
    expect(sig(d)).toBe('Equal(0,2|0,2) FormatChanged(2,3|2,3) Equal(3,5|3,5)');
  });

  test('format changes separated by unchanged separator stay distinct spans', () => {
    const d = diff(words('a', 'b', 'c'), [w('a', bold), sep(), w('b', bold), sep(), w('c')]);
    expect(sig(d)).toBe('FormatChanged(0,1|0,1) Equal(1,2|1,2) FormatChanged(2,3|2,3) Equal(3,5|3,5)');
  });

  test('contiguous format changes with no separator merge', () => {
    expect(sig(diff([w('aa'), w('bb')], [w('aa', bold), w('bb', bold)]))).toBe('FormatChanged(0,2|0,2)');
  });

  test('empty left yields one insert', () => {
    expect(sig(diff([], words('a', 'b')))).toBe('Insert(0,0|0,3)');
  });

  test('empty right yields one delete', () => {
    expect(sig(diff(words('a', 'b'), []))).toBe('Delete(0,3|0,0)');
  });

  test('empty both yields no ops', () => {
    expect(diff([], []).ops).toEqual([]);
  });

  test('repeated words insert one more', () => {
    expect(sig(diff(words('the', 'the', 'the'), words('the', 'the', 'the', 'the')))).toBe('Equal(0,5|0,5) Insert(5,5|5,7)');
  });

  test('repeated words with distinct tail', () => {
    expect(sig(diff(words('the', 'the', 'the', 'a', 'the', 'the'), words('the', 'the', 'a', 'the', 'the', 'the')))).toBe(
      'Equal(0,4|0,4) Delete(4,6|4,4) Equal(6,11|4,9) Insert(11,11|9,11)',
    );
  });

  test('determinism same inputs same ops', () => {
    const left = words('the', 'the', 'the', 'a', 'the', 'the');
    const right = words('the', 'the', 'a', 'the', 'the', 'the');
    expect(IrTokenDiffer.diff(left, right)).toEqual(IrTokenDiffer.diff(left, right));
  });

  test('real IR paragraphs diff end to end', () => {
    const leftPara = fromBodyXml('<w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p>').body.blocks[0]!;
    const rightPara = fromBodyXml('<w:p><w:r><w:t>The slow brown fox</w:t></w:r></w:p>').body.blocks[0]!;
    expect(leftPara.kind).toBe('paragraph');
    expect(rightPara.kind).toBe('paragraph');
    if (leftPara.kind !== 'paragraph' || rightPara.kind !== 'paragraph') throw new Error('expected paragraphs');
    const left = IrDiffTokenizer.tokenize(leftPara);
    const right = IrDiffTokenizer.tokenize(rightPara);
    const d = diff(left, right);
    expect(d.ops.some((o) => o.kind === 'Delete')).toBe(true);
    expect(d.ops.some((o) => o.kind === 'Insert')).toBe(true);
    expect(d.ops.some((o) => o.kind === 'Equal')).toBe(true);
  });
});
