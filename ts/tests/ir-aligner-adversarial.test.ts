import { describe, expect, test } from 'vitest';
import { IrBlockAligner } from '../src/index.js';
import { assertInvariants, count, histogram, text } from './helpers/ir-alignment-asserts.js';
import { docFromParagraphs } from './helpers/ir-test-documents.js';

const distinctClauses = (n: number): string[] =>
  Array.from({ length: n }, (_, i) => `Clause ${i}: standard wording for this section of the agreement.`);

describe('IrBlockAligner adversarial', () => {
  test('near identical 500 one word changed yields 499 unchanged 1 modified 0 moved', () => {
    const left = distinctClauses(500);
    const right = [...left];
    right[250] = 'Clause 250: REVISED wording for this section of the agreement.';
    const l = docFromParagraphs(left);
    const r = docFromParagraphs(right);
    const a = IrBlockAligner.align(l, r);
    expect(count(a, 'Unchanged')).toBe(499);
    expect(count(a, 'Modified')).toBe(1);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('identical 500 delete one yields 499 unchanged 1 deleted 0 moved 0 modified', () => {
    const l = docFromParagraphs(Array(500).fill('Standard boilerplate clause.'));
    const r = docFromParagraphs(Array(499).fill('Standard boilerplate clause.'));
    const a = IrBlockAligner.align(l, r);
    expect(count(a, 'Unchanged')).toBe(499);
    expect(count(a, 'Deleted')).toBe(1);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Modified')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('fully rewritten 200 vs 200 no throw invariants hold runtime sane', () => {
    const left = Array.from({ length: 200 }, (_, i) => `Original paragraph ${i} with its own distinct content.`);
    const right = Array.from({ length: 200 }, (_, i) => `Completely different replacement line ${i} sharing nothing.`);
    const l = docFromParagraphs(left);
    const r = docFromParagraphs(right);
    const start = performance.now();
    const a = IrBlockAligner.align(l, r);
    const elapsed = performance.now() - start;
    assertInvariants(l, r, a);
    expect(count(a, 'Modified')).toBe(0);
    expect(count(a, 'Deleted')).toBe(200);
    expect(count(a, 'Inserted')).toBe(200);
    expect(count(a, 'Unchanged')).toBe(0);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'MovedModified')).toBe(0);
    expect(elapsed).toBeLessThan(5000);
  });

  test('move 10 unique paragraph block front to back of 300 yields exactly 10 moved', () => {
    const total = 300;
    const blockSize = 10;
    const all = distinctClauses(total);
    const movedBlock = all.slice(0, blockSize);
    const right = all.slice(blockSize).concat(movedBlock);
    const l = docFromParagraphs(all);
    const r = docFromParagraphs(right);
    const a = IrBlockAligner.align(l, r);
    assertInvariants(l, r, a);
    expect(count(a, 'Moved')).toBe(blockSize);
    expect(count(a, 'Unchanged')).toBe(total - blockSize);
    expect(count(a, 'Modified')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
    const movedTexts = new Set(a.entries.filter((e) => e.kind === 'Moved').map((e) => text(e.right!)));
    expect(movedTexts).toEqual(new Set(movedBlock));
  });

  test('scale guard 500 vs 2000 wall ratio within 12x', () => {
    const small = bestSampleMs(500);
    const large = bestSampleMs(2000);
    const ratio = large / Math.max(small, 0.0001);
    expect(ratio).toBeLessThanOrEqual(12);
  });
});

function bestSampleMs(n: number): number {
  const alignsPerSample = 10;
  const baseParas = distinctClauses(n);
  const edited = [...baseParas];
  edited[Math.floor(n / 2)] = `Clause ${Math.floor(n / 2)}: REVISED wording for this section of the agreement.`;
  const l = docFromParagraphs(baseParas);
  const r = docFromParagraphs(edited);
  IrBlockAligner.align(l, r);
  let best = Number.POSITIVE_INFINITY;
  for (let i = 0; i < 5; i++) {
    const start = performance.now();
    for (let j = 0; j < alignsPerSample; j++) IrBlockAligner.align(l, r);
    best = Math.min(best, (performance.now() - start) / alignsPerSample);
  }
  void histogram;
  return best;
}
