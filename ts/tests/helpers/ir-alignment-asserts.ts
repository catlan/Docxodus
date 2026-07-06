import { expect } from 'vitest';
import {
  blockSignature,
  type IrAlignmentKind,
  type IrBlock,
  type IrBlockAlignment,
  type IrDiffSettings,
  type IrDocument,
  normalizeIrDiffSettings,
} from '../../src/index.js';

export function count(a: IrBlockAlignment, kind: IrAlignmentKind): number {
  return a.entries.filter((e) => e.kind === kind).length;
}

export function histogram(a: IrBlockAlignment): string {
  const order: IrAlignmentKind[] = [
    'Unchanged',
    'FormatOnly',
    'Modified',
    'Moved',
    'MovedModified',
    'Inserted',
    'Deleted',
    'Split',
    'Merge',
  ];
  return order.map((k) => `${k}=${count(a, k)}`).join(' ');
}

export function assertInvariants(
  left: IrDocument,
  right: IrDocument,
  a: IrBlockAlignment,
  settings: Partial<IrDiffSettings> = {},
): void {
  const s = normalizeIrDiffSettings(settings);
  const leftSeen: IrBlock[] = [];
  const rightSeen: IrBlock[] = [];
  for (const e of a.entries) {
    switch (e.kind) {
      case 'Inserted':
        expect(e.left).toBeNull();
        expect(e.right).not.toBeNull();
        break;
      case 'Deleted':
        expect(e.left).not.toBeNull();
        expect(e.right).toBeNull();
        break;
      case 'Unchanged':
        expect(e.left).not.toBeNull();
        expect(e.right).not.toBeNull();
        expect(e.left!.contentHash).toBe(e.right!.contentHash);
        expect(formatEqual(e.left!, e.right!, s)).toBe(true);
        break;
      case 'FormatOnly':
        expect(e.left).not.toBeNull();
        expect(e.right).not.toBeNull();
        expect(e.left!.contentHash).toBe(e.right!.contentHash);
        expect(formatEqual(e.left!, e.right!, s)).toBe(false);
        break;
      case 'Moved':
        expect(e.left).not.toBeNull();
        expect(e.right).not.toBeNull();
        expect(e.left!.contentHash).toBe(e.right!.contentHash);
        break;
      case 'Modified':
      case 'MovedModified':
        expect(e.left).not.toBeNull();
        expect(e.right).not.toBeNull();
        break;
      case 'Split':
        expect(e.left).not.toBeNull();
        expect(e.right).toBeNull();
        expect(e.multiBlocks?.length ?? 0).toBeGreaterThanOrEqual(2);
        break;
      case 'Merge':
        expect(e.left).toBeNull();
        expect(e.right).not.toBeNull();
        expect(e.multiBlocks?.length ?? 0).toBeGreaterThanOrEqual(2);
        break;
    }
    if (e.kind !== 'Split' && e.kind !== 'Merge') expect(e.multiBlocks).toBeNull();
    if (e.left) leftSeen.push(e.left);
    if (e.right) rightSeen.push(e.right);
    if (e.multiBlocks) {
      if (e.kind === 'Split') rightSeen.push(...e.multiBlocks);
      if (e.kind === 'Merge') leftSeen.push(...e.multiBlocks);
    }
  }
  assertSameMultiset(left.body.blocks, leftSeen);
  assertSameMultiset(right.body.blocks, rightSeen);
}

function assertSameMultiset(expected: ReadonlyArray<IrBlock>, seen: ReadonlyArray<IrBlock>): void {
  expect(seen).toHaveLength(expected.length);
  const pool = [...expected];
  for (const b of seen) {
    const idx = pool.findIndex((x) => x === b);
    expect(idx).toBeGreaterThanOrEqual(0);
    pool.splice(idx, 1);
  }
  expect(pool).toEqual([]);
}

function formatEqual(left: IrBlock, right: IrBlock, settings: IrDiffSettings): boolean {
  if (settings.formatComparison === 'ModeledOnly' && left.kind === 'paragraph' && right.kind === 'paragraph') {
    return blockSignature(left, settings) === blockSignature(right, settings);
  }
  return left.formatFingerprint === right.formatFingerprint;
}

export function text(block: IrBlock): string {
  if (block.kind !== 'paragraph') return '';
  return block.inlines
    .filter((i) => i.kind === 'textRun')
    .map((i) => i.text)
    .join('');
}
