import { expect } from 'vitest';
import {
  normalizeIrDiffSettings,
  runFormatEqual,
  type IrDiffSettings,
  type IrDiffToken,
  type IrTokenDiff,
  type IrTokenOp,
} from '../../src/index.js';

export function assertTokenDiffInvariants(
  left: ReadonlyArray<IrDiffToken>,
  right: ReadonlyArray<IrDiffToken>,
  diff: IrTokenDiff,
  settings: Partial<IrDiffSettings> = {},
): void {
  const comparison = normalizeIrDiffSettings(settings).formatComparison;
  let leftCursor = 0;
  let rightCursor = 0;
  let prev: IrTokenOp | null = null;

  for (const op of diff.ops) {
    expect(op.leftStart).toBeGreaterThanOrEqual(0);
    expect(op.leftEnd).toBeGreaterThanOrEqual(op.leftStart);
    expect(op.rightStart).toBeGreaterThanOrEqual(0);
    expect(op.rightEnd).toBeGreaterThanOrEqual(op.rightStart);
    expect(op.leftStart).toBe(leftCursor);
    expect(op.rightStart).toBe(rightCursor);

    if (op.kind === 'Insert') {
      expect(op.leftEnd).toBe(op.leftStart);
      expect(op.rightEnd - op.rightStart).toBeGreaterThan(0);
    } else if (op.kind === 'Delete') {
      expect(op.rightEnd).toBe(op.rightStart);
      expect(op.leftEnd - op.leftStart).toBeGreaterThan(0);
    } else {
      expect(op.leftEnd - op.leftStart).toBeGreaterThan(0);
      expect(op.leftEnd - op.leftStart).toBe(op.rightEnd - op.rightStart);
      for (let k = 0; k < op.leftEnd - op.leftStart; k++) {
        const l = left[op.leftStart + k]!;
        const r = right[op.rightStart + k]!;
        expect(l.matchKey).toBe(r.matchKey);
        const fmtEqual = runFormatEqual(l.format, r.format, comparison);
        if (op.kind === 'Equal') expect(fmtEqual).toBe(true);
        else expect(fmtEqual).toBe(false);
      }
    }

    if (prev) expect(prev.kind === op.kind).toBe(false);
    leftCursor = op.leftEnd;
    rightCursor = op.rightEnd;
    prev = op;
  }

  expect(leftCursor).toBe(left.length);
  expect(rightCursor).toBe(right.length);
}
