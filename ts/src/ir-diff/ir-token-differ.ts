// Port of Docxodus/Ir/Diff/IrTokenDiffer.cs

import type { IrRunFormat } from '../ir/ir-formats.js';
import type { IrDiffToken } from './ir-diff-token.js';
import {
  normalizeIrDiffSettings,
  type IrDiffSettings,
  type IrDiffSettingsOptions,
  type IrFormatComparison,
} from './ir-diff-settings.js';
import { runKey } from './ir-modeled-format.js';
import type { IrTokenDiff, IrTokenOp, IrTokenOpKind } from './ir-token-diff.js';

interface Edit {
  readonly kind: IrTokenOpKind;
  readonly left: number;
  readonly right: number;
}

export function diffIrTokens(
  left: ReadonlyArray<IrDiffToken>,
  right: ReadonlyArray<IrDiffToken>,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrTokenDiff {
  const settings = normalizeIrDiffSettings(options);
  const spans = myersSpans(left, right);
  const ops: IrTokenOp[] = [];
  for (const span of spans) {
    if (span.kind === 'Equal') splitEqualByFormat(left, right, span, ops, settings.formatComparison);
    else ops.push(span);
  }
  return { ops };
}

export const IrTokenDiffer = {
  diff: diffIrTokens,
};

function myersSpans(left: ReadonlyArray<IrDiffToken>, right: ReadonlyArray<IrDiffToken>): IrTokenOp[] {
  const n = left.length;
  const m = right.length;
  if (n === 0 && m === 0) return [];
  if (n === 0) return [{ kind: 'Insert', leftStart: 0, leftEnd: 0, rightStart: 0, rightEnd: m }];
  if (m === 0) return [{ kind: 'Delete', leftStart: 0, leftEnd: n, rightStart: 0, rightEnd: 0 }];

  const max = n + m;
  const offset = max;
  const v = Array<number>(2 * max + 1).fill(0);
  const trace: number[][] = [];
  let reached = false;

  for (let d = 0; d <= max && !reached; d++) {
    trace.push([...v]);
    for (let k = -d; k <= d; k += 2) {
      let x: number;
      if (k === -d || (k !== d && v[offset + k - 1]! < v[offset + k + 1]!)) {
        x = v[offset + k + 1]!;
      } else {
        x = v[offset + k - 1]! + 1;
      }
      let y = x - k;
      while (x < n && y < m && left[x]!.matchKey === right[y]!.matchKey) {
        x++;
        y++;
      }
      v[offset + k] = x;
      if (x >= n && y >= m) {
        reached = true;
        break;
      }
    }
  }

  return backtrace(trace, offset, n, m);
}

function backtrace(trace: ReadonlyArray<ReadonlyArray<number>>, offset: number, n: number, m: number): IrTokenOp[] {
  const rev: Edit[] = [];
  let curX = n;
  let curY = m;

  for (let d = trace.length - 1; d > 0; d--) {
    const v = trace[d]!;
    const k = curX - curY;
    const prevK = k === -d || (k !== d && v[offset + k - 1]! < v[offset + k + 1]!) ? k + 1 : k - 1;
    const prevX = v[offset + prevK]!;
    const prevY = prevX - prevK;

    while (curX > prevX && curY > prevY) {
      curX--;
      curY--;
      rev.push({ kind: 'Equal', left: curX, right: curY });
    }

    if (curX === prevX) {
      curY--;
      rev.push({ kind: 'Insert', left: -1, right: curY });
    } else {
      curX--;
      rev.push({ kind: 'Delete', left: curX, right: -1 });
    }
  }

  while (curX > 0 && curY > 0) {
    curX--;
    curY--;
    rev.push({ kind: 'Equal', left: curX, right: curY });
  }

  rev.reverse();
  return coalesce(rev);
}

function coalesce(edits: ReadonlyArray<Edit>): IrTokenOp[] {
  const spans: IrTokenOp[] = [];
  let i = 0;
  let leftCursor = 0;
  let rightCursor = 0;
  while (i < edits.length) {
    const kind = edits[i]!.kind;
    let j = i;
    while (j < edits.length && edits[j]!.kind === kind) j++;
    const len = j - i;
    if (kind === 'Equal') {
      spans.push({ kind, leftStart: leftCursor, leftEnd: leftCursor + len, rightStart: rightCursor, rightEnd: rightCursor + len });
      leftCursor += len;
      rightCursor += len;
    } else if (kind === 'Delete') {
      spans.push({ kind, leftStart: leftCursor, leftEnd: leftCursor + len, rightStart: rightCursor, rightEnd: rightCursor });
      leftCursor += len;
    } else if (kind === 'Insert') {
      spans.push({ kind, leftStart: leftCursor, leftEnd: leftCursor, rightStart: rightCursor, rightEnd: rightCursor + len });
      rightCursor += len;
    }
    i = j;
  }
  return spans;
}

function splitEqualByFormat(
  left: ReadonlyArray<IrDiffToken>,
  right: ReadonlyArray<IrDiffToken>,
  span: IrTokenOp,
  ops: IrTokenOp[],
  comparison: IrFormatComparison,
): void {
  const len = span.leftEnd - span.leftStart;
  let i = 0;
  while (i < len) {
    const changed = formatDiffers(left[span.leftStart + i]!.format, right[span.rightStart + i]!.format, comparison);
    let j = i + 1;
    while (
      j < len &&
      formatDiffers(left[span.leftStart + j]!.format, right[span.rightStart + j]!.format, comparison) === changed
    ) {
      j++;
    }
    ops.push({
      kind: changed ? 'FormatChanged' : 'Equal',
      leftStart: span.leftStart + i,
      leftEnd: span.leftStart + j,
      rightStart: span.rightStart + i,
      rightEnd: span.rightStart + j,
    });
    i = j;
  }
}

export function runFormatEqual(
  a: IrRunFormat | null,
  b: IrRunFormat | null,
  comparison: IrFormatComparison,
): boolean {
  if (comparison === 'ModeledOnly') return runKey(a) === runKey(b);
  return fullRunKey(a) === fullRunKey(b);
}

function formatDiffers(a: IrRunFormat | null, b: IrRunFormat | null, comparison: IrFormatComparison): boolean {
  return !runFormatEqual(a, b, comparison);
}

function fullRunKey(f: IrRunFormat | null): string {
  if (f === null) return '';
  return JSON.stringify(f);
}
