// Port of Docxodus/Ir/Diff/IrSplitSegmenter.cs

import type { IrParagraph } from '../ir/ir-blocks.js';
import type { IrDiffToken } from './ir-diff-token.js';
import {
  normalizeIrDiffSettings,
  type IrDiffSettings,
  type IrDiffSettingsOptions,
} from './ir-diff-settings.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import { IrTokenDiffer } from './ir-token-differ.js';
import type { IrTokenDiff } from './ir-token-diff.js';

export interface SplitScore {
  readonly coverage: number;
  readonly foreignSlack: number;
  readonly memberMatchedContent: ReadonlyArray<number>;
}

export function scoreSplitCandidate(
  singular: IrParagraph,
  run: ReadonlyArray<IrParagraph>,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): SplitScore {
  const settings = normalizeIrDiffSettings(options);
  const single = tokenizeIrParagraph(singular, settings);
  const memberTokens = run.map((p) => tokenizeIrParagraph(p, settings));
  const flat: IrDiffToken[] = [];
  const memberOfFlat: number[] = [];
  for (let m = 0; m < memberTokens.length; m++) {
    for (const t of memberTokens[m]!) {
      flat.push(t);
      memberOfFlat.push(m);
    }
  }

  const { partner, flatMatched } = lcsMatch(single, flat);
  const singleContent = countContent(single);
  let matchedContent = 0;
  const memberMatched = Array<number>(run.length).fill(0);
  for (let i = 0; i < single.length; i++) {
    const p = partner[i]!;
    if (p < 0 || !isContent(single[i]!)) continue;
    matchedContent++;
    const member = memberOfFlat[p]!;
    memberMatched[member] = memberMatched[member]! + 1;
  }

  const runContent = countContent(flat);
  let runMatchedContent = 0;
  for (let j = 0; j < flat.length; j++) {
    if (flatMatched[j] && isContent(flat[j]!)) runMatchedContent++;
  }

  return {
    coverage: singleContent === 0 ? 0 : matchedContent / singleContent,
    foreignSlack: runContent === 0 ? 0 : (runContent - runMatchedContent) / runContent,
    memberMatchedContent: memberMatched,
  };
}

export function computeSegmentDiffs(
  singular: IrParagraph,
  run: ReadonlyArray<IrParagraph>,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): ReadonlyArray<IrTokenDiff> {
  const settings = normalizeIrDiffSettings(options);
  const single = tokenizeIrParagraph(singular, settings);
  const memberTokens = run.map((p) => tokenizeIrParagraph(p, settings));
  const flat: IrDiffToken[] = [];
  const memberOfFlat: number[] = [];
  for (let m = 0; m < memberTokens.length; m++) {
    for (const t of memberTokens[m]!) {
      flat.push(t);
      memberOfFlat.push(m);
    }
  }

  const { partner } = lcsMatch(single, flat);
  const segmentOf = Array<number>(single.length).fill(0);
  let current = 0;
  for (let i = 0; i < single.length; i++) {
    if (partner[i]! >= 0) current = memberOfFlat[partner[i]!]!;
    segmentOf[i] = current;
  }
  for (let i = 1; i < single.length; i++) {
    if (segmentOf[i]! < segmentOf[i - 1]!) segmentOf[i] = segmentOf[i - 1]!;
  }

  const diffs: IrTokenDiff[] = [];
  let cursor = 0;
  for (let m = 0; m < run.length; m++) {
    const start = cursor;
    while (cursor < single.length && segmentOf[cursor] === m) cursor++;
    diffs.push(IrTokenDiffer.diff(single.slice(start, cursor), memberTokens[m]!, settings));
  }
  return diffs;
}

export function mirrorTokenDiff(diff: IrTokenDiff): IrTokenDiff {
  return {
    ops: diff.ops.map((o) => ({
      kind: o.kind === 'Insert' ? 'Delete' : o.kind === 'Delete' ? 'Insert' : o.kind,
      leftStart: o.rightStart,
      leftEnd: o.rightEnd,
      rightStart: o.leftStart,
      rightEnd: o.leftEnd,
    })),
  };
}

export const IrSplitSegmenter = {
  score: scoreSplitCandidate,
  computeSegmentDiffs,
  mirrorDiff: mirrorTokenDiff,
};

function isContent(t: IrDiffToken): boolean {
  return t.kind !== 'Separator' && t.kind !== 'Textbox';
}

function countContent(tokens: ReadonlyArray<IrDiffToken>): number {
  let n = 0;
  for (const t of tokens) if (isContent(t)) n++;
  return n;
}

function lcsMatch(a: ReadonlyArray<IrDiffToken>, b: ReadonlyArray<IrDiffToken>): { partner: number[]; flatMatched: boolean[] } {
  const n = a.length;
  const m = b.length;
  const dp = Array.from({ length: n + 1 }, () => Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i]![j] = a[i]!.matchKey === b[j]!.matchKey ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
    }
  }
  const partner = Array<number>(n).fill(-1);
  const flatMatched = Array<boolean>(m).fill(false);
  for (let i = 0, j = 0; i < n && j < m;) {
    if (a[i]!.matchKey === b[j]!.matchKey) {
      partner[i] = j;
      flatMatched[j] = true;
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) {
      i++;
    } else {
      j++;
    }
  }
  return { partner, flatMatched };
}
