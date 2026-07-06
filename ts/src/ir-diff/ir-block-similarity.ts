// Port of Docxodus/Ir/Diff/IrBlockSimilarity.cs

import type { IrBlock, IrParagraph, IrTable } from '../ir/ir-blocks.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import type { IrDiffSettings } from './ir-diff-settings.js';

interface MatchKeyBag {
  readonly counts: Map<string, number>;
  readonly total: number;
  readonly wordCount: number;
}

export class IrBlockSimilarity {
  private readonly bagCache = new WeakMap<IrParagraph, MatchKeyBag>();
  private readonly tableBagCache = new WeakMap<IrTable, MatchKeyBag>();

  constructor(private readonly settings: IrDiffSettings) {}

  score(left: IrBlock, right: IrBlock): number {
    if (left.kind === 'paragraph' && right.kind === 'paragraph') {
      return jaccard(this.bag(left), this.bag(right));
    }
    if (left.kind === 'table' && right.kind === 'table') {
      return jaccard(this.tableBag(left), this.tableBag(right));
    }
    return left.contentHash === right.contentHash ? 1 : 0;
  }

  wordCount(block: IrBlock): number {
    return block.kind === 'paragraph' ? this.bag(block).wordCount : 0;
  }

  private bag(paragraph: IrParagraph): MatchKeyBag {
    const cached = this.bagCache.get(paragraph);
    if (cached) return cached;
    const bag = buildBag(paragraph, this.settings);
    this.bagCache.set(paragraph, bag);
    return bag;
  }

  private tableBag(table: IrTable): MatchKeyBag {
    const cached = this.tableBagCache.get(table);
    if (cached) return cached;
    const counts = new Map<string, number>();
    let total = 0;
    let wordCount = 0;
    for (const row of table.rows) {
      for (const cell of row.cells) {
        for (const block of cell.blocks) {
          if (block.kind !== 'paragraph') continue;
          for (const token of tokenizeIrParagraph(block, this.settings)) {
            counts.set(token.matchKey, (counts.get(token.matchKey) ?? 0) + 1);
            total++;
            if (token.kind === 'Word') wordCount++;
          }
        }
      }
    }
    const bag = { counts, total, wordCount };
    this.tableBagCache.set(table, bag);
    return bag;
  }
}

function buildBag(paragraph: IrParagraph, settings: IrDiffSettings): MatchKeyBag {
  const counts = new Map<string, number>();
  let total = 0;
  let wordCount = 0;
  for (const token of tokenizeIrParagraph(paragraph, settings)) {
    counts.set(token.matchKey, (counts.get(token.matchKey) ?? 0) + 1);
    total++;
    if (token.kind === 'Word') wordCount++;
  }
  return { counts, total, wordCount };
}

function jaccard(a: MatchKeyBag, b: MatchKeyBag): number {
  if (a.total === 0 && b.total === 0) return 1;
  let intersection = 0;
  const [small, large] = a.counts.size <= b.counts.size ? [a, b] : [b, a];
  for (const [key, count] of small.counts) {
    const other = large.counts.get(key);
    if (other !== undefined) intersection += Math.min(count, other);
  }
  const union = a.total + b.total - intersection;
  return union === 0 ? 1 : intersection / union;
}
