// Diff-time modeled-only format projection used by block alignment.

import type { IrParagraph } from '../ir/ir-blocks.js';
import type {
  IrLineSpacing,
  IrRunFormat,
  IrSectionFormat,
  IrUnderline,
  IrParaFormat,
} from '../ir/ir-formats.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import type { IrDiffSettings } from './ir-diff-settings.js';

const append = (parts: string[], name: string, value: unknown): void => {
  if (value === null || value === undefined) return;
  parts.push(name, '=', String(value), ';');
};

const underlineKey = (u: IrUnderline | null): string | null =>
  u === null ? null : `${u.kind}:${u.colorHex ?? ''}`;

const lineSpacingKey = (l: IrLineSpacing | null): string | null =>
  l === null ? null : `${l.valueTwips}:${l.rule}`;

export function runKey(f: IrRunFormat | null): string {
  if (f === null) return '';
  const p: string[] = [];
  append(p, 'StyleId', f.styleId);
  append(p, 'Bold', f.bold);
  append(p, 'Italic', f.italic);
  append(p, 'Underline', underlineKey(f.underline));
  append(p, 'Strike', f.strike);
  append(p, 'DoubleStrike', f.doubleStrike);
  append(p, 'VertAlign', f.vertAlign);
  append(p, 'FontAscii', f.fontAscii);
  append(p, 'SizeHalfPoints', f.sizeHalfPoints);
  append(p, 'ColorHex', f.colorHex);
  append(p, 'Highlight', f.highlight);
  append(p, 'Caps', f.caps);
  append(p, 'SmallCaps', f.smallCaps);
  append(p, 'Vanish', f.vanish);
  return p.join('');
}

export function paraKey(f: IrParaFormat | null): string {
  if (f === null) return '';
  const p: string[] = [];
  append(p, 'PStyleId', f.styleId);
  append(p, 'Jc', f.justification);
  append(p, 'IndL', f.indentLeftTwips);
  append(p, 'IndR', f.indentRightTwips);
  append(p, 'IndFL', f.indentFirstLineTwips);
  append(p, 'SpB', f.spacingBeforeTwips);
  append(p, 'SpA', f.spacingAfterTwips);
  append(p, 'Line', lineSpacingKey(f.lineSpacing));
  append(p, 'Outline', f.outlineLevel);
  append(p, 'KeepNext', f.keepNext);
  append(p, 'KeepLines', f.keepLines);
  append(p, 'PBB', f.pageBreakBefore);
  append(p, 'NumId', f.numId);
  append(p, 'Ilvl', f.ilvl);
  return p.join('');
}

export function sectionKey(f: IrSectionFormat | null): string {
  if (f === null) return '';
  const p: string[] = [];
  append(p, 'PgW', f.pageWidthTwips);
  append(p, 'PgH', f.pageHeightTwips);
  append(p, 'Land', f.landscape);
  append(p, 'MT', f.marginTopTwips);
  append(p, 'MB', f.marginBottomTwips);
  append(p, 'ML', f.marginLeftTwips);
  append(p, 'MR', f.marginRightTwips);
  append(p, 'SType', f.sectionType);
  return p.join('');
}

export function blockSignature(paragraph: IrParagraph, settings: IrDiffSettings): string {
  const parts: string[] = [];
  for (const t of tokenizeIrParagraph(paragraph, settings)) {
    parts.push(t.matchKey, '\u241f', runKey(t.format), '\u241e');
  }
  if (settings.trackParagraphFormatChanges) parts.push('P:', paraKey(paragraph.format));
  if (settings.trackBlockFormatChanges) parts.push('S:', sectionKey(paragraph.inlineSectionFormat));
  return parts.join('');
}
