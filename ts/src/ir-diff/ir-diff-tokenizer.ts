// Port of Docxodus/Ir/Diff/IrDiffTokenizer.cs

import { anchorToString } from '../ir/ir-anchor.js';
import type { IrParagraph } from '../ir/ir-blocks.js';
import type { IrBreakKind, IrRunFormat } from '../ir/ir-formats.js';
import type { IrInline, IrTextbox } from '../ir/ir-inlines.js';
import type { IrDiffToken } from './ir-diff-token.js';

const ATOMIC_SENTINEL = '\u0001';
const INTERRUPTION_MARKER_PREFIX = '\u0001iw:';

const DEFAULT_WORD_SEPARATORS = [
  ' ',
  '-',
  ')',
  '(',
  ';',
  ',',
  '（',
  '）',
  '，',
  '、',
  '；',
  '。',
  '：',
  '的',
];

/** Tokenizer settings subset read by IrDiffTokenizer. */
export interface IrDiffTokenizerSettings {
  /** Characters that split text into words vs one-token-per-char separators. */
  readonly wordSeparators: ReadonlySet<string>;
  /** Case-fold match keys per culture, or invariant when culture is null. */
  readonly caseInsensitive: boolean;
  /** When true, NBSP splits and folds like ordinary space; non-breaking hyphen stays distinct. */
  readonly conflateBreakingAndNonbreakingSpaces: boolean;
  /** BCP-47 locale used for case folding; null means invariant/locale-independent. */
  readonly culture: string | null;
}

export const defaultIrDiffTokenizerSettings: IrDiffTokenizerSettings = {
  wordSeparators: new Set(DEFAULT_WORD_SEPARATORS),
  caseInsensitive: false,
  conflateBreakingAndNonbreakingSpaces: true,
  culture: null,
};

export type IrDiffTokenizerOptions = Partial<{
  readonly wordSeparators: Iterable<string>;
  readonly caseInsensitive: boolean;
  readonly conflateBreakingAndNonbreakingSpaces: boolean;
  readonly culture: string | null;
}>;

const normalizeSettings = (
  options: IrDiffTokenizerOptions = {},
): IrDiffTokenizerSettings => ({
  wordSeparators:
    options.wordSeparators === undefined
      ? defaultIrDiffTokenizerSettings.wordSeparators
      : new Set(options.wordSeparators),
  caseInsensitive:
    options.caseInsensitive ?? defaultIrDiffTokenizerSettings.caseInsensitive,
  conflateBreakingAndNonbreakingSpaces:
    options.conflateBreakingAndNonbreakingSpaces ??
    defaultIrDiffTokenizerSettings.conflateBreakingAndNonbreakingSpaces,
  culture: options.culture ?? defaultIrDiffTokenizerSettings.culture,
});

/** Tokenize an IrParagraph into word/separator/atomic diff tokens. */
export function tokenizeIrParagraph(
  paragraph: IrParagraph,
  options: IrDiffTokenizerOptions = {},
): ReadonlyArray<IrDiffToken> {
  const settings = normalizeSettings(options);
  const state = { charOffset: 0 };
  const tokens: IrDiffToken[] = [];
  walkInlines(paragraph.inlines, settings, null, tokens, state);
  interruptionPostPass(tokens);
  return tokens;
}

/** Class-shaped export mirroring the C# static class name. */
export const IrDiffTokenizer = {
  tokenize: tokenizeIrParagraph,
};

function interruptionPostPass(tokens: IrDiffToken[]): void {
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i]!.kind !== 'Word') continue;
    const leftWord = i;

    let j = i + 1;
    let atomCount = 0;
    while (
      j < tokens.length &&
      isZeroWidthContentAtom(tokens[j]!.kind) &&
      tokens[j]!.startChar === tokens[j - 1]!.endChar
    ) {
      atomCount++;
      j++;
    }

    if (
      atomCount === 0 ||
      j >= tokens.length ||
      tokens[j]!.kind !== 'Word' ||
      tokens[j]!.startChar !== tokens[j - 1]!.endChar
    ) {
      continue;
    }

    const rightWord = j;
    let marker = ATOMIC_SENTINEL + INTERRUPTION_MARKER_PREFIX;
    for (let a = leftWord + 1; a < rightWord; a++) {
      marker += tokens[a]!.matchKey;
    }

    tokens[leftWord] = appendMarker(tokens[leftWord]!, marker);
    tokens[rightWord] = appendMarker(tokens[rightWord]!, marker);
    i = rightWord - 1;
  }
}

function isZeroWidthContentAtom(kind: IrDiffToken['kind']): boolean {
  return kind === 'NoteRef' || kind === 'Image' || kind === 'Opaque' || kind === 'Textbox';
}

const appendMarker = (token: IrDiffToken, marker: string): IrDiffToken => ({
  ...token,
  matchKey: token.matchKey + marker,
});

function walkInlines(
  inlines: ReadonlyArray<IrInline>,
  settings: IrDiffTokenizerSettings,
  linkSuffix: string | null,
  tokens: IrDiffToken[],
  state: { charOffset: number },
): void {
  for (const inline of inlines) {
    switch (inline.kind) {
      case 'textRun':
        emitTextRun(inline.text, inline.format, settings, linkSuffix, tokens, state);
        break;
      case 'fieldRun':
        walkInlines(inline.cachedResult, settings, linkSuffix, tokens, state);
        break;
      case 'hyperlink': {
        const target =
          inline.target ?? (inline.internalTarget ? anchorToString(inline.internalTarget) : '');
        const composed =
          linkSuffix === null ? linkSuffixFor(target) : linkSuffix + linkSuffixFor(target);
        walkInlines(inline.inlines, settings, composed, tokens, state);
        break;
      }
      case 'tab':
        tokens.push({
          kind: 'Tab',
          text: '',
          matchKey: atomicKey('tab'),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: inline.format,
        });
        break;
      case 'break':
        tokens.push({
          kind: 'Break',
          text: '',
          matchKey: atomicKey('brk:' + breakKindKey(inline.breakKind)),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: null,
        });
        break;
      case 'noteRef':
        tokens.push({
          kind: 'NoteRef',
          text: '',
          matchKey: atomicKey(inline.noteKind === 'Footnote' ? 'fn' : 'en'),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: null,
        });
        break;
      case 'inlineImage':
        tokens.push({
          kind: 'Image',
          text: '',
          matchKey: atomicKey('img:' + inline.imageBytesHash),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: null,
        });
        break;
      case 'opaqueInline':
        tokens.push({
          kind: 'Opaque',
          text: '',
          matchKey: atomicKey('opq:' + inline.canonicalHash),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: null,
        });
        break;
      case 'textbox':
        tokens.push({
          kind: 'Textbox',
          text: '',
          matchKey: atomicKey('tbx:' + textboxRollKey(inline)),
          startChar: state.charOffset,
          endChar: state.charOffset,
          format: null,
        });
        break;
    }
  }
}

function emitTextRun(
  text: string,
  format: IrRunFormat,
  settings: IrDiffTokenizerSettings,
  linkSuffix: string | null,
  tokens: IrDiffToken[],
  state: { charOffset: number },
): void {
  const nbspIsSeparator = settings.conflateBreakingAndNonbreakingSpaces;
  const isSeparator = (c: string) =>
    settings.wordSeparators.has(c) || (nbspIsSeparator && c === '\u00a0');

  let i = 0;
  while (i < text.length) {
    const c = text[i]!;
    if (isSeparator(c)) {
      const start = state.charOffset + i;
      tokens.push({
        kind: 'Separator',
        text: c,
        matchKey: applyLink(normalizeWord(c, settings), linkSuffix),
        startChar: start,
        endChar: start + 1,
        format,
      });
      i++;
    } else {
      const wordStart = i;
      while (i < text.length && !isSeparator(text[i]!)) i++;
      const raw = text.slice(wordStart, i);
      const start = state.charOffset + wordStart;
      tokens.push({
        kind: 'Word',
        text: raw,
        matchKey: applyLink(normalizeWord(raw, settings), linkSuffix),
        startChar: start,
        endChar: start + raw.length,
        format,
      });
    }
  }
  state.charOffset += text.length;
}

function normalizeWord(raw: string, settings: IrDiffTokenizerSettings): string {
  let s = raw;
  if (settings.conflateBreakingAndNonbreakingSpaces && s.includes('\u00a0')) {
    s = s.replaceAll('\u00a0', ' ');
  }
  if (settings.caseInsensitive) {
    s = settings.culture === null ? s.toLocaleLowerCase('und') : s.toLocaleLowerCase(settings.culture);
  }
  return s;
}

const applyLink = (key: string, linkSuffix: string | null) =>
  linkSuffix === null ? key : key + linkSuffix;

const linkSuffixFor = (target: string) => ATOMIC_SENTINEL + 'lnk:' + target;

const atomicKey = (body: string) => ATOMIC_SENTINEL + body;

const breakKindKey = (kind: IrBreakKind): string => kind;

function textboxRollKey(textbox: IrTextbox): string {
  let key = '';
  for (const block of textbox.blocks) {
    key += block.contentHash + '.';
  }
  return key;
}
