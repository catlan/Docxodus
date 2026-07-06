import { describe, expect, test } from 'vitest';
import {
  type IrBreakKind,
  type IrDiffToken,
  type IrDiffTokenizerOptions,
  type IrInline,
  type IrParaFormat,
  type IrParagraph,
  type IrRunFormat,
  emptyIrProvenance,
  irAnchor,
  irHashCompute,
  tokenizeIrParagraph,
} from '../src/index.js';

const EMPTY_HASH = irHashCompute('');

const defaultRunFormat: IrRunFormat = {
  styleId: null,
  bold: null,
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
  unmodeledDigest: EMPTY_HASH,
};

const defaultParaFormat: IrParaFormat = {
  styleId: null,
  justification: null,
  indentLeftTwips: null,
  indentRightTwips: null,
  indentFirstLineTwips: null,
  spacingBeforeTwips: null,
  spacingAfterTwips: null,
  lineSpacing: null,
  outlineLevel: null,
  keepNext: null,
  keepLines: null,
  pageBreakBefore: null,
  numId: null,
  ilvl: null,
  unmodeledDigest: EMPTY_HASH,
};

const paragraph = (inlines: ReadonlyArray<IrInline>): IrParagraph => ({
  kind: 'paragraph',
  anchor: irAnchor('p', 'body', '00000000000000000000000000000000'),
  contentHash: EMPTY_HASH,
  formatFingerprint: EMPTY_HASH,
  source: emptyIrProvenance,
  format: defaultParaFormat,
  list: null,
  inlines,
  resolvedListMarker: null,
  inlineSectionBreakAnchor: null,
  inlineSectionFormat: null,
  pPrDigest: EMPTY_HASH,
  isListItemForLayout: false,
});

const textRun = (text: string, format = defaultRunFormat): IrInline => ({
  kind: 'textRun',
  text,
  format,
  fromInlineSdt: false,
});

const textPara = (text: string): IrParagraph => paragraph(text ? [textRun(text)] : []);
const para = (inlines: ReadonlyArray<IrInline>): IrParagraph => paragraph(inlines);
const tok = (p: IrParagraph, s?: IrDiffTokenizerOptions): ReadonlyArray<IrDiffToken> =>
  tokenizeIrParagraph(p, s);
const kinds = (tokens: ReadonlyArray<IrDiffToken>) => tokens.map((t) => t.kind);
const texts = (tokens: ReadonlyArray<IrDiffToken>) => tokens.map((t) => t.text);
const keys = (tokens: ReadonlyArray<IrDiffToken>) => tokens.map((t) => t.matchKey);

const tab = (): IrInline => ({ kind: 'tab', format: defaultRunFormat });
const br = (breakKind: IrBreakKind = 'Line'): IrInline => ({ kind: 'break', breakKind });
const footnote = (noteId = '1'): IrInline => ({
  kind: 'noteRef',
  noteKind: 'Footnote',
  noteId,
});
const hyperlink = (target: string, inlines: ReadonlyArray<IrInline>): IrInline => ({
  kind: 'hyperlink',
  target,
  internalTarget: null,
  inlines,
});
const fieldRun = (cachedResult: ReadonlyArray<IrInline>): IrInline => ({
  kind: 'fieldRun',
  instruction: ' PAGE ',
  cachedResult,
  isSimpleField: false,
});

const viRefDeo = () => para([textRun('Vi'), footnote(), textRun('deo')]);
const videoRefProvides = () => para([textRun('Video '), footnote(), textRun('provides')]);

describe('IrDiffTokenizer', () => {
  test('splitsWordsAndSeparatorsOneTokenPerSeparatorChar', () => {
    const tokens = tok(textPara('foo bar'));
    expect(tokens).toMatchObject([
      { kind: 'Word', text: 'foo' },
      { kind: 'Separator', text: ' ' },
      { kind: 'Word', text: 'bar' },
    ]);
  });

  test('multiSeparatorRunYieldsOneTokenPerChar', () => {
    const tokens = tok(textPara('a - b'));
    expect(kinds(tokens)).toEqual(['Word', 'Separator', 'Separator', 'Separator', 'Word']);
    expect(texts(tokens)).toEqual(['a', ' ', '-', ' ', 'b']);
  });

  test('leadingAndTrailingSeparatorsProduceSeparatorTokens', () => {
    expect(kinds(tok(textPara(' hi ')))).toEqual(['Separator', 'Word', 'Separator']);
  });

  test('emptyParagraphYieldsNoTokens', () => {
    expect(tok(para([]))).toEqual([]);
  });

  test('caseFoldOffKeepsDistinctKeys', () => {
    expect(tok(textPara('Foo'))[0]!.matchKey).not.toBe(tok(textPara('foo'))[0]!.matchKey);
  });

  test('caseFoldOnCollapsesKeysAndPreservesRawText', () => {
    const ci = { caseInsensitive: true };
    const upper = tok(textPara('Foo'), ci)[0]!;
    const lower = tok(textPara('foo'), ci)[0]!;
    expect(upper.matchKey).toBe(lower.matchKey);
    expect(upper.text).toBe('Foo');
  });

  test('caseFoldUsesSuppliedCulture', () => {
    const tr = { caseInsensitive: true, culture: 'tr-TR' };
    expect(tok(textPara('I'), tr)[0]!.matchKey).toBe('ı');
  });

  test('nbspConflationOnMatchesSpace', () => {
    const on = { conflateBreakingAndNonbreakingSpaces: true };
    const nbsp = tok(textPara('a\u00a0b'), on);
    const space = tok(textPara('a b'), on);
    expect(kinds(nbsp)).toEqual(['Word', 'Separator', 'Word']);
    expect(nbsp[1]!.text).toBe('\u00a0');
    expect(nbsp[1]!.matchKey).toBe(' ');
    expect(keys(nbsp)).toEqual(keys(space));
  });

  test('nbspConflationOffKeepsNbspDistinct', () => {
    const off = { conflateBreakingAndNonbreakingSpaces: false };
    expect(tok(textPara('a\u00a0b'), off)[0]!.matchKey).toBe('a\u00a0b');
    expect(tok(textPara('a b'), off)[0]!.matchKey).not.toBe(
      tok(textPara('a\u00a0b'), off)[0]!.matchKey,
    );
  });

  test('nbspConflationOnSplitsNbspAsASeparatorChar', () => {
    const on = { conflateBreakingAndNonbreakingSpaces: true };
    const tokens = tok(textPara('x\u00a0y'), on);
    expect(kinds(tokens)).toEqual(['Word', 'Separator', 'Word']);
    expect(texts(tokens)).toEqual(['x', '\u00a0', 'y']);
    expect(tokens[1]!.matchKey).toBe(' ');
  });

  test('nbspConflationOnYieldsIdenticalMatchkeySequenceToSpace', () => {
    const on = { conflateBreakingAndNonbreakingSpaces: true };
    expect(keys(tok(textPara("l'article\u00a01"), on))).toEqual(
      keys(tok(textPara("l'article 1"), on)),
    );
  });

  test('nbspConflationOffYieldsDistinctMatchkeySequenceFromSpace', () => {
    const off = { conflateBreakingAndNonbreakingSpaces: false };
    expect(keys(tok(textPara('a\u00a0b'), off))).not.toEqual(keys(tok(textPara('a b'), off)));
  });

  test('nonbreakingHyphenIsNotFoldedToSpace', () => {
    const on = { conflateBreakingAndNonbreakingSpaces: true };
    expect(tok(textPara('a\u2011b'), on)[0]!.matchKey).toBe('a\u2011b');
  });

  test('offsetsLineUpWithTextPositions', () => {
    const tokens = tok(textPara('foo bar'));
    expect([tokens[0]!.startChar, tokens[0]!.endChar]).toEqual([0, 3]);
    expect([tokens[1]!.startChar, tokens[1]!.endChar]).toEqual([3, 4]);
    expect([tokens[2]!.startChar, tokens[2]!.endChar]).toEqual([4, 7]);
  });

  test('zeroWidthAtomicsDoNotAdvanceOffset', () => {
    const tokens = tok(para([textRun('ab'), tab(), textRun('cd')]));
    const tabToken = tokens.find((t) => t.kind === 'Tab')!;
    expect([tokens[0]!.startChar, tokens[0]!.endChar]).toEqual([0, 2]);
    expect([tabToken.startChar, tabToken.endChar]).toEqual([2, 2]);
    expect([tokens[tokens.length - 1]!.startChar, tokens[tokens.length - 1]!.endChar]).toEqual([
      2, 4,
    ]);
  });

  test('tokenOffsetsMatchCommentTargetCoordinateSpace', () => {
    const tokens = tok(para([textRun('hello '), textRun('world')]));
    const worldTok = tokens.find((t) => t.text === 'world')!;
    const target = { startChar: 6, endChar: 11 };
    expect(worldTok.startChar).toBe(target.startChar);
    expect(worldTok.endChar).toBe(target.endChar);
  });

  test('tabAndBreakHaveAtomicKinds', () => {
    const tokenKinds = kinds(tok(para([tab(), br('Page')])));
    expect(tokenKinds).toContain('Tab');
    expect(tokenKinds).toContain('Break');
  });

  test('literalWordTabNeverMatchesTheTabToken', () => {
    const word = tok(textPara('tab')).find((t) => t.kind === 'Word')!;
    const tabToken = tok(para([tab()])).find((t) => t.kind === 'Tab')!;
    expect(word.matchKey).not.toBe(tabToken.matchKey);
  });

  test('breakKindsHaveDistinctKeys', () => {
    expect(tok(para([br('Page')]))[0]!.matchKey).not.toBe(tok(para([br('Line')]))[0]!.matchKey);
  });

  test('intraWordNoteRefBreaksWordEqualityWithContiguousForm', () => {
    const splitWords = tok(viRefDeo())
      .filter((t) => t.kind === 'Word')
      .map((t) => t.matchKey);
    const contiguous = tok(textPara('Video'));
    expect(splitWords).not.toContain(contiguous[0]!.matchKey);
    expect(splitWords).toHaveLength(2);
    expect(splitWords[0]).not.toBe(splitWords[1]);
  });

  test('betweenWordNoteRefLeavesWordKeysIdenticalToToday', () => {
    const withRefWordKeys = tok(videoRefProvides())
      .filter((t) => t.kind === 'Word')
      .map((t) => t.matchKey);
    const plainWordKeys = tok(textPara('Video provides'))
      .filter((t) => t.kind === 'Word')
      .map((t) => t.matchKey);
    expect(withRefWordKeys).toEqual(plainWordKeys);
  });

  test('intraWordInterruptionDoesNotChangeTheNoteRefKeyItself', () => {
    const split = tok(viRefDeo()).find((t) => t.kind === 'NoteRef')!;
    const between = tok(videoRefProvides()).find((t) => t.kind === 'NoteRef')!;
    expect(split.matchKey).toBe(between.matchKey);
  });

  test('intraWordMarkerIsDeterminedByTheInterruptingAtom', () => {
    const a = tok(viRefDeo())
      .filter((t) => t.kind === 'Word')
      .map((t) => t.matchKey);
    const b = tok(viRefDeo())
      .filter((t) => t.kind === 'Word')
      .map((t) => t.matchKey);
    expect(a).toEqual(b);
  });

  test('linkedTextDiffersFromPlainText', () => {
    const linked = para([hyperlink('https://a.example', [textRun('foo')])]);
    expect(tok(textPara('foo'))[0]!.matchKey).not.toBe(tok(linked)[0]!.matchKey);
  });

  test('sameTextDifferentTargetsDiffer', () => {
    const a = para([hyperlink('https://a.example', [textRun('foo')])]);
    const b = para([hyperlink('https://b.example', [textRun('foo')])]);
    expect(tok(a)[0]!.matchKey).not.toBe(tok(b)[0]!.matchKey);
  });

  test('linkedWordKeyIsSentinelFramedSoItCannotAliasALiteralWord', () => {
    const linked = para([hyperlink('https://a.example', [textRun('foo')])]);
    const linkedKey = tok(linked)[0]!.matchKey;
    expect(linkedKey).toContain('\u0001');
    expect(tok(textPara('foolnk'))[0]!.matchKey).not.toContain('\u0001');
  });

  test('fieldResultTokenizesTransparently', () => {
    const field = para([fieldRun([textRun('5')])]);
    expect(tok(textPara('5'))[0]!.matchKey).toBe(tok(field)[0]!.matchKey);
  });

  test('tokenCarriesGoverningRunFormat', () => {
    const boldFormat: IrRunFormat = { ...defaultRunFormat, bold: true };
    const token = tok(para([textRun('bold', boldFormat)])).find((t) => t.kind === 'Word')!;
    expect(token.format).not.toBeNull();
    expect(token.format!.bold).toBe(true);
  });

  test('atomicBreakTokenHasNullFormat', () => {
    const token = tok(para([br()])).find((t) => t.kind === 'Break')!;
    expect(token.format).toBeNull();
  });

  test('twoTokenizationsAreSequenceEqual', () => {
    const p = para([textRun('the quick-brown (fox)')]);
    expect(tok(p)).toEqual(tok(p));
  });
});
