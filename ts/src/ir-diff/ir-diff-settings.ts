// Port of Docxodus/Ir/Diff/IrDiffSettings.cs

export type IrFormatComparison = 'ModeledOnly' | 'Full';

export type RevisionGranularity = 'Fine' | 'WmlComparerCompatible';

export interface IrDiffSettings {
  readonly formatComparison: IrFormatComparison;
  readonly revisionGranularity: RevisionGranularity;
  readonly renderMoves: boolean;
  readonly simplifyMoveMarkup: boolean;
  readonly wordSeparators: ReadonlySet<string>;
  readonly caseInsensitive: boolean;
  readonly conflateBreakingAndNonbreakingSpaces: boolean;
  /** BCP-47 locale used for case folding; null means invariant/locale-independent. */
  readonly culture: string | null;
  readonly blockSimilarityThreshold: number;
  readonly moveSimilarityThreshold: number;
  readonly moveMinimumTokenCount: number;
  readonly detectSplitMerge: boolean;
  readonly splitCoverageThreshold: number;
  readonly splitForeignSlack: number;
  readonly splitMaxRunLength: number;
  readonly compareHeadersFooters: boolean;
  readonly trackBlockFormatChanges: boolean;
  readonly trackParagraphFormatChanges: boolean;
  readonly authorForRevisions: string;
  readonly deterministic: boolean;
  readonly dateTimeForRevisions: string;
}

export const deterministicEpoch = '2000-01-01T00:00:00Z';

export const defaultWordSeparators: ReadonlySet<string> = new Set([
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
]);

export const defaultIrDiffSettings: IrDiffSettings = {
  formatComparison: 'ModeledOnly',
  revisionGranularity: 'Fine',
  renderMoves: true,
  simplifyMoveMarkup: false,
  wordSeparators: defaultWordSeparators,
  caseInsensitive: false,
  conflateBreakingAndNonbreakingSpaces: true,
  culture: null,
  blockSimilarityThreshold: 0.5,
  moveSimilarityThreshold: 0.8,
  moveMinimumTokenCount: 3,
  detectSplitMerge: true,
  splitCoverageThreshold: 0.9,
  splitForeignSlack: 0.34,
  splitMaxRunLength: 8,
  compareHeadersFooters: true,
  trackBlockFormatChanges: true,
  trackParagraphFormatChanges: true,
  authorForRevisions: 'Open-Xml-PowerTools',
  deterministic: true,
  dateTimeForRevisions: deterministicEpoch,
};

export type IrDiffSettingsOptions = Partial<
  Omit<IrDiffSettings, 'wordSeparators'> & {
    readonly wordSeparators: Iterable<string>;
  }
>;

export function normalizeIrDiffSettings(options: IrDiffSettingsOptions = {}): IrDiffSettings {
  return {
    ...defaultIrDiffSettings,
    ...options,
    wordSeparators:
      options.wordSeparators === undefined
        ? defaultIrDiffSettings.wordSeparators
        : new Set(options.wordSeparators),
  };
}

export function irDiffSettingsWithWallClockRevisionDate(
  options: IrDiffSettingsOptions = {},
): IrDiffSettings {
  return normalizeIrDiffSettings({
    ...options,
    deterministic: false,
    dateTimeForRevisions: new Date().toISOString(),
  });
}
