import { zipSync } from 'fflate';
import {
  buildIrEditScript,
} from './ir-diff/ir-edit-script-builder.js';
import { writeIrEditScriptJson } from './ir-diff/ir-edit-script-json.js';
import { normalizeIrDiffSettings, type IrDiffSettingsOptions } from './ir-diff/ir-diff-settings.js';
import { renderIrMarkup } from './ir-diff/ir-markup-renderer.js';
import { renderIrRevisions } from './ir-diff/ir-revision-renderer.js';
import { revisionToWire, type IrRevisionWire } from './ir-diff/ir-revision.js';
import { readIrDocument } from './ir/ir-reader.js';

export interface DocxDiffSettingsTs {
  readonly authorForRevisions?: string;
  readonly deterministic?: boolean;
  readonly dateTimeForRevisions?: string | null;
  readonly caseInsensitive?: boolean;
  readonly culture?: string | null;
  readonly conflateBreakingAndNonbreakingSpaces?: boolean;
  readonly wordSeparators?: Iterable<string> | string;
  readonly detectMoves?: boolean;
  readonly moveSimilarityThreshold?: number;
  readonly moveMinimumWordCount?: number;
  readonly revisionGranularity?: 'Fine' | 'WmlComparerCompatible';
  readonly formatComparison?: 'ModeledOnly' | 'Full';
  readonly compareHeadersFooters?: boolean;
  readonly trackBlockFormatChanges?: boolean;
}

export function docxDiffCompareTs(
  left: Uint8Array,
  right: Uint8Array,
  settings: DocxDiffSettingsTs = {},
): Uint8Array {
  const diff = mapDocxDiffSettings(settings);
  const leftIr = readIrDocument(left, { retainSources: false });
  const rightIr = readIrDocument(right, { retainSources: false });
  const script = buildIrEditScript(leftIr, rightIr, diff);
  const parts = renderIrMarkup(left, right, script, diff);
  return zipSync(Object.fromEntries(parts), { mtime: new Date('1980-01-01T00:00:00Z') });
}

export function docxDiffGetRevisionsTs(
  left: Uint8Array,
  right: Uint8Array,
  settings: DocxDiffSettingsTs = {},
): IrRevisionWire[] {
  const diff = mapDocxDiffSettings(settings);
  const leftIr = readIrDocument(left, { retainSources: false });
  const rightIr = readIrDocument(right, { retainSources: false });
  const script = buildIrEditScript(leftIr, rightIr, diff);
  return renderIrRevisions(script, leftIr, rightIr, diff).map(revisionToWire);
}

export function docxDiffGetEditScriptJsonTs(
  left: Uint8Array,
  right: Uint8Array,
  settings: DocxDiffSettingsTs = {},
): string {
  const diff = mapDocxDiffSettings(settings);
  const leftIr = readIrDocument(left, { retainSources: false });
  const rightIr = readIrDocument(right, { retainSources: false });
  return writeIrEditScriptJson(buildIrEditScript(leftIr, rightIr, diff));
}

export function mapDocxDiffSettings(settings: DocxDiffSettingsTs = {}): IrDiffSettingsOptions {
  const explicitDate = settings.dateTimeForRevisions ?? undefined;
  const deterministic = settings.deterministic ?? true;
  const wordSeparators = typeof settings.wordSeparators === 'string'
    ? [...settings.wordSeparators]
    : settings.wordSeparators;
  const options: Record<string, unknown> = {
    deterministic,
  };
  const dateTimeForRevisions = explicitDate ?? (deterministic ? undefined : new Date().toISOString());
  if (settings.authorForRevisions !== undefined) options.authorForRevisions = settings.authorForRevisions;
  if (dateTimeForRevisions !== undefined) options.dateTimeForRevisions = dateTimeForRevisions;
  if (settings.caseInsensitive !== undefined) options.caseInsensitive = settings.caseInsensitive;
  if (settings.culture !== undefined) options.culture = settings.culture;
  if (settings.conflateBreakingAndNonbreakingSpaces !== undefined) options.conflateBreakingAndNonbreakingSpaces = settings.conflateBreakingAndNonbreakingSpaces;
  if (wordSeparators !== undefined) options.wordSeparators = wordSeparators;
  if (settings.detectMoves !== undefined) options.renderMoves = settings.detectMoves;
  if (settings.moveSimilarityThreshold !== undefined) options.moveSimilarityThreshold = settings.moveSimilarityThreshold;
  if (settings.moveMinimumWordCount !== undefined) options.moveMinimumTokenCount = settings.moveMinimumWordCount;
  if (settings.revisionGranularity !== undefined) options.revisionGranularity = settings.revisionGranularity;
  if (settings.formatComparison !== undefined) options.formatComparison = settings.formatComparison;
  if (settings.compareHeadersFooters !== undefined) options.compareHeadersFooters = settings.compareHeadersFooters;
  if (settings.trackBlockFormatChanges !== undefined) {
    options.trackBlockFormatChanges = settings.trackBlockFormatChanges;
    options.trackParagraphFormatChanges = settings.trackBlockFormatChanges;
  }
  return normalizeIrDiffSettings(options as IrDiffSettingsOptions);
}
