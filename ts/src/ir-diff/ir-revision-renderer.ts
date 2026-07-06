import { anchorToString } from '../ir/ir-anchor.js';
import type { IrBlock, IrCell, IrParagraph, IrRow, IrSectionBreak, IrTable } from '../ir/ir-blocks.js';
import type { IrParaFormat, IrRunFormat, IrSectionFormat } from '../ir/ir-formats.js';
import type { IrDocument, IrScope } from '../ir/ir-reader.js';
import { tokenizeIrParagraph } from './ir-diff-tokenizer.js';
import { defaultWordSeparators, type IrDiffSettings, type IrDiffSettingsOptions, normalizeIrDiffSettings } from './ir-diff-settings.js';
import type { IrDiffToken } from './ir-diff-token.js';
import type { IrCellOp, IrEditOp, IrEditScript, IrRowOp, IrTableDiff } from './ir-edit-script.js';
import { buildIrEditScript } from './ir-edit-script-builder.js';
import { paraKey, runKey, sectionKey } from './ir-modeled-format.js';
import type { IrTokenDiff, IrTokenOp } from './ir-token-diff.js';
import type { IrFormatChangeDetails, IrFormatChangeScope, IrRevision } from './ir-revision.js';

interface Context {
  readonly left: IrDocument;
  readonly right: IrDocument;
  readonly settings: IrDiffSettings;
  readonly moveSourceAnchor: ReadonlyMap<number, string>;
}

export function renderIrRevisions(
  script: IrEditScript,
  left: IrDocument,
  right: IrDocument,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrRevision[] {
  const settings = normalizeIrDiffSettings(options);
  const moveSourceAnchor = new Map<number, string>();
  for (const op of script.ops) {
    if (op.isMoveSource === true && op.moveGroupId !== null && op.leftAnchor !== null) {
      moveSourceAnchor.set(op.moveGroupId, op.leftAnchor);
    }
  }
  const ctx: Context = { left, right, settings, moveSourceAnchor };
  const revisions: IrRevision[] = [];
  renderBlockOpList(script.ops, ctx, revisions);

  if (script.noteOps) {
    for (const noteDiff of script.noteOps) {
      for (const op of noteDiff.ops) renderBlockOp(op, ctx, revisions);
    }
  }

  if (settings.revisionGranularity === 'Fine' && script.headerFooterOps) {
    for (const hfDiff of script.headerFooterOps) {
      for (const op of hfDiff.ops) renderBlockOp(op, ctx, revisions);
    }
  }

  if (
    settings.trackBlockFormatChanges &&
    left.body.blocks.length > 0 &&
    right.body.blocks.length > 0 &&
    left.body.blocks[left.body.blocks.length - 1]?.kind === 'sectionBreak' &&
    right.body.blocks[right.body.blocks.length - 1]?.kind === 'sectionBreak'
  ) {
    const lsec = left.body.blocks[left.body.blocks.length - 1] as IrSectionBreak;
    const rsec = right.body.blocks[right.body.blocks.length - 1] as IrSectionBreak;
    const details = sectionFormatChangeDetails(lsec.format, rsec.format);
    if (details.changedPropertyNames.length > 0) {
      revisions.push(formatRevision('', details, ctx, anchorToString(lsec.anchor), anchorToString(rsec.anchor)));
    }
  }

  if (settings.revisionGranularity === 'WmlComparerCompatible') {
    for (let i = revisions.length - 1; i >= 0; i--) {
      const r = revisions[i]!;
      if (isSectionBreakZeroWidth(r) || (r.formatChange && r.formatChange.scope !== 'Run')) revisions.splice(i, 1);
    }
  }

  return revisions;
}

export function renderIrRevisionsFromDocuments(
  left: IrDocument,
  right: IrDocument,
  options: IrDiffSettingsOptions | IrDiffSettings = {},
): IrRevision[] {
  const settings = normalizeIrDiffSettings(options);
  return renderIrRevisions(buildIrEditScript(left, right, settings), left, right, settings);
}

function author(ctx: Context): string { return ctx.settings.authorForRevisions; }
function date(ctx: Context): string { return ctx.settings.dateTimeForRevisions; }

function isSectionBreakZeroWidth(r: IrRevision): boolean {
  return (r.type === 'Inserted' || r.type === 'Deleted') && r.text.length === 0 &&
    (r.rightAnchor?.startsWith('sec:') === true || r.leftAnchor?.startsWith('sec:') === true);
}

function renderBlockOpList(ops: ReadonlyArray<IrEditOp>, ctx: Context, sink: IrRevision[]): void {
  if (ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
    for (const op of ops) renderBlockOp(op, ctx, sink);
    return;
  }
  let i = 0;
  while (i < ops.length) {
    const kind = ops[i]!.kind;
    if (kind === 'InsertBlock' || kind === 'DeleteBlock') {
      let end = i + 1;
      while (end < ops.length && ops[end]!.kind === kind) end++;
      renderInsDelRun(ops, i, end, kind === 'InsertBlock', ctx, sink);
      i = end;
    } else {
      renderBlockOp(ops[i]!, ctx, sink);
      i++;
    }
  }
}

function renderInsDelRun(
  ops: ReadonlyArray<IrEditOp>,
  start: number,
  end: number,
  insert: boolean,
  ctx: Context,
  sink: IrRevision[],
): void {
  let subStart = start;
  for (let k = start; k < end; k++) {
    const anchor = insert ? ops[k]!.rightAnchor : ops[k]!.leftAnchor;
    const doc = insert ? ctx.right : ctx.left;
    const block = anchor === null ? undefined : doc.anchorIndex.get(anchor);
    if (block && block.kind !== 'paragraph' && k > subStart) {
      flushInsDelSubRegion(ops, subStart, k, insert, ctx, sink);
      subStart = k;
    }
  }
  flushInsDelSubRegion(ops, subStart, end, insert, ctx, sink);
}

function flushInsDelSubRegion(
  ops: ReadonlyArray<IrEditOp>,
  start: number,
  end: number,
  insert: boolean,
  ctx: Context,
  sink: IrRevision[],
): void {
  if (end - start <= 1 || !subRegionHasText(ops, start, end, insert, ctx)) {
    for (let k = start; k < end; k++) renderBlockOp(ops[k]!, ctx, sink);
    return;
  }
  let text = '';
  let firstAnchor: string | null = null;
  for (let k = start; k < end; k++) {
    const anchor = insert ? ops[k]!.rightAnchor : ops[k]!.leftAnchor;
    firstAnchor ??= anchor;
    text += blockText(anchor, insert ? ctx.right : ctx.left, ctx.settings) + '\n';
  }
  sink.push(insert
    ? { type: 'Inserted', text, author: author(ctx), date: date(ctx), rightAnchor: firstAnchor }
    : { type: 'Deleted', text, author: author(ctx), date: date(ctx), leftAnchor: firstAnchor });
}

function subRegionHasText(ops: ReadonlyArray<IrEditOp>, start: number, end: number, insert: boolean, ctx: Context): boolean {
  const doc = insert ? ctx.right : ctx.left;
  for (let k = start; k < end; k++) {
    const anchor = insert ? ops[k]!.rightAnchor : ops[k]!.leftAnchor;
    const block = anchor === null ? undefined : doc.anchorIndex.get(anchor);
    if (block?.kind === 'paragraph' && tokenizeIrParagraph(block, ctx.settings).some((t) => t.kind === 'Word')) return true;
  }
  return false;
}

function renderBlockOp(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  switch (op.kind) {
    case 'EqualBlock':
      return;
    case 'InsertBlock':
      if (!isZeroWidthBlock(op.rightAnchor, ctx.right, ctx.settings)) {
        sink.push({ type: 'Inserted', text: blockText(op.rightAnchor, ctx.right, ctx.settings), author: author(ctx), date: date(ctx), rightAnchor: op.rightAnchor });
      }
      return;
    case 'DeleteBlock':
      if (!isZeroWidthBlock(op.leftAnchor, ctx.left, ctx.settings)) {
        sink.push({ type: 'Deleted', text: blockText(op.leftAnchor, ctx.left, ctx.settings), author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor });
      }
      return;
    case 'FormatOnlyBlock':
      renderFormatOnlyBlock(op, ctx, sink);
      return;
    case 'ModifyBlock':
      renderModifyBlock(op, ctx, sink);
      return;
    case 'MoveBlock':
    case 'MoveModifyBlock':
      renderMoveOp(op, ctx, sink);
      return;
    case 'SplitBlock':
      renderSplitBlock(op, ctx, sink);
      return;
    case 'MergeBlock':
      renderMergeBlock(op, ctx, sink);
      return;
  }
}

function renderSplitBlock(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  if (!op.splitMergeAnchors || !op.segmentDiffs) return;
  const leftTokens = paragraphTokens(op.leftAnchor, ctx.left, ctx.settings);
  let offset = 0;
  for (let s = 0; s < op.splitMergeAnchors.length; s++) {
    const diff = op.segmentDiffs[s] as IrTokenDiff;
    const sliceLen = segmentLeftLength(diff);
    const slice = leftTokens.slice(offset, offset + sliceLen);
    offset += sliceLen;
    if (s === 0 || ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
      renderSegmentTokenOps(diff, slice, paragraphTokens(op.splitMergeAnchors[s]!, ctx.right, ctx.settings), op.leftAnchor, op.splitMergeAnchors[s]!, ctx, sink);
    }
    if (s > 0 && ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
      sink.push({ type: 'Inserted', text: '\n', author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor, rightAnchor: op.splitMergeAnchors[s]! });
    }
  }
  if (ctx.settings.revisionGranularity === 'WmlComparerCompatible' && op.splitMergeAnchors.length > 1) {
    let text = '\n';
    for (let s = 1; s < op.splitMergeAnchors.length; s++) text += blockText(op.splitMergeAnchors[s]!, ctx.right, ctx.settings) + '\n';
    sink.push({ type: 'Inserted', text, author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor, rightAnchor: op.splitMergeAnchors[1]! });
  }
}

function renderMergeBlock(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  if (!op.splitMergeAnchors || !op.segmentDiffs) return;
  const rightTokens = paragraphTokens(op.rightAnchor, ctx.right, ctx.settings);
  let offset = 0;
  for (let m = 0; m < op.splitMergeAnchors.length; m++) {
    const diff = op.segmentDiffs[m] as IrTokenDiff;
    const sliceLen = segmentRightLength(diff);
    const slice = rightTokens.slice(offset, offset + sliceLen);
    offset += sliceLen;
    if (m === 0 || ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
      renderSegmentTokenOps(diff, paragraphTokens(op.splitMergeAnchors[m]!, ctx.left, ctx.settings), slice, op.splitMergeAnchors[m]!, op.rightAnchor, ctx, sink);
    }
    if (m > 0 && ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
      sink.push({ type: 'Deleted', text: '\n', author: author(ctx), date: date(ctx), leftAnchor: op.splitMergeAnchors[m]!, rightAnchor: op.rightAnchor });
    }
  }
  if (ctx.settings.revisionGranularity === 'WmlComparerCompatible' && op.splitMergeAnchors.length > 1) {
    let text = '\n';
    for (let m = 1; m < op.splitMergeAnchors.length; m++) text += blockText(op.splitMergeAnchors[m]!, ctx.left, ctx.settings) + '\n';
    sink.push({ type: 'Deleted', text, author: author(ctx), date: date(ctx), leftAnchor: op.splitMergeAnchors[1]!, rightAnchor: op.rightAnchor });
  }
}

function renderSegmentTokenOps(diff: IrTokenDiff, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, leftAnchor: string | null, rightAnchor: string | null, ctx: Context, sink: IrRevision[]): void {
  if (ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
    renderTokenOps(diff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
    return;
  }
  const buffer: IrRevision[] = [];
  renderTokenOps(diff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, buffer);
  sink.push(...buffer.filter((r) => !((r.type === 'Inserted' || r.type === 'Deleted') && r.text.length > 0 && r.text.trim().length === 0)));
}

function segmentLeftLength(diff: IrTokenDiff): number {
  return diff.ops.reduce((n, o) => n + (o.kind === 'Insert' ? 0 : o.leftEnd - o.leftStart), 0);
}

function segmentRightLength(diff: IrTokenDiff): number {
  return diff.ops.reduce((n, o) => n + (o.kind === 'Delete' ? 0 : o.rightEnd - o.rightStart), 0);
}

function isZeroWidthBlock(anchor: string | null, doc: IrDocument, settings: IrDiffSettings): boolean {
  if (settings.revisionGranularity !== 'WmlComparerCompatible' || anchor === null) return false;
  if (anchor.startsWith('p:fn:') || anchor.startsWith('p:en:')) return false;
  const block = doc.anchorIndex.get(anchor);
  return block?.kind === 'paragraph' && countContent(tokenizeIrParagraph(block, settings)) === 0;
}

function renderModifyBlock(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  if (op.tableDiff) {
    const tableDiff = op.tableDiff as IrTableDiff;
    if (tableDiffNeedsWholeTableFallback(tableDiff)) {
      if (op.leftAnchor) sink.push({ type: 'Deleted', text: blockText(op.leftAnchor, ctx.left, ctx.settings), author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor });
      if (op.rightAnchor) sink.push({ type: 'Inserted', text: blockText(op.rightAnchor, ctx.right, ctx.settings), author: author(ctx), date: date(ctx), rightAnchor: op.rightAnchor });
      return;
    }
    renderTableDiff(tableDiff, ctx, sink);
    const lt = resolveTable(op.leftAnchor, ctx.left);
    const rt = resolveTable(op.rightAnchor, ctx.right);
    if (ctx.settings.trackBlockFormatChanges && lt && rt) emitTableModifiedShellRevisions(lt, rt, tableDiff, ctx, sink);
    return;
  }

  if (op.tokenDiff) {
    emitParagraphScopeFormatChanged(op, ctx, sink);
    emitInlineSectionFormatChanged(op, ctx, sink);
    renderTokenOps(op.tokenDiff as IrTokenDiff, paragraphTokens(op.leftAnchor, ctx.left, ctx.settings), paragraphTokens(op.rightAnchor, ctx.right, ctx.settings), op.leftAnchor, op.rightAnchor, ctx, sink);
  }

  if (op.textboxDiffs) renderTextboxDiffs(op.textboxDiffs, ctx, sink);
}

function renderTextboxDiffs(textboxDiffs: NonNullable<IrEditOp['textboxDiffs']>, ctx: Context, sink: IrRevision[]): void {
  if (ctx.settings.revisionGranularity !== 'WmlComparerCompatible') {
    for (const diff of textboxDiffs) for (const op of diff.ops) renderBlockOp(op, ctx, sink);
    return;
  }
  const seen = new Map<string, number>();
  for (const diff of textboxDiffs) {
    const batch: IrRevision[] = [];
    for (const op of diff.ops) renderTextboxInnerOp(op, ctx, batch);
    const sig = batch.map((r) => `${typeOrdinal(r.type)}:${r.text}\u001f`).join('');
    const n = (seen.get(sig) ?? 0) + 1;
    seen.set(sig, n);
    if (n % 2 === 1) sink.push(...batch);
  }
}

function typeOrdinal(type: IrRevision['type']): number {
  return type === 'Inserted' ? 0 : type === 'Deleted' ? 1 : type === 'Moved' ? 2 : 3;
}

function renderTextboxInnerOp(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  const lb = op.leftAnchor ? ctx.left.anchorIndex.get(op.leftAnchor) : undefined;
  const rb = op.rightAnchor ? ctx.right.anchorIndex.get(op.rightAnchor) : undefined;
  if (op.kind === 'ModifyBlock' && !op.tableDiff && lb?.kind === 'paragraph' && rb?.kind === 'paragraph') {
    const delText = blockText(op.leftAnchor, ctx.left, ctx.settings);
    const insText = blockText(op.rightAnchor, ctx.right, ctx.settings);
    if (delText.length > 0) sink.push({ type: 'Deleted', text: delText, author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor, rightAnchor: op.rightAnchor });
    if (insText.length > 0) sink.push({ type: 'Inserted', text: insText, author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor, rightAnchor: op.rightAnchor });
    return;
  }
  renderBlockOp(op, ctx, sink);
}

function belowMoveMinimum(text: string, settings: IrDiffSettings): boolean {
  if (settings.revisionGranularity !== 'WmlComparerCompatible') return false;
  let words = 0;
  let inWord = false;
  for (const c of text) {
    if (/\s/u.test(c)) inWord = false;
    else if (!inWord) { inWord = true; words++; }
  }
  return words < settings.moveMinimumTokenCount;
}

function renderMoveOp(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  const isSource = op.isMoveSource === true;
  const text = isSource ? blockText(op.leftAnchor, ctx.left, ctx.settings) : blockText(op.rightAnchor, ctx.right, ctx.settings);
  if (!ctx.settings.renderMoves || belowMoveMinimum(text, ctx.settings)) {
    sink.push(isSource
      ? { type: 'Deleted', text, author: author(ctx), date: date(ctx), leftAnchor: op.leftAnchor }
      : { type: 'Inserted', text, author: author(ctx), date: date(ctx), rightAnchor: op.rightAnchor });
    return;
  }
  sink.push({
    type: 'Moved',
    text,
    author: author(ctx),
    date: date(ctx),
    moveGroupId: op.moveGroupId,
    isMoveSource: isSource,
    leftAnchor: isSource ? op.leftAnchor : null,
    rightAnchor: isSource ? null : op.rightAnchor,
  });
  if (!isSource && op.kind === 'MoveModifyBlock' && op.tokenDiff) {
    const sourceAnchor = op.moveGroupId !== null ? ctx.moveSourceAnchor.get(op.moveGroupId) ?? null : null;
    renderTokenOps(op.tokenDiff as IrTokenDiff, paragraphTokens(sourceAnchor, ctx.left, ctx.settings), paragraphTokens(op.rightAnchor, ctx.right, ctx.settings), sourceAnchor, op.rightAnchor, ctx, sink);
  }
}

function renderTokenOps(tokenDiff: IrTokenDiff, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, leftAnchor: string | null, rightAnchor: string | null, ctx: Context, sink: IrRevision[]): void {
  if (ctx.settings.revisionGranularity === 'WmlComparerCompatible') {
    renderTokenOpsCompatible(tokenDiff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
    return;
  }
  for (const op of tokenDiff.ops) {
    if (op.kind === 'Insert') {
      if (!isMaskedTextboxSpan(rightTokens, op.rightStart, op.rightEnd)) {
        sink.push({ type: 'Inserted', text: rawText(rightTokens, op.rightStart, op.rightEnd), author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
      }
    } else if (op.kind === 'Delete') {
      if (!isMaskedTextboxSpan(leftTokens, op.leftStart, op.leftEnd)) {
        sink.push({ type: 'Deleted', text: rawText(leftTokens, op.leftStart, op.leftEnd), author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
      }
    } else if (op.kind === 'FormatChanged') {
      renderFormatChangedSpan(op, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
    }
  }
}

function renderTokenOpsCompatible(tokenDiff: IrTokenDiff, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, leftAnchor: string | null, rightAnchor: string | null, ctx: Context, sink: IrRevision[]): void {
  const coarsen = isLowEqualCoverage(tokenDiff, leftTokens, rightTokens);
  const region = new Region();
  for (const op of tokenDiff.ops) {
    if (op.kind === 'Equal') {
      if (coarsen && region.open) region.holdSeparator(rawText(leftTokens, op.leftStart, op.leftEnd), rawText(rightTokens, op.rightStart, op.rightEnd));
      else if (isPureSeparatorSpan(leftTokens, op.leftStart, op.leftEnd)) {
        if (region.open) region.holdSeparator(rawText(leftTokens, op.leftStart, op.leftEnd), rawText(rightTokens, op.rightStart, op.rightEnd));
      } else region.flush(leftAnchor, rightAnchor, ctx, sink);
    } else if (op.kind === 'Insert') {
      if (!isMaskedTextboxSpan(rightTokens, op.rightStart, op.rightEnd)) region.addInsert(rawText(rightTokens, op.rightStart, op.rightEnd));
    } else if (op.kind === 'Delete') {
      if (!isMaskedTextboxSpan(leftTokens, op.leftStart, op.leftEnd)) region.addDelete(rawText(leftTokens, op.leftStart, op.leftEnd));
    } else {
      region.flush(leftAnchor, rightAnchor, ctx, sink);
      renderFormatChangedSpan(op, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
    }
  }
  region.flush(leftAnchor, rightAnchor, ctx, sink);
}

const LOW_COVERAGE_FLOOR = 0.67;
const MIN_COARSEN_CONTENT = 8;

function isLowEqualCoverage(tokenDiff: IrTokenDiff, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>): boolean {
  const leftContent = countContent(leftTokens);
  const rightContent = countContent(rightTokens);
  if (leftContent === 0 || rightContent === 0 || Math.max(leftContent, rightContent) < MIN_COARSEN_CONTENT) return false;
  let coveredLeft = 0;
  let coveredRight = 0;
  for (const op of tokenDiff.ops) {
    if (op.kind === 'Equal' || op.kind === 'FormatChanged') {
      coveredLeft += countContent(leftTokens, op.leftStart, op.leftEnd);
      coveredRight += countContent(rightTokens, op.rightStart, op.rightEnd);
    }
  }
  return Math.max(coveredLeft / leftContent, coveredRight / rightContent) < LOW_COVERAGE_FLOOR;
}

function countContent(tokens: ReadonlyArray<IrDiffToken>, start = 0, end = tokens.length): number {
  let n = 0;
  for (let i = start; i < end && i < tokens.length; i++) if (tokens[i]!.kind !== 'Separator' && tokens[i]!.kind !== 'Textbox') n++;
  return n;
}

class Region {
  private del = '';
  private ins = '';
  private hadDelete = false;
  private hadInsert = false;
  private pendingSepLeft = '';
  private pendingSepRight = '';
  private hasPendingSep = false;
  get open(): boolean { return this.hadDelete || this.hadInsert; }
  holdSeparator(left: string, right: string): void {
    this.pendingSepLeft = left;
    this.pendingSepRight = right;
    this.hasPendingSep = true;
  }
  addDelete(text: string): void { this.commitPendingSeparator(); this.del += text; this.hadDelete = true; }
  addInsert(text: string): void { this.commitPendingSeparator(); this.ins += text; this.hadInsert = true; }
  private commitPendingSeparator(): void {
    if (!this.hasPendingSep) return;
    this.del += this.pendingSepLeft;
    this.ins += this.pendingSepRight;
    this.hasPendingSep = false;
    this.pendingSepLeft = '';
    this.pendingSepRight = '';
  }
  flush(leftAnchor: string | null, rightAnchor: string | null, ctx: Context, sink: IrRevision[]): void {
    if (this.open) {
      let delText = this.del;
      let insText = this.ins;
      const emitByText = this.hadDelete && this.hadInsert && delText.length > 0 && insText.length > 0;
      if (emitByText) {
        if (delText !== insText) [delText, insText] = trimCommonWordAffixes(delText, insText);
        if (delText.length > 0) sink.push({ type: 'Deleted', text: delText, author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
        if (insText.length > 0) sink.push({ type: 'Inserted', text: insText, author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
      } else {
        if (this.hadDelete) sink.push({ type: 'Deleted', text: delText, author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
        if (this.hadInsert) sink.push({ type: 'Inserted', text: insText, author: author(ctx), date: date(ctx), leftAnchor, rightAnchor });
      }
    }
    this.del = '';
    this.ins = '';
    this.hadDelete = false;
    this.hadInsert = false;
    this.hasPendingSep = false;
    this.pendingSepLeft = '';
    this.pendingSepRight = '';
  }
}

function trimCommonWordAffixes(del: string, ins: string): [string, string] {
  const n = Math.min(del.length, ins.length);
  let prefix = 0;
  while (prefix < n && del[prefix] === ins[prefix]) prefix++;
  while (prefix > 0 && !(isWordBoundaryBefore(del, prefix) && isWordBoundaryBefore(ins, prefix))) prefix--;

  const remaining = n - prefix;
  let suffix = 0;
  while (suffix < remaining && del[del.length - 1 - suffix] === ins[ins.length - 1 - suffix]) suffix++;
  while (suffix > 0 && !(isWordBoundaryBefore(del, del.length - suffix) && isWordBoundaryBefore(ins, ins.length - suffix))) suffix--;
  return [del.slice(prefix, del.length - suffix), ins.slice(prefix, ins.length - suffix)];
}

function isWordBoundaryBefore(s: string, i: number): boolean {
  if (i <= 0 || i >= s.length) return true;
  return isOracleSplitChar(s, i - 1) || isOracleSplitChar(s, i);
}

function isOracleSplitChar(s: string, pos: number): boolean {
  const c = s[pos]!;
  if (defaultWordSeparators.has(c)) return true;
  const code = c.codePointAt(0)!;
  if (code >= 0x4e00 && code <= 0x9fff) return true;
  if (c === '.' || c === ',') {
    const prevDigit = pos > 0 && /\d/u.test(s[pos - 1]!);
    const nextDigit = pos < s.length - 1 && /\d/u.test(s[pos + 1]!);
    return !(prevDigit || nextDigit);
  }
  return false;
}

function isMaskedTextboxSpan(tokens: ReadonlyArray<IrDiffToken>, start: number, end: number): boolean {
  return start < end && tokens.slice(start, end).every((t) => t.kind === 'Textbox');
}

function isPureSeparatorSpan(tokens: ReadonlyArray<IrDiffToken>, start: number, end: number): boolean {
  return start < end && tokens.slice(start, end).every((t) => t.kind === 'Separator');
}

function renderFormatChangedSpan(op: IrTokenOp, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, leftAnchor: string | null, rightAnchor: string | null, ctx: Context, sink: IrRevision[]): void {
  const len = op.rightEnd - op.rightStart;
  let runStart = 0;
  while (runStart < len) {
    const li0 = op.leftStart + runStart;
    const ri0 = op.rightStart + runStart;
    const oldKey = runKey(leftTokens[li0]!.format);
    const newKey = runKey(rightTokens[ri0]!.format);
    let runEnd = runStart + 1;
    while (runEnd < len && runKey(leftTokens[op.leftStart + runEnd]!.format) === oldKey && runKey(rightTokens[op.rightStart + runEnd]!.format) === newKey) runEnd++;
    sink.push(formatRevision(rawText(rightTokens, op.rightStart + runStart, op.rightStart + runEnd), runFormatChangeDetails(leftTokens[li0]!.format, rightTokens[ri0]!.format), ctx, leftAnchor, rightAnchor));
    runStart = runEnd;
  }
}

function renderFormatOnlyBlock(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  const leftTokens = paragraphTokens(op.leftAnchor, ctx.left, ctx.settings);
  const rightTokens = paragraphTokens(op.rightAnchor, ctx.right, ctx.settings);
  const paraEmitted = emitParagraphScopeFormatChanged(op, ctx, sink);
  emitInlineSectionFormatChanged(op, ctx, sink);

  if (leftTokens.length === 0 && rightTokens.length === 0) {
    const lt = resolveTable(op.leftAnchor, ctx.left);
    const rt = resolveTable(op.rightAnchor, ctx.right);
    if (!paraEmitted && ctx.settings.trackBlockFormatChanges && lt && rt) emitTableFormatOnlyShellRevisions(lt, rt, ctx, sink);
    return;
  }
  if (leftTokens.length === rightTokens.length) {
    let emittedAny = false;
    let i = 0;
    while (i < leftTokens.length) {
      if (runFormatEqual(leftTokens[i]!.format, rightTokens[i]!.format, ctx.settings.formatComparison)) {
        i++;
        continue;
      }
      const oldKey = runKey(leftTokens[i]!.format);
      const newKey = runKey(rightTokens[i]!.format);
      let j = i + 1;
      while (
        j < leftTokens.length &&
        !runFormatEqual(leftTokens[j]!.format, rightTokens[j]!.format, ctx.settings.formatComparison) &&
        runKey(leftTokens[j]!.format) === oldKey &&
        runKey(rightTokens[j]!.format) === newKey
      ) j++;
      sink.push(formatRevision(rawText(rightTokens, i, j), runFormatChangeDetails(leftTokens[i]!.format, rightTokens[i]!.format), ctx, op.leftAnchor, op.rightAnchor));
      emittedAny = true;
      i = j;
    }
    if (!emittedAny && !paraEmitted) emitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);
    return;
  }
  emitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);
}

function emitWholeBlockFormatChanged(op: IrEditOp, leftTokens: ReadonlyArray<IrDiffToken>, rightTokens: ReadonlyArray<IrDiffToken>, ctx: Context, sink: IrRevision[]): void {
  const min = Math.min(leftTokens.length, rightTokens.length);
  let oldFmt: IrRunFormat | null = null;
  let newFmt: IrRunFormat | null = null;
  for (let i = 0; i < min; i++) {
    if (runKey(leftTokens[i]!.format) !== runKey(rightTokens[i]!.format)) {
      oldFmt = leftTokens[i]!.format;
      newFmt = rightTokens[i]!.format;
      break;
    }
  }
  if (oldFmt === null && newFmt === null && leftTokens.length !== rightTokens.length) {
    if (leftTokens.length > rightTokens.length) oldFmt = leftTokens[min]!.format;
    else newFmt = rightTokens[min]!.format;
  }
  sink.push(formatRevision(rawText(rightTokens, 0, rightTokens.length), runFormatChangeDetails(oldFmt, newFmt), ctx, op.leftAnchor, op.rightAnchor));
}

function paraFormatDiffers(left: IrParagraph, right: IrParagraph, settings: IrDiffSettings): boolean {
  if (!settings.trackParagraphFormatChanges) return false;
  return settings.formatComparison === 'ModeledOnly' ? paraKey(left.format) !== paraKey(right.format) : JSON.stringify(left.format) !== JSON.stringify(right.format);
}

function emitParagraphScopeFormatChanged(op: IrEditOp, ctx: Context, sink: IrRevision[]): boolean {
  const lp = op.leftAnchor ? ctx.left.anchorIndex.get(op.leftAnchor) : undefined;
  const rp = op.rightAnchor ? ctx.right.anchorIndex.get(op.rightAnchor) : undefined;
  if (lp?.kind !== 'paragraph' || rp?.kind !== 'paragraph' || !paraFormatDiffers(lp, rp, ctx.settings)) return false;
  const rightTokens = paragraphTokens(op.rightAnchor, ctx.right, ctx.settings);
  sink.push(formatRevision(rawText(rightTokens, 0, rightTokens.length), paraFormatChangeDetails(lp.format, rp.format), ctx, op.leftAnchor, op.rightAnchor));
  return true;
}

function emitInlineSectionFormatChanged(op: IrEditOp, ctx: Context, sink: IrRevision[]): void {
  if (!ctx.settings.trackBlockFormatChanges || !op.leftAnchor || !op.rightAnchor) return;
  const lp = ctx.left.anchorIndex.get(op.leftAnchor);
  const rp = ctx.right.anchorIndex.get(op.rightAnchor);
  if (lp?.kind !== 'paragraph' || rp?.kind !== 'paragraph' || !lp.inlineSectionFormat || !rp.inlineSectionFormat) return;
  const details = sectionFormatChangeDetails(lp.inlineSectionFormat, rp.inlineSectionFormat);
  if (details.changedPropertyNames.length === 0) return;
  sink.push(formatRevision('', details, ctx, lp.inlineSectionBreakAnchor ? anchorToString(lp.inlineSectionBreakAnchor) : op.leftAnchor, rp.inlineSectionBreakAnchor ? anchorToString(rp.inlineSectionBreakAnchor) : op.rightAnchor));
}

function tableDiffNeedsWholeTableFallback(td: IrTableDiff): boolean {
  return td.rowOps.some((r) => r.kind === 'ModifyRow' && r.cellOps?.some((c) => c.leftCellAnchor === null || c.rightCellAnchor === null));
}

function renderTableDiff(tableDiff: IrTableDiff, ctx: Context, sink: IrRevision[]): void {
  for (const rowOp of tableDiff.rowOps) {
    if (rowOp.kind === 'InsertRow') {
      sink.push({ type: 'Inserted', text: rowText(rowOp.rightRowAnchor, ctx.right, ctx.settings), author: author(ctx), date: date(ctx), rightAnchor: rowOp.rightRowAnchor });
    } else if (rowOp.kind === 'DeleteRow') {
      sink.push({ type: 'Deleted', text: rowText(rowOp.leftRowAnchor, ctx.left, ctx.settings), author: author(ctx), date: date(ctx), leftAnchor: rowOp.leftRowAnchor });
    } else if (rowOp.kind === 'MovedRow') {
      renderMovedRow(rowOp, ctx, sink);
    } else if (rowOp.kind === 'ModifyRow' && rowOp.cellOps) {
      for (const cellOp of rowOp.cellOps) renderCellOp(cellOp, ctx, sink);
    }
  }
}

function renderMovedRow(rowOp: IrRowOp, ctx: Context, sink: IrRevision[]): void {
  const isSource = rowOp.isMoveSource === true;
  const text = isSource ? rowText(rowOp.leftRowAnchor, ctx.left, ctx.settings) : rowText(rowOp.rightRowAnchor, ctx.right, ctx.settings);
  if (!ctx.settings.renderMoves) {
    sink.push(isSource
      ? { type: 'Deleted', text, author: author(ctx), date: date(ctx), leftAnchor: rowOp.leftRowAnchor }
      : { type: 'Inserted', text, author: author(ctx), date: date(ctx), rightAnchor: rowOp.rightRowAnchor });
    return;
  }
  sink.push({ type: 'Moved', text, author: author(ctx), date: date(ctx), moveGroupId: rowOp.moveGroupId, isMoveSource: isSource, leftAnchor: isSource ? rowOp.leftRowAnchor : null, rightAnchor: isSource ? null : rowOp.rightRowAnchor });
}

function renderCellOp(cellOp: IrCellOp, ctx: Context, sink: IrRevision[]): void {
  if (cellOp.blockOps) renderBlockOpList(cellOp.blockOps, ctx, sink);
}

function resolveTable(anchor: string | null, doc: IrDocument): IrTable | null {
  const block = anchor === null ? undefined : doc.anchorIndex.get(anchor);
  return block?.kind === 'table' ? block : null;
}

const emptyProps: Readonly<Record<string, string>> = {};

function tableShellRevision(scope: IrFormatChangeScope, changed: string, leftAnchor: string, rightAnchor: string, ctx: Context): IrRevision {
  return formatRevision('', { oldProperties: emptyProps, newProperties: emptyProps, changedPropertyNames: [changed], scope }, ctx, leftAnchor, rightAnchor);
}

function emitTableFormatOnlyShellRevisions(left: IrTable, right: IrTable, ctx: Context, sink: IrRevision[]): void {
  emitTableLevelShellRevisions(left, right, ctx, sink);
  const n = Math.min(left.rows.length, right.rows.length);
  for (let i = 0; i < n; i++) emitRowAndCellShellRevisions(left.rows[i]!, right.rows[i]!, ctx, sink);
}

function emitTableModifiedShellRevisions(left: IrTable, right: IrTable, diff: IrTableDiff, ctx: Context, sink: IrRevision[]): void {
  emitTableLevelShellRevisions(left, right, ctx, sink);
  const leftRows = new Map(left.rows.map((r) => [anchorToString(r.anchor), r]));
  const rightRows = new Map(right.rows.map((r) => [anchorToString(r.anchor), r]));
  for (const rowOp of diff.rowOps) {
    if (rowOp.kind !== 'EqualRow' && rowOp.kind !== 'ModifyRow') continue;
    const lr = rowOp.leftRowAnchor ? leftRows.get(rowOp.leftRowAnchor) : undefined;
    const rr = rowOp.rightRowAnchor ? rightRows.get(rowOp.rightRowAnchor) : undefined;
    if (lr && rr) emitRowAndCellShellRevisions(lr, rr, ctx, sink);
  }
}

function emitTableLevelShellRevisions(left: IrTable, right: IrTable, ctx: Context, sink: IrRevision[]): void {
  if (left.tblPrDigest !== right.tblPrDigest) sink.push(tableShellRevision('Table', 'shell', anchorToString(left.anchor), anchorToString(right.anchor), ctx));
  if (left.tblGridDigest !== right.tblGridDigest) sink.push(tableShellRevision('Table', 'grid', anchorToString(left.anchor), anchorToString(right.anchor), ctx));
}

function emitRowAndCellShellRevisions(left: IrRow, right: IrRow, ctx: Context, sink: IrRevision[]): void {
  if (left.trPrShellDigest !== right.trPrShellDigest) sink.push(tableShellRevision('TableRow', 'shell', anchorToString(left.anchor), anchorToString(right.anchor), ctx));
  if (left.trPrExDigest !== right.trPrExDigest) sink.push(tableShellRevision('TableRow', 'tblPrEx', anchorToString(left.anchor), anchorToString(right.anchor), ctx));
  const n = Math.min(left.cells.length, right.cells.length);
  for (let i = 0; i < n; i++) {
    if (left.cells[i]!.tcPrShellDigest !== right.cells[i]!.tcPrShellDigest) {
      sink.push(tableShellRevision('TableCell', 'shell', anchorToString(left.cells[i]!.anchor), anchorToString(right.cells[i]!.anchor), ctx));
    }
  }
}

function paragraphTokens(anchor: string | null, doc: IrDocument, settings: IrDiffSettings): ReadonlyArray<IrDiffToken> {
  const block = anchor === null ? undefined : doc.anchorIndex.get(anchor);
  return block?.kind === 'paragraph' ? tokenizeIrParagraph(block, settings) : [];
}

function blockText(anchor: string | null, doc: IrDocument, settings: IrDiffSettings): string {
  const block = anchor === null ? undefined : doc.anchorIndex.get(anchor);
  return block ? blockTextOf(block, settings) : '';
}

function blockTextOf(block: IrBlock, settings: IrDiffSettings): string {
  if (block.kind === 'paragraph') return rawText(tokenizeIrParagraph(block, settings), 0, tokenizeIrParagraph(block, settings).length);
  if (block.kind === 'table') return block.rows.map((r) => rowTextOf(r, settings)).join('');
  return '';
}

function rowText(anchor: string | null, doc: IrDocument, settings: IrDiffSettings): string {
  if (anchor === null) return '';
  const scanned = rowTextInScopes(anchor, [doc.body, ...[...doc.footnotes.notes.values()], ...[...doc.endnotes.notes.values()], ...doc.headers.map((h) => h.scope), ...doc.footers.map((f) => f.scope)], settings);
  return scanned ?? '';
}

function rowTextInScopes(anchor: string, scopes: ReadonlyArray<IrScope>, settings: IrDiffSettings): string | null {
  for (const scope of scopes) {
    const text = rowTextInBlocks(anchor, scope.blocks, settings);
    if (text !== null) return text;
  }
  return null;
}

function rowTextInBlocks(anchor: string, blocks: ReadonlyArray<IrBlock>, settings: IrDiffSettings): string | null {
  for (const block of blocks) {
    if (block.kind !== 'table') continue;
    const text = rowTextInTable(anchor, block, settings);
    if (text !== null) return text;
  }
  return null;
}

function rowTextInTable(anchor: string, table: IrTable, settings: IrDiffSettings): string | null {
  for (const row of table.rows) {
    if (anchorToString(row.anchor) === anchor) return rowTextOf(row, settings);
    for (const cell of row.cells) {
      const text = rowTextInBlocks(anchor, cell.blocks, settings);
      if (text !== null) return text;
    }
  }
  return null;
}

function rowTextOf(row: IrRow, settings: IrDiffSettings): string {
  let text = '';
  for (const cell of row.cells) text += cellText(cell, settings);
  return text;
}

function cellText(cell: IrCell, settings: IrDiffSettings): string {
  let text = '';
  for (const block of cell.blocks) if (block.kind === 'paragraph') text += rawText(tokenizeIrParagraph(block, settings), 0, tokenizeIrParagraph(block, settings).length);
  return text;
}

function rawText(tokens: ReadonlyArray<IrDiffToken>, start: number, end: number): string {
  let text = '';
  for (let i = start; i < end && i < tokens.length; i++) text += tokens[i]!.text;
  return text;
}

function formatRevision(text: string, formatChange: IrFormatChangeDetails, ctx: Context, leftAnchor: string | null, rightAnchor: string | null): IrRevision {
  return { type: 'FormatChanged', text, author: author(ctx), date: date(ctx), formatChange, leftAnchor, rightAnchor };
}

function runFormatEqual(a: IrRunFormat | null, b: IrRunFormat | null, comparison: IrDiffSettings['formatComparison']): boolean {
  return comparison === 'ModeledOnly' ? runKey(a) === runKey(b) : JSON.stringify(a) === JSON.stringify(b);
}

function runFormatChangeDetails(oldFormat: IrRunFormat | null, newFormat: IrRunFormat | null): IrFormatChangeDetails {
  const oldProperties = modeledRunProperties(oldFormat);
  const newProperties = modeledRunProperties(newFormat);
  const changedPropertyNames = changedNames(oldProperties, newProperties, [
    'style', 'bold', 'italic', 'underline', 'strikethrough', 'doubleStrikethrough',
    'verticalAlign', 'font', 'fontSize', 'color', 'highlight', 'allCaps', 'smallCaps', 'hidden',
  ]);
  return { oldProperties, newProperties, changedPropertyNames, scope: 'Run' };
}

function paraFormatChangeDetails(oldFormat: IrParaFormat, newFormat: IrParaFormat): IrFormatChangeDetails {
  const oldProperties = modeledParaProperties(oldFormat);
  const newProperties = modeledParaProperties(newFormat);
  const changedPropertyNames = changedNames(oldProperties, newProperties, [
    'style', 'justification', 'indentLeft', 'indentRight', 'indentFirstLine',
    'spacingBefore', 'spacingAfter', 'lineSpacing', 'outlineLevel',
    'keepNext', 'keepLines', 'pageBreakBefore', 'numId', 'numLevel',
  ]);
  return { oldProperties, newProperties, changedPropertyNames, scope: 'Paragraph' };
}

function sectionFormatChangeDetails(oldFormat: IrSectionFormat, newFormat: IrSectionFormat): IrFormatChangeDetails {
  const oldProperties = modeledSectionProperties(oldFormat);
  const newProperties = modeledSectionProperties(newFormat);
  const changedPropertyNames = changedNames(oldProperties, newProperties, [
    'pageWidth', 'pageHeight', 'landscape', 'marginTop', 'marginBottom',
    'marginLeft', 'marginRight', 'sectionType',
  ]);
  return { oldProperties, newProperties, changedPropertyNames, scope: 'Section' };
}

function modeledRunProperties(f: IrRunFormat | null): Record<string, string> {
  const dict: Record<string, string> = {};
  if (f === null) return dict;
  addProp(dict, 'style', f.styleId);
  addBool(dict, 'bold', f.bold);
  addBool(dict, 'italic', f.italic);
  addProp(dict, 'underline', f.underline === null ? null : f.underline.colorHex === null ? f.underline.kind : `${f.underline.kind}|${f.underline.colorHex}`);
  addBool(dict, 'strikethrough', f.strike);
  addBool(dict, 'doubleStrikethrough', f.doubleStrike);
  addProp(dict, 'verticalAlign', f.vertAlign);
  addProp(dict, 'font', f.fontAscii);
  addInt(dict, 'fontSize', f.sizeHalfPoints);
  addProp(dict, 'color', f.colorHex);
  addProp(dict, 'highlight', f.highlight);
  addBool(dict, 'allCaps', f.caps);
  addBool(dict, 'smallCaps', f.smallCaps);
  addBool(dict, 'hidden', f.vanish);
  return dict;
}

function modeledParaProperties(f: IrParaFormat | null): Record<string, string> {
  const dict: Record<string, string> = {};
  if (f === null) return dict;
  addProp(dict, 'style', f.styleId);
  addProp(dict, 'justification', f.justification);
  addInt(dict, 'indentLeft', f.indentLeftTwips);
  addInt(dict, 'indentRight', f.indentRightTwips);
  addInt(dict, 'indentFirstLine', f.indentFirstLineTwips);
  addInt(dict, 'spacingBefore', f.spacingBeforeTwips);
  addInt(dict, 'spacingAfter', f.spacingAfterTwips);
  addProp(dict, 'lineSpacing', f.lineSpacing === null ? null : `${f.lineSpacing.valueTwips}|${f.lineSpacing.rule}`);
  addInt(dict, 'outlineLevel', f.outlineLevel);
  addBool(dict, 'keepNext', f.keepNext);
  addBool(dict, 'keepLines', f.keepLines);
  addBool(dict, 'pageBreakBefore', f.pageBreakBefore);
  addInt(dict, 'numId', f.numId);
  addInt(dict, 'numLevel', f.ilvl);
  return dict;
}

function modeledSectionProperties(f: IrSectionFormat | null): Record<string, string> {
  const dict: Record<string, string> = {};
  if (f === null) return dict;
  addInt(dict, 'pageWidth', f.pageWidthTwips);
  addInt(dict, 'pageHeight', f.pageHeightTwips);
  addBool(dict, 'landscape', f.landscape);
  addInt(dict, 'marginTop', f.marginTopTwips);
  addInt(dict, 'marginBottom', f.marginBottomTwips);
  addInt(dict, 'marginLeft', f.marginLeftTwips);
  addInt(dict, 'marginRight', f.marginRightTwips);
  addProp(dict, 'sectionType', f.sectionType);
  return dict;
}

function changedNames(oldProperties: Readonly<Record<string, string>>, newProperties: Readonly<Record<string, string>>, order: ReadonlyArray<string>): string[] {
  const changed: string[] = [];
  for (const name of order) {
    const oldHas = Object.prototype.hasOwnProperty.call(oldProperties, name);
    const newHas = Object.prototype.hasOwnProperty.call(newProperties, name);
    if (oldHas !== newHas || (oldHas && newHas && oldProperties[name] !== newProperties[name])) changed.push(name);
  }
  return changed;
}

function addProp(dict: Record<string, string>, name: string, value: string | null | undefined): void {
  if (value !== null && value !== undefined) dict[name] = value;
}

function addBool(dict: Record<string, string>, name: string, value: boolean | null | undefined): void {
  if (value !== null && value !== undefined) dict[name] = value ? 'true' : 'false';
}

function addInt(dict: Record<string, string>, name: string, value: number | null | undefined): void {
  if (value !== null && value !== undefined) dict[name] = String(value);
}

export const IrRevisionRenderer = {
  render: renderIrRevisions,
  renderFromDocuments: renderIrRevisionsFromDocuments,
};
