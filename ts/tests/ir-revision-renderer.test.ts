import { describe, expect, test } from 'vitest';
import {
  buildIrEditScript,
  deterministicEpoch,
  renderIrRevisions,
  renderIrRevisionsFromDocuments,
  type IrDiffSettingsOptions,
  type IrDocument,
  type IrRevision,
} from '../src/index.js';
import { doc, fromBodyXml } from './helpers/ir-test-documents.js';

function render(left: IrDocument, right: IrDocument, settings: IrDiffSettingsOptions = {}): IrRevision[] {
  return renderIrRevisions(buildIrEditScript(left, right, settings), left, right, settings);
}

const compatible: IrDiffSettingsOptions = { revisionGranularity: 'WmlComparerCompatible' };

describe('IrRevisionRenderer', () => {
  test('insert and delete blocks carry text and one-sided anchors', () => {
    const inserted = render(doc('alpha', 'beta'), doc('alpha', 'inserted here', 'beta')).filter((r) => r.type === 'Inserted');
    expect(inserted).toHaveLength(1);
    expect(inserted[0]!.text).toBe('inserted here');
    expect(inserted[0]!.leftAnchor).toBeUndefined();
    expect(inserted[0]!.rightAnchor).toBeTruthy();

    const deleted = render(doc('alpha', 'to be removed', 'gamma'), doc('alpha', 'gamma')).filter((r) => r.type === 'Deleted');
    expect(deleted).toHaveLength(1);
    expect(deleted[0]!.text).toBe('to be removed');
    expect(deleted[0]!.rightAnchor).toBeUndefined();
    expect(deleted[0]!.leftAnchor).toBeTruthy();
  });

  test('modify block projects inserted and deleted token spans', () => {
    const revs = render(doc('the quick brown fox'), doc('the slow brown fox'));
    expect(revs.some((r) => r.type === 'Deleted' && r.text.includes('quick'))).toBe(true);
    expect(revs.some((r) => r.type === 'Inserted' && r.text.includes('slow'))).toBe(true);
    expect(revs.every((r) => r.text !== null)).toBe(true);
  });

  test('run format changes include modeled details and split heterogeneous sub-runs', () => {
    const left = fromBodyXml('<w:p><w:r><w:t>one two</w:t></w:r></w:p>');
    const right = fromBodyXml(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t></w:r>' +
      '<w:r><w:t> </w:t></w:r>' +
      '<w:r><w:rPr><w:i/></w:rPr><w:t>two</w:t></w:r></w:p>',
    );
    const fmt = render(left, right).filter((r) => r.type === 'FormatChanged');
    expect(fmt).toHaveLength(2);
    expect(fmt.some((r) => r.text === 'one' && r.formatChange?.changedPropertyNames.includes('bold'))).toBe(true);
    expect(fmt.some((r) => r.text === 'two' && r.formatChange?.changedPropertyNames.includes('italic'))).toBe(true);
  });

  test('format-only block reports a format changed revision with both anchors', () => {
    const left = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>');
    const right = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>');
    const fmt = render(left, right).filter((r) => r.type === 'FormatChanged');
    expect(fmt).toHaveLength(1);
    expect(fmt[0]!.text).toBe('beta');
    expect(fmt[0]!.leftAnchor).toBeTruthy();
    expect(fmt[0]!.rightAnchor).toBeTruthy();
    expect(fmt[0]!.formatChange?.newProperties.bold).toBe('true');
  });

  test('moves produce paired moved revisions and renderMoves false demotes them', () => {
    const left = doc('alpha', 'beta', 'gamma', 'delta');
    const right = doc('gamma', 'alpha', 'beta', 'delta');
    const moved = render(left, right).filter((r) => r.type === 'Moved');
    expect(moved).toHaveLength(2);
    expect(moved[0]!.moveGroupId).toBe(moved[1]!.moveGroupId);
    expect(moved.some((r) => r.isMoveSource === true && r.leftAnchor && !r.rightAnchor)).toBe(true);
    expect(moved.some((r) => r.isMoveSource === false && r.rightAnchor && !r.leftAnchor)).toBe(true);

    const demoted = render(left, right, { renderMoves: false });
    expect(demoted.some((r) => r.type === 'Moved')).toBe(false);
    expect(demoted.some((r) => r.type === 'Deleted')).toBe(true);
    expect(demoted.some((r) => r.type === 'Inserted')).toBe(true);
  });

  test('move-modify destination emits nested token revisions after destination move', () => {
    const revs = render(
      doc('alpha', 'beta', 'gamma', 'delta', 'the quick brown fox jumps over hounds'),
      doc('the quick brown fox jumps over dogs', 'alpha', 'beta', 'gamma', 'delta'),
    );
    const destIndex = revs.findIndex((r) => r.type === 'Moved' && r.isMoveSource === false);
    expect(destIndex).toBeGreaterThanOrEqual(0);
    const after = revs.slice(destIndex + 1);
    expect(after.some((r) => r.type === 'Deleted' && r.text.includes('hounds'))).toBe(true);
    expect(after.some((r) => r.type === 'Inserted' && r.text.includes('dogs'))).toBe(true);
  });

  test('table row and cell edits recurse into row or token revisions', () => {
    const row = (text: string) => `<w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
    const table = (rows: string) => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${rows}</w:tbl>`;
    const rowRevs = render(
      fromBodyXml('<w:p><w:r><w:t>intro</w:t></w:r></w:p>' + table(row('keep') + row('removed'))),
      fromBodyXml('<w:p><w:r><w:t>intro</w:t></w:r></w:p>' + table(row('keep') + row('added'))),
    );
    expect(rowRevs.some((r) => r.type === 'Deleted' && r.text.includes('removed'))).toBe(true);
    expect(rowRevs.some((r) => r.type === 'Inserted' && r.text.includes('added'))).toBe(true);

    const oneRow = (text: string) => table(row(`cell ${text} text here`));
    const cellRevs = render(fromBodyXml(oneRow('old')), fromBodyXml(oneRow('new')));
    expect(cellRevs.some((r) => r.type === 'Deleted' && r.text.includes('old'))).toBe(true);
    expect(cellRevs.some((r) => r.type === 'Inserted' && r.text.includes('new'))).toBe(true);
  });

  test('compatible mode coalesces adjacent word ops and trims common affixes', () => {
    const coalesced = render(doc('This is now foo bar baz'), doc('This is what are the chances'), compatible);
    expect(coalesced.filter((r) => r.type === 'Deleted')).toHaveLength(1);
    expect(coalesced.filter((r) => r.type === 'Inserted')).toHaveLength(1);
    expect(coalesced.find((r) => r.type === 'Deleted')!.text).toBe('now foo bar baz');
    expect(coalesced.find((r) => r.type === 'Inserted')!.text).toBe('what are the chances');

    const trimmed = render(
      fromBodyXml('<w:p><w:r><w:t>This is a test.</w:t></w:r></w:p>'),
      fromBodyXml('<w:p><w:r><w:t>This.</w:t></w:r></w:p>'),
      compatible,
    );
    expect(trimmed.filter((r) => r.type === 'Deleted')).toHaveLength(1);
    expect(trimmed.find((r) => r.type === 'Deleted')!.text).toBe(' is a test');
    expect(trimmed.some((r) => r.type === 'Inserted')).toBe(false);
  });

  test('author and date settings are deterministic and overridable', () => {
    const defaults = renderIrRevisionsFromDocuments(doc('alpha'), doc('alpha edited'));
    expect(defaults.length).toBeGreaterThan(0);
    expect(defaults.every((r) => r.author === 'Open-Xml-PowerTools' && r.date === deterministicEpoch)).toBe(true);

    const custom = renderIrRevisionsFromDocuments(doc('alpha'), doc('alpha edited'), {
      authorForRevisions: 'Daisy',
      dateTimeForRevisions: '2021-07-04T12:00:00Z',
    });
    expect(custom.every((r) => r.author === 'Daisy' && r.date === '2021-07-04T12:00:00Z')).toBe(true);
  });
});
