import { describe, expect, test } from 'vitest';
import {
  buildIrEditScript,
  writeIrEditScriptJson,
  type IrDocument,
  type IrEditOpKind,
} from '../src/index.js';
import { verifyIrEditScript } from './helpers/ir-edit-script-verifier.js';
import { doc, fromBodyXml } from './helpers/ir-test-documents.js';

const count = (s: ReturnType<typeof buildIrEditScript>, kind: IrEditOpKind): number => s.ops.filter((o) => o.kind === kind).length;

function buildVerified(left: IrDocument, right: IrDocument) {
  const script = buildIrEditScript(left, right);
  verifyIrEditScript(left, right, script);
  expect(writeIrEditScriptJson(script)).toBe(writeIrEditScriptJson(buildIrEditScript(left, right)));
  return script;
}

describe('IrEditScriptBuilder', () => {
  test('identity maps to equal blocks', () => {
    const s = buildVerified(doc('alpha', 'beta', 'gamma'), doc('alpha', 'beta', 'gamma'));
    expect(s.ops).toHaveLength(3);
    expect(s.ops.every((o) => o.kind === 'EqualBlock' && o.leftAnchor && o.rightAnchor && o.tokenDiff === null)).toBe(true);
  });

  test('single edit maps to modify block with token diff', () => {
    const s = buildVerified(doc('alpha', 'beta', 'gamma'), doc('alpha', 'beta edited here', 'gamma'));
    expect(count(s, 'EqualBlock')).toBe(2);
    expect(count(s, 'ModifyBlock')).toBe(1);
    const modify = s.ops.find((o) => o.kind === 'ModifyBlock')!;
    expect(modify.tokenDiff).not.toBeNull();
    expect((modify.tokenDiff as { ops: Array<{ kind: string }> }).ops.some((o) => o.kind === 'Insert')).toBe(true);
  });

  test('insert and delete anchors are one-sided and ordered', () => {
    const inserted = buildVerified(doc('alpha', 'beta'), doc('alpha', 'NEW', 'beta'));
    expect(inserted.ops[1]!.kind).toBe('InsertBlock');
    expect(inserted.ops[1]!.leftAnchor).toBeNull();
    expect(inserted.ops[1]!.rightAnchor).not.toBeNull();

    const deleted = buildVerified(doc('alpha', 'beta', 'gamma'), doc('alpha', 'gamma'));
    expect(deleted.ops.map((o) => o.kind)).toEqual(['EqualBlock', 'DeleteBlock', 'EqualBlock']);
    expect(deleted.ops[1]!.leftAnchor).not.toBeNull();
    expect(deleted.ops[1]!.rightAnchor).toBeNull();
  });

  test('format-only block maps without token diff', () => {
    const left = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>');
    const right = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>');
    const s = buildVerified(left, right);
    expect(count(s, 'FormatOnlyBlock')).toBe(1);
    expect(s.ops.find((o) => o.kind === 'FormatOnlyBlock')!.tokenDiff).toBeNull();
  });

  test('pure move emits paired source and destination with destination-order group id', () => {
    const left = doc('alpha', 'beta', 'gamma', 'delta');
    const right = doc('gamma', 'alpha', 'beta', 'delta');
    const s = buildVerified(left, right);
    const moves = s.ops.filter((o) => o.kind === 'MoveBlock');
    expect(moves).toHaveLength(2);
    const source = moves.find((o) => o.isMoveSource === true)!;
    const dest = moves.find((o) => o.isMoveSource === false)!;
    expect(source.moveGroupId).toBe(dest.moveGroupId);
    expect(source.moveGroupId).toBe(1);
    expect(s.ops[0]!.kind).toBe('MoveBlock');
    expect(s.ops[0]!.isMoveSource).toBe(false);
  });

  test('move source interleaves at left-anchored position', () => {
    const left = doc('alpha', 'beta', 'gamma', 'delta');
    const right = doc('gamma', 'alpha', 'beta', 'delta');
    const s = buildVerified(left, right);
    const sourceIdx = s.ops.findIndex((o) => o.kind === 'MoveBlock' && o.isMoveSource === true);
    const betaIdx = s.ops.findIndex((o) => o.leftAnchor === [...left.anchorIndex.keys()][1]);
    const deltaIdx = s.ops.findIndex((o) => o.leftAnchor === [...left.anchorIndex.keys()][3]);
    expect(sourceIdx).toBeGreaterThan(betaIdx);
    expect(sourceIdx).toBeLessThan(deltaIdx);
  });

  test('moved and edited paragraph is move modify with destination token diff', () => {
    const s = buildVerified(
      doc('alpha', 'beta', 'gamma', 'delta', 'the quick brown fox jumps over hounds'),
      doc('the quick brown fox jumps over dogs', 'alpha', 'beta', 'gamma', 'delta'),
    );
    const moves = s.ops.filter((o) => o.kind === 'MoveModifyBlock');
    expect(moves).toHaveLength(2);
    expect(moves.find((o) => o.isMoveSource === true)!.tokenDiff).toBeNull();
    const dest = moves.find((o) => o.isMoveSource === false)!;
    expect(dest.tokenDiff).not.toBeNull();
    expect((dest.tokenDiff as { ops: Array<{ kind: string }> }).ops.map((o) => o.kind)).toContain('Insert');
  });

  test('table modification carries nested table diff', () => {
    const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
    const table = (...rows: string[]) => `<w:tbl><w:tblPr/><w:tblGrid/>${rows.join('')}</w:tbl>`;
    const s = buildVerified(
      fromBodyXml(table(row(cell('alpha'), cell('beta')), row(cell('gamma'), cell('delta')))),
      fromBodyXml(table(row(cell('alpha'), cell('beta')), row(cell('gamma'), cell('edited')))),
    );
    const op = s.ops.find((o) => o.kind === 'ModifyBlock')!;
    expect(op.tableDiff).not.toBeNull();
  });

  test('split and merge scripts carry segment diffs', () => {
    const split = buildVerified(
      doc('aaa bbb ccc ddd. eee fff ggg hhh.', 'anchor one two three four five.'),
      doc('aaa bbb ccc ddd. ', 'eee fff ggg hhh.', 'anchor one two three four five.'),
    );
    const splitOp = split.ops.find((o) => o.kind === 'SplitBlock')!;
    expect(splitOp.splitMergeAnchors).toHaveLength(2);
    expect(splitOp.segmentDiffs).toHaveLength(2);

    const merge = buildVerified(
      doc('aaa bbb ccc ddd. ', 'eee fff ggg hhh.', 'anchor one two three four.'),
      doc('aaa bbb ccc ddd. eee fff ggg hhh.', 'anchor one two three four.'),
    );
    const mergeOp = merge.ops.find((o) => o.kind === 'MergeBlock')!;
    expect(mergeOp.splitMergeAnchors).toHaveLength(2);
    expect(mergeOp.segmentDiffs).toHaveLength(2);
  });

  test('json writer uses C# edit-script field names and compact token op arrays', () => {
    const script = buildVerified(doc('alpha'), doc('alpha edited'));
    const json = writeIrEditScriptJson(script);
    expect(json.startsWith('{\n  "operations": [')).toBe(true);
    expect(json).toContain('"tokenDiff": {');
    expect(json).toMatch(/\[\s+[0-3],\s+\d+,\s+\d+,\s+\d+,\s+\d+\s+\]/);
    expect(json).not.toContain('"leftAnchor": null');
  });
});
