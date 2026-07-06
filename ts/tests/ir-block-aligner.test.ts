import { describe, expect, test } from 'vitest';
import { IrBlockAligner, type IrAlignmentKind } from '../src/index.js';
import { assertInvariants, count, text } from './helpers/ir-alignment-asserts.js';
import { doc, docFromParagraphs, fromBodyXml } from './helpers/ir-test-documents.js';

const align = (l: ReturnType<typeof doc>, r: ReturnType<typeof doc>) => IrBlockAligner.align(l, r);

describe('IrBlockAligner', () => {
  test('identity all unchanged', () => {
    const l = doc('alpha', 'beta', 'gamma');
    const r = doc('alpha', 'beta', 'gamma');
    const a = align(l, r);
    expect(a.entries.every((e) => e.kind === 'Unchanged')).toBe(true);
    expect(a.entries).toHaveLength(3);
    assertInvariants(l, r, a);
  });

  test('single text edit is modified', () => {
    const l = doc('alpha', 'beta', 'gamma');
    const r = doc('alpha', 'BETA-edited', 'gamma');
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Modified')).toBe(1);
    expect(count(a, 'Moved')).toBe(0);
    assertInvariants(l, r, a);
  });

  test.each([
    ['start', ['alpha', 'beta'], ['NEW', 'alpha', 'beta'], 0],
    ['middle', ['alpha', 'beta'], ['alpha', 'NEW', 'beta'], 1],
    ['end', ['alpha', 'beta'], ['alpha', 'beta', 'NEW'], 2],
  ])('insert at %s', (_name, left, right, entryIndex) => {
    const l = doc(...left);
    const r = doc(...right);
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Inserted')).toBe(1);
    expect(a.entries[entryIndex]!.kind).toBe('Inserted');
    assertInvariants(l, r, a);
  });

  test.each([
    ['start', ['alpha', 'beta', 'gamma'], ['beta', 'gamma'], 0],
    ['middle', ['alpha', 'beta', 'gamma'], ['alpha', 'gamma'], 1],
    ['end', ['alpha', 'beta', 'gamma'], ['alpha', 'beta'], 2],
  ])('delete at %s', (_name, left, right, entryIndex) => {
    const l = doc(...left);
    const r = doc(...right);
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Deleted')).toBe(1);
    expect(a.entries[entryIndex]!.kind).toBe('Deleted');
    assertInvariants(l, r, a);
  });

  test('pure move yields exactly one moved rest unchanged', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta');
    const r = doc('gamma', 'alpha', 'beta', 'delta');
    const a = align(l, r);
    expect(count(a, 'Moved')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(3);
    expect(count(a, 'Modified')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
    expect(text(a.entries.find((e) => e.kind === 'Moved')!.right!)).toBe('gamma');
    assertInvariants(l, r, a);
  });

  test('move and unrelated edit classified independently', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta', 'epsilon');
    const r = doc('epsilon', 'alpha', 'beta-edited', 'gamma', 'delta');
    const a = align(l, r);
    expect(count(a, 'Moved')).toBe(1);
    expect(count(a, 'Modified')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(3);
    expect(text(a.entries.find((e) => e.kind === 'Moved')!.right!)).toBe('epsilon');
    assertInvariants(l, r, a);
  });

  test('adjacent swap of two unique paragraphs', () => {
    const l = doc('alpha', 'beta', 'gamma');
    const r = doc('beta', 'alpha', 'gamma');
    const a = align(l, r);
    expect(count(a, 'Moved')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Modified')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('bolding a paragraph is format only', () => {
    const l = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>');
    const r = fromBodyXml('<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>');
    const a = align(l, r);
    expect(count(a, 'FormatOnly')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(1);
    expect(count(a, 'Modified')).toBe(0);
    expect(count(a, 'Moved')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('boilerplate delete one of ten identical no false moves', () => {
    const l = docFromParagraphs(Array(10).fill('boilerplate'));
    const r = docFromParagraphs(Array(9).fill('boilerplate'));
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(9);
    expect(count(a, 'Deleted')).toBe(1);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Modified')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('cross gap move and edit is moved modified', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta', 'the quick brown fox jumps over hounds');
    const r = doc('the quick brown fox jumps over dogs', 'alpha', 'beta', 'gamma', 'delta');
    const a = align(l, r);
    expect(count(a, 'MovedModified')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(4);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Deleted')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    const mm = a.entries.find((e) => e.kind === 'MovedModified')!;
    expect(text(mm.left!)).toBe('the quick brown fox jumps over hounds');
    expect(text(mm.right!)).toBe('the quick brown fox jumps over dogs');
    expect(a.entries[0]!.kind).toBe('MovedModified');
    assertInvariants(l, r, a);
  });

  test('cross gap below similarity threshold stays delete insert', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta', 'the quick brown fox jumps over hounds');
    const r = doc('an entirely unrelated sentence with different words throughout', 'alpha', 'beta', 'gamma', 'delta');
    const a = align(l, r);
    expect(count(a, 'MovedModified')).toBe(0);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Deleted')).toBe(1);
    expect(count(a, 'Inserted')).toBe(1);
    expect(count(a, 'Unchanged')).toBe(4);
    assertInvariants(l, r, a);
  });

  test('cross gap below minimum token count stays delete insert', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta', 'hello world');
    const r = doc('hello earth', 'alpha', 'beta', 'gamma', 'delta');
    const a = align(l, r);
    expect(count(a, 'MovedModified')).toBe(0);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'Deleted')).toBe(1);
    expect(count(a, 'Inserted')).toBe(1);
    assertInvariants(l, r, a);
  });

  test('cross gap exact relocation residue classifies as moved not moved modified', () => {
    const l = doc('shared phrase here now', 'alpha', 'beta', 'gamma', 'shared phrase here now');
    const r = doc('shared phrase here now', 'alpha', 'beta', 'gamma');
    const a = align(l, r);
    expect(count(a, 'MovedModified')).toBe(0);
    assertInvariants(l, r, a);
  });

  test('in gap cross positioned edit pairs as modified', () => {
    const l = doc('alpha', 'the quick brown fox jumps high', 'a lazy sleepy dog rests here', 'omega');
    const r = doc('alpha', 'a lazy sleepy dog rests there', 'the quick brown fox leaps high', 'omega');
    const a = align(l, r);
    expect(count(a, 'Modified')).toBe(2);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Deleted')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    const modifies = a.entries.filter((e) => e.kind === 'Modified');
    expect(modifies.some((e) => text(e.left!).includes('quick brown fox') && text(e.right!).includes('quick brown fox'))).toBe(true);
    expect(modifies.some((e) => text(e.left!).includes('lazy sleepy dog') && text(e.right!).includes('lazy sleepy dog'))).toBe(true);
    assertInvariants(l, r, a);
  });

  test('boilerplate adversarial yields zero false moves', () => {
    const standard = Array(8).fill('Standard clause.');
    const l = doc(...standard, 'unique closing remark goes here');
    const r = doc(...standard.slice(0, 7), 'unique closing remark goes here');
    const a = align(l, r);
    expect(count(a, 'Moved')).toBe(0);
    expect(count(a, 'MovedModified')).toBe(0);
    expect(count(a, 'Deleted')).toBe(1);
    assertInvariants(l, r, a);
  });

  test('cross gap move detection is deterministic', () => {
    const l = doc('alpha', 'beta', 'gamma', 'the quick brown fox jumps over hounds');
    const r = doc('the quick brown fox jumps over dogs', 'alpha', 'beta', 'gamma');
    const a1 = align(l, r);
    const a2 = align(l, r);
    expect(a1.entries).toEqual(a2.entries);
    expect(count(a1, 'MovedModified')).toBe(1);
    assertInvariants(l, r, a1);
  });

  test('table cell edit makes table block modified', () => {
    const table = (cell: string) =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const l = fromBodyXml('<w:p><w:r><w:t>intro</w:t></w:r></w:p>' + table('cell-old'));
    const r = fromBodyXml('<w:p><w:r><w:t>intro</w:t></w:r></w:p>' + table('cell-new'));
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(1);
    expect(count(a, 'Modified')).toBe(1);
    expect(a.entries.find((e) => e.kind === 'Modified')!.left!.kind).toBe('table');
    assertInvariants(l, r, a);
  });

  test('unrelated tables in a gap pair as modified', () => {
    const row = (cell: string) => `<w:tr><w:tc><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc></w:tr>`;
    const table = (a: string, b: string) =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${row(a)}${row(b)}</w:tbl>`;
    const l = fromBodyXml(`<w:p><w:r><w:t>head</w:t></w:r></w:p>${table('Apple', 'Banana')}<w:p><w:r><w:t>tail</w:t></w:r></w:p>`);
    const r = fromBodyXml(`<w:p><w:r><w:t>head</w:t></w:r></w:p>${table('Xylophone', 'Zebra')}<w:p><w:r><w:t>tail</w:t></w:r></w:p>`);
    const a = align(l, r);
    expect(count(a, 'Unchanged')).toBe(2);
    expect(count(a, 'Modified')).toBe(1);
    expect(count(a, 'Deleted')).toBe(0);
    expect(count(a, 'Inserted')).toBe(0);
    expect(a.entries.find((e) => e.kind === 'Modified')!.left!.kind).toBe('table');
    assertInvariants(l, r, a);
  });

  test.each([
    ['empty left all inserted', fromBodyXml(''), doc('alpha', 'beta'), 'Inserted' as IrAlignmentKind, 2],
    ['empty right all deleted', doc('alpha', 'beta'), fromBodyXml(''), 'Deleted' as IrAlignmentKind, 2],
  ])('%s', (_name, l, r, kind, n) => {
    const a = align(l, r);
    expect(count(a, kind)).toBe(n);
    expect(a.entries).toHaveLength(n);
    assertInvariants(l, r, a);
  });

  test('both empty no entries', () => {
    const l = fromBodyXml('');
    const r = fromBodyXml('');
    const a = align(l, r);
    expect(a.entries).toEqual([]);
    assertInvariants(l, r, a);
  });

  test('two align calls are sequence equal', () => {
    const l = doc('alpha', 'beta', 'gamma', 'delta', 'boilerplate', 'boilerplate');
    const r = doc('gamma', 'alpha', 'beta-edited', 'boilerplate', 'delta', 'NEW');
    const a1 = align(l, r);
    const a2 = align(l, r);
    expect(a1.entries).toEqual(a2.entries);
    assertInvariants(l, r, a1);
  });
});
