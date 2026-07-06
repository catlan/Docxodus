// Behavioral tests for the deterministic unid scheme — the properties
// documented on C# UnidHelper (two opens of the same bytes produce
// identical unids; editing a paragraph changes only its own unid;
// duplicate-content siblings are disambiguated by dup index).

import { describe, expect, test } from 'vitest';
import { parseXml, type XElement } from '../src/xml/xelement.js';
import { W } from '../src/ir/names.js';
import { assignToAllElementsDeterministic } from '../src/ir/unid-helper.js';

const DOC = (body: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function bodyParagraphUnids(xml: string): string[] {
  const root = parseXml(xml);
  const unids = assignToAllElementsDeterministic(root);
  const body = root.element(W.body)!;
  return [...body.elements(W.p)].map((p) => unids.get(p)!);
}

describe('deterministic unids', () => {
  test('two parses of the same bytes produce identical unids', () => {
    const xml = DOC(P('alpha') + P('beta') + P('gamma'));
    expect(bodyParagraphUnids(xml)).toEqual(bodyParagraphUnids(xml));
  });

  test('unids are 32 lowercase hex chars', () => {
    for (const unid of bodyParagraphUnids(DOC(P('alpha')))) {
      expect(unid).toMatch(/^[0-9a-f]{32}$/);
    }
  });

  test('editing one paragraph changes only that paragraph', () => {
    const before = bodyParagraphUnids(DOC(P('alpha') + P('beta') + P('gamma')));
    const after = bodyParagraphUnids(DOC(P('alpha') + P('CHANGED') + P('gamma')));
    expect(after[0]).toBe(before[0]);
    expect(after[1]).not.toBe(before[1]);
    expect(after[2]).toBe(before[2]);
  });

  test('inserting a unique paragraph does not shift sibling unids', () => {
    const before = bodyParagraphUnids(DOC(P('alpha') + P('gamma')));
    const after = bodyParagraphUnids(DOC(P('alpha') + P('inserted') + P('gamma')));
    expect(after[0]).toBe(before[0]);
    expect(after[2]).toBe(before[1]);
  });

  test('duplicate-content siblings get distinct unids via dup index', () => {
    const unids = bodyParagraphUnids(DOC(P('same') + P('same') + P('same')));
    expect(new Set(unids).size).toBe(3);
  });

  test('style id participates in the signature', () => {
    const styled = (style: string) =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const a = bodyParagraphUnids(DOC(styled('Heading1')));
    const b = bodyParagraphUnids(DOC(styled('Heading2')));
    expect(a[0]).not.toBe(b[0]);
  });

  test('a persisted pt:Unid attribute wins and still counts for dup index', () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:pt14="http://powertools.codeplex.com/2011"><w:body><w:p pt14:Unid="feedfacefeedfacefeedfacefeedface"><w:r><w:t>same</w:t></w:r></w:p>${P('same')}</w:body></w:document>`;
    const root = parseXml(xml);
    const unids = assignToAllElementsDeterministic(root);
    const body = root.element(W.body)!;
    const paragraphs = [...body.elements(W.p)];
    const first = paragraphs[0] as XElement;
    const second = paragraphs[1] as XElement;
    expect(unids.get(first)).toBeUndefined(); // attribute wins, not in map
    expect(first.attribute({ ns: 'http://powertools.codeplex.com/2011', local: 'Unid' })).toBe(
      'feedfacefeedfacefeedfacefeedface',
    );
    // The second 'same' paragraph is dup index 1, matching the
    // all-assigned second slot rather than colliding into index 0.
    const bothAssigned = bodyParagraphUnids(DOC(P('same') + P('same')));
    expect(unids.get(second)).toBe(bothAssigned[1]);
  });
});
