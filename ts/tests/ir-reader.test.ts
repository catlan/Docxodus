import { describe, expect, test } from 'vitest';
import { anchorToString, parseXml, readIrDocument, type IrParagraph } from '../src/index.js';

const enc = new TextEncoder();

const DOC = (body: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`;

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function parts(body: string, styles?: string, rels?: string): Map<string, Uint8Array> {
  const map = new Map<string, Uint8Array>();
  map.set('word/document.xml', enc.encode(DOC(body)));
  if (styles) {
    map.set(
      'word/styles.xml',
      enc.encode(`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`),
    );
  }
  if (rels) {
    map.set(
      'word/_rels/document.xml.rels',
      enc.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`),
    );
  }
  return map;
}

const paragraphs = (body: string, styles?: string): IrParagraph[] =>
  readIrDocument(parts(body, styles)).body.blocks.filter((b): b is IrParagraph => b.kind === 'paragraph');

const text = (p: IrParagraph): string =>
  p.inlines.filter((i) => i.kind === 'textRun').map((i) => i.text).join('');

describe('IrReader stage A', () => {
  test('simple paragraphs produce paragraph blocks and anchors', () => {
    const paras = paragraphs(P('Hello world') + P('Second line'));
    expect(paras).toHaveLength(2);
    expect(text(paras[0]!)).toBe('Hello world');
    expect(text(paras[1]!)).toBe('Second line');
    for (const p of paras) {
      expect(p.anchor.kind).toBe('p');
      expect(p.anchor.scope).toBe('body');
      expect(p.anchor.unid).toMatch(/^[0-9a-f]{32}$/);
    }
  });

  test('twice yields identical anchors and hashes', () => {
    const input = parts(P('Same bytes') + P('Twice over'));
    const a = readIrDocument(input).body.blocks;
    const b = readIrDocument(input).body.blocks;
    expect(a.map((x) => anchorToString(x.anchor))).toEqual(b.map((x) => anchorToString(x.anchor)));
    expect(a.map((x) => x.contentHash)).toEqual(b.map((x) => x.contentHash));
    expect(a.map((x) => x.formatFingerprint)).toEqual(b.map((x) => x.formatFingerprint));
  });

  test('bold run maps run format', () => {
    const para = paragraphs('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>')[0]!;
    const run = para.inlines[0]!;
    expect(run.kind).toBe('textRun');
    if (run.kind === 'textRun') expect(run.format.bold).toBe(true);
  });

  test('body-level bookmark markers are dropped', () => {
    const ir = readIrDocument(parts(P('before') + '<w:bookmarkStart w:id="7" w:name="sec"/><w:bookmarkEnd w:id="7"/>' + P('after')));
    expect(ir.body.blocks).toHaveLength(2);
    expect(ir.body.blocks.some((b) => b.kind === 'opaqueBlock')).toBe(false);
    expect(text(ir.body.blocks[0] as IrParagraph)).toBe('before');
    expect(text(ir.body.blocks[1] as IrParagraph)).toBe('after');
  });

  test('adjacent equal runs coalesce', () => {
    const para = paragraphs('<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>')[0]!;
    expect(para.inlines.filter((i) => i.kind === 'textRun')).toHaveLength(1);
    expect(text(para)).toBe('Hello world');
  });

  test('tab and break become typed inlines', () => {
    const para = paragraphs('<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br w:type="page"/></w:r></w:p>')[0]!;
    expect(para.inlines.some((i) => i.kind === 'tab')).toBe(true);
    const br = para.inlines.find((i) => i.kind === 'break');
    expect(br).toMatchObject({ kind: 'break', breakKind: 'Page' });
  });

  test('table structure and anchors', () => {
    const ir = readIrDocument(parts(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>' +
        '<w:tr><w:tc>' + P('R0C0') + '</w:tc><w:tc>' + P('R0C1') + '</w:tc></w:tr>' +
        '<w:tr><w:tc>' + P('R1C0') + '</w:tc><w:tc>' + P('R1C1') + '</w:tc></w:tr></w:tbl>',
    ));
    const table = ir.body.blocks[0]!;
    expect(table.kind).toBe('table');
    if (table.kind !== 'table') return;
    expect(table.anchor.kind).toBe('tbl');
    expect(table.rows).toHaveLength(2);
    for (const row of table.rows) {
      expect(row.anchor.kind).toBe('tr');
      expect(row.cells).toHaveLength(2);
      for (const cell of row.cells) {
        expect(cell.anchor.kind).toBe('tc');
        expect(cell.blocks[0]!.kind).toBe('paragraph');
        expect(ir.anchorIndex.get(anchorToString(cell.blocks[0]!.anchor))).toBe(cell.blocks[0]);
      }
    }
  });

  test('nested table recurses and indexes inner table', () => {
    const ir = readIrDocument(parts('<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc>' + P('inner') + '</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>'));
    const outer = ir.body.blocks[0]!;
    expect(outer.kind).toBe('table');
    if (outer.kind !== 'table') return;
    const inner = outer.rows[0]!.cells[0]!.blocks[0]!;
    expect(inner.kind).toBe('table');
    expect(ir.anchorIndex.get(anchorToString(inner.anchor))).toBe(inner);
  });

  test('unknown elements become opaque', () => {
    const ir = readIrDocument(parts('<w:unknownBlock w:id="1"/><w:p><w:r><w:ptab w:relativeTo="margin"/></w:r></w:p>'), { revisionView: 'raw' });
    expect(ir.body.blocks.some((b) => b.kind === 'opaqueBlock')).toBe(true);
    const para = ir.body.blocks.find((b): b is IrParagraph => b.kind === 'paragraph')!;
    expect(para.inlines.some((i) => i.kind === 'opaqueInline')).toBe(true);
  });

  test('content hash ignores formatting but fingerprint changes', () => {
    const plain = paragraphs(P('hello'))[0]!;
    const bold = paragraphs('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r></w:p>')[0]!;
    expect(plain.contentHash).toBe(bold.contentHash);
    expect(plain.formatFingerprint).not.toBe(bold.formatFingerprint);
  });

  test('accepted revision view is default and fail mode still rejects body revisions', () => {
    const input = parts('<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="1"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>');
    expect(text(readIrDocument(input).body.blocks[0] as IrParagraph)).toBe('kept inserted');
    expect(() =>
      readIrDocument(input, { revisionView: 'failIfPresent' }),
    ).toThrow(/revision-free XML/);
  });

  test('trailing sectPr becomes section break', () => {
    const ir = readIrDocument(parts(P('body') + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'));
    const sec = ir.body.blocks[1]!;
    expect(sec.kind).toBe('sectionBreak');
    if (sec.kind === 'sectionBreak') expect(sec.format.pageWidthTwips).toBe(12240);
  });

  test('style-inherited list item is classified as li', () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="ListBase"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="MyListPara"><w:basedOn w:val="ListBase"/></w:style>';
    const para = paragraphs('<w:p><w:pPr><w:pStyle w:val="MyListPara"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>', styles)[0]!;
    expect(para.anchor.kind).toBe('li');
  });

  test('unmodeled formatting flips fingerprint only', () => {
    const plain = paragraphs(P('same'))[0]!;
    const runUnmodeled = paragraphs('<w:p><w:r><w:rPr><w:rFonts w:hAnsi="Arial"/></w:rPr><w:t>same</w:t></w:r></w:p>')[0]!;
    const paraUnmodeled = paragraphs('<w:p><w:pPr><w:kinsoku/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>')[0]!;
    expect(runUnmodeled.contentHash).toBe(plain.contentHash);
    expect(runUnmodeled.formatFingerprint).not.toBe(plain.formatFingerprint);
    expect(paraUnmodeled.contentHash).toBe(plain.contentHash);
    expect(paraUnmodeled.formatFingerprint).not.toBe(plain.formatFingerprint);
  });

  test('proofErr does not affect hashes', () => {
    const without = paragraphs(P('spell'))[0]!;
    const withProof = paragraphs('<w:p><w:proofErr w:type="spellStart"/><w:r><w:t>spell</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>')[0]!;
    expect(withProof.contentHash).toBe(without.contentHash);
    expect(withProof.formatFingerprint).toBe(without.formatFingerprint);
  });

  test('hyperlinks, fields, and note references are modeled', () => {
    const rel = '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/"/>';
    const para = readIrDocument(parts(
      '<w:p><w:hyperlink r:id="rId5"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
        '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
        '<w:r><w:footnoteReference w:id="9"/></w:r></w:p>',
      undefined,
      rel,
    )).body.blocks[0] as IrParagraph;
    expect(para.inlines[0]).toMatchObject({ kind: 'hyperlink', target: 'https://example.test/' });
    expect(para.inlines[1]).toMatchObject({ kind: 'fieldRun', instruction: ' PAGE ', isSimpleField: true });
    expect(para.inlines[2]).toMatchObject({ kind: 'noteRef', noteKind: 'Footnote', noteId: '9' });
  });
});

describe('xelement attribute iteration', () => {
  test('attributes() yields correctly split namespace-qualified names', () => {
    // Regression: the iterator split map keys on NUL while the map was
    // space-keyed — every canonicalized attribute name was garbage.
    const root = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="00AB12" plain="x"/>',
    );
    const names = [...root.attributes()].map(([name]) => `${name.ns}|${name.local}`);
    expect(names).toContain(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main|rsidR',
    );
    expect(names).toContain('|plain');
  });
});
