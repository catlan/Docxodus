// Synthesize corpus fixtures for renderer paths no organic WC document exercises,
// so the differential parity gates (edit-script JSON / revisions / markup part-map)
// witness them against the C# oracle:
//
//   WC-InsHyperlink-Before/After   — an INSERTED paragraph carrying a hyperlink whose
//                                    relationship id exists only in the RIGHT package
//                                    (external-relationship recreation path)
//   WC-InsHeader-Before/After      — a RIGHT-only default header story
//                                    (insertHeaderFooterStory: fresh part, reference attach)
//   WC-InsEvenHeader-Before/After  — a RIGHT-only EVEN header story; the BEFORE package
//                                    has NO settings part (settings wiring + w:evenAndOddHeaders
//                                    + CT_Settings ordering path)
//   WC-InsFirstHeader-Before/After — a RIGHT-only FIRST-page header story
//                                    (w:titlePg insertion in CT_SectPr order)
//   WC-InsStyledPara-Before/After  — an inserted paragraph naming a RIGHT-only style
//                                    (CopyMissingStyles carry)
//   WC-InsNumbered-Before/After    — an inserted numbered paragraph; the BEFORE package has
//                                    NO numbering part (CopyMissingNumbering carry incl.
//                                    fresh-part creation + package wiring)
//
// Deterministic output (fixed zip mtime). Re-run: node ts/scripts/make-parity-fixtures.mjs
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const WC = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'TestFiles', 'WC');
const base = unzipSync(new Uint8Array(readFileSync(join(WC, 'WC001-Digits.docx'))));

const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

const HDR_NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

function clone(parts) {
  return Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, v.slice()]));
}

function edit(parts, name, fn) {
  parts[name] = strToU8(fn(strFromU8(parts[name])));
}

function save(parts, name) {
  writeFileSync(join(WC, name), zipSync(parts, { mtime: new Date('1980-01-01T00:00:00Z') }));
  console.log(`wrote ${name}`);
}

// --- 1. right-only hyperlink relationship -------------------------------------------------
{
  const before = clone(base);
  save(before, 'WC-InsHyperlink-Before.docx');

  const after = clone(base);
  edit(after, 'word/document.xml', (xml) => xml.replace(
    '<w:sectPr ',
    '<w:p><w:hyperlink r:id="rId100" w:history="1"><w:r><w:t>Inserted example link</w:t></w:r></w:hyperlink></w:p><w:sectPr ',
  ));
  edit(after, 'word/_rels/document.xml.rels', (xml) => xml.replace(
    '</Relationships>',
    '<Relationship Id="rId100" Type="' + REL_HYPERLINK + '" Target="http://example.com/synth" TargetMode="External"/></Relationships>',
  ));
  save(after, 'WC-InsHyperlink-After.docx');
}

// --- 2. right-only DEFAULT header story ---------------------------------------------------
{
  const before = clone(base);
  save(before, 'WC-InsHeader-Before.docx');

  const after = clone(base);
  after['word/header1.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${HDR_NS}><w:p><w:r><w:t>Inserted default header</w:t></w:r></w:p></w:hdr>`,
  );
  edit(after, 'word/_rels/document.xml.rels', (xml) => xml.replace(
    '</Relationships>',
    '<Relationship Id="rId101" Type="' + REL_HEADER + '" Target="header1.xml"/></Relationships>',
  ));
  edit(after, '[Content_Types].xml', (xml) => xml.replace(
    '</Types>',
    '<Override PartName="/word/header1.xml" ContentType="' + CT_HEADER + '"/></Types>',
  ));
  edit(after, 'word/document.xml', (xml) => xml.replace(
    /(<w:sectPr[^>]*>)/,
    '$1<w:headerReference w:type="default" r:id="rId101"/>',
  ));
  save(after, 'WC-InsHeader-After.docx');
}

// --- 3. right-only EVEN header story; BEFORE has NO settings part -------------------------
{
  const before = clone(base);
  delete before['word/settings.xml'];
  edit(before, 'word/_rels/document.xml.rels', (xml) =>
    xml.replace(/<Relationship [^>]*Target="settings\.xml"[^>]*\/>/, ''));
  edit(before, '[Content_Types].xml', (xml) =>
    xml.replace(/<Override PartName="\/word\/settings\.xml"[^>]*\/>/, ''));
  save(before, 'WC-InsEvenHeader-Before.docx');

  const after = clone(base);
  after['word/header1.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${HDR_NS}><w:p><w:r><w:t>Inserted even header</w:t></w:r></w:p></w:hdr>`,
  );
  edit(after, 'word/_rels/document.xml.rels', (xml) => xml.replace(
    '</Relationships>',
    '<Relationship Id="rId101" Type="' + REL_HEADER + '" Target="header1.xml"/></Relationships>',
  ));
  edit(after, '[Content_Types].xml', (xml) => xml.replace(
    '</Types>',
    '<Override PartName="/word/header1.xml" ContentType="' + CT_HEADER + '"/></Types>',
  ));
  edit(after, 'word/document.xml', (xml) => xml.replace(
    /(<w:sectPr[^>]*>)/,
    '$1<w:headerReference w:type="even" r:id="rId101"/>',
  ));
  // The AFTER document opts into even/odd headers (required for the even story to be live).
  edit(after, 'word/settings.xml', (xml) => xml.replace(
    /(<w:settings [^>]*>)/,
    '$1<w:evenAndOddHeaders/>',
  ));
  save(after, 'WC-InsEvenHeader-After.docx');
}

// --- 4. right-only FIRST-page header story (w:titlePg insertion) --------------------------
{
  const before = clone(base);
  save(before, 'WC-InsFirstHeader-Before.docx');

  const after = clone(base);
  after['word/header1.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${HDR_NS}><w:p><w:r><w:t>Inserted first-page header</w:t></w:r></w:p></w:hdr>`,
  );
  edit(after, 'word/_rels/document.xml.rels', (xml) => xml.replace(
    '</Relationships>',
    '<Relationship Id="rId101" Type="' + REL_HEADER + '" Target="header1.xml"/></Relationships>',
  ));
  edit(after, '[Content_Types].xml', (xml) => xml.replace(
    '</Types>',
    '<Override PartName="/word/header1.xml" ContentType="' + CT_HEADER + '"/></Types>',
  ));
  // titlePg makes the first-page story live; the sectPr already ends with w:docGrid, which
  // FOLLOWS titlePg in CT_SectPr — the renderer must insert the flag BEFORE it.
  edit(after, 'word/document.xml', (xml) => xml
    .replace(/(<w:sectPr[^>]*>)/, '$1<w:headerReference w:type="first" r:id="rId101"/>')
    .replace('<w:docGrid', '<w:titlePg/><w:docGrid'));
  save(after, 'WC-InsFirstHeader-After.docx');
}

// --- 5. inserted paragraph naming a RIGHT-only style (styles carry) -----------------------
{
  const before = clone(base);
  save(before, 'WC-InsStyledPara-Before.docx');

  const after = clone(base);
  edit(after, 'word/styles.xml', (xml) => xml.replace(
    '</w:styles>',
    '<w:style w:type="paragraph" w:styleId="SynthEmphatic"><w:name w:val="Synth Emphatic"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:i/></w:rPr></w:style></w:styles>',
  ));
  edit(after, 'word/document.xml', (xml) => xml.replace(
    '<w:sectPr ',
    '<w:p><w:pPr><w:pStyle w:val="SynthEmphatic"/></w:pPr><w:r><w:t>Inserted styled paragraph</w:t></w:r></w:p><w:sectPr ',
  ));
  save(after, 'WC-InsStyledPara-After.docx');
}

// --- 6. inserted numbered paragraph; BEFORE has NO numbering part (numbering carry) -------
{
  const before = clone(base);
  save(before, 'WC-InsNumbered-Before.docx');

  const after = clone(base);
  after['word/numbering.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:numbering ${HDR_NS}><w:abstractNum w:abstractNumId="7"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="42"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
  );
  edit(after, 'word/_rels/document.xml.rels', (xml) => xml.replace(
    '</Relationships>',
    '<Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
  ));
  edit(after, '[Content_Types].xml', (xml) => xml.replace(
    '</Types>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
  ));
  edit(after, 'word/document.xml', (xml) => xml.replace(
    '<w:sectPr ',
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/></w:numPr></w:pPr><w:r><w:t>Inserted numbered item</w:t></w:r></w:p><w:sectPr ',
  ));
  save(after, 'WC-InsNumbered-After.docx');
}
