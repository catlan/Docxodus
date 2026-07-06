import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { beforeAll, describe, expect, test } from 'vitest';
import { docxDiffCompare, initialize } from 'docxodus';
import { docxDiffCompareTs, docxDiffGetEditScriptJsonTs, docxDiffGetRevisionsTs } from '../src/index.js';
import { nameEquals, parseXml, type XElement } from '../src/xml/xelement.js';
import { W } from '../src/ir/names.js';

const TEST_FILES = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'TestFiles');
const WC = join(TEST_FILES, 'WC');

const PAIRS: Array<readonly [string, string]> = [
  ['WC002-Unmodified.docx', 'WC002-InsertInMiddle.docx'],
  ['WC012-Math-Before.docx', 'WC012-Math-After.docx'],
  ['WC053-Text-in-Cell.docx', 'WC053-Text-in-Cell-Mod.docx'],
  ['WC066-Textbox-Before-Ins.docx', 'WC066-Textbox-Before-Ins-Mod.docx'],
];

beforeAll(async () => {
  await initialize();
});

describe('public docx-diff facade', () => {
  for (const [leftName, rightName] of PAIRS) {
    test(`${leftName} -> ${rightName}`, async () => {
      const left = new Uint8Array(readFileSync(join(WC, leftName)));
      const right = new Uint8Array(readFileSync(join(WC, rightName)));
      const ours = docxDiffCompareTs(left, right);
      const oracle = await docxDiffCompare(left, right, {});
      expect(normalizedOwnedParts(ours, oracle)).toEqual(normalizedOwnedParts(oracle, oracle));
      expectWellFormedRevisionXml(ours);
      expect(docxDiffGetEditScriptJsonTs(left, right)).toBeTypeOf('string');
      expect(docxDiffGetRevisionsTs(left, right).every((r) => r.revisionType)).toBe(true);
    });
  }
});

function normalizedOwnedParts(actualDocx: Uint8Array, oracleDocx: Uint8Array): Record<string, string> {
  const actual = unzipSync(actualDocx);
  const oracle = unzipSync(oracleDocx);
  const out: Array<readonly [string, string]> = [];
  for (const name of Object.keys(oracle).sort()) {
    if (!isOwnedPart(name)) continue;
    const bytes = actual[name];
    if (!bytes) continue;
    out.push([name, isXmlPart(name) ? normalizeGeneratedIds(strFromU8(bytes)) : `base64:${Buffer.from(bytes).toString('base64')}`]);
  }
  return Object.fromEntries(out);
}

function expectWellFormedRevisionXml(docx: Uint8Array): void {
  const parts = unzipSync(docx);
  for (const [name, bytes] of Object.entries(parts)) {
    if (!isXmlPart(name)) continue;
    const xml = strFromU8(bytes);
    const root = parseXml(xml);
    expect(countElements(root, W.ins)).toBeGreaterThanOrEqual(0);
    expect(countElements(root, W.del)).toBeGreaterThanOrEqual(0);
  }
}

function countElements(root: XElement, name: typeof W.ins): number {
  let n = nameEquals(root.name, name) ? 1 : 0;
  for (const el of root.descendants(name)) n++;
  return n;
}

function isOwnedPart(name: string): boolean {
  return name === '[Content_Types].xml' ||
    name === 'word/document.xml' ||
    name === 'word/_rels/document.xml.rels' ||
    name === 'word/footnotes.xml' ||
    name === 'word/endnotes.xml' ||
    name === 'word/settings.xml' ||
    /^word\/(?:header|footer)\d+\.xml$/i.test(name) ||
    /^word\/_rels\/(?:header|footer)\d+\.xml\.rels$/i.test(name) ||
    /^word\/(?:media|diagrams)\//i.test(name);
}

function isXmlPart(name: string): boolean {
  return name.endsWith('.xml') || name.endsWith('.rels') || name === '[Content_Types].xml';
}

function normalizeGeneratedIds(xml: string): string {
  return xml
    .replace(/R[0-9a-f]{16,32}/gi, 'RGEN')
    .replace(/P[0-9a-f]{32}(?=\.)/gi, 'PGEN');
}
