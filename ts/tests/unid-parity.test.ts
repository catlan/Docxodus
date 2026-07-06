// THE load-bearing v2.1 gate: the TS deterministic-unid computation must
// reproduce the C# engine's unids EXACTLY — anchors are the join key of
// the whole differential-testing strategy (byte-comparable edit-script
// JSON between the C# WASM engine and this port).
//
// Method: a DocxDiff self-compare emits an EqualBlock op per body block,
// whose anchors carry the C# IR's unids. The TS side unzips the same
// document, parses word/document.xml, assigns deterministic unids, and
// the body-level w:p/w:tbl unids must match the script's block-op
// anchors in order.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { beforeAll, describe, expect, test } from 'vitest';
import { initialize, docxDiffGetEditScript } from 'docxodus';

import { parseXml } from '../src/xml/xelement.js';
import { W } from '../src/ir/names.js';
import { assignToAllElementsDeterministic } from '../src/ir/unid-helper.js';

const TEST_FILES = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  'TestFiles',
);

// Revision-free fixtures of increasing structure (plain paragraphs,
// digits, tables). Revision-bearing docs need the reader's
// revision-view pass first — later phase.
const FIXTURES = [
  'CA/CA001-Plain.docx',
  'WC/WC001-Digits.docx',
  'WC/WC002-Unmodified.docx',
  'WC/WC024-Table-Before.docx',
  'WC/WC027-Twenty-Paras-Before.docx',
];

const tsBodyBlockUnids = (docx: Uint8Array): string[] => {
  const xml = strFromU8(unzipSync(docx)['word/document.xml']!);
  const root = parseXml(xml);
  const unids = assignToAllElementsDeterministic(root);
  const body = root.element(W.body);
  if (!body) throw new Error('no w:body');
  const result: string[] = [];
  for (const block of body.elements()) {
    if (
      (block.name.local === 'p' || block.name.local === 'tbl') &&
      block.name.ns === W.p.ns
    ) {
      result.push(unids.get(block)!);
    }
  }
  return result;
};

const BLOCK_KINDS = new Set(['p', 'h', 'li', 'tbl']);

const csBodyBlockUnids = async (docx: Uint8Array): Promise<string[]> => {
  const json = await docxDiffGetEditScript(docx, docx, {});
  const script = JSON.parse(json) as {
    operations?: Array<{ kind: string; leftAnchor?: string }>;
  };
  const result: string[] = [];
  for (const op of script.operations ?? []) {
    const anchor = op.leftAnchor;
    if (!anchor) continue;
    const [kind, scope, unid] = anchor.split(':');
    if (scope === 'body' && kind && unid && BLOCK_KINDS.has(kind)) {
      result.push(unid);
    }
  }
  return result;
};

beforeAll(async () => {
  await initialize();
});

describe('unid parity with the C# IR engine', () => {
  for (const fixture of FIXTURES) {
    test(fixture, async () => {
      const docx = new Uint8Array(readFileSync(join(TEST_FILES, fixture)));
      const cs = await csBodyBlockUnids(docx);
      expect(cs.length).toBeGreaterThan(0);
      const ts = tsBodyBlockUnids(docx);
      expect(ts).toEqual(cs);
    });
  }
});
