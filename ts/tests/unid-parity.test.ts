// THE load-bearing v2.1 gate: the TS deterministic-unid computation must
// reproduce the C# engine's unids EXACTLY — anchors are the join key of
// the whole differential-testing strategy (byte-comparable edit-script
// JSON between the C# WASM engine and this port).
//
// Method: a DocxDiff self-compare emits an EqualBlock op per body block,
// whose anchors carry the C# IR's full `kind:scope:unid` strings. The TS
// side reads the same document through the Stage-A reader and the
// body-level w:p/w:tbl anchors must match in order, including h/li kind
// classification.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { beforeAll, describe, expect, test } from 'vitest';
import { initialize, docxDiffGetEditScript } from 'docxodus';

import { anchorToString, readIrDocument } from '../src/index.js';

const TEST_FILES = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  'TestFiles',
);

// Revision-free fixtures of increasing structure (plain paragraphs,
// digits, tables, headings, and list items). Revision-bearing docs need
// the reader's revision-view pass first — later phase.
const FIXTURES = [
  'CA/CA001-Plain.docx',
  'WC/WC001-Digits.docx',
  'WC/WC002-Unmodified.docx',
  'WC/WC024-Table-Before.docx',
  'WC/WC027-Twenty-Paras-Before.docx',
  'WC/WC004-Large.docx',
  'WC/WC-BodyBookmarks-Before.docx',
];

const tsBodyBlockAnchors = (docx: Uint8Array): string[] =>
  readIrDocument(docx).body.blocks
    .filter((block) => BLOCK_KINDS.has(block.anchor.kind))
    .map((block) => anchorToString(block.anchor));

const BLOCK_KINDS = new Set(['p', 'h', 'li', 'tbl']);

const csBodyBlockAnchors = async (docx: Uint8Array): Promise<string[]> => {
  const json = await docxDiffGetEditScript(docx, docx, {});
  const script = JSON.parse(json) as {
    operations?: Array<{ kind: string; leftAnchor?: string }>;
  };
  const result: string[] = [];
  for (const op of script.operations ?? []) {
    const anchor = op.leftAnchor;
    if (!anchor) continue;
    const [kind, scope, unid] = anchor.split(':');
    if (scope === 'body' && kind && unid && BLOCK_KINDS.has(kind)) result.push(anchor);
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
      const cs = await csBodyBlockAnchors(docx);
      expect(cs.length).toBeGreaterThan(0);
      const ts = tsBodyBlockAnchors(docx);
      expect(ts).toEqual(cs);
    });
  }
});
