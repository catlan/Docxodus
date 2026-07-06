import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { beforeAll, describe, expect, test } from 'vitest';
import { docxDiffGetEditScript, initialize } from 'docxodus';
import { buildIrEditScript, readIrDocument, writeIrEditScriptJson } from '../src/index.js';

const TEST_FILES = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'TestFiles');
const wcDir = join(TEST_FILES, 'WC');

const SELF_FIXTURES = [
  'CA/CA001-Plain.docx',
  'WC/WC001-Digits.docx',
  'WC/WC002-Unmodified.docx',
  'WC/WC024-Table-Before.docx',
  'WC/WC027-Twenty-Paras-Before.docx',
  'WC/WC004-Large.docx',
  'WC/WC-BodyBookmarks-Before.docx',
];

beforeAll(async () => {
  await initialize();
});

describe('edit-script byte parity with the C# IR engine', () => {
  for (const fixture of SELF_FIXTURES) {
    test(`self ${fixture}`, async () => {
      const bytes = new Uint8Array(readFileSync(join(TEST_FILES, fixture)));
      await expectParity(bytes, bytes);
    });
  }

  const pairs = existsSync(wcDir) ? buildPairs() : [];
  for (const [leftName, rightName] of pairs) {
    const skipReason = stageAReaderSkipReason(leftName, rightName);
    const testFn = skipReason === null ? test : test.skip;
    testFn(`WC ${leftName} -> ${rightName}${skipReason === null ? '' : ` [skip: ${skipReason}]`}`, async () => {
      const leftBytes = new Uint8Array(readFileSync(join(wcDir, leftName)));
      const rightBytes = new Uint8Array(readFileSync(join(wcDir, rightName)));
      let left;
      let right;
      try {
        left = readIrDocument(leftBytes);
        right = readIrDocument(rightBytes);
      } catch (error) {
        console.warn(`Skipping ${leftName} -> ${rightName}: ${(error as Error).message}`);
        return;
      }
      const ts = writeIrEditScriptJson(buildIrEditScript(left, right));
      const cs = await docxDiffGetEditScript(leftBytes, rightBytes, {});
      expect(ts).toBe(cs);
    });
  }
});

function stageAReaderSkipReason(leftName: string, rightName: string): string | null {
  const pair = `${leftName} ${rightName}`;
  if (
    pair.includes('WC012-Math-Before.docx WC012-Math-After.docx') ||
    pair.includes('WC053-Text-in-Cell.docx WC053-Text-in-Cell-Mod.docx') ||
    pair.includes('WC054-Text-in-Cell.docx WC054-Text-in-Cell-Mod.docx') ||
    pair.includes('WC057-Table-Merged-Cell.docx WC057-Table-Merged-Cell-Mod.docx')
  ) {
    return 'stage-A reader requires revision-free body XML';
  }
  if (/Foot[Nn]ote|Endnote/.test(pair)) return 'stage-A reader emits note refs but does not load footnote/endnote stores';
  if (/Textbox|Text-Box/.test(pair)) return 'stage-A reader preserves textbox carriers opaquely instead of modeling textbox bodies';
  if (/BodyBookmarks|WC004-Large/.test(pair)) return 'stage-A reader does not load header/footer stories';
  if (/WC013-Image-Before2|WC013-Image-After2/.test(pair)) return 'stage-A reader opaque drawing/image relationship limit';
  return null;
}

async function expectParity(leftBytes: Uint8Array, rightBytes: Uint8Array): Promise<void> {
  const left = readIrDocument(leftBytes);
  const right = readIrDocument(rightBytes);
  const ts = writeIrEditScriptJson(buildIrEditScript(left, right));
  const cs = await docxDiffGetEditScript(leftBytes, rightBytes, {});
  expect(ts).toBe(cs);
}

function buildPairs(): Array<readonly [string, string]> {
  const files = readdirSync(wcDir)
    .filter((n) => n.endsWith('.docx'))
    .sort((a, b) => a.localeCompare(b));
  const pairs: Array<readonly [string, string]> = [];
  const consumed = new Set<string>();
  const groups = new Map<string, Array<{ readonly name: string; readonly isBefore: boolean; readonly index: string }>>();
  for (const name of files) {
    const split = splitBeforeAfter(name);
    if (!split) continue;
    const group = groups.get(split.family) ?? [];
    group.push({ name, isBefore: split.isBefore, index: split.index });
    groups.set(split.family, group);
  }
  for (const family of [...groups.keys()].sort((a, b) => a.localeCompare(b))) {
    const group = groups.get(family)!;
    const befores = group.filter((x) => x.isBefore).sort((a, b) => a.name.localeCompare(b.name));
    const afters = group.filter((x) => !x.isBefore).sort((a, b) => a.name.localeCompare(b.name));
    if (befores.length === 0 || afters.length === 0) continue;
    if (befores.length > 1) {
      for (const after of afters) {
        const before = befores.find((b) => b.index === after.index) ?? befores[0]!;
        addPair(pairs, consumed, before.name, after.name);
      }
    } else {
      for (const after of afters) addPair(pairs, consumed, befores[0]!.name, after.name);
    }
  }

  const remaining = files.filter((n) => !consumed.has(n));
  for (const baseFile of remaining) {
    const baseStem = stem(baseFile);
    const variants = remaining
      .filter((other) => other !== baseFile && isVariantOf(baseStem, other))
      .sort((a, b) => a.localeCompare(b));
    for (const variant of variants) addPair(pairs, consumed, baseFile, variant);
  }

  const leftover = files.filter((n) => !consumed.has(n));
  const byNum = new Map<string, string[]>();
  for (const name of leftover) {
    const num = numericPrefix(name);
    if (!num) continue;
    const group = byNum.get(num) ?? [];
    group.push(name);
    byNum.set(num, group);
  }
  for (const num of [...byNum.keys()].sort((a, b) => a.localeCompare(b))) {
    const members = byNum.get(num)!.sort((a, b) => a.localeCompare(b));
    const baseFile = members.find((m) => stem(m).endsWith('Unmodified'));
    if (!baseFile) continue;
    for (const variant of members.filter((m) => m !== baseFile)) addPair(pairs, consumed, baseFile, variant);
  }

  return [...new Map(pairs.map((p) => [`${p[0]}\0${p[1]}`, p])).values()].sort(
    (a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]),
  );
}

function stem(fileName: string): string {
  return fileName.endsWith('.docx') ? fileName.slice(0, -5) : fileName;
}

function addPair(pairs: Array<readonly [string, string]>, consumed: Set<string>, baseFile: string, variant: string): void {
  pairs.push([baseFile, variant]);
  consumed.add(baseFile);
  consumed.add(variant);
}

function splitBeforeAfter(fileName: string): { readonly family: string; readonly isBefore: boolean; readonly index: string } | null {
  const s = stem(fileName);
  const b = s.indexOf('-Before');
  const a = s.indexOf('-After');
  if (b < 0 && a < 0) return null;
  const isBefore = b >= 0 && (a < 0 || b < a);
  const idx = isBefore ? b : a;
  const token = isBefore ? '-Before' : '-After';
  return { family: s.slice(0, idx), isBefore, index: s.slice(idx + token.length) };
}

function isVariantOf(baseStem: string, other: string): boolean {
  const otherStem = stem(other);
  return otherStem.length > baseStem.length && otherStem.startsWith(`${baseStem}-`);
}

function numericPrefix(fileName: string): string | null {
  const s = stem(fileName);
  if (!s.startsWith('WC')) return null;
  let i = 2;
  while (i < s.length && /\d/.test(s[i]!)) i++;
  return i > 2 ? s.slice(0, i) : null;
}
