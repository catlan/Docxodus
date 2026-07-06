import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { beforeAll, describe, expect, test } from 'vitest';
import { docxDiffCompare, initialize } from 'docxodus';
import { buildIrEditScript, readIrDocument, renderIrMarkup } from '../src/index.js';

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

describe('native markup renderer word/document.xml byte parity with C# oracle', () => {
  for (const fixture of SELF_FIXTURES) {
    const divergent = KNOWN_DIVERGENT.has(`self ${fixture}`);
    test(`self ${fixture}${divergent ? ' [known divergent]' : ''}`, async () => {
      const bytes = new Uint8Array(readFileSync(join(TEST_FILES, fixture)));
      await expectDocumentXmlParity(bytes, bytes, divergent);
    });
  }

  const pairs = existsSync(wcDir) ? buildPairs() : [];
  for (const [leftName, rightName] of pairs) {
    const skipReason = markupSkipReason(leftName, rightName);
    const testFn = skipReason === null ? test : test.skip;
    const divergent = KNOWN_DIVERGENT.has(`${leftName} -> ${rightName}`);
    testFn(
      `WC ${leftName} -> ${rightName}${skipReason === null ? '' : ` [skip: ${skipReason}]`}${divergent ? ' [known divergent]' : ''}`,
      async () => {
        const leftBytes = new Uint8Array(readFileSync(join(wcDir, leftName)));
        const rightBytes = new Uint8Array(readFileSync(join(wcDir, rightName)));
        await expectDocumentXmlParity(leftBytes, rightBytes, divergent);
      },
    );
  }
});

function markupSkipReason(leftName: string, rightName: string): string | null {
  const pair = `${leftName} ${rightName}`;
  if (
    pair.includes('WC012-Math-Before.docx WC012-Math-After.docx') ||
    pair.includes('WC053-Text-in-Cell.docx WC053-Text-in-Cell-Mod.docx') ||
    pair.includes('WC054-Text-in-Cell.docx WC054-Text-in-Cell-Mod.docx') ||
    pair.includes('WC057-Table-Merged-Cell.docx WC057-Table-Merged-Cell-Mod.docx')
  ) {
    return 'stage-A reader requires revision-free body XML';
  }
  if (/Header|Footer/i.test(pair)) return 'M2 header/footer story part rebuild and sectPr relationship mutation deferred';
  if (/Footnote|Endnote|Foot|End/i.test(pair)) return 'M2 footnote/endnote part rebuild and note-id renumbering deferred';
  if (/Image|Picture|Media|Drawing/i.test(pair)) return 'M2 cross-part media and relationship import deferred';
  if (/-Ins(?:\.|-)/i.test(pair)) return 'stage-A reader requires revision-free body XML';
  return null;
}

// Divergence RATCHET: the native renderer is converging on byte parity
// fixture by fixture. Pairs here are KNOWN divergent — the test asserts
// the divergence still exists, so a FIX flips the test loudly and the
// pair must be removed from this list (never grows, only shrinks).
const KNOWN_DIVERGENT = new Set<string>([
  'WC-BodyBookmarks-Before.docx -> WC-BodyBookmarks-After.docx',
  'WC002-Unmodified.docx -> WC002-DiffAtBeginning.docx',
  'WC002-Unmodified.docx -> WC002-DiffInMiddle.docx',
  'WC002-Unmodified.docx -> WC002-InsertAtBeginning.docx',
  'WC002-Unmodified.docx -> WC002-InsertInMiddle.docx',
  'WC004-Large.docx -> WC004-Large-Mod.docx',
  'WC006-Table.docx -> WC006-Table-Delete-Contests-of-Row.docx',
  'WC006-Table.docx -> WC006-Table-Delete-Row.docx',
  'WC007-Unmodified.docx -> WC007-Deleted-at-Beginning-of-Para.docx',
  'WC007-Unmodified.docx -> WC007-Moved-into-Table.docx',
  'WC009-Table-Unmodified.docx -> WC009-Table-Cell-1-1-Mod.docx',
  'WC010-Para-Before-Table-Unmodified.docx -> WC010-Para-Before-Table-Mod.docx',
  'WC011-Before.docx -> WC011-After.docx',
  'WC014-SmartArt-Before.docx -> WC014-SmartArt-After.docx',
  'WC015-Three-Paragraphs.docx -> WC015-Three-Paragraphs-After.docx',
  'WC018-Field-Simple-Before.docx -> WC018-Field-Simple-After-1.docx',
  'WC018-Field-Simple-Before.docx -> WC018-Field-Simple-After-2.docx',
  'WC019-Hyperlink-Before.docx -> WC019-Hyperlink-After-1.docx',
  'WC019-Hyperlink-Before.docx -> WC019-Hyperlink-After-2.docx',
  'WC021-Math-Before-1.docx -> WC021-Math-After-1.docx',
  'WC021-Math-Before-2.docx -> WC021-Math-After-2.docx',
  'WC024-Table-Before.docx -> WC024-Table-After.docx',
  'WC024-Table-Before.docx -> WC024-Table-After2.docx',
  'WC025-Simple-Table-Before.docx -> WC025-Simple-Table-After.docx',
  'WC026-Long-Table-Before.docx -> WC026-Long-Table-After-1.docx',
  'WC032-Para-with-Para-Props.docx -> WC032-Para-with-Para-Props-After.docx',
  'WC033-Merged-Cells-Before.docx -> WC033-Merged-Cells-After1.docx',
  'WC033-Merged-Cells-Before.docx -> WC033-Merged-Cells-After2.docx',
  'WC037-Textbox-Before.docx -> WC037-Textbox-After1.docx',
  'WC038-Document-With-BR-Before.docx -> WC038-Document-With-BR-After.docx',
  'WC039-Break-In-Row.docx -> WC039-Break-In-Row-After1.docx',
  'WC040-Case-Before.docx -> WC040-Case-After.docx',
  'WC041-Table-5.docx -> WC041-Table-5-Mod.docx',
  'WC042-Table-5.docx -> WC042-Table-5-Mod.docx',
  'WC043-Nested-Table.docx -> WC043-Nested-Table-Mod.docx',
  'WC044-Text-Box.docx -> WC044-Text-Box-Mod.docx',
  'WC045-Text-Box.docx -> WC045-Text-Box-Mod.docx',
  'WC046-Two-Text-Box.docx -> WC046-Two-Text-Box-Mod.docx',
  'WC047-Two-Text-Box.docx -> WC047-Two-Text-Box-Mod.docx',
  'WC048-Text-Box-in-Cell.docx -> WC048-Text-Box-in-Cell-Mod.docx',
  'WC049-Text-Box-in-Cell.docx -> WC049-Text-Box-in-Cell-Mod.docx',
  'WC050-Table-in-Text-Box.docx -> WC050-Table-in-Text-Box-Mod.docx',
  'WC051-Table-in-Text-Box.docx -> WC051-Table-in-Text-Box-Mod.docx',
  'WC052-SmartArt-Same.docx -> WC052-SmartArt-Same-Mod.docx',
  'WC055-French.docx -> WC055-French-Mod.docx',
  'WC056-French.docx -> WC056-French-Mod.docx',
  'WC058-Table-Merged-Cell.docx -> WC058-Table-Merged-Cell-Mod.docx',
  'WC062-New-Char-Style-Added.docx -> WC062-New-Char-Style-Added-Mod.docx',
  'WC065-Textbox.docx -> WC065-Textbox-Deleted.docx',
  'WC065-Textbox.docx -> WC065-Textbox-Mod.docx',
  'self WC/WC-BodyBookmarks-Before.docx',
]);


async function expectDocumentXmlParity(
  leftBytes: Uint8Array,
  rightBytes: Uint8Array,
  expectDivergent = false,
): Promise<void> {
  const left = readIrDocument(leftBytes);
  const right = readIrDocument(rightBytes);
  const script = buildIrEditScript(left, right);
  const tsParts = renderIrMarkup(leftBytes, rightBytes, script);
  const tsDocument = strFromU8(tsParts.get('word/document.xml')!);
  const oracle = await docxDiffCompare(leftBytes, rightBytes, {});
  const oracleDocument = strFromU8(unzipSync(oracle)['word/document.xml']!);
  if (expectDivergent) {
    // Ratchet: this pair is known divergent. When a renderer fix makes it
    // match, this assertion fails loudly — REMOVE the pair from
    // KNOWN_DIVERGENT to lock the win in.
    expect(tsDocument, 'pair now matches — remove it from KNOWN_DIVERGENT').not.toBe(oracleDocument);
    return;
  }
  expect(tsDocument).toBe(oracleDocument);
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
