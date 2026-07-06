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

describe('native markup renderer full-part parity with C# oracle', () => {
  for (const fixture of SELF_FIXTURES) {
    const divergent = KNOWN_DIVERGENT.has(`self ${fixture}`);
    test(`self ${fixture}${divergent ? ' [known divergent]' : ''}`, async () => {
      const bytes = new Uint8Array(readFileSync(join(TEST_FILES, fixture)));
      await expectPartMapParity(bytes, bytes, divergent);
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
        await expectPartMapParity(leftBytes, rightBytes, divergent);
      },
    );
  }
});

function markupSkipReason(leftName: string, rightName: string): string | null {
  void leftName;
  void rightName;
  return null;
}

// Divergence RATCHET: the native renderer is converging on byte parity
// fixture by fixture. Pairs here are KNOWN divergent — the test asserts
// the divergence still exists, so a FIX flips the test loudly and the
// pair must be removed from this list (never grows, only shrinks).
const KNOWN_DIVERGENT = new Set<string>([]);


async function expectPartMapParity(
  leftBytes: Uint8Array,
  rightBytes: Uint8Array,
  expectDivergent = false,
): Promise<void> {
  const left = readIrDocument(leftBytes);
  const right = readIrDocument(rightBytes);
  const script = buildIrEditScript(left, right);
  const tsParts = renderIrMarkup(leftBytes, rightBytes, script);
  const oracle = await docxDiffCompare(leftBytes, rightBytes, {});
  const leftParts = new Map(Object.entries(unzipSync(leftBytes)));
  const oracleParts = new Map(Object.entries(unzipSync(oracle)));
  const rewritten = new Set([...oracleRewrittenPartNames(leftParts, oracleParts)].map(normalizeGeneratedPartName));
  const tsSnapshot = normalizedPartSnapshot(tsParts, rewritten);
  const oracleSnapshot = normalizedPartSnapshot(oracleParts, rewritten);
  if (expectDivergent) {
    // Ratchet: this pair is known divergent. When a renderer fix makes it
    // match, this assertion fails loudly — REMOVE the pair from
    // KNOWN_DIVERGENT to lock the win in.
    expect(tsSnapshot, 'pair now matches — remove it from KNOWN_DIVERGENT').not.toEqual(oracleSnapshot);
    return;
  }
  expect(tsSnapshot).toEqual(oracleSnapshot);
}

function normalizedPartSnapshot(parts: ReadonlyMap<string, Uint8Array>, names: ReadonlySet<string>): Record<string, string> {
  const entries: Array<readonly [string, string]> = [];
  for (const [name, bytes] of parts) {
    if (name.endsWith('/')) continue;
    if (name === '_rels/.rels') continue;
    const normalizedName = normalizeGeneratedPartName(name);
    if (!names.has(normalizedName)) continue;
    const value = isXmlPart(name)
      ? normalizeGeneratedIds(strFromU8(bytes))
      : `base64:${Buffer.from(bytes).toString('base64')}`;
    entries.push([normalizedName, value]);
  }
  entries.sort((a, b) => a[0].localeCompare(b[0]));
  return Object.fromEntries(entries);
}

function oracleRewrittenPartNames(leftParts: ReadonlyMap<string, Uint8Array>, oracleParts: ReadonlyMap<string, Uint8Array>): Set<string> {
  const names = new Set<string>();
  for (const [name, oracleBytes] of oracleParts) {
    if (name.endsWith('/') || name === '_rels/.rels') continue;
    if (!isMarkupRendererOwnedPart(name)) continue;
    const leftBytes = leftParts.get(name);
    if (!leftBytes || !sameBytes(leftBytes, oracleBytes)) names.add(name);
  }
  for (const name of leftParts.keys()) if (!oracleParts.has(name)) names.add(name);
  return names;
}

function isMarkupRendererOwnedPart(name: string): boolean {
  return name === '[Content_Types].xml' ||
    name === 'word/document.xml' ||
    name === 'word/_rels/document.xml.rels' ||
    name === 'word/footnotes.xml' ||
    name === 'word/endnotes.xml' ||
    name === 'word/settings.xml' ||
    /^word\/(?:header|footer)\d+\.xml$/i.test(name) ||
    /^word\/_rels\/(?:header|footer)\d+\.xml\.rels$/i.test(name) ||
    /^word\/(?:footnotes|endnotes)\.xml$/i.test(name) ||
    /^word\/(?:media|diagrams)\//i.test(name) ||
    /^word\/(?:media|diagrams)\/_rels\//i.test(name);
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function isXmlPart(name: string): boolean {
  return name.endsWith('.xml') || name.endsWith('.rels') || name === '[Content_Types].xml';
}

function normalizeGeneratedPartName(name: string): string {
  return name.replace(/P[0-9a-f]{32}(?=\.)/gi, 'PGEN');
}

function normalizeGeneratedIds(xml: string): string {
  const normalized = xml
    .replace(/R[0-9a-f]{16,32}/gi, 'RGEN')
    .replace(/P[0-9a-f]{32}(?=\.)/gi, 'PGEN');
  return normalized;
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
