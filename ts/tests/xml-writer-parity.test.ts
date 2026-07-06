import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { beforeAll, describe, expect, test } from 'vitest';
import { docxDiffCompare, initialize } from 'docxodus';
import { parseXml, readIrDocument, writeXmlPart } from '../src/index.js';

const TEST_FILES = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'TestFiles');
const wcDir = join(TEST_FILES, 'WC');
const identityFixtures = [
  ...(existsSync(wcDir) ? readdirSync(wcDir).filter((n) => n.endsWith('.docx')).map((n) => join(wcDir, n)) : []),
  join(TEST_FILES, 'CA', 'CA001-Plain.docx'),
].sort((a, b) => a.localeCompare(b));

const redlinePairs = buildPairs().slice(0, 10);

beforeAll(async () => {
  await initialize();
});

describe('XML writer byte parity', () => {
  test('retainSources exposes parsed part roots and opt-in source elements', () => {
    const bytes = new Uint8Array(readFileSync(join(TEST_FILES, 'CA', 'CA001-Plain.docx')));
    const retained = readIrDocument(bytes, { retainSources: true });
    const defaultDoc = readIrDocument(bytes);
    expect(defaultDoc.parsedPartRoots).toBeUndefined();
    expect(retained.parsedPartRoots?.get('/word/document.xml')?.name.local).toBe('document');
    expect(retained.body.blocks[0]?.source.element?.name.local).toBe('p');
    expect(retained.body.blocks[0]?.source.partUri).toBe('/word/document.xml');
    expect(defaultDoc.body.blocks[0]?.source.element).toBeNull();
  });

  for (const fixture of identityFixtures) {
    const label = relative(TEST_FILES, fixture);
    test(`identity write ${label}`, () => {
      const zip = unzipSync(new Uint8Array(readFileSync(fixture)));
      assertPartRoundTrips(zip, 'word/document.xml', label);
      if (zip['word/styles.xml']) assertPartRoundTrips(zip, 'word/styles.xml', label);
    });
  }

  for (const [leftName, rightName] of redlinePairs) {
    test(`redline write ${leftName} -> ${rightName}`, async () => {
      const leftBytes = new Uint8Array(readFileSync(join(wcDir, leftName)));
      const rightBytes = new Uint8Array(readFileSync(join(wcDir, rightName)));
      const redline = await docxDiffCompare(leftBytes, rightBytes, {});
      const zip = unzipSync(redline);
      assertPartRoundTrips(zip, 'word/document.xml', `${leftName} -> ${rightName}`);
    });
  }
});

function assertPartRoundTrips(zip: Record<string, Uint8Array>, partName: string, label: string): void {
  const bytes = zip[partName];
  expect(bytes, `${label} missing ${partName}`).toBeTruthy();
  const original = strFromU8(bytes!);
  const actual = writeXmlPart(parseXml(original));
  if (actual !== original) throw new Error(formatMismatch(label, partName, original, actual));
}

function formatMismatch(label: string, partName: string, expected: string, actual: string): string {
  const index = firstDiff(expected, actual);
  const start = Math.max(0, index - 60);
  const end = Math.min(Math.max(expected.length, actual.length), index + 120);
  return [
    `${label} ${partName} XML writer mismatch at byte/char ${index}`,
    `expected length ${expected.length}, actual length ${actual.length}`,
    `expected: ${visible(expected.slice(start, end))}`,
    `actual:   ${visible(actual.slice(start, end))}`,
  ].join('\n');
}

function firstDiff(left: string, right: string): number {
  const len = Math.min(left.length, right.length);
  for (let i = 0; i < len; i++) if (left.charCodeAt(i) !== right.charCodeAt(i)) return i;
  return len;
}

function visible(s: string): string {
  return s.replaceAll('\r', '\\r').replaceAll('\n', '\\n').replaceAll('\t', '\\t');
}

function buildPairs(): Array<readonly [string, string]> {
  if (!existsSync(wcDir)) return [];
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
    for (const after of afters) {
      const before = befores.find((b) => b.index === after.index) ?? befores[0]!;
      addPair(pairs, consumed, before.name, after.name);
    }
  }
  return [...new Map(pairs.map((p) => [`${p[0]}\0${p[1]}`, p])).values()].sort(
    (a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]),
  );
}

function addPair(pairs: Array<readonly [string, string]>, consumed: Set<string>, baseFile: string, variant: string): void {
  pairs.push([baseFile, variant]);
  consumed.add(baseFile);
  consumed.add(variant);
}

function splitBeforeAfter(fileName: string): { readonly family: string; readonly isBefore: boolean; readonly index: string } | null {
  const s = fileName.endsWith('.docx') ? fileName.slice(0, -5) : fileName;
  const b = s.indexOf('-Before');
  const a = s.indexOf('-After');
  if (b < 0 && a < 0) return null;
  const isBefore = b >= 0 && (a < 0 || b < a);
  const idx = isBefore ? b : a;
  const token = isBefore ? '-Before' : '-After';
  return { family: s.slice(0, idx), isBefore, index: s.slice(idx + token.length) };
}
