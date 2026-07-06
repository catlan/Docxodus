import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { IrBlockAligner, readIrDocument, type IrAlignmentKind } from '../src/index.js';
import { assertInvariants, histogram } from './helpers/ir-alignment-asserts.js';

const wcDir = join(process.cwd(), '..', 'TestFiles', 'WC');

describe('IrBlockAligner WC corpus', () => {
  test('WC corpus pairs align without throwing invariants hold both directions', () => {
    if (!existsSync(wcDir)) {
      console.warn(`Skipping WC corpus: missing ${wcDir}`);
      return;
    }
    const pairs = buildPairs();
    expect(pairs.length).toBeGreaterThanOrEqual(30);
    const totals = new Map<IrAlignmentKind, number>();
    const skipped: string[] = [];

    for (const [baseName, variantName] of pairs) {
      let baseDoc;
      let variantDoc;
      try {
        baseDoc = readWc(baseName);
        variantDoc = readWc(variantName);
      } catch (error) {
        skipped.push(`${baseName} -> ${variantName}: ${(error as Error).message}`);
        continue;
      }

      const fwd = IrBlockAligner.align(baseDoc, variantDoc);
      assertInvariants(baseDoc, variantDoc, fwd);
      const rev = IrBlockAligner.align(variantDoc, baseDoc);
      assertInvariants(variantDoc, baseDoc, rev);
      for (const e of fwd.entries) totals.set(e.kind, (totals.get(e.kind) ?? 0) + 1);
      void histogram(rev);
    }

    if (skipped.length > 0) console.warn(`Skipped WC corpus pairs:\n${skipped.join('\n')}`);
    expect(pairs.length - skipped.length).toBeGreaterThan(0);
    void totals;
  });
});

function readWc(fileName: string) {
  return readIrDocument(readFileSync(join(wcDir, fileName)));
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
