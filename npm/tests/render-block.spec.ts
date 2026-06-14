import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Browser-level proof of the feasibility gate (docs/architecture/ir_editor_feasibility.md):
// across the real WASM boundary, RenderBlockHtml(anchor) must match the full
// render's data-anchor element — same tag, same text. This is the editor's
// incremental per-block re-render path running in the actual runtime.
test.describe('RenderBlockHtml (WASM bridge) — incremental block render', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('single-block render matches the full render per anchor', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const dc = (window as any).Docxodus.DocumentConverter;

      // Full render WITH stampAnchors (final arg) — the editor's initial render.
      const fullHtml: string = dc.ConvertDocxToHtmlComplete(
        bin, 'Document', 'docx-', false, '', -1, 'comment-', 0, 1.0, 'page-',
        false, 0, 'annot-', false, false, false, true, true, false, null, /*stampAnchors*/ true);

      const norm = (s: string) => s.replace(/\s+/g, ' ').trim();
      const doc = new DOMParser().parseFromString(fullHtml, 'text/html');
      const candidates = Array.from(doc.querySelectorAll('[data-anchor]'))
        .filter((e) =>
          /^(P|H1|H2|H3|H4|H5|H6)$/.test(e.tagName) &&
          norm(e.textContent || '').length > 0 &&
          e.querySelector('img') === null)
        .slice(0, 8);

      const checks = candidates.map((el) => {
        const unid = el.getAttribute('data-anchor') as string;
        // The editor passes the bare data-anchor straight back; RenderBlockHtml
        // keys on the unid tail so a bare unid is accepted.
        const blockHtml: string = dc.RenderBlockHtml(bin, unid, 'docx-', false);
        const bel = new DOMParser().parseFromString(blockHtml, 'text/html').body.firstElementChild;
        return {
          fullTag: el.tagName,
          fullText: norm(el.textContent || ''),
          blockTag: bel ? bel.tagName : '(none)',
          blockText: bel ? norm(bel.textContent || '') : '(none)',
        };
      });

      return { count: candidates.length, checks };
    }, Array.from(bytes));

    expect(result.count).toBeGreaterThan(0);
    for (const c of result.checks) {
      expect(c.blockTag).toBe(c.fullTag);
      expect(c.blockText).toBe(c.fullText);
    }
  });
});
