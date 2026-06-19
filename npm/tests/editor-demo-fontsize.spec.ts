import { test, expect } from '@playwright/test';

// Issue 8 — arbitrary font size (demo). The size control was a <select> capped at 48pt; it is
// now an editable combobox accepting any positive value. Driving the real demo (editor.html):
// typing 72 sets a 72pt run (~96px), well beyond the old cap.
test.describe('Demo — arbitrary font size', () => {
  test('typing 72 in the size field sets a 72pt run (beyond the old 48 cap)', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    // Type "BIG" into the first paragraph.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.keyboard.type('BIG');

    // Select the paragraph, then set the size field to 72 and apply.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.fill('#fontsize', '72');
    await page.locator('#fontsize').dispatchEvent('change');

    const px = await page.evaluate(() => {
      const span = document.querySelector('#editor p[data-anchor] span') as HTMLElement | null;
      return span ? parseFloat(getComputedStyle(span).fontSize) : 0;
    });
    // 72pt ≈ 96px — comfortably beyond the old 48pt (64px) ceiling.
    expect(px).toBeGreaterThan(90);
  });
});
