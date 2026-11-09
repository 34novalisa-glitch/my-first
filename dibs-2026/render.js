// Export PNGs: node render.js  → export/post-1080x1350.png, export/story-1080x1920.png
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const browser = await chromium.launch();
  const file = 'file://' + path.join(__dirname, 'post.html');
  for (const [name, q, h] of [['post-1080x1350', '', 1350], ['story-1080x1920', '?story', 1920]]) {
    const page = await browser.newPage({ viewport: { width: 1080, height: h }, deviceScaleFactor: 1 });
    await page.goto(file + q);
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(300);
    await page.locator('#canvas').screenshot({ path: path.join(__dirname, 'export', name + '.png') });
    await page.close();
  }
  await browser.close();
})();
