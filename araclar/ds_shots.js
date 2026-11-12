// Deterministic renderer: seeks all CSS animations to a given time, then screenshots.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const SEEK = `window.__anims = window.__anims || document.getAnimations();
window.__seek = (t) => { for (const a of window.__anims) { a.pause(); a.currentTime = t; } };`;

(async () => {
  const [mode, ...args] = process.argv.slice(2);
  const browser = await chromium.launch();
  if (mode === 'illust') {
    // args: json file with [{html, png, w, h, t}]
    const jobs = JSON.parse(fs.readFileSync(args[0], 'utf8'));
    for (const j of jobs) {
      const ctx = await browser.newContext({ viewport: { width: j.w, height: j.h }, deviceScaleFactor: j.scale || 2 });
      const page = await ctx.newPage();
      await page.goto('file://' + path.resolve(j.html));
      await page.evaluate(() => document.fonts.ready);
      await page.evaluate(SEEK);
      await page.evaluate((t) => window.__seek(t), j.t);
      await page.screenshot({ path: j.png, omitBackground: !j.opaque });
      await ctx.close();
    }
  } else if (mode === 'frames') {
    const [html, outdir, dur, fps] = [args[0], args[1], parseFloat(args[2]), parseInt(args[3])];
    fs.mkdirSync(outdir, { recursive: true });
    const ctx = await browser.newContext({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
    const page = await ctx.newPage();
    await page.goto('file://' + path.resolve(html));
    await page.evaluate(() => document.fonts.ready);
    await page.evaluate(SEEK);
    const n = Math.round(dur * fps);
    for (let i = 0; i < n; i++) {
      await page.evaluate((t) => window.__seek(t), (i / fps) * 1000);
      await page.screenshot({ path: path.join(outdir, String(i).padStart(5, '0') + '.png'), omitBackground: true });
    }
    await ctx.close();
  }
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
