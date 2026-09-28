#!/usr/bin/env node
'use strict';
/**
 * Regenerates the screenshots in docs/ from the local preview.
 * Needs Playwright: npm install --no-save playwright && npx playwright install chromium
 */
const path = require('path');

let chromium;
try {
  ({ chromium } = require('playwright'));
} catch (err) {
  console.error('Playwright is not installed. Run: npm install --no-save playwright && npx playwright install chromium');
  process.exit(1);
}

process.env.DELAY = '0';
const preview = require('./preview');
const OUT = path.join(__dirname, '..', 'docs');

(async () => {
  await new Promise((resolve) => preview.server.listen(0, resolve));
  const base = 'http://localhost:' + preview.server.address().port;
  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  const shot = async (url, file, viewport, prepare, options) => {
    const context = await browser.newContext({ viewport, deviceScaleFactor: (options && options.scale) || 1 });
    const page = await context.newPage();
    await page.goto(base + url);
    if (prepare) await prepare(page);
    await page.waitForTimeout(400);
    await page.screenshot(Object.assign({ path: path.join(OUT, file) }, options && options.screenshot));
    await context.close();
    console.log('docs/' + file);
  };

  await shot('/', 'app-desktop.png', { width: 1280, height: 780 }, (p) => p.waitForSelector('.row'));
  await shot('/', 'app-phone.png', { width: 390, height: 780 }, (p) => p.waitForSelector('.row'), { scale: 2 });
  await shot('/?part=P-0004', 'part-phone.png', { width: 390, height: 780 }, async (p) => {
    await p.waitForSelector('#part-history .act');
    await p.evaluate(() => document.activeElement.blur());
  }, { scale: 2 });

  // A part runs low: the email that goes out.
  const app = preview.getApp();
  app.context.doGet({ parameter: {} });
  app.env.sent.length = 0;
  app.run('apiAdjustStock', {}, { id: 'P-0003', mode: 'remove', amount: 3 });
  await shot('/emails/0', 'email.png', { width: 680, height: 420 }, null, { scale: 2, screenshot: { fullPage: true } });

  await browser.close();
  preview.server.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
