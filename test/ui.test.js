'use strict';
/**
 * Browser tests: the real Index.html and Code.js, through the local preview
 * server, in headless Chromium. Skipped when Playwright isn't installed
 * (npm install --no-save playwright && npx playwright install chromium).
 */
const test = require('node:test');
const assert = require('node:assert/strict');

let chromium = null;
try {
  ({ chromium } = require('playwright'));
} catch (err) {
  chromium = null;
}

process.env.DELAY = '0';
const skip = chromium ? false : 'Playwright is not installed';

test.describe('web app in a browser', { skip }, () => {
  let preview;
  let browser;
  let base;

  test.before(async () => {
    preview = require('../dev/preview');
    await new Promise((resolve) => preview.server.listen(0, resolve));
    base = 'http://localhost:' + preview.server.address().port;
    browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  });

  test.after(async () => {
    if (browser) await browser.close();
    if (preview) preview.server.close();
  });

  async function open(path, viewport) {
    const context = await browser.newContext({ viewport: viewport || { width: 1280, height: 900 } });
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    await page.goto(base + path);
    page.errors = errors;
    return page;
  }

  const serverPart = (name) => {
    const app = preview.getApp();
    return app.run('apiGetData', {}).parts.find((p) => p.name === name);
  };
  const row = (page, id) => page.locator('.row[data-id="' + id + '"]');

  test('lists the parts with their status and totals', async () => {
    preview.reset();
    const page = await open('/');
    await page.waitForSelector('.row');
    assert.equal(await page.locator('.row').count(), 12);
    assert.deepEqual(await page.locator('.stat .stat-num').allTextContents(), ['12', '4', '1', '1']);
    assert.match(await row(page, 'P-0004').textContent(), /V-belt A42\s*Out/);
    assert.equal(await row(page, 'P-0002').locator('.badge.low').count(), 1);
    assert.equal(await row(page, 'P-0006').locator('.badge.ordered').count(), 1);
    assert.equal(await row(page, 'P-0001').locator('a.icon-btn').getAttribute('href'), 'https://www.mcmaster.com/91292A112/');

    await page.click('.stat.low');
    assert.equal(await page.locator('.row').count(), 4);
    await page.click('.stat.low');
    await page.fill('#search', 'grainger belt');
    await page.waitForTimeout(200);
    assert.deepEqual(await page.locator('.row .name').allTextContents(), ['V-belt A42']);
    await page.fill('#search', '');
    await page.selectOption('#category-filter', 'Electrical');
    await page.waitForFunction(() => document.querySelectorAll('.row').length === 3);
    assert.deepEqual(await page.locator('.row .name').allTextContents(),
      ['5 A fast-acting fuse, 5 × 20 mm', '14 AWG THHN wire, black', 'Cable ties, 8 in, black UV']);
    assert.deepEqual(page.errors, []);
  });

  test('tapping − several times sends one change and emails when it runs low', async () => {
    preview.reset();
    const page = await open('/');
    await page.waitForSelector('.row');
    const bearing = row(page, 'P-0003');
    for (let i = 0; i < 3; i++) await bearing.locator('.step').first().click();
    assert.equal((await bearing.locator('.qty strong').textContent()).trim(), '3', 'updates on screen straight away');
    await page.waitForSelector('.toast:has-text("Low-stock email sent")', { timeout: 5000 });
    assert.equal(serverPart('6203-2RS sealed ball bearing').quantity, 3);
    const app = preview.getApp();
    assert.equal(app.env.sent.length, 1);
    assert.match(app.env.sent[0].subject, /Low stock: 6203-2RS sealed ball bearing \(3 ea left\)/);
    const log = app.run('apiGetActivity', {}, { partId: 'P-0003' }).entries;
    assert.deepEqual([log[0].action, log[1].action, log[1].change], ['Alert emailed', 'Used', -3]);
    assert.equal(await row(page, 'P-0003').locator('.badge.low').count(), 1);
    assert.equal(await page.locator('.stat.low .stat-num').textContent(), '5');
    assert.deepEqual(page.errors, []);
  });

  test('adding, editing and deleting a part', async () => {
    preview.reset();
    const page = await open('/');
    await page.waitForSelector('.row');
    await page.click('#add-btn');
    await page.click('#part-save');
    assert.match(await page.textContent('#part-form-error'), /Give the part a name/);
    await page.fill('#f-name', 'Air compressor inline filter');
    await page.fill('#f-partNumber', '91292A112');
    assert.match(await page.textContent('#dup-hint'), /M3 × 8 mm socket head cap screw \(P-0001\) already has this part number/);
    await page.fill('#f-partNumber', 'ACF-38');
    await page.fill('#f-location', 'Shelf D-1');
    await page.fill('#f-quantity', '4');
    await page.fill('#f-minQty', '2');
    await page.fill('#f-link', 'not a link');
    await page.click('#part-save');
    await page.waitForSelector('#part-form-error:has-text("doesn\'t look like a web link")');
    await page.fill('#f-link', 'www.grainger.com/product/ACF38');
    await page.click('#part-save');
    await page.waitForSelector('.toast:has-text("Added Air compressor inline filter (P-0013)")');
    assert.equal(await page.locator('#part-dialog').evaluate((d) => d.open), false);
    const added = serverPart('Air compressor inline filter');
    assert.deepEqual([added.id, added.location, added.quantity, added.minQty, added.links[0]],
      ['P-0013', 'Shelf D-1', 4, 2, 'https://www.grainger.com/product/ACF38']);

    await row(page, 'P-0013').locator('.row-main').click();
    await page.click('#drawer-panel button:has-text("Edit")');
    await page.fill('#f-location', 'Shelf D-2');
    await page.click('#part-save');
    await page.waitForSelector('.toast:has-text("Saved.")');
    assert.equal(serverPart('Air compressor inline filter').location, 'Shelf D-2');
    const edited = preview.getApp().run('apiGetActivity', {}, { partId: 'P-0013' }).entries[0];
    assert.equal(edited.note, 'Changed Location', 'only the changed field was sent');

    await page.click('#drawer-panel button:has-text("Delete")');
    await page.click('#confirm-ok');
    await page.waitForSelector('.toast:has-text("Deleted Air compressor inline filter")');
    assert.equal(serverPart('Air compressor inline filter'), undefined);
    assert.equal(await page.locator('#drawer').isHidden(), true);
    assert.deepEqual(page.errors, []);
  });

  test('the part panel: restock, set count and mark as ordered', async () => {
    preview.reset();
    const page = await open('/?part=P-0002');
    await page.waitForSelector('#drawer-panel .stock-card');
    assert.equal(await page.textContent('#drawer-title'), 'M3 hex nut, stainless');
    assert.equal(await page.locator('#drawer-panel a.btn.primary').getAttribute('href'), 'https://www.mcmaster.com/94150A325/');
    await page.click('#drawer-panel button:has-text("Mark as ordered")');
    await page.waitForSelector('.ordered-note');
    assert.equal(serverPart('M3 hex nut, stainless').ordered, true);

    await page.click('.segmented button:has-text("Restock")');
    await page.fill('#adjust-amount', '200');
    await page.fill('#adjust-note', 'PO 5521');
    assert.match(await page.textContent('.adjust-preview'), /New quantity: 238/);
    await page.click('#drawer-panel button:has-text("Apply")');
    await page.waitForSelector('.toast:has-text("Restocked 200")');
    const nut = serverPart('M3 hex nut, stainless');
    assert.deepEqual([nut.quantity, nut.ordered, nut.alerted], [238, false, false], 'restocking clears the order and alert marks');
    await page.waitForSelector('#part-history .act');
    assert.match(await page.textContent('#part-history'), /Restocked\s*\+200/);
    assert.match(await page.textContent('#part-history'), /PO 5521/);

    await page.click('.segmented button:has-text("Set count")');
    await page.fill('#adjust-amount', '230');
    await page.click('#drawer-panel button:has-text("Apply")');
    await page.waitForSelector('.toast:has-text("Count set to 230")');
    await page.keyboard.press('Escape');
    assert.equal(await page.locator('#drawer').isHidden(), true);
    assert.equal((await row(page, 'P-0002').locator('.qty strong').textContent()).trim(), '230');
    assert.deepEqual(page.errors, []);
  });

  test('an access code keeps visitors out until they enter it', async () => {
    preview.reset({ code: 'bolt-7731' });
    const page = await open('/?as=visitor');
    await page.waitForSelector('#code-input');
    assert.equal(await page.locator('#app').isHidden(), true);
    await page.fill('#code-input', 'wrong-code');
    await page.click('.screen-card button[type=submit]');
    await page.waitForSelector('.error-text:has-text("isn\'t right")');
    await page.fill('#code-input', 'bolt-7731');
    await page.click('.screen-card button[type=submit]');
    await page.waitForSelector('.row');
    assert.equal(await page.locator('.row').count(), 12);

    await page.reload();
    await page.waitForSelector('.row');
    assert.equal(await page.locator('#code-input').count(), 0, 'the code is remembered on this device');

    await page.click('.banner button:has-text("Add my name")');
    await page.fill('#name-input', 'Maria');
    await page.click('#name-form button[type=submit]');
    await row(page, 'P-0001').locator('.step').last().click();
    await page.waitForFunction(() => !document.querySelector('.qty.saving'));
    await page.waitForTimeout(100);
    const entry = preview.getApp().run('apiGetActivity', {}, { partId: 'P-0001' }).entries[0];
    assert.deepEqual([entry.action, entry.by], ['Restocked', 'Maria']);

    await page.click('.tab[data-view="settings"]');
    assert.equal(await page.locator('#s-recipients').isDisabled(), true, 'visitors can\'t change settings');
    assert.match(await page.textContent('#settings'), /Only shop@example\.com can change these/);
    assert.equal(await page.locator('#s-code').inputValue(), '', 'the code isn\'t revealed');
    assert.deepEqual(page.errors, []);
  });

  test('the owner can change settings and send a test email', async () => {
    preview.reset();
    const page = await open('/');
    await page.click('.tab[data-view="settings"]');
    await page.fill('#s-recipients', 'buyer@example.com, not-an-email');
    await page.click('#settings button:has-text("Save settings")');
    await page.waitForSelector('.toast.error:has-text("isn\'t a valid email")');
    await page.fill('#s-recipients', 'buyer@example.com');
    await page.selectOption('#s-reminder', 'Daily');
    await page.selectOption('#s-hour', '7');
    await page.fill('#s-appname', 'Shop 2 Parts');
    await page.click('#settings button:has-text("Save settings")');
    await page.waitForSelector('.toast:has-text("Settings saved")');
    assert.equal(await page.textContent('#app-name'), 'Shop 2 Parts');
    const settings = preview.getApp().run('apiGetData', {}).settings;
    assert.deepEqual([settings.recipients, settings.reminder, settings.reminderHour, settings.appName],
      [['buyer@example.com'], 'Daily', 7, 'Shop 2 Parts']);
    await page.click('#settings button:has-text("Send test email")');
    await page.waitForSelector('.toast:has-text("Test email sent to buyer@example.com")');
    assert.match(preview.getApp().env.sent.pop().subject, /\[Shop 2 Parts\] Test/);
    assert.deepEqual(page.errors, []);
  });

  test('works on a phone, in the spreadsheet sidebar and with no parts', async () => {
    preview.reset();
    let page = await open('/', { width: 375, height: 740 });
    await page.waitForSelector('.row');
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(overflow <= 0, 'no sideways scrolling on a phone');
    await row(page, 'P-0003').locator('.row-main').click();
    await page.waitForSelector('#drawer-panel .stock-card');
    const panel = await page.locator('#drawer-panel').boundingBox();
    assert.equal(Math.round(panel.width), 375, 'the part panel fills the screen');

    page = await open('/sidebar');
    const frame = page.frameLocator('iframe');
    await frame.locator('.row').first().waitFor();
    assert.equal(await frame.locator('.row').count(), 12);

    preview.reset({ empty: true });
    page = await open('/');
    await page.waitForSelector('.empty');
    assert.match(await page.textContent('.empty'), /No parts yet/);
    await page.click('.empty button:has-text("Add a part")');
    assert.equal(await page.locator('#part-dialog').evaluate((d) => d.open), true);
    assert.deepEqual(page.errors, []);
  });
});
