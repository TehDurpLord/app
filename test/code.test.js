'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp } = require('./gas-fakes');
const { sheetSpec, isUpToDate, OUT } = require('../dev/build-sheet');
const { unzip } = require('../dev/xlsx');

const OWNER = 'owner@example.com';
const LAYOUT = ['Part Name', 'Part Number', 'Location', 'Quantity', 'Min Qty', 'Order Link', 'Ordered', 'Notes', 'Alert Sent'];
const REORDER_FORMULA = '=IFNA(SORT(FILTER(Inventory!A:F,((Inventory!A:A<>"")+(Inventory!B:B<>""))>0,' +
  'ISNUMBER(Inventory!D:D),ISNUMBER(Inventory!E:E),Inventory!D:D<=Inventory!E:E,Inventory!G:G<>TRUE,' +
  'Inventory!G:G<>"Yes"),4,TRUE),"Nothing needs reordering right now.")';

/**
 * A spreadsheet that has been through setup, with the setup email cleared.
 * The clock is 6 AM on a Monday in New York, before the 8 AM reminder, so
 * only the reminder tests get reminders.
 */
function freshApp(opts) {
  const app = createApp(Object.assign({ now: '2026-09-28T10:00:00Z' }, opts));
  app.run('setup');
  app.env.sent.length = 0;
  return app;
}

function inventory(app) {
  return app.sheet('Inventory');
}

function col(app, header) {
  const i = inventory(app).dump()[0].indexOf(header);
  if (i === -1) throw new Error('No column ' + header);
  return i + 1;
}

/** What Sheets does when someone edits cells by hand: the change, then the edit trigger. */
function edit(app, range, oldValue) {
  app.context.onInventoryEdit({ range, oldValue, user: { getEmail: () => '' } });
}

/** Types a new row into the Inventory tab. values: { header: value } */
function addPart(app, values) {
  const sheet = inventory(app);
  const headers = sheet.dump()[0];
  const row = sheet.getLastRow() + 1;
  const range = sheet.getRange(row, 1, 1, headers.length);
  range.setValues([headers.map((h) => (h in values ? values[h] : ''))]);
  edit(app, range);
  return row;
}

function setCell(app, row, header, value) {
  const range = inventory(app).getRange(row, col(app, header));
  const old = range.getValue();
  range.setValue(value);
  edit(app, range, old);
}

function cell(app, row, header) {
  return inventory(app).getRange(row, col(app, header)).getValue();
}

function settingRow(app, label) {
  const rows = app.sheet('Settings').dump();
  const i = rows.findIndex((r) => r[0] === label);
  if (i === -1) throw new Error('No setting ' + label);
  return i + 1;
}

function setting(app, label) {
  return app.sheet('Settings').dump()[settingRow(app, label) - 1][1];
}

function changeSetting(app, label, value) {
  app.sheet('Settings').getRange(settingRow(app, label), 2).setValue(value);
}

function status(app) {
  const row = app.sheet('Settings').dump()[settingRow(app, 'Status') - 1];
  return { value: row[1], help: row[2] };
}

const BOLT = { 'Part Name': 'Hex bolt M8', 'Part Number': 'HB-8', Location: 'Bin 3', Quantity: 10, 'Min Qty': 5,
  'Order Link': 'https://www.mcmaster.com/91292A112' };

// ---------------------------------------------------------------------------

test('setup builds the three tabs on a blank spreadsheet, emails a test, and can run again', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' });
  const report = app.run('setup');
  assert.deepEqual(app.ss.getSheets().map((s) => s.getName()), ['Inventory', 'Reorder', 'Settings'],
    'the blank Sheet1 becomes Inventory');

  const sheet = inventory(app);
  assert.deepEqual(sheet.dump()[0], LAYOUT);
  assert.equal(sheet.getFrozenRows(), 1);
  assert.equal(sheet.getConditionalFormatRules().length, 2);
  assert.equal(sheet.getRange(2, 7).getDataValidation().type, 'CHECKBOX', 'Ordered is tick boxes');
  assert.equal(sheet.getRange(500, 7).getDataValidation().type, 'CHECKBOX');
  assert.equal(sheet.getRange(2, 4).getDataValidation().type, 'NUMBER_GTE');
  assert.match(sheet.getRange(1, 5).getNote(), /drops to this number or below/);
  assert.equal(sheet.getLastRow(), 1, 'no example rows');

  const reorder = app.sheet('Reorder');
  assert.deepEqual(reorder.dump()[1], ['Part Name', 'Part Number', 'Location', 'Quantity', 'Min Qty', 'Order Link']);
  assert.match(reorder.dump()[0][0], /^Parts at or below their Min Qty/);
  assert.equal(reorder.getRange(3, 1).getFormula(), REORDER_FORMULA);
  assert.equal(reorder.getFrozenRows(), 2);

  assert.deepEqual(status(app).value, 'On');
  assert.match(status(app).help, /sent from owner@example\.com/);
  assert.equal(setting(app, 'Send emails to'), OWNER);
  assert.equal(setting(app, 'Email right away when a part runs low'), 'On');
  assert.equal(setting(app, 'Reminder email'), 'Weekdays');
  assert.equal(setting(app, 'Reminder hour (0-23)'), 8);

  assert.deepEqual(app.env.triggers.map((t) => t.fn + ':' + t.type).sort(), ['hourlyCheck:CLOCK', 'onInventoryEdit:ON_EDIT']);
  assert.equal(app.env.sent.length, 1);
  assert.equal(app.env.sent[0].subject, '[Parts Inventory] Low-stock emails are on');
  assert.match(app.env.sent[0].body, /Nothing is low right now\./);
  assert.deepEqual(report.testEmailTo, [OWNER]);
  assert.ok(app.env.logs.some((l) => /Done\. The low-stock emails are on/.test(l[1])), 'the editor log says it worked');

  const settingsRows = app.sheet('Settings').dump().length;
  app.run('setup');
  assert.equal(app.ss.getSheets().length, 3);
  assert.equal(app.env.triggers.length, 2, 'no duplicate triggers');
  assert.equal(sheet.getConditionalFormatRules().length, 2, 'no duplicate highlight rules');
  assert.equal(app.sheet('Settings').dump().length, settingsRows, 'no duplicate settings rows');
  assert.equal(reorder.getRange(3, 1).getFormula(), REORDER_FORMULA);
  assert.equal(app.env.sent.length, 2, 'every setup sends a test email');
});

test('setup finishes the ready-made sheet without losing anything typed into it', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' });
  // What Google Drive makes of template/parts-inventory.xlsx.
  sheetSpec().sheets.forEach((spec, i) => {
    const sheet = i === 0 ? app.ss.getSheets()[0].setName(spec.name) : app.ss.insertSheet(spec.name);
    spec.rows.forEach((row, r) => row.forEach((c, j) => {
      const value = c && typeof c === 'object' ? (c.f !== undefined ? '=' + c.f : c.v) : c;
      if (value === '' || value === null || value === undefined) return;
      sheet.getRange(r + 1, j + 1).setValue(typeof value === 'string' && value[0] !== '=' ? "'" + value : value);
    }));
  });
  assert.equal(status(app).value, 'Off');
  assert.match(status(app).help, /^Not turned on yet/);
  assert.equal(setting(app, 'Send emails to'), '');
  const sheet = inventory(app);
  sheet.getRange(2, 1, 3, 9).setValues([
    ['Hex bolt M8', 'HB-8', 'Bin 3', 2, 10, 'https://www.mcmaster.com/91292A112', '', 'Ask for the zinc ones', ''],
    ['Air filter', 'AF-9', 'Shelf 1', 0, 2, 'https://www.grainger.com/product/1', 'Yes', '', ''],
    ['Gloves', '', '', 50, 10, '', '', '', ''],
  ]);

  const report = app.run('setup');
  assert.deepEqual(app.ss.getSheets().map((s) => s.getName()), ['Inventory', 'Reorder', 'Settings']);
  assert.deepEqual(report.addedColumns, []);
  assert.deepEqual(sheet.dump().slice(0, 4).map((r) => r.slice(0, 8)), [
    LAYOUT.slice(0, 8),
    ['Hex bolt M8', 'HB-8', 'Bin 3', 2, 10, 'https://www.mcmaster.com/91292A112', '', 'Ask for the zinc ones'],
    ['Air filter', 'AF-9', 'Shelf 1', 0, 2, 'https://www.grainger.com/product/1', true, ''],
    ['Gloves', '', '', 50, 10, '', '', ''],
  ], '"Yes" becomes a ticked box and nothing else changes');
  assert.equal(sheet.getRange(3, 7).getDataValidation().type, 'CHECKBOX');
  assert.match(sheet.getRange(1, 5).getNote(), /drops to this number or below/, 'header tips are added');
  assert.equal(app.sheet('Reorder').getRange(3, 1).getFormula(), REORDER_FORMULA);
  assert.equal(setting(app, 'Send emails to'), OWNER, 'filled in with the account that ran setup');
  assert.equal(status(app).value, 'On');

  const mail = app.env.sent[0];
  assert.equal(mail.subject, '[Parts Inventory] Low-stock emails are on');
  assert.match(mail.body, /These parts are low right now:/);
  assert.match(mail.body, /Air filter: OUT OF STOCK \(min 2, ticked as ordered\)[\s\S]*Hex bolt M8: 2 on hand/,
    'out of stock first');
  assert.doesNotMatch(mail.body, /Gloves/);
});

test('a part that runs low is emailed once, with its order link, until it is restocked', () => {
  const app = freshApp();
  const row = addPart(app, BOLT);
  assert.equal(app.env.sent.length, 0);

  setCell(app, row, 'Quantity', 5);
  assert.equal(app.env.sent.length, 1);
  const mail = app.env.sent[0];
  assert.equal(mail.to, OWNER);
  assert.equal(mail.subject, '[Parts Inventory] Low stock: Hex bolt M8 (5 left)');
  assert.match(mail.htmlBody, /href="https:\/\/www\.mcmaster\.com\/91292A112"[^>]*>Order from mcmaster\.com/);
  assert.match(mail.htmlBody, new RegExp('#gid=' + inventory(app).getSheetId() + '&amp;range=A2'), 'links to the row');
  assert.match(mail.body, /Order: https:\/\/www\.mcmaster\.com\/91292A112/);
  assert.ok(cell(app, row, 'Alert Sent') instanceof Date);

  setCell(app, row, 'Quantity', 3);
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1, 'not emailed again while it stays low');

  setCell(app, row, 'Quantity', 20);
  assert.equal(cell(app, row, 'Alert Sent'), '', 'restocking resets it');
  setCell(app, row, 'Quantity', 4);
  assert.equal(app.env.sent.length, 2, 'running low again sends a new email');

  setCell(app, row, 'Min Qty', '');
  assert.equal(cell(app, row, 'Alert Sent'), '', 'no Min Qty means no emails');
  assert.equal(app.env.sent.length, 2);
});

test('several parts in one paste share one email, out of stock first', () => {
  const app = freshApp();
  const sheet = inventory(app);
  const range = sheet.getRange(2, 1, 3, 5);
  range.setValues([
    ['Gasket', '', '', 1, 5],
    ['V-belt', '', '', 0, 2],
    ['Spring', '', '', 40, 5],
  ]);
  edit(app, range);
  assert.equal(app.env.sent.length, 1);
  assert.equal(app.env.sent[0].subject, '[Parts Inventory] Low stock: 2 parts need ordering');
  assert.match(app.env.sent[0].htmlBody, /V-belt[\s\S]*Out of stock[\s\S]*Gasket/);
  assert.match(app.env.sent[0].htmlBody, /No order link saved for this part yet/);
  assert.doesNotMatch(app.env.sent[0].htmlBody, /Spring/);

  const row = addPart(app, { 'Part Name': 'Fuse', Quantity: 3, 'Min Qty': 1 });
  setCell(app, row, 'Quantity', 0);
  assert.equal(app.env.sent[1].subject, '[Parts Inventory] Out of stock: Fuse');
});

test('edits on other tabs and in the header row are ignored, and the hourly check catches the rest', () => {
  const app = freshApp();
  inventory(app).getRange(2, 1, 1, 5).setValues([['Gasket', '', '', 1, 5]]);
  edit(app, app.sheet('Settings').getRange(2, 2));
  edit(app, app.sheet('Reorder').getRange(3, 1));
  edit(app, inventory(app).getRange(1, 1));
  assert.equal(app.env.sent.length, 0);
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1, 'a change the edit check never saw (a formula, a script) still gets emailed');
});

test('ticked parts leave the reminder and untick themselves once restocked', () => {
  const app = freshApp();
  const bolt = addPart(app, Object.assign({}, BOLT, { Quantity: 2 }));
  const filter = addPart(app, { 'Part Name': 'Air filter', Quantity: 0, 'Min Qty': 2, Ordered: 'yes' });
  const tape = addPart(app, { 'Part Name': 'Tape', Quantity: 1, 'Min Qty': 2, Ordered: 'PO 4471' });
  setCell(app, bolt, 'Ordered', true);
  assert.equal(app.env.sent.length, 3, 'ordered or not, the first email still goes out');

  app.run('menuSendReorderList');
  const reminder = app.env.sent[3];
  assert.equal(reminder.subject, '[Parts Inventory] Reminder: 1 part needs ordering');
  assert.match(reminder.body, /Tape: 1 on hand[\s\S]*Already ordered:[\s\S]*Air filter[\s\S]*Hex bolt M8/,
    '"PO 4471" isn\'t a tick, so Tape still needs ordering');

  [bolt, filter, tape].forEach((row) => setCell(app, row, 'Quantity', 50));
  assert.equal(cell(app, bolt, 'Ordered'), false, 'the box unticks itself');
  assert.equal(cell(app, filter, 'Ordered'), '', '"yes" is cleared');
  assert.equal(cell(app, tape, 'Ordered'), 'PO 4471', 'other notes are left alone');
  assert.equal(cell(app, bolt, 'Alert Sent'), '');
});

test('ticking Ordered before a part runs low keeps the tick until it has run low and been restocked', () => {
  const app = freshApp();
  const row = addPart(app, Object.assign({}, BOLT, { Quantity: 6 }));
  setCell(app, row, 'Ordered', true);
  app.run('hourlyCheck');
  setCell(app, row, 'Quantity', 7);
  assert.equal(cell(app, row, 'Ordered'), true, 'ordering ahead is kept');

  setCell(app, row, 'Quantity', 5);
  assert.equal(app.env.sent.length, 1);
  assert.match(app.env.sent[0].htmlBody, /ticked as ordered/);
  setCell(app, row, 'Quantity', 40);
  assert.equal(cell(app, row, 'Ordered'), false, 'unticked once the shortage is over');
});

test('Ordered columns that hold quantities, dates or formulas are left alone', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' });
  const sheet = app.ss.insertSheet('Inventory');
  sheet.getRange(1, 1, 1, 4).setValues([['Part Name', 'Quantity', 'Min Qty', 'Ordered']]);
  sheet.getRange(2, 1, 2, 4).setValues([['Gasket', 1, 5, 120], ['Spring', 1, 5, '=HYPERLINK("https://x.example","Yes")']]);
  app.run('setup');
  assert.equal(sheet.getRange(2, 4).getDataValidation(), null, 'no tick boxes over your numbers');
  assert.equal(sheet.getRange(2, 4).getValue(), 120);

  const formulaApp = freshApp();
  const row = addPart(formulaApp, { 'Part Name': 'Spring', Quantity: 1, 'Min Qty': 5 });
  const orderedCell = inventory(formulaApp).getRange(row, col(formulaApp, 'Ordered'));
  orderedCell.setValue('=HYPERLINK("https://x.example","Yes")');
  setCell(formulaApp, row, 'Quantity', 40);
  assert.equal(orderedCell.getFormula(), '=HYPERLINK("https://x.example","Yes")', 'a formula is never overwritten');
  assert.equal(cell(formulaApp, row, 'Alert Sent'), '');
});

test('a tab called "inventory" in lower case works the same', () => {
  const app = freshApp();
  const sheet = inventory(app).setName('inventory');
  sheet.getRange(2, 1, 1, 5).setValues([['Gasket', '', '', 1, 5]]);
  edit(app, sheet.getRange(2, 4));
  assert.equal(app.env.sent.length, 1);
  app.run('setup');
  assert.deepEqual(app.ss.getSheets().map((s) => s.getName()), ['inventory', 'Reorder', 'Settings']);
});

test('a reminder time typed as "2 PM" is read as 14', () => {
  const app = freshApp();
  // Sheets stores a typed time as a time value, which Apps Script hands over as a Date.
  changeSetting(app, 'Reminder hour (0-23)', new Date('2026-09-28T18:00:00Z')); // 2 PM in New York
  assert.equal(app.context.getSettings_(app.ss).reminderHour, 14);
});

test('the reminder goes out once a day at the chosen hour, weekdays only by default', () => {
  const app = freshApp();
  changeSetting(app, 'Email right away when a part runs low', 'Off');
  addPart(app, { 'Part Name': 'Gloves', Quantity: 1, 'Min Qty': 5 });
  const tape = addPart(app, { 'Part Name': 'Tape', Quantity: 0, 'Min Qty': 2 });
  setCell(app, tape, 'Ordered', true);
  const reminders = () => app.env.sent.filter((m) => /Reminder/.test(m.subject));
  assert.equal(app.env.sent.length, 0, 'instant emails are off');

  app.setTime('2026-09-28T11:30:00Z'); // Monday 7:30 AM in New York
  app.run('hourlyCheck');
  assert.equal(reminders().length, 0, 'too early');

  app.setTime('2026-09-28T12:10:00Z'); // 8:10 AM
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1);
  assert.equal(reminders()[0].subject, '[Parts Inventory] Reminder: 1 part needs ordering');
  assert.match(reminders()[0].htmlBody, /Gloves[\s\S]*Already ordered[\s\S]*Tape/);

  app.setTime('2026-09-28T15:10:00Z');
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1, 'once a day');

  app.setTime('2026-10-03T13:00:00Z'); // Saturday
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1, 'not on weekends');

  changeSetting(app, 'Reminder email', 'Daily');
  changeSetting(app, 'Reminder hour (0-23)', 6);
  app.run('hourlyCheck');
  assert.equal(reminders().length, 2, 'Daily includes Saturday');

  changeSetting(app, 'Reminder email', 'Weekly');
  app.setTime('2026-10-04T13:00:00Z'); // Sunday
  app.run('hourlyCheck');
  app.setTime('2026-10-05T13:00:00Z'); // Monday
  app.run('hourlyCheck');
  assert.equal(reminders().length, 3, 'Weekly is Mondays');

  changeSetting(app, 'Reminder email', 'Off');
  app.setTime('2026-10-06T13:00:00Z');
  app.run('hourlyCheck');
  assert.equal(reminders().length, 3, 'Off means off');
  assert.equal(app.env.sent.length, 3);
});

test('nothing goes out with instant emails off or no one listed, and failures show in Settings', () => {
  const app = freshApp();
  changeSetting(app, 'Email right away when a part runs low', 'Off');
  const row = addPart(app, { 'Part Name': 'Gasket', Quantity: 1, 'Min Qty': 5 });
  assert.equal(app.env.sent.length, 0);
  assert.equal(cell(app, row, 'Alert Sent'), '', 'not marked, so it goes out once turned back on');
  changeSetting(app, 'Email right away when a part runs low', 'On');

  changeSetting(app, 'Send emails to', '');
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 0);
  changeSetting(app, 'Send emails to', 'buyer@example.com, ' + OWNER);

  app.env.mailFails = 'Service unavailable: Gmail';
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 0);
  assert.equal(cell(app, row, 'Alert Sent'), '');
  assert.equal(status(app).value, 'Problem');
  assert.match(status(app).help, /couldn't be sent: Service unavailable: Gmail\. It will be tried again within the hour\./);

  app.env.mailFails = null;
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1);
  assert.equal(app.env.sent[0].to, 'buyer@example.com,' + OWNER);
  assert.equal(status(app).value, 'On', 'back to normal after the next email');

  changeSetting(app, 'Send emails to', '');
  addPart(app, { 'Part Name': 'Fuse', Quantity: 0, 'Min Qty': 1 });
  assert.equal(status(app).value, 'Problem');
  assert.match(status(app).help, /no one is listed next to "Send emails to"/);
  changeSetting(app, 'Send emails to', OWNER);
  app.run('menuSendTestEmail');
  assert.equal(status(app).value, 'On', 'a test email that goes through clears it');

  inventory(app).getRange(1, col(app, 'Alert Sent')).setValue('');
  addPart(app, { 'Part Name': 'Belt', Quantity: 0, 'Min Qty': 1 });
  assert.match(status(app).help, /no "Alert Sent" column/, 'a deleted Alert Sent column is reported, not silent');
});

test('the daily email limit is respected', () => {
  const app = freshApp({ quota: 1 });
  app.env.quota = 0;
  addPart(app, { 'Part Name': 'Gasket', Quantity: 1, 'Min Qty': 5 });
  assert.equal(app.env.sent.length, 0);
  assert.match(status(app).help, /daily email limit/);
});

test('email content is escaped and unsafe links are dropped', () => {
  const app = freshApp();
  addPart(app, {
    'Part Name': '<img src=x onerror=alert(1)> "Seal" & Co',
    'Part Number': '=HYPERLINK("http://evil.example","x")',
    Quantity: 1,
    'Min Qty': 5,
    'Order Link': 'javascript:alert(1) grainger.com/product/AF9 http://bad"host.com',
    Notes: '<script>alert(1)</script>',
  });
  const mail = app.env.sent[0];
  assert.doesNotMatch(mail.htmlBody, /<img|<script|javascript:/);
  assert.match(mail.htmlBody, /&lt;img src=x onerror=alert\(1\)&gt; &quot;Seal&quot; &amp; Co/);
  assert.match(mail.htmlBody, /href="https:\/\/grainger\.com\/product\/AF9"/);
  assert.doesNotMatch(mail.htmlBody, /bad&quot;host|bad"host/);
});

test('an existing sheet with its own headers, a title row, linked text and a Total row works as is', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' });
  const sheet = app.ss.insertSheet('Inventory');
  sheet.getRange(1, 1).setValue('Maintenance shop parts');
  sheet.getRange(2, 1, 1, 7).setValues([['Item', 'Part #', 'Qty', 'Reorder Point', 'Vendor', 'Link', 'Shelf Photo']]);
  sheet.getRange(3, 1, 1, 7).setValues([['V-belt A42', "'00042", 3, 2, 'Grainger', '', 'photo-1.jpg']]);
  sheet.getRange(3, 6).setLinkForTest('Buy here', 'https://www.grainger.com/product/A42');
  sheet.getRange(4, 1, 1, 7).setValues([['Air filter', 'AF-9', 8, 4, 'Uline', '=HYPERLINK("https://www.uline.com/AF9","Uline page")', 'photo-2.jpg']]);
  sheet.getRange(5, 1, 1, 7).setValues([['Total', '', '=SUM(C3:C4)', '', '', '', '']]);
  sheet.getRange(2, 3).setNote('On the shelf, not counting the truck');
  const report = app.run('setup');

  assert.deepEqual(report.addedColumns, ['Ordered', 'Alert Sent'], 'only what the emails need is added');
  assert.equal(sheet.getRange(2, 3).getNote(), 'On the shelf, not counting the truck', 'your notes stay');
  assert.match(sheet.getRange(2, 4).getNote(), /drops to this number or below/);
  assert.deepEqual(sheet.dump()[1], ['Item', 'Part #', 'Qty', 'Reorder Point', 'Vendor', 'Link', 'Shelf Photo', 'Ordered', 'Alert Sent']);
  assert.deepEqual(app.ss.getSheets().map((s) => s.getName()), ['Sheet1', 'Inventory', 'Reorder', 'Settings']);
  assert.deepEqual(app.sheet('Reorder').dump()[1].slice(0, 5), ['Item', 'Part #', 'Qty', 'Reorder Point', 'Link'],
    'the Reorder tab uses your headers');
  assert.equal(app.sheet('Reorder').getRange(3, 1).getFormula(),
    '=IFNA(SORT(FILTER({Inventory!A:A,Inventory!B:B,Inventory!C:C,Inventory!D:D,Inventory!F:F},' +
    '((Inventory!A:A<>"")+(Inventory!B:B<>""))>0,ISNUMBER(Inventory!C:C),ISNUMBER(Inventory!D:D),' +
    'Inventory!C:C<=Inventory!D:D,Inventory!H:H<>TRUE,Inventory!H:H<>"Yes"),3,TRUE),"Nothing needs reordering right now.")');
  app.env.sent.length = 0;

  const range = sheet.getRange(3, 3, 2, 1);
  range.setValues([[2], [1]]);
  edit(app, range);
  const mail = app.env.sent[0];
  assert.equal(mail.subject, '[Parts Inventory] Low stock: 2 parts need ordering');
  assert.match(mail.htmlBody, /href="https:\/\/www\.grainger\.com\/product\/A42"/, 'link behind linked text');
  assert.match(mail.htmlBody, /href="https:\/\/www\.uline\.com\/AF9"/, 'link inside =HYPERLINK()');
  assert.match(mail.htmlBody, /Part # 00042/);
  assert.match(mail.htmlBody, /Grainger/, 'the Vendor column is shown');
  assert.doesNotMatch(mail.htmlBody, /Total/);
  assert.deepEqual(sheet.getRange(3, 1, 1, 7).getValues()[0], ['V-belt A42', '00042', 2, 2, 'Grainger', 'Buy here', 'photo-1.jpg']);
});

test('columns of yours that look like the script\'s are never taken over', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' });
  const sheet = app.ss.insertSheet('Inventory');
  sheet.getRange(1, 1, 1, 6).setValues([['Part Name', 'Quantity', 'Min Qty', 'Ordered On', 'Alerted', 'Updated']]);
  sheet.getRange(2, 1, 2, 6).setValues([
    ['Gasket', 1, 5, 'yes', 'no', 'yes'],
    ['Spring', 40, 5, 120, 'yes', 'yes'],
  ]);
  app.run('setup');
  assert.deepEqual(sheet.dump()[0].slice(6), ['Order Link', 'Ordered', 'Alert Sent']);
  app.run('hourlyCheck');
  setCell(app, 2, 'Quantity', 50);
  app.run('menuSendReorderList');
  assert.deepEqual(sheet.getRange(2, 4, 2, 3).getValues(), [['yes', 'no', 'yes'], [120, 'yes', 'yes']],
    'your columns are untouched');
});

test('a tab of yours called Reorder is left alone', () => {
  const app = createApp();
  app.ss.insertSheet('Reorder').getRange(1, 1, 1, 2).setValues([['Vendor', 'Phone']]);
  const report = app.run('setup');
  assert.equal(report.reorder, 'skipped');
  assert.deepEqual(app.sheet('Reorder').dump(), [['Vendor', 'Phone']]);
  assert.ok(app.env.logs.some((l) => /already have a tab named "Reorder"/.test(l[1])));
});

test('settings are read forgivingly', () => {
  const app = createApp();
  const parse = (key, raw) => app.context.parseSettingValue_(app.eval('SETTING_BY_KEY.' + key), raw);
  [['On', true], ['off', false], [true, true], ['yes', true], ['NO', false], ['', true], ['maybe', true]]
    .forEach(([raw, want]) => assert.equal(parse('alertsEnabled', raw), want, JSON.stringify(raw)));
  [[8, 8], ['8', 8], ['8 AM', 8], ['2 PM', 14], ['2pm', 14], ['12 AM', 0], ['12 PM', 12], ['14:00', 14], [23, 23],
    [24, 8], ['noon', 8], ['13 PM', 8], ['', 8]]
    .forEach(([raw, want]) => assert.equal(parse('reminderHour', raw), want, JSON.stringify(raw)));
  assert.equal(parse('reminder', 'weekly (mondays)'), 'Weekly');
  assert.equal(parse('reminder', 'sometimes'), 'Weekdays');
  assert.deepEqual(app.value('parseSettingValue_(SETTING_BY_KEY.recipients, "a@x.com; B@y.org, A@x.com, junk, <c@z.io>")'),
    ['a@x.com', 'B@y.org', 'c@z.io']);
});

test('the menu works from the spreadsheet', () => {
  const app = freshApp({ ui: true });
  app.context.onOpen();
  const menu = app.env.menus[0];
  assert.equal(menu.name, 'Inventory');
  assert.deepEqual(menu.items.filter((i) => i !== '---').map((i) => i.fn), ['menuSendReorderList', 'menuSendTestEmail', 'setup']);
  menu.items.filter((i) => i !== '---').forEach((i) => assert.equal(typeof app.context[i.fn], 'function'));

  app.run('menuSendReorderList');
  assert.match(app.env.alerts.pop().prompt, /Nothing is low right now/);
  addPart(app, { 'Part Name': 'Gasket', Quantity: 1, 'Min Qty': 5 });
  app.run('menuSendReorderList');
  assert.match(app.env.alerts.pop().prompt, /Emailed the list of 1 low part to owner@example\.com/);
  app.run('menuSendTestEmail');
  assert.match(app.env.alerts.pop().prompt, /Sent a test email to owner@example\.com/);
  assert.equal(app.env.sent.length, 3);

  changeSetting(app, 'Send emails to', '');
  app.run('menuSendTestEmail');
  assert.match(app.env.alerts.pop().prompt, /No one is listed next to "Send emails to"/, 'problems are shown, not thrown');

  app.env.mailFails = 'Mail is down';
  app.run('setup');
  assert.match(app.env.alerts.pop().prompt,
    /The tabs are ready and the automatic checks are on, but the test email couldn't be sent: Mail is down\. Fix the/);
  assert.equal(setting(app, 'Send emails to'), OWNER, 'setup fills in an empty address list');
});

test('a coworker can run setup too, and two sets of checks still send each email once', () => {
  const app = freshApp();
  app.env.owner = 'coworker@example.com';
  const report = app.run('setup');
  assert.equal(report.emailsFrom, 'coworker@example.com');
  assert.equal(app.env.triggers.length, 4, 'each person has their own edit and hourly checks');
  assert.match(status(app).help, /sent from coworker@example\.com/);
  app.env.sent.length = 0;

  const row = addPart(app, { 'Part Name': 'Gasket', Quantity: 10, 'Min Qty': 5 });
  const range = inventory(app).getRange(row, col(app, 'Quantity'));
  range.setValue(2);
  edit(app, range);
  edit(app, range); // both people's edit checks run
  app.setTime('2026-09-28T12:10:00Z'); // 8:10 AM, reminder time
  app.run('hourlyCheck');
  app.run('hourlyCheck');
  assert.deepEqual(app.env.sent.map((m) => m.subject), [
    '[Parts Inventory] Low stock: Gasket (2 left)',
    '[Parts Inventory] Reminder: 1 part needs ordering',
  ]);
});

test('a very long list of low parts still fits in one email', () => {
  const app = freshApp();
  const rows = Array.from({ length: 400 }, (_, i) =>
    ['Part ' + i + ' ' + 'x'.repeat(150), 'PN-' + i, 'Aisle ' + i, 1, 5, 'https://www.example.com/parts/' + i + '?ref=' + 'y'.repeat(120)]);
  inventory(app).getRange(2, 1, rows.length, 6).setValues(rows);
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1);
  assert.match(app.env.sent[0].htmlBody, /And 375 more/);
  assert.match(app.env.sent[0].htmlBody, /Plus 225 more/);
});

test('pasting the code anywhere but the spreadsheet gives a clear message', () => {
  const app = createApp();
  app.context.SpreadsheetApp.getActiveSpreadsheet = () => null;
  assert.throws(() => app.run('setup'), /has to be added from inside the spreadsheet/);
});

test('setup is the first function in the file, so the editor\'s Run button runs it', () => {
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Code.js'), 'utf8');
  assert.equal(/^function (\w+)/m.exec(source)[1], 'setup');
  assert.match(source, /\/\*\* @OnlyCurrentDoc\b/, 'Google asks for this spreadsheet only');
  assert.doesNotMatch(source, /[^\x00-\x7f]/, 'plain ASCII, so copying it from a doc or web page can\'t mangle it');
});

test('helpers: header matching, ordered values, links and safe cells', () => {
  const app = createApp();
  assert.deepEqual(app.value('mapColumns_(["Item", "SKU", "Qty On Hand", "Min", "URL", "Ordered?", "Alert sent"])'),
    { name: 1, partNumber: 2, quantity: 3, minQty: 4, link: 5, ordered: 6, alertSentAt: 7 });
  assert.equal(app.value('mapColumns_(["Part Name", "On Order"])').ordered, undefined, '"On Order" is usually a quantity');
  assert.deepEqual(app.value('mapColumns_(["Part Name", "Alerted", "Alert"])'), { name: 1 }, 'the script\'s own column matches exactly');
  const ordered = (v) => app.context.isOrdered_(v);
  assert.deepEqual([true, 'yes', ' YES '].map(ordered), [true, true, true]);
  assert.deepEqual([false, '', 'no', 'x', 'PO 4471', 1, null, new Date('2026-09-01')].map(ordered),
    [false, false, false, false, false, false, false, false], 'the same as the Reorder tab\'s formula');
  assert.equal(app.context.normalizeUrl_('mcmaster.com/91292A112'), 'https://mcmaster.com/91292A112');
  assert.equal(app.context.normalizeUrl_('javascript:alert(1)'), '');
  assert.equal(app.context.normalizeUrl_('not a link'), '');
  assert.equal(app.context.safeCell_('=SUM(A1)'), "'=SUM(A1)");
  assert.equal(app.context.safeCell_('00123'), "'00123");
  assert.equal(app.context.safeCell_('Shelf A'), 'Shelf A');
});

test('the ready-made sheet file matches the script', () => {
  assert.ok(isUpToDate(), 'template/parts-inventory.xlsx is out of date. Run: npm run build');
  const files = unzip(require('fs').readFileSync(OUT));
  const workbook = files['xl/workbook.xml'].toString();
  assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1]), ['Inventory', 'Reorder', 'Settings']);
  const text = (name) => files[name].toString().replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
  LAYOUT.forEach((header) => assert.ok(text('xl/worksheets/sheet1.xml').includes('>' + header + '<'), header));
  assert.ok(text('xl/worksheets/sheet2.xml').includes('<f>' + REORDER_FORMULA.slice(1) + '</f>'));
  createApp().value('SETTINGS').forEach((def) => assert.ok(text('xl/worksheets/sheet3.xml').includes('>' + def.label + '<'), def.label));
});
