'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp } = require('./gas-fakes');

const OWNER = 'owner@example.com';
const ctx = { code: '', actor: 'Sam' };

/**
 * A spreadsheet that has been through setup and whose web app has been opened
 * once (so the app link is known). The clock is set to 6 AM on a Monday in New
 * York, before the 8 AM reminder, so only the reminder tests get reminders.
 */
function freshApp(opts) {
  const app = createApp(Object.assign({
    serviceUrl: 'https://script.google.com/macros/s/abc123/exec',
    now: '2026-09-28T10:00:00Z',
  }, opts));
  app.run('setup');
  app.context.doGet({ parameter: {} });
  return app;
}

function addPart(app, fields) {
  return app.run('apiSavePart', ctx, Object.assign({ name: 'Widget', quantity: 10, minQty: 5 }, fields)).part;
}

function adjust(app, id, mode, amount, note) {
  return app.run('apiAdjustStock', ctx, { id, mode, amount, note });
}

function logRows(app) {
  return app.sheet('Activity Log').dump().slice(1);
}

function settingValue(app, label) {
  const row = app.sheet('Settings').dump().find((r) => r[0] === label);
  return row ? row[1] : undefined;
}

// ---------------------------------------------------------------------------

test('setup creates the tabs, default settings and triggers, and can run again', () => {
  const app = freshApp();
  const names = app.ss.getSheets().map((s) => s.getName());
  assert.deepEqual(names, ['Inventory', 'Settings', 'Activity Log'], 'the blank Sheet1 is reused for Inventory');

  const headers = app.sheet('Inventory').dump()[0];
  assert.deepEqual(headers, app.value('FIELDS.map(f => f.header)'));
  assert.equal(app.sheet('Inventory').getFrozenRows(), 1);
  assert.equal(app.sheet('Inventory').getConditionalFormatRules().length, 2);

  assert.equal(settingValue(app, 'Alert email recipients'), OWNER);
  assert.equal(settingValue(app, 'Email as soon as a part runs low'), true);
  assert.equal(settingValue(app, 'Reminder email'), 'Weekdays');
  assert.equal(settingValue(app, 'Reminder hour (0-23)'), 8);
  assert.equal(settingValue(app, 'App name'), 'Parts Inventory');

  const triggers = app.env.triggers.map((t) => t.fn + ':' + t.type).sort();
  assert.deepEqual(triggers, ['hourlyCheck:CLOCK', 'onInventoryEdit:ON_EDIT']);

  app.run('setup');
  assert.equal(app.env.triggers.length, 2, 'no duplicate triggers');
  assert.equal(app.sheet('Inventory').getConditionalFormatRules().length, 2, 'no duplicate highlight rules');
  assert.equal(app.sheet('Settings').dump().length, 8, 'no duplicate settings rows');
  assert.equal(app.ss.getSheets().length, 3);
});

test('setup keeps other tabs and puts Inventory first when the spreadsheet already has data', () => {
  const app = createApp();
  app.ss.getSheets()[0].getRange(1, 1).setValue('something else');
  app.run('setup');
  assert.deepEqual(app.ss.getSheets().map((s) => s.getName()), ['Inventory', 'Sheet1', 'Settings', 'Activity Log']);
});

test('setup fills in an empty Inventory tab that was made by hand', () => {
  const app = createApp();
  app.ss.insertSheet('Inventory');
  app.run('setup');
  assert.deepEqual(app.sheet('Inventory').dump()[0], app.value('FIELDS.map(f => f.header)'));
  assert.equal(app.sheet('Inventory').getConditionalFormatRules().length, 2);
});

test('setup does not add a second set of triggers for another person', () => {
  const app = freshApp();
  app.env.owner = 'coworker@example.com';
  app.env.active = 'coworker@example.com';
  const report = app.run('setup');
  assert.equal(report.triggers.installed, false);
  assert.equal(report.triggers.owner, OWNER);
  assert.equal(app.env.triggers.length, 2);
});

test('the web app sets the spreadsheet up by itself on first use', () => {
  const app = createApp({ serviceUrl: 'https://script.google.com/macros/s/abc123/exec' });
  const output = app.context.doGet({ parameter: {} });
  assert.ok(app.sheet('Inventory') && app.sheet('Settings') && app.sheet('Activity Log'));
  assert.equal(app.env.triggers.length, 2);
  assert.equal(output.getTitle(), 'Parts Inventory');
  assert.deepEqual(output.meta, [{ name: 'viewport', content: 'width=device-width, initial-scale=1' }]);
  assert.equal(settingValue(app, 'Web app URL'), 'https://script.google.com/macros/s/abc123/exec', 'app link remembered');
});

test('adding a part keeps text as text, normalizes links and logs it', () => {
  const app = freshApp();
  const part = addPart(app, {
    name: '=IMPORTXML("http://evil.example","//a")',
    partNumber: '00123',
    location: '3-4',
    category: 'TRUE',
    unit: 'ea',
    quantity: '12',
    minQty: '5',
    reorderQty: 50,
    unitCost: '0.125',
    link: 'mcmaster.com/91292A112  https://www.grainger.com/product/1234',
    notes: 'line one\nline two',
  });
  assert.equal(part.id, 'P-0001');
  assert.equal(part.name, '=IMPORTXML("http://evil.example","//a")');
  assert.equal(part.partNumber, '00123');
  assert.equal(part.location, '3-4');
  assert.equal(part.category, 'TRUE');
  assert.equal(part.quantity, 12);
  assert.equal(part.unitCost, 0.125);
  assert.deepEqual(part.links, ['https://mcmaster.com/91292A112', 'https://www.grainger.com/product/1234']);
  assert.equal(part.notes, 'line one\nline two');
  assert.equal(part.status, 'ok');
  assert.equal(part.updatedBy, 'owner@example.com', 'the owner is identified by Google');
  assert.deepEqual(app.sheet('Inventory').formulas(), [], 'nothing typed into the app becomes a formula');

  const log = logRows(app);
  assert.equal(log.length, 1);
  assert.equal(log[0][1], 'P-0001');
  assert.equal(log[0][3], 'Added');
  assert.equal(log[0][5], 12);
  assert.deepEqual(app.sheet('Activity Log').formulas(), []);
});

test('part input is checked', () => {
  const app = freshApp();
  const save = (fields) => () => app.run('apiSavePart', ctx, fields);
  assert.throws(save({ name: '  ', quantity: 1 }), /Give the part a name/);
  assert.throws(save({ name: 'A', quantity: -1 }), /can't be negative/);
  assert.throws(save({ name: 'A', quantity: 'lots' }), /must be a number/);
  assert.throws(save({ name: 'A', quantity: 1, link: 'javascript:alert(1)' }), /doesn't look like a web link/);
  assert.throws(save({ name: 'A', quantity: 1, link: 'not a link' }), /doesn't look like a web link/);
  assert.throws(save({ name: 'A', quantity: 1, link: 'a.com b.com c.com d.com e.com f.com' }), /at most 5/);
  assert.throws(save({ name: 'x'.repeat(201), quantity: 1 }), /too long/);
  assert.equal(addPart(app, { name: 'No qty given', quantity: undefined }).quantity, 0);
});

test('editing a part only writes what changed and logs count changes separately', () => {
  const app = freshApp();
  const part = addPart(app, { location: 'A1' });
  const res = app.run('apiSavePart', ctx, { id: part.id, name: 'Widget', location: 'B2', quantity: 7 });
  assert.equal(res.part.location, 'B2');
  assert.equal(res.part.quantity, 7);
  const log = logRows(app);
  assert.deepEqual(log.slice(1).map((r) => [r[3], r[4], r[5], r[7]]), [
    ['Count set', -3, 7, ''],
    ['Edited', '', 7, 'Changed Location'],
  ]);
  const same = app.run('apiSavePart', ctx, { id: part.id, location: 'B2' });
  assert.equal(same.unchanged, true);
  assert.equal(logRows(app).length, 3);
  assert.throws(() => app.run('apiSavePart', ctx, { id: 'P-9999', name: 'Ghost' }), /wasn't found/);
});

test('a part that runs low is emailed once, with its order link, until it is restocked', () => {
  const app = freshApp();
  const part = addPart(app, {
    name: 'M3 x 8 socket head screw', partNumber: '91292A112', location: 'Bin A-12', supplier: 'McMaster-Carr',
    quantity: 10, minQty: 5, reorderQty: 100, unit: 'ea', link: 'https://www.mcmaster.com/91292A112/',
  });

  assert.equal(adjust(app, part.id, 'remove', 4).alert, null);
  assert.equal(app.env.sent.length, 0);

  let res = adjust(app, part.id, 'remove', 1, 'Line 3 repair');
  assert.equal(res.part.quantity, 5);
  assert.equal(res.part.status, 'low');
  assert.equal(res.part.alerted, true);
  assert.ok(res.part.alertSentAt);
  assert.deepEqual(res.alert, { sent: true, count: 1, recipients: [OWNER], error: '', skipped: '' });
  assert.equal(app.env.sent.length, 1);
  const mail = app.env.sent[0];
  assert.equal(mail.to, OWNER);
  assert.equal(mail.name, 'Parts Inventory');
  assert.equal(mail.subject, '[Parts Inventory] Low stock: M3 x 8 socket head screw (5 ea left)');
  assert.match(mail.htmlBody, /href="https:\/\/www\.mcmaster\.com\/91292A112\/"/);
  assert.match(mail.htmlBody, /Order from mcmaster\.com/);
  assert.match(mail.htmlBody, /href="https:\/\/script\.google\.com\/macros\/s\/abc123\/exec\?part=P-0001"/);
  assert.match(mail.body, /Order: https:\/\/www\.mcmaster\.com\/91292A112\//);
  assert.match(mail.body, /min 5, order 100 ea/);

  adjust(app, part.id, 'remove', 3);
  adjust(app, part.id, 'remove', 2);
  assert.equal(app.env.sent.length, 1, 'no repeat emails while it stays low (even at zero)');

  res = adjust(app, part.id, 'add', 100);
  assert.equal(res.part.alerted, false, 'restocking re-arms the alert');
  assert.equal(res.part.status, 'ok');
  res = adjust(app, part.id, 'set', 1);
  assert.equal(res.alert.sent, true);
  assert.equal(app.env.sent.length, 2);

  const actions = logRows(app).map((r) => r[3]);
  assert.deepEqual(actions, ['Added', 'Used', 'Used', 'Alert emailed', 'Used', 'Used', 'Restocked', 'Count set', 'Alert emailed']);
  const used = logRows(app)[2];
  assert.deepEqual([used[4], used[5], used[6], used[7]], [-1, 5, OWNER, 'Line 3 repair']);
});

test('out of stock gets its own subject, and several parts share one email', () => {
  const app = freshApp();
  const part = addPart(app, { name: 'Fuse 5A', quantity: 1, minQty: 0 });
  adjust(app, part.id, 'remove', 1);
  assert.equal(app.env.sent[0].subject, '[Parts Inventory] Out of stock: Fuse 5A');

  // Two parts edited straight into the sheet go low at the same time.
  const a = addPart(app, { name: 'Belt', quantity: 10, minQty: 2 });
  const b = addPart(app, { name: 'Filter', quantity: 10, minQty: 2 });
  const sheet = app.sheet('Inventory');
  const rows = sheet.dump();
  const qtyCol = rows[0].indexOf('Quantity') + 1;
  [a, b].forEach((p) => sheet.getRange(rows.findIndex((r) => r[0] === p.id) + 1, qtyCol).setValue(1));
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 2);
  assert.equal(app.env.sent[1].subject, '[Parts Inventory] Low stock: 2 parts need ordering');
  assert.match(app.env.sent[1].htmlBody, /Belt[\s\S]*Filter/);
});

test('can\'t take out more than is on hand', () => {
  const app = freshApp();
  const part = addPart(app, { quantity: 3 });
  assert.throws(() => adjust(app, part.id, 'remove', 5), /Only 3 on hand/);
  assert.throws(() => adjust(app, part.id, 'remove', 0), /greater than zero/);
  assert.throws(() => adjust(app, part.id, 'fly', 1), /Unknown stock change/);
  assert.equal(adjust(app, part.id, 'set', 0).part.quantity, 0);
  assert.equal(adjust(app, part.id, 'add', 0.25).part.quantity, 0.25);
});

test('nothing is emailed when alerts are off or no one is listed, and the app is told why', () => {
  const app = freshApp();
  app.run('apiSaveSettings', ctx, { alertsEnabled: false });
  const part = addPart(app, { quantity: 10, minQty: 5 });
  let res = adjust(app, part.id, 'remove', 6);
  assert.equal(res.alert.sent, false);
  assert.match(res.alert.skipped, /turned off/);
  assert.equal(res.part.alerted, false, 'not marked, so it goes out once alerts are back on');

  app.run('apiSaveSettings', ctx, { alertsEnabled: true, recipients: '' });
  res = adjust(app, part.id, 'remove', 1);
  assert.match(res.alert.skipped, /No alert email recipients/);
  assert.equal(app.env.sent.length, 0);

  app.run('apiSaveSettings', ctx, { recipients: 'buyer@example.com, lead@example.com' });
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1);
  assert.equal(app.env.sent[0].to, 'buyer@example.com,lead@example.com');
});

test('a failed email is reported and retried by the hourly check', () => {
  const app = freshApp();
  const part = addPart(app, { quantity: 10, minQty: 5 });
  app.env.mailFails = 'Mail service unavailable';
  const res = adjust(app, part.id, 'remove', 8);
  assert.equal(res.alert.sent, false);
  assert.match(res.alert.error, /Mail service unavailable/);
  assert.equal(res.part.alerted, false);
  let data = app.run('apiGetData', ctx);
  assert.match(data.info.lastAlertError.message, /Mail service unavailable/);

  app.env.mailFails = null;
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1);
  data = app.run('apiGetData', ctx);
  assert.equal(data.info.lastAlertError, null);
  assert.ok(data.info.lastCheck);
});

test('the daily email limit is respected', () => {
  const app = freshApp({ quota: 0 });
  const part = addPart(app, { quantity: 10, minQty: 5 });
  const res = adjust(app, part.id, 'remove', 8);
  assert.match(res.alert.error, /daily email limit/);
});

test('email content is escaped and unsafe links from the sheet are dropped', () => {
  const app = freshApp();
  const sheet = app.sheet('Inventory');
  const headers = sheet.dump()[0];
  const row = headers.map(() => '');
  row[headers.indexOf('Part Name')] = '<img src=x onerror=alert(1)> "quoted" & co';
  row[headers.indexOf('Quantity')] = 1;
  row[headers.indexOf('Min Qty')] = 5;
  row[headers.indexOf('Order Link')] = 'javascript:alert(1) https://ok.example.com/p?a=1&b=<2>';
  row[headers.indexOf('Notes')] = '<script>alert(2)</script>';
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  app.run('hourlyCheck');
  const html = app.env.sent[0].htmlBody;
  assert.ok(!html.includes('<img src=x'), 'name is escaped');
  assert.ok(html.includes('&lt;img src=x onerror=alert(1)&gt; &quot;quoted&quot; &amp; co'));
  assert.ok(!html.includes('<script>'));
  assert.ok(!/javascript:/i.test(html), 'javascript: links never reach the email');
  assert.ok(!html.includes('ok.example.com/p?a=1&b=<2>'), 'links with markup in them are dropped');
});

test('hand edits in the spreadsheet get an ID, a timestamp, a log entry and an alert', () => {
  const app = freshApp();
  const sheet = app.sheet('Inventory');
  const headers = sheet.dump()[0];
  const col = (h) => headers.indexOf(h) + 1;
  sheet.getRange(2, col('Part Name')).setValue('Hand-typed part');
  sheet.getRange(2, col('Min Qty')).setValue(3);
  sheet.getRange(2, col('Quantity')).setValue(2);
  const range = sheet.getRange(2, col('Quantity'));
  app.as('');
  app.context.onInventoryEdit({ range, oldValue: '10', value: '2', user: { getEmail: () => 'editor@example.com' } });

  const row = sheet.dump()[1];
  assert.equal(row[col('ID') - 1], 'P-0001');
  assert.equal(row[col('Updated By') - 1], 'editor@example.com');
  assert.ok(row[col('Last Updated') - 1] instanceof Date);
  assert.ok(row[col('Alert Sent') - 1] instanceof Date);
  assert.equal(app.env.sent.length, 1);
  const log = logRows(app);
  assert.deepEqual(log[0].slice(1, 8), ['P-0001', 'Hand-typed part', 'Count set', -8, 2, 'editor@example.com', 'Edited in the spreadsheet']);
  assert.equal(log[1][3], 'Alert emailed');

  // Editing a column the app manages doesn't touch the timestamp or send anything.
  const before = sheet.getRange(2, col('Last Updated')).getValue().getTime();
  app.context.onInventoryEdit({ range: sheet.getRange(2, col('Alert Sent')), user: { getEmail: () => '' } });
  assert.equal(sheet.getRange(2, col('Last Updated')).getValue().getTime(), before);
  // Edits on other tabs are ignored.
  app.context.onInventoryEdit({ range: app.sheet('Settings').getRange(2, 2) });
  assert.equal(app.env.sent.length, 1);
});

test('copied rows get new IDs, and IDs are never reused', () => {
  const app = freshApp();
  const a = addPart(app, { name: 'First' });
  addPart(app, { name: 'Second' });
  const sheet = app.sheet('Inventory');
  const firstRow = sheet.dump()[1];
  sheet.getRange(4, 1, 1, firstRow.length).setValues([firstRow.map((v) => (typeof v === 'string' ? "'" + v : v))]);
  const data = app.run('apiGetData', ctx);
  assert.deepEqual(data.parts.map((p) => p.id), ['P-0001', 'P-0002', 'P-0003']);
  app.run('apiDeletePart', ctx, 'P-0003');
  app.run('apiDeletePart', ctx, a.id);
  assert.equal(addPart(app, { name: 'Third' }).id, 'P-0004');
  assert.deepEqual(app.run('apiGetData', ctx).parts.map((p) => p.name), ['Second', 'Third']);
  assert.equal(logRows(app).filter((r) => r[3] === 'Deleted').length, 2);
});

test('marking a part as ordered keeps it out of reminders until it is restocked', () => {
  const app = freshApp();
  const part = addPart(app, { quantity: 2, minQty: 5 });
  let res = app.run('apiSetOrdered', ctx, part.id, true);
  assert.equal(res.part.ordered, true);
  assert.ok(res.part.orderedAt);
  res = adjust(app, part.id, 'add', 1);
  assert.equal(res.part.ordered, true, 'still low, still on order');
  res = adjust(app, part.id, 'add', 10);
  assert.equal(res.part.ordered, false, 'cleared once restocked');
  res = app.run('apiSetOrdered', ctx, part.id, false);
  assert.equal(res.part.ordered, false);
});

test('the reminder goes out once a day at the chosen hour, weekdays only by default', () => {
  const app = freshApp();
  app.run('apiSaveSettings', ctx, { alertsEnabled: false });
  const low = addPart(app, { name: 'Gloves', quantity: 1, minQty: 5 });
  const ordered = addPart(app, { name: 'Tape', quantity: 0, minQty: 2 });
  app.run('apiSetOrdered', ctx, ordered.id, true);
  const reminders = () => app.env.sent.filter((m) => /Reminder/.test(m.subject));

  app.setTime('2026-09-28T11:30:00Z'); // Monday 7:30 AM in New York
  app.run('hourlyCheck');
  assert.equal(reminders().length, 0, 'too early');

  app.setTime('2026-09-28T12:10:00Z'); // 8:10 AM
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1);
  const mail = reminders()[0];
  assert.equal(mail.subject, '[Parts Inventory] Reminder: 1 part needs ordering');
  assert.match(mail.htmlBody, /Gloves[\s\S]*Already on order[\s\S]*Tape/);

  app.setTime('2026-09-28T15:10:00Z');
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1, 'once a day');

  app.setTime('2026-10-03T13:00:00Z'); // Saturday
  app.run('hourlyCheck');
  assert.equal(reminders().length, 1, 'not on weekends');

  app.run('apiSaveSettings', ctx, { reminder: 'Daily', reminderHour: 6 });
  app.run('hourlyCheck');
  assert.equal(reminders().length, 2, 'daily includes Saturday');

  app.run('apiSetOrdered', ctx, low.id, true);
  app.setTime('2026-10-04T13:00:00Z');
  app.run('hourlyCheck');
  assert.equal(reminders().length, 2, 'nothing to order, so no reminder');

  app.run('apiSaveSettings', ctx, { reminder: 'Off' });
  app.setTime('2026-10-05T13:00:00Z');
  app.run('apiSetOrdered', ctx, low.id, false);
  app.run('hourlyCheck');
  assert.equal(reminders().length, 2, 'off means off');

  const now = app.run('apiSendReminderNow', ctx);
  assert.deepEqual(now, { sent: true, count: 1, onOrder: 1, recipients: [OWNER] });
});

test('an access code locks the web app for everyone except the owner', () => {
  const app = freshApp();
  addPart(app, { name: 'Secret sauce' });
  app.run('apiSaveSettings', ctx, { accessCode: 'bolt-7731' });
  assert.equal(app.run('apiGetData', ctx).settings.accessCode, 'bolt-7731', 'the owner sees it');

  app.as(''); // someone outside the company opens the link
  const boot = JSON.parse(/const BOOT = (.*?);<\/script>/s.exec(app.context.doGet({ parameter: {} }).getContent())[1]);
  assert.equal(boot.locked, true);
  assert.equal(boot.data, undefined, 'no inventory in the page before the code is entered');
  assert.throws(() => app.run('apiGetData', { code: '' }), /ACCESS_DENIED: Enter the access code/);
  assert.throws(() => app.run('apiGetData', { code: 'guess' }), /ACCESS_DENIED: That access code/);
  assert.throws(() => app.run('apiAdjustStock', { code: 'guess' }, { id: 'P-0001', mode: 'remove', amount: 1 }), /ACCESS_DENIED/);

  const data = app.run('apiGetData', { code: 'bolt-7731', actor: 'Visitor' });
  assert.equal(data.parts[0].name, 'Secret sauce');
  assert.equal(data.settings.accessCode, '', 'the code itself is never sent to other people');
  assert.equal(data.settings.accessCodeSet, true);
  assert.equal(data.user.isOwner, false);
  const res = app.run('apiAdjustStock', { code: 'bolt-7731', actor: 'Visitor' }, { id: 'P-0001', mode: 'remove', amount: 1 });
  assert.equal(res.part.updatedBy, 'Visitor');
  assert.throws(() => app.run('apiSaveSettings', { code: 'bolt-7731' }, { recipients: 'me@evil.example' }), /Only owner@example.com/);
  assert.throws(() => app.run('apiSendTestEmail', { code: 'bolt-7731' }), /Only owner@example.com/);

  for (let i = 0; i < 18; i++) assert.throws(() => app.run('apiGetData', { code: 'guess' + i }), /ACCESS_DENIED/);
  assert.throws(() => app.run('apiGetData', { code: 'bolt-7731' }), /ACCESS_LOCKED/, 'locked after 20 wrong codes');
});

test('settings are checked before they are saved', () => {
  const app = freshApp();
  const save = (input) => () => app.run('apiSaveSettings', ctx, input);
  assert.throws(save({ recipients: 'buyer@example.com, not-an-email' }), /isn't a valid email/);
  assert.throws(save({ reminder: 'Hourly' }), /how often/);
  assert.throws(save({ reminderHour: 24 }), /0 to 23/);
  assert.throws(save({ accessCode: '123' }), /at least 6/);
  assert.throws(save({ appName: ' ' }), /can't be empty/);
  assert.throws(save({ appUrl: 'javascript:alert(1)' }), /https/);
  const res = app.run('apiSaveSettings', ctx, {
    recipients: 'Buyer@Example.com; lead@example.com buyer@example.com', reminder: 'Weekly', reminderHour: 14,
    appName: 'Shop 2 Parts', accessCode: '',
  });
  assert.deepEqual(res.settings.recipients, ['Buyer@Example.com', 'lead@example.com']);
  assert.equal(settingValue(app, 'Alert email recipients'), 'Buyer@Example.com, lead@example.com');
  assert.equal(settingValue(app, 'App name'), 'Shop 2 Parts', 'text with digits is stored as text');
  assert.equal(settingValue(app, 'Reminder hour (0-23)'), 14);
});

test('an existing sheet with its own headers, a title row and linked text works as is', () => {
  const app = createApp();
  const sheet = app.ss.insertSheet('Inventory');
  sheet.getRange(1, 1).setValue('Maintenance shop parts');
  sheet.getRange(2, 1, 1, 7).setValues([['Item', 'Part #', 'Qty', 'Reorder Point', 'Vendor', 'Link', 'Shelf Photo']]);
  sheet.getRange(3, 1, 1, 7).setValues([['V-belt A42', "'00042", 3, 2, 'Grainger', '', 'photo-1.jpg']]);
  sheet.getRange(3, 6).setLinkForTest('Buy here', 'https://www.grainger.com/product/A42');
  sheet.getRange(4, 1, 1, 7).setValues([['Air filter', 'AF-9', 8, 4, 'Uline', '=HYPERLINK("https://www.uline.com/AF9","Uline page")', 'photo-2.jpg']]);
  sheet.getRange(5, 1, 1, 7).setValues([['Total', '', '', '', '', '', '']]);
  app.run('setup');

  const headers = sheet.dump()[1];
  assert.deepEqual(headers.slice(7), ['ID', 'Ordered On', 'Alert Sent', 'Last Updated', 'Updated By'], 'only core columns are added');
  const data = app.run('apiGetData', ctx);
  assert.equal(data.fields.partNumber, true);
  assert.equal(data.fields.category, false, 'optional columns that are missing stay hidden');
  const [belt, filter] = data.parts;
  assert.deepEqual(
    [belt.id, belt.name, belt.partNumber, belt.quantity, belt.minQty, belt.supplier, belt.links],
    ['P-0001', 'V-belt A42', '00042', 3, 2, 'Grainger', ['https://www.grainger.com/product/A42']]);
  assert.deepEqual(filter.links, ['https://www.uline.com/AF9']);
  assert.equal(data.parts.length, 2, 'the Total row is not a part');

  const res = adjust(app, belt.id, 'remove', 1);
  assert.equal(res.part.quantity, 2);
  assert.equal(res.alert.sent, true);
  assert.deepEqual(sheet.getRange(3, 1, 1, 7).getValues()[0], ['V-belt A42', '00042', 2, 2, 'Grainger', 'Buy here', 'photo-1.jpg']);
  assert.match(app.env.sent[0].htmlBody, /href="https:\/\/www\.grainger\.com\/product\/A42"/);

  addPart(app, { name: 'Fan belt B30' });
  assert.deepEqual(sheet.dump().slice(2).map((r) => r[0]), ['V-belt A42', 'Air filter', 'Fan belt B30', 'Total'],
    'new parts go above the Total row');
});

test('the app never takes over look-alike columns of yours, or clears them', () => {
  const app = createApp({ now: '2026-09-28T10:00:00Z' }); // 6 AM in New York, before the reminder
  const sheet = app.ss.insertSheet('Inventory');
  sheet.getRange(1, 1, 1, 7).setValues([['Part Name', 'Quantity', 'Min Qty', 'Ordered', 'Alerted', 'Updated', 'Part ID']]);
  sheet.getRange(2, 1, 2, 7).setValues([
    ['Gasket', 1, 5, false, 'no', 'yes', 'G-1'],
    ['Spring', 40, 5, 120, 'yes', 'yes', 'G-1'],
  ]);
  app.run('setup');
  assert.deepEqual(sheet.dump()[0].slice(7), ['ID', 'Order Link', 'Ordered On', 'Alert Sent', 'Last Updated', 'Updated By'],
    'the app adds its own columns instead of using yours');
  const [gasket, spring] = app.run('apiGetData', ctx).parts;
  assert.deepEqual([gasket.ordered, gasket.alerted, spring.ordered], [false, false, false]);
  app.run('hourlyCheck');
  adjust(app, spring.id, 'add', 1);
  app.run('apiSendReminderNow', ctx);
  assert.deepEqual(sheet.getRange(2, 4, 2, 4).getValues(), [[false, 'no', 'yes', 'G-1'], [120, 'yes', 'yes', 'G-1']],
    'your columns are untouched');
  assert.equal(app.env.sent.length, 2, 'an alert, and a reminder that lists the gasket as still to order');
  assert.match(app.env.sent[1].subject, /Reminder: 1 part needs ordering/);
});

test('an unticked box or "no" in Ordered On does not count as ordered', () => {
  const app = freshApp();
  const part = addPart(app, { quantity: 1, minQty: 5 });
  const sheet = app.sheet('Inventory');
  const col = sheet.dump()[0].indexOf('Ordered On') + 1;
  [false, 'no', 0].forEach((value) => {
    sheet.getRange(2, col).setValue(value);
    assert.equal(app.run('apiGetData', ctx).parts[0].ordered, false, JSON.stringify(value));
  });
  [true, 'yes', 'Oct 2'].forEach((value) => {
    sheet.getRange(2, col).setValue(value);
    assert.equal(app.run('apiGetData', ctx).parts[0].ordered, true, JSON.stringify(value));
  });
  assert.equal(part.id, 'P-0001');
});

test('the hourly check can\'t be used to keep the inventory busy', () => {
  const app = freshApp();
  app.as('');
  app.run('hourlyCheck');
  const first = app.env.props.LAST_CHECK;
  assert.ok(first, 'one call from the page runs');
  app.setTime('2026-09-28T10:05:00Z');
  app.run('hourlyCheck');
  app.run('hourlyCheck', { triggerUid: 'made-up' });
  assert.equal(app.env.props.LAST_CHECK, first, 'more calls within 10 minutes do nothing');
  app.context.hourlyCheck(app.triggerEvent('hourlyCheck'));
  assert.notEqual(app.env.props.LAST_CHECK, first, 'the real hourly trigger always runs');
});

test('columns formatted as Plain text keep part numbers exactly, with no stray apostrophe', () => {
  const app = freshApp();
  const sheet = app.sheet('Inventory');
  const headers = sheet.dump()[0];
  ['Part Number', 'Part Name'].forEach((h) => sheet.getRange(2, headers.indexOf(h) + 1, 998, 1).setNumberFormat('@'));
  const part = addPart(app, { name: '=cmd|calc', partNumber: '00123', location: '3-4' });
  assert.equal(part.partNumber, '00123');
  assert.equal(part.location, '3-4');
  assert.deepEqual(sheet.formulas(), [], 'still no formulas');
  const edited = app.run('apiSavePart', ctx, { id: part.id, partNumber: '0042-B' }).part;
  assert.equal(edited.partNumber, '0042-B');
});

test('a very long list of low parts still fits in one email', () => {
  const app = freshApp();
  const sheet = app.sheet('Inventory');
  const headers = sheet.dump()[0];
  const rows = [];
  for (let i = 1; i <= 200; i++) {
    const row = headers.map(() => '');
    row[headers.indexOf('Part Name')] = 'Bolt size ' + i;
    row[headers.indexOf('Quantity')] = 1;
    row[headers.indexOf('Min Qty')] = 10;
    row[headers.indexOf('Order Link')] = 'https://www.example.com/bolt/' + i;
    row[headers.indexOf('Notes')] = 'Some notes about this bolt that make the email a bit longer.';
    rows.push(row.map((v) => (typeof v === 'string' && /\d/.test(v) ? "'" + v : v)));
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  app.run('hourlyCheck');
  assert.equal(app.env.sent.length, 1);
  const mail = app.env.sent[0];
  assert.match(mail.subject, /200 parts need ordering/);
  assert.match(mail.htmlBody, /And 175 more/);
  assert.match(mail.htmlBody, /Plus 25 more\. Open the app/);
  assert.match(mail.body, /\.\.\.plus 25 more/);
  assert.equal(app.run('apiGetData', ctx).parts.filter((p) => p.alerted).length, 200, 'every part counts as emailed');
});

test('new parts go right under the list, never over a totals row', () => {
  const app = freshApp();
  addPart(app, { name: 'One' });
  const sheet = app.sheet('Inventory');
  const headers = sheet.dump()[0];
  sheet.getRange(4, headers.indexOf('Quantity') + 1).setValue(999); // a stray total two rows down
  addPart(app, { name: 'Two' });
  addPart(app, { name: 'Three' });
  const names = sheet.dump().slice(1).map((r) => r[1]);
  assert.deepEqual(names, ['One', 'Two', 'Three', '']);
  assert.equal(sheet.dump()[4][headers.indexOf('Quantity')], 999);
});

test('the page gets its data embedded safely, and ?part= is passed through', () => {
  const app = freshApp();
  addPart(app, { name: '</script><script>alert(1)</script>', notes: '  <!-- --> &amp;' });
  const html = app.context.doGet({ parameter: { part: 'P-0001' } }).getContent();
  const script = /const BOOT = (.*?);<\/script>/s.exec(html)[1];
  assert.ok(!script.includes('</script>'), 'no way to close the script tag early');
  const boot = JSON.parse(script);
  assert.equal(boot.part, 'P-0001');
  assert.equal(boot.data.parts[0].name, '</script><script>alert(1)</script>');
  assert.equal(boot.data.user.isOwner, true);
  assert.equal(typeof boot.data.parts[0].updatedAt, 'string');
});

test('activity can be read back, newest first, and filtered to one part', () => {
  const app = freshApp();
  const a = addPart(app, { name: 'A' });
  const b = addPart(app, { name: 'B' });
  adjust(app, a.id, 'remove', 1);
  adjust(app, b.id, 'add', 2);
  const all = app.run('apiGetActivity', ctx, { limit: 3 }).entries;
  assert.deepEqual(all.map((e) => [e.partId, e.action, e.change]), [['P-0002', 'Restocked', 2], ['P-0001', 'Used', -1], ['P-0002', 'Added', 10]]);
  assert.equal(typeof all[0].at, 'string');
  const onlyA = app.run('apiGetActivity', ctx, { partId: 'p-0001' }).entries;
  assert.deepEqual(onlyA.map((e) => e.action), ['Used', 'Added']);
});

test('test email works with or without parts', () => {
  const app = freshApp();
  assert.deepEqual(app.run('apiSendTestEmail', ctx), { recipients: [OWNER] });
  assert.equal(app.env.sent[0].subject, '[Parts Inventory] Test: low-stock emails are working');
  assert.match(app.env.sent[0].htmlBody, /Example part/);
});

test('menu and setup functions refuse to run through the web app', () => {
  const app = freshApp();
  app.as('');
  ['setup', 'menuCheckNow', 'menuSendTestEmail', 'menuSendReminder', 'menuOpenApp', 'menuOpenSidebar', 'menuHelp']
    .forEach((name) => assert.throws(() => app.run(name), /only be run from the spreadsheet/, name));
  assert.equal(app.run('onInventoryEdit', { range: { getSheet: 'nope' } }), undefined, 'fake events are ignored');
});

test('spreadsheet menu items work from the spreadsheet', () => {
  const app = freshApp({ ui: true });
  assert.equal(app.env.dialogs[0].title, 'Parts Inventory is ready');
  app.context.onOpen();
  const menu = app.env.menus[0];
  assert.equal(menu.name, 'Inventory');
  const handlers = menu.items.filter((i) => i !== '---').map((i) => i.fn);
  handlers.forEach((fn) => assert.equal(typeof app.context[fn], 'function', fn + ' exists'));

  app.run('apiSaveSettings', ctx, { alertsEnabled: false });
  const part = addPart(app, { quantity: 1, minQty: 5 });
  app.run('menuCheckNow');
  assert.match(app.env.alerts.pop().prompt, /1 part is low, but no email was sent/);
  app.run('apiSaveSettings', ctx, { alertsEnabled: true });
  app.run('menuCheckNow');
  assert.match(app.env.alerts.pop().prompt, /Emailed a low-stock alert for 1 part/);
  app.run('menuSendReminder');
  assert.match(app.env.alerts.pop().prompt, /Sent a reminder listing 1 low part/);
  app.run('menuSendTestEmail');
  assert.match(app.env.alerts.pop().prompt, /Sent a test email/);
  app.run('menuOpenApp');
  assert.match(app.env.dialogs.pop().html, /script\.google\.com\/macros\/s\/abc123\/exec/);
  app.run('menuOpenSidebar');
  assert.match(app.env.sidebars[0].getContent(), new RegExp(part.id));
  app.run('menuHelp');
  assert.equal(app.env.dialogs.pop().title, 'Parts Inventory help');
});

test('only the intended functions can be called from the page', () => {
  const app = createApp();
  const publicFns = app.functionNames().filter((k) => !/_$/.test(k)).sort();
  assert.deepEqual(publicFns, [
    'apiAdjustStock', 'apiDeletePart', 'apiGetActivity', 'apiGetData', 'apiSavePart', 'apiSaveSettings',
    'apiSendReminderNow', 'apiSendTestEmail', 'apiSetOrdered', 'doGet', 'hourlyCheck', 'menuCheckNow',
    'menuHelp', 'menuOpenApp', 'menuOpenSidebar', 'menuSendReminder', 'menuSendTestEmail', 'onInventoryEdit',
    'onOpen', 'setup',
  ]);
});

test('helpers: header matching, links and safe cells', () => {
  const app = createApp();
  const cols = app.eval('mapColumns_')(['Part #', 'Item', 'Qty on hand', 'Par level', 'U/M', 'Product Link', 'Qty']);
  assert.deepEqual(JSON.parse(JSON.stringify(cols)), { name: 2, partNumber: 1, quantity: 3, minQty: 4, unit: 5, link: 6 });

  const url = app.eval('normalizeUrl_');
  assert.equal(url('www.digikey.com/en/products/detail/x'), 'https://www.digikey.com/en/products/detail/x');
  assert.equal(url('http://example.com/a,'), 'http://example.com/a');
  assert.equal(url('ftp://example.com'), '');
  assert.equal(url('javascript:alert(1)'), '');
  assert.equal(url('data:text/html,hi'), '');
  assert.equal(url('https://exa mple.com'), '');
  assert.equal(url('https://example.com/"onmouseover='), '');

  const safe = app.eval('safeCell_');
  assert.equal(safe('Widget'), 'Widget');
  assert.equal(safe('00123'), "'00123");
  assert.equal(safe('=1+1'), "'=1+1");
  assert.equal(safe("'quoted"), "''quoted");
  assert.equal(safe(''), '');
  assert.equal(safe(5), 5);
});

test('install/Code.gs is up to date and works as the only file', () => {
  const fs = require('fs');
  const { buildBundle, OUT } = require('../dev/build-bundle');
  assert.equal(fs.readFileSync(OUT, 'utf8'), buildBundle(), 'install/Code.gs is out of date: run npm run build');

  const app = createApp({ codeFile: OUT, serviceUrl: 'https://script.google.com/macros/s/abc123/exec', now: '2026-09-28T10:00:00Z' });
  app.run('setup');
  const html = app.context.doGet({ parameter: {} }).getContent();
  assert.match(html, /<div id="app" hidden>/);
  assert.match(html, /const BOOT = \{"mode":"webapp"/);
  const part = addPart(app, { quantity: 5, minQty: 3 });
  adjust(app, part.id, 'remove', 2);
  assert.equal(app.env.sent.length, 1, 'the low-stock email goes out');
});

