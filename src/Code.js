/**
 * Parts Inventory: low-stock emails for this Google Sheet.
 *
 * To turn the emails on (once):
 *   1. In the spreadsheet, choose Extensions > Apps Script.
 *   2. Delete everything in the editor, paste in this whole file, and click Save.
 *   3. Click Run, and allow access when Google asks.
 * Run it again any time to repair the tabs. More help is in README.md.
 *
 * From then on:
 *   - When a part's Quantity drops to its Min Qty or below, the people listed
 *     in the Settings tab get an email with the part's order link. Each part
 *     is emailed once per shortage, and again only after it has been
 *     restocked and runs low again.
 *   - A reminder lists everything that is still low and not ticked as
 *     Ordered (weekday mornings, unless Settings says otherwise).
 *   - Ordered unticks itself once the part has been restocked.
 */

// Asks Google for access to this spreadsheet only, rather than all of your files.
/** @OnlyCurrentDoc */

/**
 * Sets up the Inventory, Reorder and Settings tabs, turns on the automatic
 * emails and sends a test email. Safe to run again at any time. It's the
 * first function in the file so that the editor's Run button runs it.
 */
function setup() {
  const ss = getSpreadsheet_();
  const report = withLock_(function () {
    return setupSpreadsheet_(ss);
  });
  const ui = getUi_();
  report.emailsFrom = installTriggers_(ss);
  const settings = getSettings_(ss);
  try {
    sendTestEmail_(ss, settings);
    report.testEmailTo = settings.recipients.slice();
    PropertiesService.getScriptProperties().deleteProperty(PROP.lastError);
  } catch (err) {
    report.emailError = errorMessage_(err);
    PropertiesService.getScriptProperties().setProperty(PROP.lastError, report.emailError.slice(0, 300));
  }
  updateStatus_(ss);
  const lines = setupSummary_(report);
  if (ui) {
    ss.setActiveSheet(sheetNamed_(ss, SHEET_NAMES.inventory));
    ui.alert('Parts Inventory', lines.join('\n\n'), ui.ButtonSet.OK);
  } else {
    lines.forEach(function (line) { console.log(line); });
  }
  return report;
}

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const SHEET_NAMES = {
  inventory: 'Inventory',
  reorder: 'Reorder',
  settings: 'Settings',
};

/**
 * Columns of the Inventory tab. Each one is found by its header (case, spaces
 * and punctuation don't matter, and the aliases work too), so columns can be
 * moved, and columns of your own are left alone.
 *   layout:  part of a new Inventory tab
 *   core:    added by setup when it's missing, because the emails need it
 *   managed: filled in by the script (grey header). It only matches its exact
 *            header, never an alias, so the script never writes in a column
 *            of yours.
 */
const FIELDS = [
  {
    key: 'name', header: 'Part Name', layout: true, core: true, width: 260,
    aliases: ['Name', 'Part', 'Item', 'Item Name', 'Description', 'Part Description', 'Item Description'],
  },
  {
    key: 'partNumber', header: 'Part Number', layout: true, width: 140,
    aliases: ['Part #', 'Part No', 'PN', 'P/N', 'SKU', 'MPN', 'Item Number', 'Item #', 'Model', 'Model Number',
      'Catalog Number', 'Mfr Part Number', 'Manufacturer Part Number'],
  },
  {
    key: 'location', header: 'Location', layout: true, width: 120,
    aliases: ['Bin', 'Shelf', 'Bin Location', 'Storage', 'Storage Location', 'Area'],
  },
  {
    key: 'quantity', header: 'Quantity', type: 'number', layout: true, core: true, width: 90,
    aliases: ['Qty On Hand', 'Quantity On Hand', 'On Hand', 'Qty', 'Current Qty', 'Current Stock', 'In Stock', 'Stock', 'Count'],
    note: 'How many are on hand right now.',
  },
  {
    key: 'minQty', header: 'Min Qty', type: 'number', layout: true, core: true, width: 90,
    aliases: ['Min', 'Minimum', 'Min Quantity', 'Minimum Qty', 'Minimum Quantity', 'Min Stock', 'Minimum Stock',
      'Reorder Point', 'Reorder Level', 'Reorder At', 'Low Stock At', 'Alert At', 'Alert Level', 'Threshold',
      'Par', 'Par Level', 'Safety Stock'],
    note: 'When Quantity drops to this number or below, the row turns yellow and an email goes out. Leave it blank for no emails.',
  },
  {
    key: 'link', header: 'Order Link', layout: true, core: true, width: 300,
    aliases: ['Link', 'Links', 'URL', 'Part Link', 'Product Link', 'Reorder Link', 'Purchase Link', 'Buy Link',
      'Order URL', 'Supplier Link', 'Website', 'Web Link'],
    note: 'The web page to reorder the part from. It goes in the email. For a second supplier, add its link after a space.',
  },
  {
    key: 'ordered', header: 'Ordered', type: 'checkbox', layout: true, core: true, width: 80,
    aliases: ['Ordered?'],
    note: 'Tick it once more is on the way. Ticked parts are left out of the reminder emails, and the box unticks itself when the part is restocked.',
  },
  {
    key: 'notes', header: 'Notes', layout: true, width: 260,
    aliases: ['Note', 'Comments', 'Comment', 'Remarks'],
  },
  {
    key: 'reorderQty', header: 'Reorder Qty', type: 'number',
    aliases: ['Reorder Quantity', 'Order Qty', 'Order Quantity', 'Reorder Amount', 'Qty To Order', 'Quantity To Order'],
  },
  {
    key: 'unit', header: 'Unit',
    aliases: ['Units', 'UOM', 'Unit Of Measure', 'U/M'],
  },
  {
    key: 'supplier', header: 'Supplier',
    aliases: ['Vendor', 'Supplier Name', 'Vendor Name', 'Source', 'Store'],
  },
  {
    key: 'alertSentAt', header: 'Alert Sent', type: 'date', layout: true, core: true, managed: true, width: 160,
    note: 'Filled in when the low-stock email for this part goes out, and cleared once the part is restocked. Clear it yourself to get emailed about the part again.',
  },
];

const FIELD_BY_KEY = FIELDS.reduce(function (map, field) {
  map[field.key] = field;
  return map;
}, {});

/** Columns shown in the Reorder tab, when the Inventory tab has them. */
const REORDER_KEYS = ['name', 'partNumber', 'location', 'quantity', 'minQty', 'link'];
const REORDER_TITLE = 'Parts at or below their Min Qty that aren\'t ticked as Ordered. This list updates by itself.';
const NOTHING_TO_REORDER = 'Nothing needs reordering right now.';

/** Rows of the Settings tab. Each one is found by the label in column A. */
const SETTINGS = [
  {
    key: 'status', label: 'Status', type: 'status', default: 'Off',
    aliases: ['Email alerts', 'Email status'],
  },
  {
    key: 'recipients', label: 'Send emails to', type: 'emails',
    aliases: ['Alert email recipients', 'Alert emails', 'Recipients', 'Send alerts to', 'Email to'],
    help: 'Who gets the low-stock emails. Separate several addresses with commas.',
  },
  {
    key: 'alertsEnabled', label: 'Email right away when a part runs low', type: 'onoff', default: true,
    aliases: ['Email as soon as a part runs low', 'Instant alerts', 'Alerts enabled', 'Send alerts'],
    help: 'On: an email goes out as soon as a part drops to its Min Qty or below. Each part is emailed once, and again only after it has been restocked and runs low again.',
  },
  {
    key: 'reminder', label: 'Reminder email', type: 'choice', choices: ['Off', 'Daily', 'Weekdays', 'Weekly'], default: 'Weekdays',
    aliases: ['Reminder', 'Summary email', 'Digest'],
    help: 'A list of everything that is still low and not ticked as Ordered: Off, Daily, Weekdays, or Weekly (Mondays).',
  },
  {
    key: 'reminderHour', label: 'Reminder hour (0-23)', type: 'hour', default: 8,
    aliases: ['Reminder hour', 'Reminder time'],
    help: 'When the reminder goes out, in this spreadsheet\'s time zone (File > Settings). 8 is 8 AM, 14 is 2 PM.',
  },
];

const SETTING_BY_KEY = SETTINGS.reduce(function (map, def) {
  map[def.key] = def;
  return map;
}, {});

/** What the Status row of the Settings tab says. */
const STATUS_TEXT = {
  off: 'Not turned on yet. To turn them on, open Extensions > Apps Script in this spreadsheet, paste in the Parts Inventory code, click Save, then click Run.',
  on: 'Emails are sent from {owner}. The sheet is checked after every edit and once an hour.',
  problem: 'The last email couldn\'t be sent: {error}. It will be tried again within the hour.',
};

/** Keys used in Script Properties. */
const PROP = {
  triggersOwner: 'TRIGGERS_OWNER',
  lastReminderDate: 'LAST_REMINDER_DATE',
  lastError: 'LAST_ERROR',
};

const TRIGGER_HANDLERS = ['onInventoryEdit', 'hourlyCheck'];
const MAX_LINKS = 5;
const EMAIL_CARDS = 25; // parts shown in full in one email
const EMAIL_LIST = 150; // parts listed by name after those (Gmail caps an email at about 200 KB)
const DATE_TIME_FORMAT = 'mmm d, yyyy h:mm am/pm';
const YES_TEXT = /^(yes|y|x|ordered|true)$/i;
const NO_TEXT = /^(no|n|false|not ordered|-)$/i;

const COLORS = {
  header: '#1f2937',
  managedHeader: '#6b7280',
  helpText: '#6b7280',
  accent: '#0f766e',
  reorderTab: '#b45309',
  settingsTab: '#6b7280',
  outRow: '#fde2e1',
  lowRow: '#fff1c2',
};

// ---------------------------------------------------------------------------
// Spreadsheet menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory')
    .addItem('Email the reorder list now', 'menuSendReorderList')
    .addItem('Send a test email', 'menuSendTestEmail')
    .addSeparator()
    .addItem('Set up / repair', 'setup')
    .addToUi();
}

function menuSendReorderList() {
  runMenu_(function (ss) {
    const settings = getSettings_(ss);
    const result = withLock_(function () {
      return sendReminder_(ss, readInventory_(ss), settings, true);
    });
    if (result.sent) clearError_(ss);
    alert_('Reorder list', result.sent
      ? 'Emailed the list of ' + countLabel_(result.count + result.onOrder, 'low part', 'low parts') +
        ' to ' + settings.recipients.join(', ') + '.'
      : 'Nothing is low right now, so there was nothing to send.');
  });
}

function menuSendTestEmail() {
  runMenu_(function (ss) {
    const settings = getSettings_(ss);
    sendTestEmail_(ss, settings);
    clearError_(ss);
    alert_('Test email', 'Sent a test email to ' + settings.recipients.join(', ') +
      '. If it doesn\'t arrive in a minute or two, check the spam folder.');
  });
}

function runMenu_(fn) {
  const ss = getSpreadsheet_();
  try {
    fn(ss);
  } catch (err) {
    alert_('Parts Inventory', errorMessage_(err));
  }
}

// ---------------------------------------------------------------------------
// Automatic triggers (turned on by setup)
// ---------------------------------------------------------------------------

/** Runs after every edit in the spreadsheet. */
function onInventoryEdit(e) {
  try {
    if (!e || !e.range || typeof e.range.getSheet !== 'function') return;
    const sheet = e.range.getSheet();
    if (sheet.getName().toLowerCase() !== SHEET_NAMES.inventory.toLowerCase()) return;
    const ss = sheet.getParent();
    withLock_(function () {
      const inv = readInventory_(ss);
      const startRow = e.range.getRow();
      const endRow = startRow + e.range.getNumRows() - 1;
      const edited = inv.parts.filter(function (p) { return p.row >= startRow && p.row <= endRow; });
      if (edited.length) processAlerts_(inv, getSettings_(ss), edited);
    });
  } catch (err) {
    console.error('onInventoryEdit failed: ' + errorMessage_(err));
  }
}

/** Runs every hour: catches anything the edit check missed and sends the reminder. */
function hourlyCheck() {
  const ss = getSpreadsheet_();
  if (!sheetNamed_(ss, SHEET_NAMES.inventory)) return;
  withLock_(function () {
    const inv = readInventory_(ss);
    const settings = getSettings_(ss);
    processAlerts_(inv, settings);
    sendReminderIfDue_(ss, inv, settings, new Date());
  });
}

/**
 * Turns on the edit and hourly checks under the account running setup, and
 * returns that account. Triggers belong to the person who made them, so a
 * coworker's setup adds a second set. That doesn't double the emails: both
 * sets share Alert Sent and the reminder date, under the same script lock.
 */
function installTriggers_(ss) {
  const me = effectiveEmail_() || 'the account that ran setup';
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onInventoryEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('hourlyCheck').timeBased().everyHours(1).create();
  PropertiesService.getScriptProperties().setProperty(PROP.triggersOwner, me);
  return me;
}

// ---------------------------------------------------------------------------
// Low-stock emails
// ---------------------------------------------------------------------------

/** At or below its Min Qty. A blank Min Qty means no emails for the part. */
function needsReorder_(p) {
  return p.minQty !== null && p.quantity !== null && p.quantity <= p.minQty;
}

function isOut_(p) {
  return p.quantity !== null && p.quantity <= 0;
}

/**
 * Emails the parts that have just dropped to or below their Min Qty (one email
 * for all of them), and clears Alert Sent and Ordered on restocked parts.
 * Call while holding the lock. `parts` defaults to every part.
 */
function processAlerts_(inv, settings, parts) {
  const result = { sent: false, parts: [], error: '', skipped: '' };
  const newlyLow = [];
  const clearAlert = [];
  const clearOrdered = [];
  (parts || inv.parts).forEach(function (p) {
    if (needsReorder_(p)) {
      if (!p.alerted) newlyLow.push(p);
    } else if (p.quantity !== null && p.alerted) {
      // Restocked after running low: start over, ready for the next shortage.
      clearAlert.push(p);
      if (p.ordered && !p.orderedFormula) clearOrdered.push(p);
    }
  });
  if (clearAlert.length && inv.cols.alertSentAt) {
    writeCells_(inv.sheet, inv.cols.alertSentAt, clearAlert.map(function (p) { return { row: p.row, value: '' }; }));
    clearAlert.forEach(function (p) { p.alerted = false; });
  }
  if (clearOrdered.length) {
    // A ticked box goes back to unticked; "yes" is cleared.
    writeCells_(inv.sheet, inv.cols.ordered, clearOrdered.map(function (p) {
      return { row: p.row, value: p.orderedRaw === true ? false : '' };
    }));
    clearOrdered.forEach(function (p) { p.ordered = false; });
  }
  if (!newlyLow.length) return result;

  newlyLow.sort(byUrgency_);
  result.parts = newlyLow.map(displayName_);
  if (!settings.alertsEnabled) {
    result.skipped = 'Emailing right away is turned off in the Settings tab.';
    return result;
  }
  if (!inv.cols.alertSentAt) {
    result.skipped = 'the Inventory tab has no "Alert Sent" column. Choose Inventory > Set up / repair';
  } else if (!settings.recipients.length) {
    result.skipped = 'no one is listed next to "Send emails to" in the Settings tab';
  }
  if (result.skipped) {
    recordError_(inv.ss, result.skipped);
    return result;
  }
  try {
    sendEmail_(inv.ss, settings, buildEmail_(inv, settings, newlyLow, 'alert'));
  } catch (err) {
    result.error = errorMessage_(err);
    recordError_(inv.ss, result.error);
    return result;
  }
  const now = new Date();
  writeCells_(inv.sheet, inv.cols.alertSentAt, newlyLow.map(function (p) { return { row: p.row, value: now }; }));
  newlyLow.forEach(function (p) { p.alerted = true; });
  clearError_(inv.ss);
  result.sent = true;
  return result;
}

/** Sends the scheduled reminder when it is due (at most once a day). */
function sendReminderIfDue_(ss, inv, settings, now) {
  if (settings.reminder === 'Off' || !settings.recipients.length) return null;
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  if (Number(Utilities.formatDate(now, tz, 'H')) < settings.reminderHour) return null;
  const dow = dayOfWeek_(today);
  if (settings.reminder === 'Weekdays' && (dow === 0 || dow === 6)) return null;
  if (settings.reminder === 'Weekly' && dow !== 1) return null;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP.lastReminderDate) === today) return null;
  try {
    const result = sendReminder_(ss, inv, settings, false);
    props.setProperty(PROP.lastReminderDate, today);
    if (result.sent) clearError_(ss);
    return result;
  } catch (err) {
    recordError_(ss, errorMessage_(err));
    return null;
  }
}

/**
 * Emails the low parts that aren't ticked as Ordered, plus the ones that are,
 * for context. With force, also sends when every low part is ordered.
 */
function sendReminder_(ss, inv, settings, force) {
  const low = inv.parts.filter(needsReorder_).sort(byUrgency_);
  const toOrder = low.filter(function (p) { return !p.ordered; });
  const onOrder = low.filter(function (p) { return p.ordered; });
  if (!toOrder.length && !(force && onOrder.length)) return { sent: false, count: 0, onOrder: onOrder.length };
  sendEmail_(ss, settings, buildEmail_(inv, settings, toOrder, 'reminder', onOrder));
  return { sent: true, count: toOrder.length, onOrder: onOrder.length };
}

function sendTestEmail_(ss, settings) {
  const inv = readInventory_(ss);
  sendEmail_(ss, settings, buildEmail_(inv, settings, inv.parts.filter(needsReorder_).sort(byUrgency_), 'test'));
}

function sendEmail_(ss, settings, email) {
  if (!settings.recipients.length) {
    throw new Error('No one is listed next to "Send emails to" in the Settings tab');
  }
  if (MailApp.getRemainingDailyQuota() < settings.recipients.length) {
    throw new Error('Google\'s daily email limit for this account has been reached');
  }
  MailApp.sendEmail({
    to: settings.recipients.join(','),
    subject: email.subject,
    body: email.text,
    htmlBody: email.html,
    name: appName_(ss),
  });
}

function buildEmail_(inv, settings, parts, kind, onOrder) {
  onOrder = onOrder || [];
  const ss = inv.ss;
  let subject;
  let heading;
  let intro;
  if (kind === 'test') {
    subject = 'Low-stock emails are on';
    heading = 'Low-stock emails are on';
    intro = 'When a part drops to its Min Qty or below, an email like this goes to ' + settings.recipients.join(', ') +
      ' with the part\'s order link. ' + (parts.length
      ? 'These parts are low right now:'
      : 'Nothing is low right now.');
  } else if (kind === 'reminder') {
    subject = parts.length
      ? 'Reminder: ' + countLabel_(parts.length, 'part needs', 'parts need') + ' ordering'
      : 'Reminder: ' + countLabel_(onOrder.length, 'low part is', 'low parts are') + ' on order';
    heading = parts.length ? 'Parts that still need ordering' : 'Low parts that are on order';
    intro = parts.length
      ? 'These parts are at or below their Min Qty and aren\'t ticked as Ordered yet.'
      : 'Every low part is ticked as Ordered.';
  } else if (parts.length === 1) {
    const p = parts[0];
    const out = isOut_(p);
    subject = (out ? 'Out of stock: ' : 'Low stock: ') + displayName_(p) +
      (out ? '' : ' (' + formatQty_(p.quantity, p.unit) + ' left)');
    heading = displayName_(p) + (out ? ' is out of stock' : ' is running low');
    intro = out
      ? 'There are none left.'
      : 'Only ' + formatQty_(p.quantity, p.unit) + ' left. The Min Qty is ' + formatNumber_(p.minQty) + '.';
  } else {
    subject = 'Low stock: ' + parts.length + ' parts need ordering';
    heading = parts.length + ' parts are running low';
    intro = 'These parts have dropped to or below their Min Qty.';
  }
  subject = '[' + appName_(ss) + '] ' + subject;

  const sheetUrl = ss.getUrl();
  const rowUrl = function (p) { return sheetUrl + '#gid=' + inv.sheet.getSheetId() + '&range=A' + p.row; };
  const reorderSheet = sheetNamed_(ss, SHEET_NAMES.reorder);
  const font = '-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif';
  const orderedCards = onOrder.length
    ? '<h2 style="font-size:16px;margin:24px 0 10px;color:#1f2937;">Already ordered</h2>' + partsHtml_(onOrder, rowUrl)
    : '';
  const footerLinks = [
    linkHtml_(sheetUrl, 'Open the spreadsheet'),
    reorderSheet ? linkHtml_(sheetUrl + '#gid=' + reorderSheet.getSheetId(), 'See the Reorder tab') : '',
  ].filter(String).join(' &nbsp;&middot;&nbsp; ');
  const footerNote = 'You get these emails because your address is in the Settings tab of "' + appName_(ss) + '".';
  const html =
    '<div style="background:#f3f4f6;padding:24px 12px;font-family:' + font + ';color:#1f2937;">' +
    '<div style="max-width:600px;margin:0 auto;">' +
    '<div style="font-size:13px;font-weight:600;color:' + COLORS.accent + ';margin-bottom:6px;">' +
    escapeHtml_(appName_(ss)) + '</div>' +
    '<h1 style="font-size:22px;line-height:1.3;margin:0 0 8px;color:#111827;">' + escapeHtml_(heading) + '</h1>' +
    '<p style="font-size:15px;line-height:1.5;margin:0 0 20px;color:#374151;">' + escapeHtml_(intro) + '</p>' +
    partsHtml_(parts, rowUrl) + orderedCards +
    '<p style="font-size:14px;margin:24px 0 8px;">' + footerLinks + '</p>' +
    '<p style="font-size:12px;line-height:1.5;margin:0;color:#6b7280;">' + escapeHtml_(footerNote) + '</p>' +
    '</div></div>';

  const lines = [heading, '', intro, ''];
  pushPartsText_(lines, parts);
  if (onOrder.length) {
    lines.push('Already ordered:', '');
    pushPartsText_(lines, onOrder);
  }
  lines.push('Spreadsheet: ' + sheetUrl, '', footerNote);
  return { subject: subject, html: html, text: lines.join('\n') };
}

function partsHtml_(parts, rowUrl) {
  const html = parts.slice(0, EMAIL_CARDS).map(function (p) { return partCardHtml_(p, rowUrl(p)); }).join('');
  const rest = parts.slice(EMAIL_CARDS);
  if (!rest.length) return html;
  const listed = rest.slice(0, EMAIL_LIST);
  const items = listed.map(function (p) {
    const name = p.links[0]
      ? '<a href="' + escapeHtml_(p.links[0]) + '" style="color:' + COLORS.accent + ';">' + escapeHtml_(displayName_(p)) + '</a>'
      : escapeHtml_(displayName_(p));
    return '<li style="margin:0 0 6px;">' + name + ' <span style="color:#6b7280;">' + escapeHtml_(shortStock_(p)) + '</span></li>';
  }).join('');
  const more = rest.length - listed.length;
  return html +
    '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:8px;padding:14px 16px;margin:0 0 12px;">' +
    '<div style="font-size:15px;font-weight:600;margin-bottom:8px;color:#111827;">And ' + rest.length + ' more</div>' +
    '<ul style="margin:0;padding-left:18px;font-size:14px;line-height:1.4;">' + items + '</ul>' +
    (more ? '<div style="font-size:13px;color:#6b7280;margin-top:8px;">Plus ' + more + ' more. See the Reorder tab for all of them.</div>' : '') +
    '</div>';
}

function pushPartsText_(lines, parts) {
  parts.slice(0, EMAIL_CARDS).forEach(function (p) { lines.push(partText_(p), ''); });
  const rest = parts.slice(EMAIL_CARDS);
  rest.slice(0, EMAIL_LIST).forEach(function (p) { lines.push('* ' + displayName_(p) + ': ' + shortStock_(p)); });
  if (rest.length > EMAIL_LIST) lines.push('...plus ' + (rest.length - EMAIL_LIST) + ' more. See the Reorder tab for all of them.');
  if (rest.length) lines.push('');
}

function shortStock_(p) {
  return (isOut_(p) ? 'out of stock' : formatQty_(p.quantity, p.unit) + ' left') +
    (p.minQty !== null ? ' (min ' + formatNumber_(p.minQty) + ')' : '');
}

function partCardHtml_(p, rowUrl) {
  const out = isOut_(p);
  const color = out ? '#b91c1c' : '#b45309';
  const status = out ? 'Out of stock' : formatQty_(p.quantity, p.unit) + ' on hand';
  const details = [
    p.minQty !== null ? 'min ' + formatNumber_(p.minQty) : '',
    p.reorderQty !== null ? 'order ' + formatQty_(p.reorderQty, p.unit) : '',
    p.ordered ? 'ticked as ordered' : '',
  ].filter(String).join(' \u00b7 ');
  const meta = [
    p.partNumber ? 'Part # ' + p.partNumber : '',
    p.location ? 'Location: ' + p.location : '',
    p.supplier,
  ].filter(String).join(' \u00b7 ');
  const buttons = p.links.map(function (url, i) {
    if (i === 0) {
      return '<a href="' + escapeHtml_(url) + '" style="display:inline-block;background:' + COLORS.accent +
        ';color:#ffffff;text-decoration:none;font-weight:600;font-size:14px;padding:10px 16px;border-radius:6px;' +
        'margin:8px 8px 0 0;">Order from ' + escapeHtml_(hostOf_(url)) + ' &rarr;</a>';
    }
    return '<a href="' + escapeHtml_(url) + '" style="display:inline-block;color:' + COLORS.accent +
      ';font-size:14px;padding:10px 4px;margin:8px 8px 0 0;">or ' + escapeHtml_(hostOf_(url)) + '</a>';
  }).join('');
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
    'style="border-collapse:separate;background:#ffffff;border:1px solid #e5e7eb;border-left:4px solid ' + color +
    ';border-radius:8px;margin:0 0 12px;"><tr><td style="padding:14px 16px;">' +
    '<div style="font-size:16px;font-weight:600;line-height:1.35;color:#111827;">' + escapeHtml_(displayName_(p)) + '</div>' +
    (meta ? '<div style="font-size:13px;color:#6b7280;margin-top:2px;">' + escapeHtml_(meta) + '</div>' : '') +
    '<div style="font-size:14px;margin-top:8px;"><span style="color:' + color + ';font-weight:700;">' +
    escapeHtml_(status) + '</span>' +
    (details ? '<span style="color:#6b7280;"> &middot; ' + escapeHtml_(details) + '</span>' : '') + '</div>' +
    (p.notes ? '<div style="font-size:13px;color:#4b5563;margin-top:6px;white-space:pre-wrap;">' +
      escapeHtml_(truncate_(p.notes, 300)) + '</div>' : '') +
    (buttons || '<div style="font-size:13px;color:#6b7280;margin-top:8px;">No order link saved for this part yet.</div>') +
    '<div style="margin-top:10px;"><a href="' + escapeHtml_(rowUrl) +
    '" style="font-size:13px;color:#6b7280;">Show it in the spreadsheet</a></div>' +
    '</td></tr></table>';
}

function partText_(p) {
  const qty = isOut_(p) ? 'OUT OF STOCK' : formatQty_(p.quantity, p.unit) + ' on hand';
  const extra = [
    p.minQty !== null ? 'min ' + formatNumber_(p.minQty) : '',
    p.reorderQty !== null ? 'order ' + formatQty_(p.reorderQty, p.unit) : '',
    p.ordered ? 'ticked as ordered' : '',
  ].filter(String).join(', ');
  const lines = ['* ' + displayName_(p) + ': ' + qty + (extra ? ' (' + extra + ')' : '')];
  const meta = [p.partNumber ? 'Part # ' + p.partNumber : '', p.location, p.supplier].filter(String).join(' / ');
  if (meta) lines.push('  ' + meta);
  p.links.forEach(function (url) { lines.push('  Order: ' + url); });
  return lines.join('\n');
}

/** Out of stock first, then the rest in spreadsheet order. */
function byUrgency_(a, b) {
  return (isOut_(b) ? 1 : 0) - (isOut_(a) ? 1 : 0) || a.row - b.row;
}

function recordError_(ss, message) {
  PropertiesService.getScriptProperties().setProperty(PROP.lastError, String(message).slice(0, 300));
  updateStatus_(ss);
}

function clearError_(ss) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP.lastError) === null) return;
  props.deleteProperty(PROP.lastError);
  updateStatus_(ss);
}

// ---------------------------------------------------------------------------
// Inventory tab
// ---------------------------------------------------------------------------

function readInventory_(ss) {
  const sheet = sheetNamed_(ss, SHEET_NAMES.inventory);
  if (!sheet) throw new Error('The Inventory tab is missing. Choose Inventory > Set up / repair.');
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : [];
  const headerIndex = findHeaderRow_(values);
  const cols = mapColumns_(values[headerIndex] || []);
  const headerRow = headerIndex + 1;
  const firstDataRow = headerRow + 1;
  const dataRows = lastRow - firstDataRow + 1;
  const links = cols.link && dataRows > 0 ? readLinkUrls_(sheet, firstDataRow, cols.link, dataRows) : [];
  const orderedFormulas = cols.ordered && dataRows > 0
    ? sheet.getRange(firstDataRow, cols.ordered, dataRows, 1).getFormulas()
    : [];
  const parts = [];
  for (let i = headerIndex + 1; i < values.length; i++) {
    const part = rowToPart_(values[i], cols, i + 1, links[i - headerIndex - 1]);
    if (!part) continue;
    part.orderedFormula = !!(orderedFormulas[i - headerIndex - 1] && orderedFormulas[i - headerIndex - 1][0]);
    parts.push(part);
  }
  return { ss: ss, sheet: sheet, values: values, headerRow: headerRow, cols: cols, parts: parts };
}

function rowToPart_(row, cols, rowNumber, extraLinks) {
  const get = function (key) { return cols[key] ? row[cols[key] - 1] : ''; };
  const name = cellText_(get('name'));
  const partNumber = cellText_(get('partNumber'));
  if (!name && !partNumber) return null;
  // A "Total" row under the list isn't a part.
  if (!partNumber && /^(grand\s*|sub\s*)?totals?\s*:?$/i.test(name)) return null;
  const links = [];
  const seen = {};
  parseLinks_(cellText_(get('link'))).concat(extraLinks || []).forEach(function (url) {
    const clean = normalizeUrl_(url);
    // Sheets turns typed URLs into links, sometimes as http:// or with a
    // trailing slash, so compare without those to avoid listing a link twice.
    const key = clean.replace(/^https?:\/\//i, '').replace(/\/+$/, '').toLowerCase();
    if (clean && !seen[key]) {
      seen[key] = true;
      links.push(clean);
    }
  });
  const orderedRaw = get('ordered');
  return {
    row: rowNumber,
    name: name,
    partNumber: partNumber,
    location: cellText_(get('location')),
    quantity: toNumberOrNull_(get('quantity')),
    minQty: toNumberOrNull_(get('minQty')),
    reorderQty: toNumberOrNull_(get('reorderQty')),
    unit: cellText_(get('unit')),
    supplier: cellText_(get('supplier')),
    notes: cellText_(get('notes')),
    links: links.slice(0, MAX_LINKS),
    ordered: isOrdered_(orderedRaw),
    orderedRaw: orderedRaw,
    alerted: isMarked_(get('alertSentAt')),
  };
}

/**
 * A ticked box (or "yes") counts as ordered, the same as in the Reorder tab's
 * formula. Anything else doesn't, and is never cleared.
 */
function isOrdered_(value) {
  return value === true || (typeof value === 'string' && /^yes$/i.test(value.trim()));
}

/** Link URLs hidden behind rich text links or =HYPERLINK() formulas in the link column. */
function readLinkUrls_(sheet, firstRow, col, numRows) {
  const out = [];
  const range = sheet.getRange(firstRow, col, numRows, 1);
  const add = function (i, url) {
    if (!url) return;
    out[i] = out[i] || [];
    out[i].push(url);
  };
  try {
    range.getRichTextValues().forEach(function (r, i) {
      const cell = r[0];
      if (!cell) return;
      cell.getRuns().forEach(function (run) { add(i, run.getLinkUrl()); });
      if (!out[i]) add(i, cell.getLinkUrl());
    });
  } catch (err) {
    // Rich text isn't available for every kind of cell; plain URLs still work.
  }
  try {
    range.getFormulas().forEach(function (r, i) {
      const match = /HYPERLINK\(\s*"([^"]+)"/i.exec(r[0] || '');
      if (match) add(i, match[1]);
    });
  } catch (err) {
    // Same as above.
  }
  return out;
}

/** The header row is the first of the top 5 rows that looks like one (usually row 1). */
function findHeaderRow_(values) {
  let best = 0;
  let bestScore = 0;
  for (let i = 0; i < Math.min(values.length, 5); i++) {
    const cols = mapColumns_(values[i]);
    const score = Object.keys(cols).length;
    if (score > bestScore && (cols.name || cols.partNumber) && score >= 2) {
      best = i;
      bestScore = score;
    }
  }
  return best;
}

function normalizeHeader_(text) {
  return String(text == null ? '' : text).toLowerCase()
    .replace(/#/g, ' number ')
    .replace(/\bno\b\.?/g, ' number ')
    .replace(/[^a-z0-9]+/g, '');
}

/** Maps field keys to 1-based column numbers, using the headers and their aliases. */
function mapColumns_(headers) {
  const normalized = (headers || []).map(normalizeHeader_);
  const taken = {};
  const cols = {};
  FIELDS.forEach(function (field) {
    const names = [field.header].concat(field.managed ? [] : field.aliases || []).map(normalizeHeader_);
    for (let n = 0; n < names.length; n++) {
      for (let i = 0; i < normalized.length; i++) {
        if (!taken[i] && normalized[i] && normalized[i] === names[n]) {
          cols[field.key] = i + 1;
          taken[i] = true;
          return;
        }
      }
    }
  });
  return cols;
}

/** Writes values into one column. updates = [{ row, value }]; one write per run of neighbouring rows. */
function writeCells_(sheet, col, updates) {
  const sorted = updates.slice().sort(function (a, b) { return a.row - b.row; });
  let i = 0;
  while (i < sorted.length) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1].row === sorted[j].row + 1) j++;
    const run = sorted.slice(i, j + 1);
    sheet.getRange(run[0].row, col, run.length, 1).setValues(run.map(function (u) { return [u.value]; }));
    i = j + 1;
  }
}

function ensureRows_(sheet, lastNeededRow) {
  const max = sheet.getMaxRows();
  if (lastNeededRow > max) sheet.insertRowsAfter(max, lastNeededRow - max);
}

function isEmptySheet_(sheet) {
  return sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;
}

// ---------------------------------------------------------------------------
// Settings tab
// ---------------------------------------------------------------------------

function getSettings_(ss) {
  const found = readSettingsRows_(sheetNamed_(ss, SHEET_NAMES.settings));
  const tz = ss.getSpreadsheetTimeZone();
  const settings = {};
  SETTINGS.forEach(function (def) {
    settings[def.key] = parseSettingValue_(def, found[def.key] ? found[def.key].value : undefined, tz);
  });
  return settings;
}

function readSettingsRows_(sheet) {
  const found = {};
  if (!sheet || sheet.getLastRow() < 1) return found;
  sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues().forEach(function (row, i) {
    const def = findSettingDef_(row[0]);
    if (def && !found[def.key]) found[def.key] = { row: i + 1, value: row[1], help: row[2] };
  });
  return found;
}

function findSettingDef_(label) {
  const n = normalizeHeader_(label);
  if (!n) return null;
  return SETTINGS.find(function (def) {
    return [def.label].concat(def.aliases || []).some(function (name) { return normalizeHeader_(name) === n; });
  }) || null;
}

function parseSettingValue_(def, raw, tz) {
  if (def.type === 'emails') return uniqueEmails_(parseEmails_(raw).filter(isEmail_));
  if (def.type === 'onoff') {
    if (typeof raw === 'boolean') return raw;
    const text = String(raw == null ? '' : raw).trim();
    if (/^(on|yes|y|true|1)$/i.test(text)) return true;
    if (/^(off|no|n|false|0)$/i.test(text)) return false;
    return def.default;
  }
  if (def.type === 'choice') {
    const text = String(raw == null ? '' : raw).trim().toLowerCase();
    return def.choices.find(function (c) { return text && text.indexOf(c.toLowerCase()) === 0; }) || def.default;
  }
  if (def.type === 'hour') return parseHour_(raw, def.default, tz);
  return cellText_(raw) || def.default;
}

/**
 * 8, "8", "8 AM", "2 PM" or "14:00" as an hour from 0 to 23. Sheets turns a
 * typed time like 2 PM into a time value, which arrives here as a Date.
 */
function parseHour_(raw, fallback, tz) {
  if (isDate_(raw)) return isNaN(raw.getTime()) ? fallback : Number(Utilities.formatDate(raw, tz || 'UTC', 'H'));
  if (typeof raw === 'number') return raw >= 0 && raw < 24 ? Math.floor(raw) : fallback;
  const m = /^(\d{1,2})(?::00)?\s*([ap])?\.?\s*m?\.?$/i.exec(String(raw == null ? '' : raw).trim());
  if (!m) return fallback;
  let hour = Number(m[1]);
  if (m[2]) {
    if (hour < 1 || hour > 12) return fallback;
    hour = hour % 12 + (m[2].toLowerCase() === 'p' ? 12 : 0);
  }
  return hour <= 23 ? hour : fallback;
}

function writeSettings_(ss, values) {
  const sheet = sheetNamed_(ss, SHEET_NAMES.settings);
  const found = readSettingsRows_(sheet);
  Object.keys(values).forEach(function (key) {
    const def = SETTING_BY_KEY[key];
    const row = (found[key] && found[key].row) || appendSettingRow_(sheet, def);
    sheet.getRange(row, 2).setValue(settingCellValue_(def, values[key]));
  });
}

function settingCellValue_(def, value) {
  if (def.type === 'emails') return safeCell_((value || []).join(', '));
  if (def.type === 'onoff') return value ? 'On' : 'Off';
  if (def.type === 'hour') return Number(value);
  return safeCell_(String(value == null ? '' : value));
}

function appendSettingRow_(sheet, def) {
  const row = Math.max(sheet.getLastRow(), 1) + 1;
  ensureRows_(sheet, row);
  sheet.getRange(row, 1, 1, 3).setValues([[def.label, '', def.help || ''].map(safeCell_)]);
  sheet.getRange(row, 1).setFontWeight('bold');
  sheet.getRange(row, 3).setFontColor(COLORS.helpText).setWrap(true);
  const cell = sheet.getRange(row, 2);
  if (def.type === 'onoff') cell.setDataValidation(listRule_(['On', 'Off']));
  if (def.type === 'choice') cell.setDataValidation(listRule_(def.choices));
  if (def.type === 'hour') {
    cell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 23).setAllowInvalid(false).setHelpText('A whole number from 0 to 23.').build());
  }
  if (def.type === 'status') cell.setFontWeight('bold');
  return row;
}

function listRule_(choices) {
  return SpreadsheetApp.newDataValidation().requireValueInList(choices, true).setAllowInvalid(false).build();
}

/** Shows in the Settings tab whether the emails are on, and the last problem if there is one. */
function updateStatus_(ss) {
  const sheet = sheetNamed_(ss, SHEET_NAMES.settings);
  if (!sheet) return;
  const props = PropertiesService.getScriptProperties();
  const owner = props.getProperty(PROP.triggersOwner);
  const error = props.getProperty(PROP.lastError);
  let value = 'Off';
  let help = STATUS_TEXT.off;
  if (owner && error) {
    value = 'Problem';
    help = STATUS_TEXT.problem.replace('{error}', error.replace(/\.$/, ''));
  } else if (owner) {
    value = 'On';
    help = STATUS_TEXT.on.replace('{owner}', owner);
  }
  const found = readSettingsRows_(sheet);
  const row = found.status ? found.status.row : appendSettingRow_(sheet, SETTING_BY_KEY.status);
  const current = found.status ? [found.status.value, found.status.help] : ['', ''];
  if (current[0] === value && current[1] === help) return;
  sheet.getRange(row, 2, 1, 2).setValues([[value, safeCell_(help)]]);
}

// ---------------------------------------------------------------------------
// Setting up the tabs
// ---------------------------------------------------------------------------

/** Creates the tabs, or adds what's missing. Call while holding the lock. */
function setupSpreadsheet_(ss) {
  const report = { created: [], addedColumns: [], reorder: '' };
  let sheet = sheetNamed_(ss, SHEET_NAMES.inventory);
  if (!sheet) {
    const sheets = ss.getSheets();
    sheet = sheets.length === 1 && isEmptySheet_(sheets[0])
      ? sheets[0].setName(SHEET_NAMES.inventory)
      : ss.insertSheet(SHEET_NAMES.inventory, 0);
    report.created.push(SHEET_NAMES.inventory);
  }
  if (isEmptySheet_(sheet)) {
    createInventoryLayout_(sheet);
  } else {
    const found = readInventory_(ss);
    const missing = FIELDS.filter(function (f) { return f.core && !found.cols[f.key]; });
    if (missing.length) addColumns_(sheet, found.headerRow, missing);
    report.addedColumns = missing.map(function (f) { return f.header; });
  }
  const inv = readInventory_(ss);
  const first = inv.headerRow + 1;
  ensureRows_(sheet, first);
  const rows = sheet.getMaxRows() - inv.headerRow;
  if (inv.cols.ordered) makeCheckboxes_(sheet, first, inv.cols.ordered, rows);
  sheet.getRange(first, inv.cols.alertSentAt, rows, 1).setNumberFormat(DATE_TIME_FORMAT);
  applyStockHighlighting_(sheet, inv.headerRow, inv.cols);
  addHeaderNotes_(inv);

  report.reorder = setupReorderSheet_(ss, inv);
  setupSettingsSheet_(ss, report);
  return report;
}

function createInventoryLayout_(sheet) {
  const fields = FIELDS.filter(function (f) { return f.layout; });
  sheet.getRange(1, 1, 1, fields.length).setValues([fields.map(function (f) { return f.header; })]);
  styleHeaders_(sheet, 1, 1, fields, true);
  sheet.setFrozenRows(1);
  sheet.setTabColor(COLORS.accent);
  fields.forEach(function (f, i) { formatNewColumn_(sheet, 1, i + 1, f); });
}

function addColumns_(sheet, headerRow, fields) {
  const start = sheet.getLastColumn() + 1;
  const needed = start + fields.length - 1;
  if (sheet.getMaxColumns() < needed) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  sheet.getRange(headerRow, start, 1, fields.length).setValues([fields.map(function (f) { return f.header; })]);
  styleHeaders_(sheet, headerRow, start, fields, true);
  fields.forEach(function (f, i) { formatNewColumn_(sheet, headerRow, start + i, f); });
}

/** Number checks for a column the script just added. Columns of yours are left as they are. */
function formatNewColumn_(sheet, headerRow, col, field) {
  ensureRows_(sheet, headerRow + 1);
  const range = sheet.getRange(headerRow + 1, col, sheet.getMaxRows() - headerRow, 1);
  if (field.type === 'number') {
    range.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Enter a number (0 or more).').build());
  }
  if (field.key === 'link') range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function styleHeaders_(sheet, headerRow, startCol, fields, withNotes) {
  fields.forEach(function (field, i) {
    const cell = sheet.getRange(headerRow, startCol + i);
    cell.setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground(field.managed ? COLORS.managedHeader : COLORS.header);
    if (withNotes && field.note) cell.setNote(field.note);
    if (field.width) sheet.setColumnWidth(startCol + i, field.width);
  });
}

/** Explains the columns the emails depend on, on hover. A note that's already there stays. */
function addHeaderNotes_(inv) {
  FIELDS.forEach(function (field) {
    const col = inv.cols[field.key];
    if (!field.note || !col) return;
    const cell = inv.sheet.getRange(inv.headerRow, col);
    if (!cell.getNote()) cell.setNote(field.note);
  });
}

/**
 * Turns the Ordered column into tick boxes, keeping what's already marked.
 * Only when it holds nothing but yes/no answers: a column of quantities,
 * dates, notes or formulas is left as it is.
 */
function makeCheckboxes_(sheet, firstRow, col, numRows) {
  const range = sheet.getRange(firstRow, col, numRows, 1);
  const formulas = range.getFormulas();
  const updates = [];
  const yesNoOnly = range.getValues().every(function (r, i) {
    const value = r[0];
    if (formulas[i][0]) return false;
    if (value === '' || value === null || typeof value === 'boolean') return true;
    const text = typeof value === 'string' ? value.trim() : '';
    if (YES_TEXT.test(text)) updates.push({ row: firstRow + i, value: true });
    else if (NO_TEXT.test(text)) updates.push({ row: firstRow + i, value: false });
    else return false;
    return true;
  });
  if (!yesNoOnly) return false;
  writeCells_(sheet, col, updates);
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  return true;
}

/** Colours low rows yellow and out-of-stock rows red, right in the spreadsheet. */
function applyStockHighlighting_(sheet, headerRow, cols) {
  if (!cols.quantity) return;
  const first = headerRow + 1;
  ensureRows_(sheet, first);
  const range = sheet.getRange(first, 1, sheet.getMaxRows() - headerRow, sheet.getMaxColumns());
  const formulas = stockRuleFormulas_(cols, first);
  const colors = [COLORS.outRow, COLORS.lowRow];
  const rules = formulas.map(function (formula, i) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + formula)
      .setBackground(colors[i])
      .setRanges([range])
      .build();
  });
  const others = sheet.getConditionalFormatRules().filter(function (rule) { return !isStockRule_(rule); });
  sheet.setConditionalFormatRules(rules.concat(others));
}

/** The out-of-stock rule, then the low-stock rule (when there's a Min Qty column). */
function stockRuleFormulas_(cols, firstRow) {
  const qty = '$' + columnLetter_(cols.quantity) + firstRow;
  const out = ['AND(ISNUMBER(' + qty + '),' + qty + '<=0)'];
  if (cols.minQty) {
    const min = '$' + columnLetter_(cols.minQty) + firstRow;
    out.push('AND(ISNUMBER(' + qty + '),ISNUMBER(' + min + '),' + qty + '<=' + min + ')');
  }
  return out;
}

function isStockRule_(rule) {
  const condition = rule.getBooleanCondition();
  if (!condition) return false;
  const values = condition.getCriteriaValues();
  return /^=AND\(ISNUMBER\(\$[A-Z]+\d+\),/.test(String((values && values[0]) || ''));
}

/**
 * The Reorder tab: a FILTER formula that lists the low parts that aren't
 * ticked as Ordered. Returns 'created', 'updated', or 'skipped' when a tab of
 * yours already uses the name.
 */
function setupReorderSheet_(ss, inv) {
  const keys = REORDER_KEYS.filter(function (key) { return inv.cols[key]; });
  let sheet = sheetNamed_(ss, SHEET_NAMES.reorder);
  if (sheet && !isReorderSheet_(sheet)) return 'skipped';
  const created = !sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.reorder, inv.sheet.getIndex());
  if (!isEmptySheet_(sheet)) {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }
  const headerValues = inv.values[inv.headerRow - 1] || [];
  const headers = keys.map(function (key) { return cellText_(headerValues[inv.cols[key] - 1]) || FIELD_BY_KEY[key].header; });
  ensureRows_(sheet, 3);
  sheet.getRange(1, 1).setValue(REORDER_TITLE).setFontWeight('bold');
  sheet.getRange(2, 1, 1, headers.length).setValues([headers.map(safeCell_)]);
  styleHeaders_(sheet, 2, 1, keys.map(function (key) { return FIELD_BY_KEY[key]; }), false);
  sheet.getRange(3, 1).setFormula(reorderFormula_(inv.cols, keys));
  sheet.setFrozenRows(2);
  sheet.setTabColor(COLORS.reorderTab);
  return created ? 'created' : 'updated';
}

/** An empty tab, or one this script made. */
function isReorderSheet_(sheet) {
  if (isEmptySheet_(sheet)) return true;
  if (cellText_(sheet.getRange(1, 1).getValue()) === REORDER_TITLE) return true;
  return /FILTER\(/i.test(sheet.getRange(3, 1).getFormula()) && sheet.getRange(3, 1).getFormula().indexOf(SHEET_NAMES.inventory) !== -1;
}

/**
 * =IFNA(SORT(FILTER(...)), "Nothing needs reordering right now."), most urgent
 * first. Whole-column ranges (A:A) keep it working as rows are added, and
 * also survive Google's .xlsx import, which rejects ranges like A2:A.
 */
function reorderFormula_(cols, keys) {
  const tab = SHEET_NAMES.inventory;
  const column = function (key) {
    const letter = columnLetter_(cols[key]);
    return tab + '!' + letter + ':' + letter;
  };
  const numbers = keys.map(function (key) { return cols[key]; });
  const contiguous = numbers.every(function (n, i) { return i === 0 || n === numbers[i - 1] + 1; });
  const data = contiguous
    ? tab + '!' + columnLetter_(numbers[0]) + ':' + columnLetter_(numbers[numbers.length - 1])
    : '{' + keys.map(column).join(',') + '}';
  const named = ['name', 'partNumber'].filter(function (key) { return cols[key]; })
    .map(function (key) { return column(key) + '<>""'; });
  const conditions = [
    named.length > 1 ? '((' + named.join(')+(') + '))>0' : named[0],
    'ISNUMBER(' + column('quantity') + ')',
    'ISNUMBER(' + column('minQty') + ')',
    column('quantity') + '<=' + column('minQty'),
  ];
  if (cols.ordered) conditions.push(column('ordered') + '<>TRUE', column('ordered') + '<>"Yes"');
  const list = 'SORT(FILTER(' + data + ',' + conditions.join(',') + '),' + (keys.indexOf('quantity') + 1) + ',TRUE)';
  return '=IFNA(' + list + ',"' + NOTHING_TO_REORDER + '")';
}

function setupSettingsSheet_(ss, report) {
  let sheet = sheetNamed_(ss, SHEET_NAMES.settings);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.settings);
    sheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(COLORS.header);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 270);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 560);
    sheet.setTabColor(COLORS.settingsTab);
    report.created.push(SHEET_NAMES.settings);
  }
  const found = readSettingsRows_(sheet);
  SETTINGS.forEach(function (def) {
    if (found[def.key]) return;
    const row = appendSettingRow_(sheet, def);
    if (def.type !== 'emails') sheet.getRange(row, 2).setValue(settingCellValue_(def, def.default));
  });
  if (!getSettings_(ss).recipients.length && effectiveEmail_()) {
    writeSettings_(ss, { recipients: [effectiveEmail_()] });
  }
}

function setupSummary_(report) {
  const lines = [];
  if (report.emailError) {
    lines.push('The tabs are ready and the automatic checks are on, but the test email couldn\'t be sent: ' +
      report.emailError.replace(/\.$/, '') + '. Fix the "Send emails to" row in the Settings tab, then choose ' +
      'Inventory > Send a test email.');
  } else {
    lines.push('Done. The low-stock emails are on, and a test email went to ' + report.testEmailTo.join(', ') + '.');
  }
  if (report.addedColumns.length) {
    lines.push('Added these columns to the Inventory tab: ' + report.addedColumns.join(', ') + '.');
  }
  if (report.reorder === 'skipped') {
    lines.push('You already have a tab named "' + SHEET_NAMES.reorder + '", so it was left alone and the reorder list wasn\'t added.');
  }
  lines.push('List your parts in the Inventory tab. Give each part a Min Qty to get emails about it.');
  if (!getUi_()) lines.push('You can close this tab and go back to the spreadsheet.');
  return lines;
}

// ---------------------------------------------------------------------------
// Spreadsheet, people and locking
// ---------------------------------------------------------------------------

/** The tab with this name. Sheets treats "inventory" and "Inventory" as the same name. */
function sheetNamed_(ss, name) {
  const want = name.toLowerCase();
  return ss.getSheets().find(function (sheet) { return sheet.getName().toLowerCase() === want; }) || null;
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('This code has to be added from inside the spreadsheet: open it and choose Extensions > Apps Script.');
  }
  return ss;
}

function appName_(ss) {
  return ss.getName() || 'Parts Inventory';
}

function effectiveEmail_() {
  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The inventory is busy. Please try again in a moment.');
  try {
    const result = fn();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function getUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (err) {
    return null;
  }
}

function alert_(title, message) {
  const ui = getUi_();
  if (ui) ui.alert(title, message, ui.ButtonSet.OK);
  else console.log(title + ': ' + message);
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function displayName_(p) {
  return p.name || p.partNumber || 'Unnamed part';
}

function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

/** A date, a ticked box or text like "yes" counts as set; blank, unticked or "no" doesn't. */
function isMarked_(value) {
  if (value === true) return true;
  if (isDate_(value)) return !isNaN(value.getTime());
  if (typeof value === 'number') return value > 0;
  if (typeof value === 'string') return !!value.trim() && !NO_TEXT.test(value.trim()) && value.trim() !== '0';
  return false;
}

function cellText_(value) {
  if (value === null || value === undefined) return '';
  if (isDate_(value)) return isNaN(value.getTime()) ? '' : value.toISOString().slice(0, 10);
  return String(value).trim();
}

function toNumberOrNull_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;
  if (value === null || value === undefined || typeof value === 'boolean' || isDate_(value)) return null;
  let s = String(value).trim().replace(/[\s$\u20ac\u00a3\u00a5]/g, '');
  if (!s) return null;
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
  else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : null;
}

function round_(n, decimals) {
  const f = Math.pow(10, decimals);
  return Math.round(n * f) / f;
}

function formatNumber_(n) {
  if (n === null || n === undefined || n === '') return '';
  const rounded = round_(Number(n), 3);
  const pieces = String(Math.abs(rounded)).split('.');
  pieces[0] = pieces[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (rounded < 0 ? '-' : '') + pieces.join('.');
}

function formatQty_(n, unit) {
  return formatNumber_(n === null ? 0 : n) + (unit ? ' ' + unit : '');
}

function countLabel_(n, singular, plural) {
  return n + ' ' + (n === 1 ? singular : plural);
}

function truncate_(text, max) {
  const s = String(text);
  return s.length > max ? s.slice(0, max - 1) + '\u2026' : s;
}

function dayOfWeek_(isoDate) {
  const p = isoDate.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Sheets reads text the way it reads typing: "00123" becomes 123, "3-4" becomes
 * a date and "=..." becomes a formula. A leading apostrophe keeps it as text
 * (the apostrophe itself isn't stored).
 */
function safeCell_(value) {
  if (typeof value !== 'string' || value === '') return value;
  return /^[=+\-@'#]/.test(value) || /\d/.test(value) || /^(true|false)$/i.test(value) ? "'" + value : value;
}

function parseLinks_(text) {
  return String(text || '').split(/\s+/).map(normalizeUrl_).filter(String);
}

/** Returns a clean http(s) URL, adding https:// when it's missing, or '' if it isn't a web link. */
function normalizeUrl_(text) {
  let s = String(text == null ? '' : text).trim().replace(/[,;.]+$/, '');
  if (!s) return '';
  if (!/^[a-z][a-z0-9+.-]*:/i.test(s)) {
    if (!/^[a-z0-9-]+(\.[a-z0-9-]+)+([/?#:]|$)/i.test(s)) return '';
    s = 'https://' + s;
  }
  if (!/^https?:\/\//i.test(s) || /[\s<>"\\`]/.test(s)) return '';
  const host = s.replace(/^https?:\/\//i, '').split(/[/?#]/)[0];
  if (!/^[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z0-9-]{2,}(:\d+)?$/i.test(host)) return '';
  return s;
}

function hostOf_(url) {
  return String(url).replace(/^https?:\/\//i, '').split(/[/?#:]/)[0].replace(/^www\./i, '');
}

function linkHtml_(url, text) {
  return '<a href="' + escapeHtml_(url) + '" style="color:' + COLORS.accent + ';">' + escapeHtml_(text) + '</a>';
}

function parseEmails_(text) {
  return String(text == null ? '' : text).split(/[\s,;]+/)
    .map(function (s) { return s.replace(/^<|>$/g, '').trim(); })
    .filter(String);
}

function isEmail_(text) {
  return String(text).length <= 254 && /^[^\s@<>(),;:"[\]]+@[^\s@<>(),;:"[\]]+\.[^\s@<>(),;:"[\]]{2,}$/.test(String(text));
}

function uniqueEmails_(list) {
  const seen = {};
  return list.filter(function (email) {
    const key = email.toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function escapeHtml_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function errorMessage_(err) {
  const message = err && err.message ? err.message : String(err);
  return message.replace(/^Exception:\s*/, '');
}
