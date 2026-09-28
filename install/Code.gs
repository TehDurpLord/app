// Parts Inventory: the whole app in one file. Paste all of it into Code.gs in the Apps Script editor.
// Built from src/Code.js and src/Index.html by dev/build-bundle.js; edit those, not this.

/**
 * Parts Inventory
 * Track parts in a Google Sheet and get an email with the order links when
 * something runs low.
 *
 * This is the server side of the app. It runs as a Google Apps Script bound to
 * the spreadsheet: it keeps the parts in the "Inventory" tab, serves the web
 * app (Index.html) and sends the low-stock emails. Setup steps are in
 * README.md, and in the spreadsheet under Inventory > Help.
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const APP_VERSION = '1.0.0';

const SHEET_NAMES = {
  inventory: 'Inventory',
  settings: 'Settings',
  log: 'Activity Log',
};

/**
 * Columns of the Inventory tab. Each column is found by its header (case,
 * spaces and punctuation don't matter, and the aliases work too), so columns
 * can be moved around and any extra columns of your own are left alone.
 *   core:    added automatically when missing, because the app needs it
 *   managed: kept up to date by the app (grey header in the sheet). These only
 *            match their exact header, never an alias, because the app writes
 *            and clears them and must never take over a column of yours.
 * The other columns are optional: the app only shows them when they exist.
 */
const FIELDS = [
  {
    key: 'id', header: 'ID', type: 'text', core: true, managed: true, width: 80,
    note: 'Assigned by the app. Leave it blank when you add a row by hand.',
  },
  {
    key: 'name', header: 'Part Name', type: 'text', core: true, width: 260,
    aliases: ['Name', 'Part', 'Item', 'Item Name', 'Description', 'Part Description', 'Item Description'],
  },
  {
    key: 'partNumber', header: 'Part Number', type: 'text', width: 140,
    aliases: ['Part #', 'Part No', 'PN', 'P/N', 'SKU', 'MPN', 'Item Number', 'Item #', 'Model', 'Model Number',
      'Catalog Number', 'Mfr Part Number', 'Manufacturer Part Number'],
  },
  {
    key: 'category', header: 'Category', type: 'text', width: 130,
    aliases: ['Type', 'Group', 'Class', 'Part Type', 'Item Type'],
  },
  {
    key: 'location', header: 'Location', type: 'text', width: 120,
    aliases: ['Bin', 'Shelf', 'Bin Location', 'Storage', 'Storage Location', 'Area'],
  },
  {
    key: 'quantity', header: 'Quantity', type: 'number', core: true, width: 90,
    aliases: ['Qty On Hand', 'Quantity On Hand', 'On Hand', 'Qty', 'Current Qty', 'Current Stock', 'In Stock', 'Stock', 'Count'],
    note: 'How many are on hand right now.',
  },
  {
    key: 'minQty', header: 'Min Qty', type: 'number', core: true, width: 90,
    aliases: ['Min', 'Minimum', 'Min Quantity', 'Minimum Qty', 'Minimum Quantity', 'Min Stock', 'Minimum Stock',
      'Reorder Point', 'Reorder Level', 'Reorder At', 'Low Stock At', 'Alert At', 'Alert Level', 'Threshold',
      'Par', 'Par Level', 'Safety Stock'],
    note: 'Low-stock level. An email goes out when Quantity drops to this number or below. Leave it blank for no emails.',
  },
  {
    key: 'reorderQty', header: 'Reorder Qty', type: 'number', width: 100,
    aliases: ['Reorder Quantity', 'Order Qty', 'Order Quantity', 'Reorder Amount', 'Qty To Order', 'Quantity To Order'],
    note: 'How many to order when it runs low (shown in the email).',
  },
  {
    key: 'unit', header: 'Unit', type: 'text', width: 70,
    aliases: ['Units', 'UOM', 'Unit Of Measure', 'U/M'],
  },
  {
    key: 'supplier', header: 'Supplier', type: 'text', width: 140,
    aliases: ['Vendor', 'Supplier Name', 'Vendor Name', 'Source', 'Store'],
  },
  {
    key: 'link', header: 'Order Link', type: 'link', core: true, width: 260,
    aliases: ['Link', 'Links', 'URL', 'Part Link', 'Product Link', 'Reorder Link', 'Purchase Link', 'Buy Link',
      'Order URL', 'Supplier Link', 'Website', 'Web Link'],
    note: 'Web page to reorder the part from. Put extra links (other suppliers) on separate lines.',
  },
  {
    key: 'unitCost', header: 'Unit Cost', type: 'number', width: 90,
    aliases: ['Cost', 'Price', 'Unit Price', 'Cost Each', 'Price Each'],
  },
  {
    key: 'notes', header: 'Notes', type: 'text', width: 240,
    aliases: ['Note', 'Comments', 'Comment', 'Remarks'],
  },
  {
    key: 'orderedAt', header: 'Ordered On', type: 'date', core: true, managed: true, width: 110,
    note: 'Set by "Mark as ordered" in the app. Cleared automatically once the part is restocked.',
  },
  {
    key: 'alertSentAt', header: 'Alert Sent', type: 'date', core: true, managed: true, width: 140,
    note: 'When the low-stock email for this part went out. Cleared once the part is restocked, so the next time it runs low a new email goes out.',
  },
  {
    key: 'updatedAt', header: 'Last Updated', type: 'date', core: true, managed: true, width: 140,
  },
  {
    key: 'updatedBy', header: 'Updated By', type: 'text', core: true, managed: true, width: 180,
  },
];

const FIELD_BY_KEY = FIELDS.reduce(function (map, field) {
  map[field.key] = field;
  return map;
}, {});

const REMINDER_CHOICES = ['Off', 'Daily', 'Weekdays', 'Weekly'];

/** Rows of the Settings tab. Found by the label in column A. */
const SETTINGS = [
  {
    key: 'recipients', label: 'Alert email recipients', type: 'emails',
    aliases: ['Alert emails', 'Recipients', 'Send alerts to', 'Email to'],
    help: 'Who gets the low-stock emails. Separate several addresses with commas.',
  },
  {
    key: 'alertsEnabled', label: 'Email as soon as a part runs low', type: 'boolean', default: true,
    aliases: ['Instant alerts', 'Alerts enabled', 'Send alerts'],
    help: 'Sends an email the moment a part drops to or below its Min Qty. Each part is emailed about once, and again only after it has been restocked and runs low again.',
  },
  {
    key: 'reminder', label: 'Reminder email', type: 'choice', choices: REMINDER_CHOICES, default: 'Weekdays',
    aliases: ['Reminder', 'Summary email', 'Digest'],
    help: 'A summary of the parts that are still low and not ordered yet: Off, Daily, Weekdays, or Weekly (Mondays).',
  },
  {
    key: 'reminderHour', label: 'Reminder hour (0-23)', type: 'hour', default: 8,
    aliases: ['Reminder hour', 'Reminder time'],
    help: 'Hour of the day for the reminder, in this spreadsheet\'s time zone (File > Settings). 8 = 8 AM, 14 = 2 PM.',
  },
  {
    key: 'appName', label: 'App name', type: 'text', default: 'Parts Inventory',
    aliases: ['Title', 'Company name'],
    help: 'Shown at the top of the app and in email subjects.',
  },
  {
    key: 'accessCode', label: 'App access code', type: 'text', default: '',
    aliases: ['Access code', 'Passcode'],
    help: 'Optional. When set, people must enter this code the first time they open the web app (6 characters or more). The person who deployed the app never needs it.',
  },
  {
    key: 'appUrl', label: 'Web app URL', type: 'url', default: '',
    aliases: ['App URL', 'App link'],
    help: 'Link to the web app, used in emails. Filled in automatically the first time the web app is opened. If it stays empty, paste the web app URL (it ends in /exec) here.',
  },
];

const LOG_HEADERS = ['Timestamp', 'Part ID', 'Part', 'Action', 'Change', 'Qty After', 'By', 'Note'];

/** Keys used in Script Properties. */
const PROP = {
  spreadsheetId: 'SPREADSHEET_ID',
  nextId: 'NEXT_ID',
  triggersOwner: 'TRIGGERS_OWNER',
  lastCheck: 'LAST_CHECK',
  lastReminderDate: 'LAST_REMINDER_DATE',
  lastAlertError: 'LAST_ALERT_ERROR',
};

const TRIGGER_HANDLERS = ['onInventoryEdit', 'hourlyCheck'];
const TEXT_LIMITS = { name: 200, partNumber: 100, category: 100, location: 100, unit: 30, supplier: 150, notes: 2000 };
const NUMBER_INPUTS = { quantity: 'Quantity', minQty: 'Min Qty', reorderQty: 'Reorder Qty', unitCost: 'Unit cost' };
const MAX_LINKS = 5;
const EMAIL_CARDS = 25; // parts shown in full in one email
const EMAIL_LIST = 150; // parts listed by name after those (Gmail caps an email at about 200 KB)
const MAX_ACCESS_FAILURES = 20;

const COLORS = {
  header: '#1f2937',
  managedHeader: '#6b7280',
  accent: '#0f766e',
  outRow: '#fde2e1',
  lowRow: '#fff1c2',
};

// ---------------------------------------------------------------------------
// Spreadsheet menu and web app entry points
// ---------------------------------------------------------------------------

/** Adds the Inventory menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory')
    .addItem('Open inventory app', 'menuOpenApp')
    .addItem('Open in sidebar', 'menuOpenSidebar')
    .addSeparator()
    .addItem('Check stock and send alerts now', 'menuCheckNow')
    .addItem('Send reminder email now', 'menuSendReminder')
    .addItem('Send test email', 'menuSendTestEmail')
    .addSeparator()
    .addItem('Set up / repair', 'setup')
    .addItem('Help', 'menuHelp')
    .addToUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && !ss.getSheetByName(SHEET_NAMES.inventory)) {
      ss.toast('Choose Inventory > Set up / repair to get started.', 'Parts Inventory', 15);
    }
  } catch (err) {
    // The hint is optional; never let it break the menu.
  }
}

/** Serves the web app. ?part=P-0001 opens that part. */
function doGet(e) {
  const params = (e && e.parameter) || {};
  const boot = { mode: 'webapp', part: String(params.part || '').slice(0, 40) };
  let title = 'Parts Inventory';
  try {
    const ss = getSpreadsheet_();
    ensureSetup_(ss);
    let settings = getSettings_(ss);
    if (rememberAppUrl_(ss, settings)) settings = getSettings_(ss);
    title = settings.appName;
    if (needsAccessCode_(settings)) {
      boot.locked = true;
      boot.appName = settings.appName;
    } else {
      boot.data = buildAppData_(ss, settings, currentUser_());
    }
  } catch (err) {
    boot.error = errorMessage_(err);
  }
  return renderApp_(boot, title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function renderApp_(boot, title) {
  // The one-file install (install/Code.gs) carries the page in INDEX_HTML_;
  // otherwise it's the Index.html file next to this one.
  const template = typeof INDEX_HTML_ === 'string'
    ? HtmlService.createTemplate(INDEX_HTML_)
    : HtmlService.createTemplateFromFile('Index');
  template.bootJson = jsonForHtml_(jsonSafe_(boot));
  return template.evaluate().setTitle(title || 'Parts Inventory');
}

/**
 * Creates or repairs the Inventory, Settings and Activity Log tabs and turns
 * on the automatic alerts. Safe to run again at any time.
 */
function setup() {
  requireSheetUser_();
  const ss = getSpreadsheet_();
  const report = withLock_(function () {
    return setupSpreadsheet_(ss);
  });
  const ui = getUi_();
  let triggers = installTriggers_(ss, false);
  if (!triggers.installed && ui) {
    const answer = ui.alert(
      'Automatic alerts',
      'Automatic alerts were turned on by ' + triggers.owner + '. Turn them on under your account as well? ' +
      'Only do this if ' + triggers.owner + ' no longer has access to this spreadsheet.',
      ui.ButtonSet.YES_NO);
    if (answer === ui.Button.YES) triggers = installTriggers_(ss, true);
  }
  report.triggers = triggers;
  if (ui) {
    ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.inventory));
    showDialog_(setupDoneHtml_(ss, report), 'Parts Inventory is ready', 560, 500);
  } else {
    console.log('Setup finished: ' + JSON.stringify(report));
  }
  return report;
}

function menuOpenApp() {
  runMenu_(function (ss) {
    const url = getSettings_(ss).appUrl || currentServiceUrl_();
    if (!url) {
      showDialog_(deployHelpHtml_(), 'Deploy the web app first', 560, 520);
      return;
    }
    showDialog_(openLinkHtml_(url), 'Open inventory app', 380, 150);
  });
}

function menuOpenSidebar() {
  runMenu_(function (ss) {
    ensureSetup_(ss);
    const settings = getSettings_(ss);
    const boot = { mode: 'sidebar', data: buildAppData_(ss, settings, currentUser_()) };
    SpreadsheetApp.getUi().showSidebar(renderApp_(boot, settings.appName));
  });
}

function menuCheckNow() {
  runMenu_(function (ss) {
    ensureSetup_(ss);
    const result = withLock_(function () {
      const inv = ensureCoreColumns_(readInventory_(ss));
      assignMissingIds_(inv);
      return processAlerts_(inv, getSettings_(ss));
    });
    alert_('Stock check', describeAlertResult_(result));
  });
}

function menuSendReminder() {
  runMenu_(function (ss) {
    ensureSetup_(ss);
    const settings = getSettings_(ss);
    const result = withLock_(function () {
      return sendReminder_(ss, readInventory_(ss), settings, true);
    });
    alert_('Reminder email', result.sent
      ? 'Sent a reminder listing ' + countLabel_(result.count + result.onOrder, 'low part', 'low parts') +
        ' to ' + settings.recipients.join(', ') + '.'
      : 'Nothing is low right now, so no reminder was sent.');
  });
}

function menuSendTestEmail() {
  runMenu_(function (ss) {
    ensureSetup_(ss);
    const result = sendTestEmail_(ss, getSettings_(ss));
    alert_('Test email', 'Sent a test email to ' + result.recipients.join(', ') +
      '. If it doesn\'t arrive in a minute or two, check the spam folder.');
  });
}

function menuHelp() {
  runMenu_(function (ss) {
    showDialog_(helpHtml_(ss), 'Parts Inventory help', 600, 560);
  });
}

// ---------------------------------------------------------------------------
// Web app API: functions called from Index.html with google.script.run.
// Each call passes ctx = { code, actor }: the access code (when the owner set
// one) and the name the person typed into the app, which goes into the
// activity log when Google doesn't tell us who they are.
// ---------------------------------------------------------------------------

function apiGetData(ctx) {
  return runApi_(ctx, function (c) {
    return buildAppData_(c.ss, c.settings, c.user);
  });
}

/** Adds a part (no id) or updates the fields sent for an existing part. */
function apiSavePart(ctx, input) {
  return runApi_(ctx, function (c) {
    input = input || {};
    const isNew = !input.id;
    const clean = cleanPartInput_(input, isNew);
    return withLock_(function () {
      const inv = ensureCoreColumns_(readInventory_(c.ss));
      const now = new Date();
      if (isNew) {
        const row = nextEmptyRow_(inv);
        const fields = Object.assign({}, clean, {
          id: nextId_(inv),
          updatedAt: now,
          updatedBy: c.actor,
          orderedAt: null,
          alertSentAt: null,
        });
        writeFields_(inv, row, fields);
        const created = readPartAt_(inv, row);
        logActivity_(c.ss, [{ part: created, action: 'Added', change: created.quantity, qtyAfter: created.quantity }], c.actor);
        const alert = processAlerts_(inv, c.settings, [created]);
        return { part: partForClient_(created), alert: alertForClient_(alert) };
      }

      const part = findPart_(inv, input.id);
      const changes = {};
      Object.keys(clean).forEach(function (key) {
        if (!sameValue_(part[key], clean[key])) changes[key] = clean[key];
      });
      if (!Object.keys(changes).length) return { part: partForClient_(part), alert: null, unchanged: true };
      writeFields_(inv, part.row, Object.assign({}, changes, { updatedAt: now, updatedBy: c.actor }));
      const updated = readPartAt_(inv, part.row);
      const entries = [];
      if ('quantity' in changes) {
        entries.push({
          part: updated, action: 'Count set',
          change: round_((updated.quantity || 0) - (part.quantity || 0), 3), qtyAfter: updated.quantity,
        });
      }
      const edited = Object.keys(changes).filter(function (key) { return key !== 'quantity'; });
      if (edited.length) {
        entries.push({
          part: updated, action: 'Edited', qtyAfter: updated.quantity,
          note: 'Changed ' + edited.map(function (key) { return FIELD_BY_KEY[key].header; }).join(', '),
        });
      }
      logActivity_(c.ss, entries, c.actor);
      const alert = processAlerts_(inv, c.settings, [updated]);
      return { part: partForClient_(updated), alert: alertForClient_(alert) };
    });
  });
}

/** request = { id, mode: 'add' | 'remove' | 'set', amount, note } */
function apiAdjustStock(ctx, request) {
  return runApi_(ctx, function (c) {
    const req = request || {};
    const mode = req.mode;
    if (['add', 'remove', 'set'].indexOf(mode) === -1) throw new Error('Unknown stock change.');
    const amount = parseInputNumber_(req.amount, 'Amount', 3);
    if (amount === null) throw new Error('Enter an amount.');
    if (mode !== 'set' && amount === 0) throw new Error('Enter an amount greater than zero.');
    const note = cleanString_(req.note, false).slice(0, 300);
    return withLock_(function () {
      const inv = ensureCoreColumns_(readInventory_(c.ss));
      const part = findPart_(inv, req.id);
      const current = part.quantity || 0;
      let next = mode === 'add' ? current + amount : mode === 'remove' ? current - amount : amount;
      next = round_(next, 3);
      if (next < 0) {
        throw new Error('Only ' + formatQty_(current, part.unit) + ' on hand, so ' + formatNumber_(amount) +
          ' can\'t be taken out. Use "Set count" if the real count is different.');
      }
      if (next > 1e9) throw new Error('That quantity is too large.');
      writeFields_(inv, part.row, { quantity: next, updatedAt: new Date(), updatedBy: c.actor });
      const updated = readPartAt_(inv, part.row);
      const action = mode === 'add' ? 'Restocked' : mode === 'remove' ? 'Used' : 'Count set';
      logActivity_(c.ss, [{ part: updated, action: action, change: round_(next - current, 3), qtyAfter: next, note: note }], c.actor);
      const alert = processAlerts_(inv, c.settings, [updated]);
      return { part: partForClient_(updated), alert: alertForClient_(alert) };
    });
  });
}

/** Marks a part as ordered (so reminders stop listing it) or clears that mark. */
function apiSetOrdered(ctx, id, ordered) {
  return runApi_(ctx, function (c) {
    return withLock_(function () {
      const inv = ensureCoreColumns_(readInventory_(c.ss));
      const part = findPart_(inv, id);
      writeFields_(inv, part.row, { orderedAt: ordered ? new Date() : null, updatedAt: new Date(), updatedBy: c.actor });
      const updated = readPartAt_(inv, part.row);
      logActivity_(c.ss, [{ part: updated, action: ordered ? 'Marked ordered' : 'Order mark cleared', qtyAfter: updated.quantity }], c.actor);
      return { part: partForClient_(updated) };
    });
  });
}

function apiDeletePart(ctx, id) {
  return runApi_(ctx, function (c) {
    return withLock_(function () {
      const inv = readInventory_(c.ss);
      const part = findPart_(inv, id);
      if (inv.sheet.getMaxRows() - inv.headerRow <= 1) {
        // Sheets can't delete the last row below the header, so blank it out instead.
        inv.sheet.getRange(part.row, 1, 1, inv.sheet.getLastColumn()).clearContent();
      } else {
        inv.sheet.deleteRow(part.row);
      }
      logActivity_(c.ss, [{ part: part, action: 'Deleted', change: part.quantity ? -part.quantity : null, qtyAfter: 0 }], c.actor);
      return { id: part.id };
    });
  });
}

/** options = { partId, limit } */
function apiGetActivity(ctx, options) {
  return runApi_(ctx, function (c) {
    const opts = options || {};
    const limit = Math.min(Math.max(Number(opts.limit) || 100, 1), 500);
    return { entries: readActivity_(c.ss, String(opts.partId || ''), limit) };
  });
}

function apiSaveSettings(ctx, input) {
  return runApi_(ctx, function (c) {
    requireOwner_(c.user);
    const clean = cleanSettingsInput_(input || {});
    withLock_(function () {
      writeSettings_(c.ss, clean);
    });
    return { settings: settingsForClient_(getSettings_(c.ss), c.user) };
  });
}

function apiSendTestEmail(ctx) {
  return runApi_(ctx, function (c) {
    requireOwner_(c.user);
    return sendTestEmail_(c.ss, c.settings);
  });
}

function apiSendReminderNow(ctx) {
  return runApi_(ctx, function (c) {
    requireOwner_(c.user);
    return withLock_(function () {
      return sendReminder_(c.ss, readInventory_(c.ss), c.settings, true);
    });
  });
}

function runApi_(ctx, handler) {
  const ss = getSpreadsheet_();
  ensureSetup_(ss);
  const settings = getSettings_(ss);
  const user = checkAccess_(ctx, settings);
  return jsonSafe_(handler({ ss: ss, settings: settings, user: user, actor: actorName_(ctx, user) }));
}

function buildAppData_(ss, settings, user) {
  let inv = readInventory_(ss);
  if (needsIdRepair_(inv)) {
    inv = withLock_(function () {
      const fresh = ensureCoreColumns_(readInventory_(ss));
      assignMissingIds_(fresh);
      return fresh;
    });
  }
  const props = PropertiesService.getScriptProperties();
  const fields = {};
  FIELDS.forEach(function (field) {
    fields[field.key] = !!field.core || !!inv.cols[field.key];
  });
  let quota = null;
  try {
    quota = MailApp.getRemainingDailyQuota();
  } catch (err) {
    quota = null;
  }
  return {
    version: APP_VERSION,
    parts: inv.parts.map(partForClient_),
    fields: fields,
    settings: settingsForClient_(settings, user),
    user: { email: user.email, isOwner: user.isOwner },
    info: {
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      timeZone: ss.getSpreadsheetTimeZone(),
      sender: user.owner,
      emailQuota: quota,
      lastCheck: props.getProperty(PROP.lastCheck) || '',
      alertsInstalled: !!props.getProperty(PROP.triggersOwner),
      lastAlertError: parseJson_(props.getProperty(PROP.lastAlertError)),
    },
    serverTime: new Date().toISOString(),
  };
}

function partForClient_(p) {
  return {
    id: p.id,
    name: p.name,
    partNumber: p.partNumber,
    category: p.category,
    location: p.location,
    quantity: p.quantity,
    minQty: p.minQty,
    reorderQty: p.reorderQty,
    unit: p.unit,
    supplier: p.supplier,
    link: p.link,
    links: p.links,
    unitCost: p.unitCost,
    notes: p.notes,
    ordered: p.ordered,
    orderedAt: p.orderedAt,
    alerted: p.alerted,
    alertSentAt: p.alertSentAt,
    updatedAt: p.updatedAt,
    updatedBy: p.updatedBy,
    status: stockStatus_(p),
    needsReorder: needsReorder_(p),
  };
}

function settingsForClient_(s, user) {
  return {
    recipients: s.recipients,
    alertsEnabled: s.alertsEnabled,
    reminder: s.reminder,
    reminderHour: s.reminderHour,
    appName: s.appName,
    appUrl: s.appUrl,
    accessCodeSet: !!s.accessCode,
    accessCode: user.isOwner ? s.accessCode : '',
  };
}

function alertForClient_(result) {
  if (!result || !result.parts.length) return null;
  return {
    sent: result.sent,
    count: result.parts.length,
    recipients: result.recipients,
    error: result.error,
    skipped: result.skipped,
  };
}

// ---------------------------------------------------------------------------
// Automatic triggers (installed by setup)
// ---------------------------------------------------------------------------

/** Runs when someone edits the spreadsheet by hand. */
function onInventoryEdit(e) {
  try {
    if (!e || !e.range || typeof e.range.getSheet !== 'function') return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAMES.inventory) return;
    const ss = sheet.getParent();
    withLock_(function () {
      handleInventoryEdit_(ss, e);
    });
  } catch (err) {
    console.error('onInventoryEdit failed: ' + errorMessage_(err));
  }
}

function handleInventoryEdit_(ss, e) {
  const inv = ensureCoreColumns_(readInventory_(ss));
  const startRow = e.range.getRow();
  const endRow = startRow + e.range.getNumRows() - 1;
  const startCol = e.range.getColumn();
  const endCol = startCol + e.range.getNumColumns() - 1;
  if (endRow <= inv.headerRow) return;

  assignMissingIds_(inv);
  const edited = inv.parts.filter(function (p) { return p.row >= startRow && p.row <= endRow; });
  if (!edited.length) return;

  const touched = function (key) {
    const col = inv.cols[key];
    return !!col && col >= startCol && col <= endCol;
  };
  const managedCols = FIELDS.filter(function (f) { return f.managed && inv.cols[f.key]; })
    .map(function (f) { return inv.cols[f.key]; });
  let userEdit = false;
  for (let col = startCol; col <= endCol; col++) {
    if (managedCols.indexOf(col) === -1) userEdit = true;
  }

  let editor = '';
  try {
    editor = e.user && typeof e.user.getEmail === 'function' ? e.user.getEmail() : '';
  } catch (err) {
    editor = '';
  }
  editor = editor || 'Spreadsheet edit';

  if (userEdit) {
    const now = new Date();
    writeColumnValues_(inv, 'updatedAt', edited.map(function (p) { return { row: p.row, value: now }; }));
    writeColumnValues_(inv, 'updatedBy', edited.map(function (p) { return { row: p.row, value: editor }; }));
  }

  if (edited.length === 1 && startRow === endRow && startCol === endCol && touched('quantity')) {
    const part = edited[0];
    const before = toNumberOrNull_(e.oldValue);
    if (before !== part.quantity) {
      logActivity_(ss, [{
        part: part, action: 'Count set',
        change: before !== null && part.quantity !== null ? round_(part.quantity - before, 3) : null,
        qtyAfter: part.quantity, note: 'Edited in the spreadsheet',
      }], editor);
    }
  }

  if (touched('quantity') || touched('minQty')) {
    processAlerts_(inv, getSettings_(ss), edited);
  }
}

/** Runs every hour: catches anything missed and sends the reminder email. */
function hourlyCheck(e) {
  // Public functions can also be called from the web app page. The hourly
  // trigger and the owner run this freely; anyone else at most every 10
  // minutes, so nobody can keep the inventory busy by calling it in a loop.
  if (!isOwnTrigger_(e) && !currentUser_().isOwner) {
    const cache = CacheService.getScriptCache();
    if (cache.get('HOURLY_CHECK_CALLED')) return;
    cache.put('HOURLY_CHECK_CALLED', '1', 600);
  }
  const ss = getSpreadsheet_();
  if (!ss.getSheetByName(SHEET_NAMES.inventory)) return;
  withLock_(function () {
    const inv = ensureCoreColumns_(readInventory_(ss));
    assignMissingIds_(inv);
    const settings = getSettings_(ss);
    processAlerts_(inv, settings);
    sendReminderIfDue_(ss, inv, settings, new Date());
    PropertiesService.getScriptProperties().setProperty(PROP.lastCheck, new Date().toISOString());
  });
}

/** True for the event of one of this script's own triggers. */
function isOwnTrigger_(e) {
  const uid = e && e.triggerUid ? String(e.triggerUid) : '';
  if (!uid) return false;
  return ScriptApp.getProjectTriggers().some(function (t) { return String(t.getUniqueId()) === uid; });
}

function installTriggers_(ss, force) {
  const props = PropertiesService.getScriptProperties();
  const me = effectiveEmail_();
  const owner = props.getProperty(PROP.triggersOwner);
  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) !== -1;
  });
  // Triggers belong to the person who created them and are invisible to
  // everyone else, so don't quietly add a second set for another person.
  if (owner && owner !== me && !existing.length && !force) return { installed: false, owner: owner };
  existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onInventoryEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('hourlyCheck').timeBased().everyHours(1).create();
  props.setProperty(PROP.triggersOwner, me);
  return { installed: true, owner: me };
}

// ---------------------------------------------------------------------------
// Low-stock alerts and emails
// ---------------------------------------------------------------------------

/** True when the part is at or below its Min Qty (blank Min Qty = no alerts). */
function needsReorder_(p) {
  return p.minQty !== null && p.quantity !== null && p.quantity <= p.minQty;
}

function isOut_(p) {
  return p.quantity !== null && p.quantity <= 0;
}

function stockStatus_(p) {
  if (isOut_(p)) return 'out';
  if (needsReorder_(p)) return 'low';
  return 'ok';
}

/**
 * Emails parts that have just dropped to or below their Min Qty (one email for
 * all of them) and clears the alert and "ordered" marks of restocked parts.
 * Call while holding the lock. `parts` defaults to every part.
 */
function processAlerts_(inv, settings, parts) {
  const result = { sent: false, parts: [], recipients: [], error: '', skipped: '' };
  const newlyLow = [];
  (parts || inv.parts).forEach(function (p) {
    if (needsReorder_(p)) {
      if (!p.alerted) newlyLow.push(p);
      return;
    }
    const clear = {};
    if (p.alerted) clear.alertSentAt = null;
    if (p.ordered && stockStatus_(p) === 'ok') clear.orderedAt = null;
    if (Object.keys(clear).length) {
      writeFields_(inv, p.row, clear);
      if ('alertSentAt' in clear) { p.alertSentAt = null; p.alerted = false; }
      if ('orderedAt' in clear) { p.orderedAt = null; p.ordered = false; }
    }
  });
  if (!newlyLow.length) return result;

  newlyLow.sort(byUrgency_);
  result.parts = newlyLow.map(function (p) { return p.id; });
  if (!inv.cols.alertSentAt) {
    result.skipped = 'The Inventory tab has no "Alert Sent" column. Run Inventory > Set up / repair.';
    return result;
  }
  if (!settings.alertsEnabled) {
    result.skipped = 'Instant emails are turned off in Settings.';
    return result;
  }
  if (!settings.recipients.length) {
    result.skipped = 'No alert email recipients are set.';
    return result;
  }
  try {
    sendLowStockEmail_(inv.ss, settings, newlyLow, 'alert');
  } catch (err) {
    result.error = errorMessage_(err);
    recordAlertError_(result.error);
    return result;
  }
  const now = new Date();
  writeColumnValues_(inv, 'alertSentAt', newlyLow.map(function (p) { return { row: p.row, value: now }; }));
  newlyLow.forEach(function (p) {
    p.alertSentAt = now;
    p.alerted = true;
  });
  logActivity_(inv.ss, newlyLow.map(function (p) {
    return { part: p, action: 'Alert emailed', qtyAfter: p.quantity, note: 'To ' + settings.recipients.join(', ') };
  }), 'Automatic');
  PropertiesService.getScriptProperties().deleteProperty(PROP.lastAlertError);
  result.sent = true;
  result.recipients = settings.recipients.slice();
  return result;
}

/** Sends the scheduled reminder when it is due (at most once a day). */
function sendReminderIfDue_(ss, inv, settings, now) {
  if (settings.reminder === 'Off') return null;
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  if (Number(Utilities.formatDate(now, tz, 'H')) < settings.reminderHour) return null;
  const dow = dayOfWeek_(today);
  if (settings.reminder === 'Weekdays' && (dow === 0 || dow === 6)) return null;
  if (settings.reminder === 'Weekly' && dow !== 1) return null;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP.lastReminderDate) === today) return null;
  if (!settings.recipients.length) return null;
  try {
    const result = sendReminder_(ss, inv, settings, false);
    props.setProperty(PROP.lastReminderDate, today);
    return result;
  } catch (err) {
    recordAlertError_(errorMessage_(err));
    return null;
  }
}

/**
 * Emails the parts that are low and not marked as ordered (plus the ones on
 * order, for context). With force, also sends when every low part is ordered.
 */
function sendReminder_(ss, inv, settings, force) {
  const low = inv.parts.filter(needsReorder_).sort(byUrgency_);
  const toOrder = low.filter(function (p) { return !p.ordered; });
  const onOrder = low.filter(function (p) { return p.ordered; });
  if (!toOrder.length && !(force && onOrder.length)) return { sent: false, count: 0, onOrder: onOrder.length };
  if (!settings.recipients.length) throw new Error('Add an alert email recipient in Settings first.');
  sendLowStockEmail_(ss, settings, toOrder, 'reminder', { onOrder: onOrder });
  return { sent: true, count: toOrder.length, onOrder: onOrder.length, recipients: settings.recipients.slice() };
}

function sendTestEmail_(ss, settings) {
  if (!settings.recipients.length) throw new Error('Add an alert email recipient in Settings first.');
  const inv = readInventory_(ss);
  let sample = inv.parts.filter(needsReorder_).slice(0, 3);
  if (!sample.length) sample = inv.parts.filter(function (p) { return p.links.length; }).slice(0, 1);
  if (!sample.length) sample = [samplePart_()];
  sendLowStockEmail_(ss, settings, sample, 'test');
  return { recipients: settings.recipients.slice() };
}

function sendLowStockEmail_(ss, settings, parts, kind, extra) {
  const email = buildEmail_(ss, settings, parts, kind, extra);
  if (MailApp.getRemainingDailyQuota() < settings.recipients.length) {
    throw new Error('Google\'s daily email limit for this account has been reached. The app will try again later.');
  }
  MailApp.sendEmail({
    to: settings.recipients.join(','),
    subject: email.subject,
    body: email.text,
    htmlBody: email.html,
    name: settings.appName,
  });
  return email;
}

function buildEmail_(ss, settings, parts, kind, extra) {
  const onOrder = (extra && extra.onOrder) || [];
  const tz = ss.getSpreadsheetTimeZone();
  const appUrl = settings.appUrl;
  let subject;
  let heading;
  let intro;
  if (kind === 'test') {
    subject = 'Test: low-stock emails are working';
    heading = 'Low-stock emails are working';
    intro = 'This is a test. When a part drops to or below its Min Qty, an email like this goes to ' +
      settings.recipients.join(', ') + ' with the links for reordering it.';
  } else if (kind === 'reminder') {
    subject = parts.length
      ? 'Reminder: ' + countLabel_(parts.length, 'part needs', 'parts need') + ' ordering'
      : 'Reminder: ' + countLabel_(onOrder.length, 'low part is', 'low parts are') + ' on order';
    heading = parts.length ? 'Parts that still need ordering' : 'Low parts that are on order';
    intro = parts.length
      ? 'These parts are at or below their minimum and haven\'t been marked as ordered yet.'
      : 'Every low part has been marked as ordered.';
  } else if (parts.length === 1) {
    const p = parts[0];
    const out = isOut_(p);
    subject = (out ? 'Out of stock: ' : 'Low stock: ') + displayName_(p) +
      (out ? '' : ' (' + formatQty_(p.quantity, p.unit) + ' left)');
    heading = displayName_(p) + (out ? ' is out of stock' : ' is running low');
    intro = out
      ? 'There are none left.'
      : 'Only ' + formatQty_(p.quantity, p.unit) + ' left. The minimum is ' + formatNumber_(p.minQty) + '.';
  } else {
    subject = 'Low stock: ' + parts.length + ' parts need ordering';
    heading = parts.length + ' parts are running low';
    intro = 'These parts have dropped to or below their minimum stock level.';
  }
  subject = '[' + settings.appName + '] ' + subject;

  const font = '-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif';
  const cards = partsHtml_(parts, appUrl, tz);
  const orderedCards = onOrder.length
    ? '<h2 style="font-size:16px;margin:24px 0 10px;color:#1f2937;">Already on order</h2>' + partsHtml_(onOrder, appUrl, tz)
    : '';
  const footerLinks = [
    appUrl ? linkHtml_(appUrl, 'Open the inventory app') : '',
    linkHtml_(ss.getUrl(), 'Open the spreadsheet'),
  ].filter(String).join(' &nbsp;&middot;&nbsp; ');
  const footerNote = 'You get these emails because your address is listed under "Alert email recipients" ' +
    'in the Settings tab of "' + ss.getName() + '".';
  const html =
    '<div style="background:#f3f4f6;padding:24px 12px;font-family:' + font + ';color:#1f2937;">' +
    '<div style="max-width:600px;margin:0 auto;">' +
    '<div style="font-size:13px;font-weight:600;color:' + COLORS.accent + ';margin-bottom:6px;">' +
    escapeHtml_(settings.appName) + '</div>' +
    '<h1 style="font-size:22px;line-height:1.3;margin:0 0 8px;color:#111827;">' + escapeHtml_(heading) + '</h1>' +
    '<p style="font-size:15px;line-height:1.5;margin:0 0 20px;color:#374151;">' + escapeHtml_(intro) + '</p>' +
    cards + orderedCards +
    '<p style="font-size:14px;margin:24px 0 8px;">' + footerLinks + '</p>' +
    '<p style="font-size:12px;line-height:1.5;margin:0;color:#6b7280;">' + escapeHtml_(footerNote) + '</p>' +
    '</div></div>';

  const lines = [heading, '', intro, ''];
  pushPartsText_(lines, parts);
  if (onOrder.length) {
    lines.push('Already on order:', '');
    pushPartsText_(lines, onOrder);
  }
  if (appUrl) lines.push('Inventory app: ' + appUrl);
  lines.push('Spreadsheet: ' + ss.getUrl(), '', footerNote);
  return { subject: subject, html: html, text: lines.join('\n') };
}

function partsHtml_(parts, appUrl, tz) {
  const html = parts.slice(0, EMAIL_CARDS).map(function (p) { return partCardHtml_(p, appUrl, tz); }).join('');
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
    (more ? '<div style="font-size:13px;color:#6b7280;margin-top:8px;">Plus ' + more + ' more. Open the app to see them all.</div>' : '') +
    '</div>';
}

function pushPartsText_(lines, parts) {
  parts.slice(0, EMAIL_CARDS).forEach(function (p) { lines.push(partText_(p), ''); });
  const rest = parts.slice(EMAIL_CARDS);
  rest.slice(0, EMAIL_LIST).forEach(function (p) { lines.push('* ' + displayName_(p) + ': ' + shortStock_(p)); });
  if (rest.length > EMAIL_LIST) lines.push('...plus ' + (rest.length - EMAIL_LIST) + ' more. Open the app to see them all.');
  if (rest.length) lines.push('');
}

function shortStock_(p) {
  return (isOut_(p) ? 'out of stock' : formatQty_(p.quantity, p.unit) + ' left') +
    (p.minQty !== null ? ' (min ' + formatNumber_(p.minQty) + ')' : '');
}

function partCardHtml_(p, appUrl, tz) {
  const out = isOut_(p);
  const color = out ? '#b91c1c' : '#b45309';
  const status = out ? 'Out of stock' : formatQty_(p.quantity, p.unit) + ' on hand';
  const details = [
    p.minQty !== null ? 'min ' + formatNumber_(p.minQty) : '',
    p.reorderQty !== null ? 'order ' + formatQty_(p.reorderQty, p.unit) : '',
    p.ordered ? 'ordered' + (p.orderedAt ? ' ' + Utilities.formatDate(p.orderedAt, tz, 'MMM d') : '') : '',
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
    (appUrl ? '<div style="margin-top:10px;"><a href="' + escapeHtml_(partUrl_(appUrl, p.id)) +
      '" style="font-size:13px;color:#6b7280;">View in the app</a></div>' : '') +
    '</td></tr></table>';
}

function partText_(p) {
  const qty = isOut_(p) ? 'OUT OF STOCK' : formatQty_(p.quantity, p.unit) + ' on hand';
  const extra = [
    p.minQty !== null ? 'min ' + formatNumber_(p.minQty) : '',
    p.reorderQty !== null ? 'order ' + formatQty_(p.reorderQty, p.unit) : '',
    p.ordered ? 'ordered' : '',
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

function samplePart_() {
  return {
    id: 'P-0000', name: 'Example part (you have no parts yet)', partNumber: 'ABC-123', location: 'Shelf 1',
    supplier: 'Example Supplier', quantity: 2, minQty: 5, reorderQty: 20, unit: 'ea', notes: '',
    links: ['https://www.example.com/'], ordered: false, orderedAt: null,
  };
}

function describeAlertResult_(r) {
  const n = r.parts.length;
  if (r.sent) return 'Emailed a low-stock alert for ' + countLabel_(n, 'part', 'parts') + ' to ' + r.recipients.join(', ') + '.';
  if (r.error) return 'Couldn\'t send the alert email: ' + r.error;
  if (r.skipped) return countLabel_(n, 'part is', 'parts are') + ' low, but no email was sent. ' + r.skipped;
  return 'No new low-stock parts. Parts that were already emailed about aren\'t emailed again until they are ' +
    'restocked. Use "Send reminder email now" for a list of everything that is low.';
}

function recordAlertError_(message) {
  PropertiesService.getScriptProperties().setProperty(PROP.lastAlertError,
    JSON.stringify({ at: new Date().toISOString(), message: String(message).slice(0, 500) }));
}

// ---------------------------------------------------------------------------
// Inventory tab: reading and writing
// ---------------------------------------------------------------------------

function readInventory_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.inventory);
  if (!sheet) throw new Error('The Inventory tab is missing. Open the spreadsheet and choose Inventory > Set up / repair.');
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : [];
  const headerIndex = findHeaderRow_(values);
  const cols = mapColumns_(values[headerIndex] || []);
  const headerRow = headerIndex + 1;
  const firstDataRow = headerRow + 1;
  const links = cols.link && lastRow >= firstDataRow
    ? readLinkUrls_(sheet, firstDataRow, cols.link, lastRow - firstDataRow + 1)
    : [];
  const parts = [];
  for (let i = headerIndex + 1; i < values.length; i++) {
    const part = rowToPart_(values[i], cols, i + 1, links[i - headerIndex - 1]);
    if (part) parts.push(part);
  }
  return { ss: ss, sheet: sheet, values: values, headerRow: headerRow, cols: cols, parts: parts };
}

function readPartAt_(inv, row) {
  const values = inv.sheet.getRange(row, 1, 1, Math.max(inv.sheet.getLastColumn(), 1)).getValues()[0];
  const links = inv.cols.link ? readLinkUrls_(inv.sheet, row, inv.cols.link, 1)[0] : null;
  return rowToPart_(values, inv.cols, row, links);
}

function rowToPart_(row, cols, rowNumber, extraLinks) {
  const get = function (key) { return cols[key] ? row[cols[key] - 1] : ''; };
  const id = cellText_(get('id'));
  const name = cellText_(get('name'));
  const partNumber = cellText_(get('partNumber'));
  if (!id && !name && !partNumber) return null;
  // A "Total" row under the list isn't a part.
  if (!id && !partNumber && /^(grand\s*|sub\s*)?totals?\s*:?$/i.test(name)) return null;
  const link = cellText_(get('link'));
  const links = [];
  const seen = {};
  parseLinks_(link).concat(extraLinks || []).forEach(function (url) {
    const clean = normalizeUrl_(url);
    // Sheets turns typed URLs into links, sometimes as http:// or with a
    // trailing slash, so compare without those to avoid listing a link twice.
    const key = clean.replace(/^https?:\/\//i, '').replace(/\/+$/, '').toLowerCase();
    if (clean && !seen[key]) {
      seen[key] = true;
      links.push(clean);
    }
  });
  const orderedRaw = get('orderedAt');
  const alertRaw = get('alertSentAt');
  return {
    row: rowNumber,
    id: id,
    name: name,
    partNumber: partNumber,
    category: cellText_(get('category')),
    location: cellText_(get('location')),
    quantity: toNumberOrNull_(get('quantity')),
    minQty: toNumberOrNull_(get('minQty')),
    reorderQty: toNumberOrNull_(get('reorderQty')),
    unit: cellText_(get('unit')),
    supplier: cellText_(get('supplier')),
    link: link,
    links: links.slice(0, MAX_LINKS),
    unitCost: toNumberOrNull_(get('unitCost')),
    notes: cellText_(get('notes')),
    // A date, a ticked box or text like "yes" counts as set; blank, unticked or "no" doesn't.
    ordered: isMarked_(orderedRaw),
    orderedAt: toDateOrNull_(orderedRaw),
    alerted: isMarked_(alertRaw),
    alertSentAt: toDateOrNull_(alertRaw),
    updatedAt: toDateOrNull_(get('updatedAt')),
    updatedBy: cellText_(get('updatedBy')),
  };
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
    const names = [field.header].concat(field.aliases || []).map(normalizeHeader_);
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

/** Adds any missing core columns to the end of the Inventory tab. Returns a fresh read if it did. */
function ensureCoreColumns_(inv) {
  const missing = FIELDS.filter(function (f) { return f.core && !inv.cols[f.key]; });
  if (!missing.length) return inv;
  addColumns_(inv.sheet, inv.headerRow, missing);
  return readInventory_(inv.ss);
}

function addColumns_(sheet, headerRow, fields) {
  const start = sheet.getLastColumn() + 1;
  const needed = start + fields.length - 1;
  if (sheet.getMaxColumns() < needed) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  sheet.getRange(headerRow, start, 1, fields.length).setValues([fields.map(function (f) { return f.header; })]);
  styleHeaders_(sheet, headerRow, start, fields);
}

/** Writes several fields of one row. Only fields that have a column are written. */
function writeFields_(inv, row, fields) {
  const cells = Object.keys(fields)
    .filter(function (key) { return FIELD_BY_KEY[key] && inv.cols[key]; })
    .map(function (key) { return { key: key, col: inv.cols[key], value: fields[key] }; })
    .sort(function (a, b) { return a.col - b.col; });
  // One write per run of neighbouring columns keeps this fast without
  // touching any of your own columns in between.
  let i = 0;
  while (i < cells.length) {
    let j = i;
    while (j + 1 < cells.length && cells[j + 1].col === cells[j].col + 1) j++;
    const run = cells.slice(i, j + 1);
    const range = inv.sheet.getRange(row, run[0].col, 1, run.length);
    const plain = plainTextCells_(range)[0];
    range.setValues([run.map(function (c, k) { return toCell_(c.key, c.value, plain[k]); })]);
    i = j + 1;
  }
}

/** Writes one field for several rows with a single range write. updates = [{ row, value }] */
function writeColumnValues_(inv, key, updates) {
  const col = inv.cols[key];
  if (!col || !updates.length) return;
  let min = Infinity;
  let max = -Infinity;
  updates.forEach(function (u) {
    min = Math.min(min, u.row);
    max = Math.max(max, u.row);
  });
  const range = inv.sheet.getRange(min, col, max - min + 1, 1);
  const plain = plainTextCells_(range);
  if (updates.length === 1) {
    range.setValue(toCell_(key, updates[0].value, plain[0][0]));
    return;
  }
  const block = range.getValues().map(function (r, k) {
    return [typeof r[0] === 'string' && !plain[k][0] ? safeCell_(r[0]) : r[0]];
  });
  updates.forEach(function (u) { block[u.row - min][0] = toCell_(key, u.value, plain[u.row - min][0]); });
  range.setValues(block);
}

function toCell_(key, value, plainText) {
  const field = FIELD_BY_KEY[key];
  if (value === null || value === undefined || value === '') return '';
  if (field.type === 'number') return Number(value);
  if (field.type === 'date') return isDate_(value) ? value : new Date(value);
  const text = String(value);
  // Plain-text cells keep what's written as is, apostrophes included, so there
  // only a leading "=" needs guarding.
  if (plainText) return text.charAt(0) === '=' ? "'" + text : text;
  return safeCell_(text);
}

/** Which cells are formatted as Plain text (Format > Number > Plain text). */
function plainTextCells_(range) {
  let formats = null;
  try {
    formats = range.getNumberFormats();
  } catch (err) {
    formats = null;
  }
  const out = [];
  for (let r = 0; r < range.getNumRows(); r++) {
    out.push([]);
    for (let c = 0; c < range.getNumColumns(); c++) out[r].push(!!formats && formats[r][c] === '@');
  }
  return out;
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

/** The first empty row after the last part. */
function nextEmptyRow_(inv) {
  let lastPartRow = inv.headerRow;
  inv.parts.forEach(function (p) { lastPartRow = Math.max(lastPartRow, p.row); });
  const row = lastPartRow + 1;
  const existing = inv.values[row - 1];
  const occupied = existing && Object.keys(inv.cols).some(function (key) {
    return !isBlank_(existing[inv.cols[key] - 1]);
  });
  // Something like a totals row right under the list: add the part above it.
  if (occupied) inv.sheet.insertRowAfter(lastPartRow);
  ensureRows_(inv.sheet, row);
  return row;
}

function ensureRows_(sheet, lastNeededRow) {
  const max = sheet.getMaxRows();
  if (lastNeededRow > max) sheet.insertRowsAfter(max, lastNeededRow - max);
}

function findPart_(inv, id) {
  const key = String(id || '').trim().toUpperCase();
  const part = key ? inv.parts.find(function (p) { return p.id.toUpperCase() === key; }) : null;
  if (!part) throw new Error('That part wasn\'t found. It may have been deleted or changed. Refresh and try again.');
  return part;
}

function needsIdRepair_(inv) {
  if (!inv.cols.id) return inv.parts.length > 0;
  const seen = {};
  return inv.parts.some(function (p) {
    const key = p.id.toUpperCase();
    if (!key || seen[key]) return true;
    seen[key] = true;
    return false;
  });
}

/** Gives every part a unique ID (rows added by hand, or copied rows that share one). */
function assignMissingIds_(inv) {
  if (!inv.cols.id) return 0;
  const seen = {};
  const updates = [];
  let next = firstFreeIdNumber_(inv);
  inv.parts.forEach(function (p) {
    const key = p.id.toUpperCase();
    if (key && !seen[key]) {
      seen[key] = true;
      return;
    }
    let id;
    do {
      id = formatId_(next++);
    } while (seen[id]);
    seen[id] = true;
    p.id = id;
    updates.push({ row: p.row, value: id });
  });
  if (!updates.length) return 0;
  writeColumnValues_(inv, 'id', updates);
  PropertiesService.getScriptProperties().setProperty(PROP.nextId, String(next));
  return updates.length;
}

function nextId_(inv) {
  const n = firstFreeIdNumber_(inv);
  PropertiesService.getScriptProperties().setProperty(PROP.nextId, String(n + 1));
  return formatId_(n);
}

/** IDs are never reused, even after a part is deleted, so old log entries stay unambiguous. */
function firstFreeIdNumber_(inv) {
  let n = Number(PropertiesService.getScriptProperties().getProperty(PROP.nextId)) || 1;
  inv.parts.forEach(function (p) {
    const match = /^P-(\d+)$/i.exec(p.id);
    if (match) n = Math.max(n, Number(match[1]) + 1);
  });
  return n;
}

function formatId_(n) {
  return 'P-' + String(n).padStart(4, '0');
}

// ---------------------------------------------------------------------------
// Input checks
// ---------------------------------------------------------------------------

function cleanPartInput_(input, isNew) {
  const out = {};
  Object.keys(TEXT_LIMITS).forEach(function (key) {
    if (!(key in input)) return;
    const value = cleanString_(input[key], key === 'notes');
    if (value.length > TEXT_LIMITS[key]) {
      throw new Error(FIELD_BY_KEY[key].header + ' is too long (' + TEXT_LIMITS[key] + ' characters at most).');
    }
    out[key] = value;
  });
  if ((isNew || 'name' in out) && !out.name) throw new Error('Give the part a name.');
  Object.keys(NUMBER_INPUTS).forEach(function (key) {
    if (!(key in input)) return;
    const value = parseInputNumber_(input[key], NUMBER_INPUTS[key], key === 'unitCost' ? 4 : 3);
    if (value === null && key === 'quantity') throw new Error('Enter the quantity on hand (0 if there are none).');
    out[key] = value;
  });
  if ('link' in input) out.link = cleanLinks_(input.link);
  if (isNew && !('quantity' in out)) out.quantity = 0;
  return out;
}

function parseInputNumber_(value, label, decimals) {
  if (value === null || value === undefined || String(value).trim() === '') return null;
  const n = typeof value === 'number' ? value : Number(String(value).replace(/,/g, '').trim());
  if (!isFinite(n)) throw new Error(label + ' must be a number.');
  if (n < 0) throw new Error(label + ' can\'t be negative.');
  if (n > 1e9) throw new Error(label + ' is too large.');
  return round_(n, decimals);
}

function cleanLinks_(value) {
  const tokens = String(value == null ? '' : value).split(/\s+/).filter(String);
  if (tokens.length > MAX_LINKS) throw new Error('Add at most ' + MAX_LINKS + ' order links.');
  const urls = tokens.map(function (token) {
    const url = normalizeUrl_(token);
    if (!url) {
      throw new Error('"' + truncate_(token, 60) + '" doesn\'t look like a web link. ' +
        'Paste the full address, like https://www.example.com/part-page');
    }
    return url;
  });
  const joined = urls.join('\n');
  if (joined.length > 2000) throw new Error('The order links are too long.');
  return joined;
}

function cleanSettingsInput_(input) {
  const out = {};
  if ('recipients' in input) {
    const list = Array.isArray(input.recipients) ? input.recipients.map(String) : parseEmails_(input.recipients);
    const bad = list.filter(function (email) { return !isEmail_(email); });
    if (bad.length) throw new Error('"' + truncate_(bad[0], 80) + '" isn\'t a valid email address.');
    if (list.length > 20) throw new Error('Use 20 alert email addresses at most.');
    out.recipients = uniqueEmails_(list);
  }
  if ('alertsEnabled' in input) out.alertsEnabled = input.alertsEnabled === true;
  if ('reminder' in input) {
    if (REMINDER_CHOICES.indexOf(input.reminder) === -1) throw new Error('Pick how often to send the reminder.');
    out.reminder = input.reminder;
  }
  if ('reminderHour' in input) {
    const hour = Number(input.reminderHour);
    if (!(hour >= 0 && hour <= 23 && Math.floor(hour) === hour)) throw new Error('The reminder hour must be 0 to 23.');
    out.reminderHour = hour;
  }
  if ('appName' in input) {
    const name = cleanString_(input.appName, false);
    if (!name) throw new Error('The app name can\'t be empty.');
    if (name.length > 60) throw new Error('Keep the app name under 60 characters.');
    out.appName = name;
  }
  if ('accessCode' in input) {
    const code = String(input.accessCode == null ? '' : input.accessCode).trim();
    if (code && code.length < 6) throw new Error('Use at least 6 characters for the access code.');
    if (code.length > 64) throw new Error('Keep the access code under 64 characters.');
    out.accessCode = code;
  }
  if ('appUrl' in input) {
    const raw = String(input.appUrl == null ? '' : input.appUrl).trim();
    const url = raw ? normalizeUrl_(raw) : '';
    if (raw && !(url && /^https:\/\//i.test(url))) throw new Error('The web app URL should start with https://');
    out.appUrl = url;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Settings tab
// ---------------------------------------------------------------------------

function getSettings_(ss) {
  const found = readSettingsRows_(ss.getSheetByName(SHEET_NAMES.settings));
  const settings = {};
  SETTINGS.forEach(function (def) {
    settings[def.key] = parseSettingValue_(def, found[def.key] ? found[def.key].value : undefined);
  });
  return settings;
}

function readSettingsRows_(sheet) {
  const found = {};
  if (!sheet || sheet.getLastRow() < 1) return found;
  sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues().forEach(function (row, i) {
    const def = findSettingDef_(row[0]);
    if (def && !found[def.key]) found[def.key] = { row: i + 1, value: row[1] };
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

function parseSettingValue_(def, raw) {
  if (def.type === 'emails') return uniqueEmails_(parseEmails_(raw).filter(isEmail_));
  if (def.type === 'boolean') return isBlank_(raw) ? def.default : toBoolean_(raw);
  if (def.type === 'choice') {
    const text = String(raw == null ? '' : raw).trim().toLowerCase();
    return def.choices.find(function (c) { return text && text.indexOf(c.toLowerCase()) === 0; }) || def.default;
  }
  if (def.type === 'hour') {
    const hour = typeof raw === 'number' ? raw : parseInt(raw, 10);
    return hour >= 0 && hour <= 23 ? Math.floor(hour) : def.default;
  }
  if (def.type === 'url') return normalizeUrl_(raw) || '';
  return cellText_(raw) || def.default;
}

function writeSettings_(ss, values) {
  const sheet = ss.getSheetByName(SHEET_NAMES.settings) || createSettingsSheet_(ss);
  const found = readSettingsRows_(sheet);
  Object.keys(values).forEach(function (key) {
    const def = SETTINGS.find(function (d) { return d.key === key; });
    if (!def) return;
    let row = found[key] && found[key].row;
    if (!row) row = appendSettingRow_(sheet, def);
    sheet.getRange(row, 2).setValue(settingCellValue_(def, values[key]));
  });
}

function settingCellValue_(def, value) {
  if (def.type === 'emails') return safeCell_((value || []).join(', '));
  if (def.type === 'boolean') return value === true;
  if (def.type === 'hour') return Number(value);
  return safeCell_(String(value == null ? '' : value));
}

function defaultSettingValue_(def) {
  if (def.type === 'emails') {
    const me = effectiveEmail_();
    return me ? [me] : [];
  }
  return def.default;
}

function appendSettingRow_(sheet, def) {
  const row = sheet.getLastRow() + 1;
  ensureRows_(sheet, row);
  sheet.getRange(row, 1, 1, 3).setValues([[def.label, '', def.help].map(safeCell_)]);
  sheet.getRange(row, 1).setFontWeight('bold');
  sheet.getRange(row, 3).setFontColor('#6b7280').setWrap(true);
  const cell = sheet.getRange(row, 2);
  if (def.type === 'boolean') cell.insertCheckboxes();
  if (def.type === 'choice') {
    cell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(def.choices, true).setAllowInvalid(false).build());
  }
  if (def.type === 'hour') {
    cell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 23).setAllowInvalid(false).setHelpText('A whole number from 0 to 23.').build());
  }
  return row;
}

// ---------------------------------------------------------------------------
// Activity Log tab
// ---------------------------------------------------------------------------

/** entries = [{ part, action, change, qtyAfter, note }] */
function logActivity_(ss, entries, by) {
  if (!entries || !entries.length) return;
  const sheet = ss.getSheetByName(SHEET_NAMES.log) || createLogSheet_(ss);
  const now = new Date();
  const rows = entries.map(function (e) {
    return [
      now,
      safeCell_(e.part.id || ''),
      safeCell_(displayName_(e.part)),
      safeCell_(e.action),
      typeof e.change === 'number' ? e.change : '',
      typeof e.qtyAfter === 'number' ? e.qtyAfter : '',
      safeCell_(by || ''),
      safeCell_(e.note || ''),
    ];
  });
  const start = sheet.getLastRow() + 1;
  ensureRows_(sheet, start + rows.length - 1);
  sheet.getRange(start, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

function readActivity_(ss, partId, limit) {
  const sheet = ss.getSheetByName(SHEET_NAMES.log);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  // When filtering by part, look further back so a quiet part still shows its history.
  const count = Math.min(last - 1, partId ? 5000 : limit);
  const width = Math.min(LOG_HEADERS.length, sheet.getMaxColumns());
  const values = sheet.getRange(last - count + 1, 1, count, width).getValues();
  const wanted = partId.toUpperCase();
  const entries = [];
  for (let i = values.length - 1; i >= 0 && entries.length < limit; i--) {
    const r = values[i];
    if (wanted && cellText_(r[1]).toUpperCase() !== wanted) continue;
    entries.push({
      at: toDateOrNull_(r[0]),
      partId: cellText_(r[1]),
      part: cellText_(r[2]),
      action: cellText_(r[3]),
      change: toNumberOrNull_(r[4]),
      qtyAfter: toNumberOrNull_(r[5]),
      by: cellText_(r[6]),
      note: cellText_(r[7]),
    });
  }
  return entries;
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

/** Creates the tabs the app needs, or adds what's missing. Call while holding the lock. */
function setupSpreadsheet_(ss) {
  const report = { created: [], addedColumns: [], addedSettings: [], missingOptional: [] };
  const isEmpty = function (s) { return s.getLastRow() === 0 && s.getLastColumn() === 0; };
  let sheet = ss.getSheetByName(SHEET_NAMES.inventory);
  if (!sheet || isEmpty(sheet)) {
    if (!sheet) {
      const sheets = ss.getSheets();
      sheet = sheets.length === 1 && isEmpty(sheets[0])
        ? sheets[0].setName(SHEET_NAMES.inventory)
        : ss.insertSheet(SHEET_NAMES.inventory, 0);
    }
    createInventoryLayout_(sheet);
    report.created.push(SHEET_NAMES.inventory);
  } else {
    const inv = readInventory_(ss);
    const missing = FIELDS.filter(function (f) { return f.core && !inv.cols[f.key]; });
    if (missing.length) addColumns_(sheet, inv.headerRow, missing);
    report.addedColumns = missing.map(function (f) { return f.header; });
    report.missingOptional = FIELDS.filter(function (f) { return !f.core && !inv.cols[f.key]; })
      .map(function (f) { return f.header; });
  }
  const inv = readInventory_(ss);
  applyStockHighlighting_(sheet, inv.headerRow, inv.cols);
  assignMissingIds_(inv);

  const settingsSheet = ss.getSheetByName(SHEET_NAMES.settings);
  if (!settingsSheet) {
    createSettingsSheet_(ss);
    report.created.push(SHEET_NAMES.settings);
  } else {
    const found = readSettingsRows_(settingsSheet);
    SETTINGS.forEach(function (def) {
      if (found[def.key]) return;
      const row = appendSettingRow_(settingsSheet, def);
      settingsSheet.getRange(row, 2).setValue(settingCellValue_(def, defaultSettingValue_(def)));
      report.addedSettings.push(def.label);
    });
  }

  if (!ss.getSheetByName(SHEET_NAMES.log)) {
    createLogSheet_(ss);
    report.created.push(SHEET_NAMES.log);
  }
  PropertiesService.getScriptProperties().setProperty(PROP.spreadsheetId, ss.getId());
  return report;
}

/**
 * Sets the spreadsheet up the first time the web app or sidebar is used,
 * so the Set up menu item is optional.
 */
function ensureSetup_(ss) {
  if (ss.getSheetByName(SHEET_NAMES.inventory) && ss.getSheetByName(SHEET_NAMES.settings)) return false;
  withLock_(function () {
    setupSpreadsheet_(ss);
  });
  if (!PropertiesService.getScriptProperties().getProperty(PROP.triggersOwner)) {
    try {
      installTriggers_(ss, false);
    } catch (err) {
      console.warn('Could not turn on automatic alerts: ' + errorMessage_(err));
    }
  }
  return true;
}

function createInventoryLayout_(sheet) {
  sheet.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS.map(function (f) { return f.header; })]);
  styleHeaders_(sheet, 1, 1, FIELDS);
  sheet.setFrozenRows(1);
  sheet.setTabColor(COLORS.accent);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  ensureRows_(sheet, rows + 1);
  const col = function (key) { return FIELDS.indexOf(FIELD_BY_KEY[key]) + 1; };
  const numberRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('Enter a number (0 or more).')
    .build();
  ['quantity', 'minQty', 'reorderQty', 'unitCost'].forEach(function (key) {
    sheet.getRange(2, col(key), rows, 1).setDataValidation(numberRule);
  });
  sheet.getRange(2, col('unitCost'), rows, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, col('orderedAt'), rows, 1).setNumberFormat('yyyy-mm-dd');
  ['alertSentAt', 'updatedAt'].forEach(function (key) {
    sheet.getRange(2, col(key), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  });
  sheet.getRange(2, col('link'), rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function styleHeaders_(sheet, headerRow, startCol, fields) {
  fields.forEach(function (field, i) {
    const cell = sheet.getRange(headerRow, startCol + i);
    cell.setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground(field.managed ? COLORS.managedHeader : COLORS.header);
    if (field.note) cell.setNote(field.note);
    if (field.width) sheet.setColumnWidth(startCol + i, field.width);
  });
}

/** Colours low rows amber and out-of-stock rows red, right in the spreadsheet. */
function applyStockHighlighting_(sheet, headerRow, cols) {
  if (!cols.quantity) return;
  const first = headerRow + 1;
  ensureRows_(sheet, first);
  const qty = '$' + columnLetter_(cols.quantity) + first;
  const range = sheet.getRange(first, 1, sheet.getMaxRows() - headerRow, sheet.getMaxColumns());
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + qty + '),' + qty + '<=0)')
      .setBackground(COLORS.outRow)
      .setRanges([range])
      .build(),
  ];
  if (cols.minQty) {
    const min = '$' + columnLetter_(cols.minQty) + first;
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + qty + '),ISNUMBER(' + min + '),' + qty + '<=' + min + ')')
      .setBackground(COLORS.lowRow)
      .setRanges([range])
      .build());
  }
  const others = sheet.getConditionalFormatRules().filter(function (rule) { return !isStockRule_(rule); });
  sheet.setConditionalFormatRules(rules.concat(others));
}

function isStockRule_(rule) {
  const condition = rule.getBooleanCondition();
  if (!condition) return false;
  const values = condition.getCriteriaValues();
  return /^=AND\(ISNUMBER\(\$[A-Z]+\d+\),/.test(String((values && values[0]) || ''));
}

function createSettingsSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_NAMES.settings);
  sheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(COLORS.header);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 560);
  sheet.setTabColor('#6b7280');
  SETTINGS.forEach(function (def) {
    const row = appendSettingRow_(sheet, def);
    sheet.getRange(row, 2).setValue(settingCellValue_(def, defaultSettingValue_(def)));
  });
  return sheet;
}

function createLogSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_NAMES.log);
  sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(COLORS.header);
  sheet.setFrozenRows(1);
  [150, 80, 240, 130, 70, 80, 200, 320].forEach(function (width, i) { sheet.setColumnWidth(i + 1, width); });
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.setTabColor('#9ca3af');
  return sheet;
}

// ---------------------------------------------------------------------------
// Access, users and locking
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    ss = null;
  }
  if (ss) {
    if (props.getProperty(PROP.spreadsheetId) !== ss.getId()) props.setProperty(PROP.spreadsheetId, ss.getId());
    return ss;
  }
  const id = props.getProperty(PROP.spreadsheetId);
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('Couldn\'t find the spreadsheet. Open it and choose Inventory > Set up / repair.');
}

/**
 * email:   who is using the app, when Google shares it (blank for people
 *          outside your Google Workspace domain)
 * owner:   the account the script runs as, which sends the emails
 * isOwner: true for that account, and for anyone using the spreadsheet menu
 *          or sidebar (those run under their own account)
 */
function currentUser_() {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    email = '';
  }
  const owner = effectiveEmail_();
  return { email: email, owner: owner, isOwner: !!email && !!owner && email.toLowerCase() === owner.toLowerCase() };
}

function effectiveEmail_() {
  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

function needsAccessCode_(settings) {
  return !!settings.accessCode && !currentUser_().isOwner;
}

function checkAccess_(ctx, settings) {
  const user = currentUser_();
  if (!settings.accessCode || user.isOwner) return user;
  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get('ACCESS_FAILURES')) || 0;
  if (failures >= MAX_ACCESS_FAILURES) {
    throw new Error('ACCESS_LOCKED: Too many wrong access codes were tried. Wait 10 minutes and try again.');
  }
  const code = String((ctx && ctx.code) || '').trim();
  if (code && safeEqual_(code, settings.accessCode)) return user;
  if (code) cache.put('ACCESS_FAILURES', String(failures + 1), 600);
  throw new Error('ACCESS_DENIED: ' + (code ? 'That access code isn\'t right.' : 'Enter the access code to open this app.'));
}

function requireOwner_(user) {
  if (!user.isOwner) {
    throw new Error('Only ' + (user.owner || 'the person who set up the app') + ' can change this.');
  }
}

/** Menu and setup functions are public, so make sure they aren't being called through the web app. */
function requireSheetUser_() {
  if (!currentUser_().isOwner) throw new Error('This can only be run from the spreadsheet.');
}

function actorName_(ctx, user) {
  if (user.email) return user.email;
  const name = cleanString_(ctx && ctx.actor, false).slice(0, 60);
  return name || 'Web app user';
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

/** Saves the web app's /exec link the first time the app is opened, for use in emails. */
function rememberAppUrl_(ss, settings) {
  const url = currentServiceUrl_();
  if (!url || !/\/exec$/.test(url) || settings.appUrl === url) return false;
  if (settings.appUrl && !/^https:\/\/script\.google\.com\//.test(settings.appUrl)) return false;
  try {
    writeSettings_(ss, { appUrl: url });
    return true;
  } catch (err) {
    return false;
  }
}

function currentServiceUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

// ---------------------------------------------------------------------------
// Spreadsheet dialogs
// ---------------------------------------------------------------------------

function runMenu_(fn) {
  requireSheetUser_();
  const ss = getSpreadsheet_();
  try {
    fn(ss);
  } catch (err) {
    alert_('Parts Inventory', errorMessage_(err));
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

function showDialog_(bodyHtml, title, width, height) {
  const html = '<div style="font:14px/1.5 Arial,Helvetica,sans-serif;color:#1f2937;">' + bodyHtml + '</div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(width).setHeight(height), title);
}

function openLinkHtml_(url) {
  return '<p><a href="' + escapeHtml_(url) + '" target="_blank" rel="noopener" ' +
    'onclick="google.script.host.close()" style="font-size:16px;font-weight:bold;color:' + COLORS.accent +
    ';">Open the inventory app &rarr;</a></p>' +
    '<p style="color:#6b7280;">Tip: bookmark it, or add it to your phone\'s home screen.</p>' +
    '<script>var w = window.open(' + jsonForHtml_(url) + ', "_blank"); if (w) google.script.host.close();</script>';
}

function deployStepsHtml_() {
  return '<ol style="padding-left:20px;">' +
    '<li>Click <b>Extensions &gt; Apps Script</b>.</li>' +
    '<li>Click <b>Deploy &gt; New deployment</b>. Next to "Select type", click the gear and choose <b>Web app</b>.</li>' +
    '<li>Set <b>Execute as</b> to <b>Me</b>.</li>' +
    '<li>Set <b>Who has access</b>: <i>Anyone within your company</i> if you use Google Workspace; ' +
    'otherwise <i>Anyone with a Google account</i>, or <i>Anyone</i> for coworkers without Google accounts ' +
    '(then set an <b>App access code</b> in the Settings tab).</li>' +
    '<li>Click <b>Deploy</b>, allow access when Google asks, then copy the <b>Web app URL</b> and open it.</li>' +
    '</ol>' +
    '<p>The link is saved automatically the first time you open it, so emails can link back to the app.</p>';
}

function deployHelpHtml_() {
  return '<p>To use the app on a phone or computer, publish it as a web app once:</p>' + deployStepsHtml_() +
    '<p>You can also use <b>Inventory &gt; Open in sidebar</b> right here in the spreadsheet.</p>';
}

function setupDoneHtml_(ss, report) {
  const settings = getSettings_(ss);
  const items = [];
  if (report.created.length) items.push('Created the ' + listText_(report.created) + ' tab' + (report.created.length > 1 ? 's' : '') + '.');
  if (report.addedColumns.length) items.push('Added the ' + listText_(report.addedColumns) + ' column' + (report.addedColumns.length > 1 ? 's' : '') + ' to the Inventory tab.');
  if (report.addedSettings.length) items.push('Added ' + listText_(report.addedSettings) + ' to the Settings tab.');
  if (!items.length) items.push('Everything was already in place.');
  if (report.missingOptional.length) {
    items.push('Optional columns you can add to the Inventory tab if you want them: ' + listText_(report.missingOptional) + '.');
  }
  items.push(report.triggers && report.triggers.installed
    ? 'Automatic alerts are on: the stock is checked whenever the Inventory tab is edited, and every hour.'
    : 'Automatic alerts were not changed; they run under ' + escapeHtml_((report.triggers && report.triggers.owner) || 'another account') + '.');
  items.push(settings.recipients.length
    ? 'Low-stock emails go to ' + escapeHtml_(settings.recipients.join(', ')) + '. Change that in the Settings tab.'
    : 'Add who should get low-stock emails in the Settings tab.');
  return '<ul style="padding-left:20px;">' + items.map(function (item) { return '<li>' + item + '</li>'; }).join('') + '</ul>' +
    '<p><b>Next:</b> add parts in the Inventory tab, or publish the app for your phone and computer:</p>' + deployStepsHtml_();
}

function helpHtml_(ss) {
  const url = getSettings_(ss).appUrl;
  return '<p><b>Inventory tab</b>: one row per part. Set <b>Min Qty</b> to get an email when the quantity drops to that ' +
    'number or below, and put the page to reorder from in <b>Order Link</b>. Grey columns are filled in by the app.</p>' +
    '<p><b>Settings tab</b>: who gets the emails, the reminder schedule and the app name.</p>' +
    '<p><b>Activity Log tab</b>: every change made through the app, and the emails sent.</p>' +
    (url
      ? '<p><b>Web app</b>: <a href="' + escapeHtml_(url) + '" target="_blank" rel="noopener">' + escapeHtml_(url) + '</a></p>'
      : '<p><b>Web app</b>: not published yet.</p>' + deployStepsHtml_()) +
    '<p><b>Changed the code?</b> In Apps Script, use Deploy &gt; Manage deployments, edit the web app and pick ' +
    '"New version", so the web app picks up the change.</p>' +
    '<p style="color:#6b7280;">Parts Inventory ' + APP_VERSION + '</p>';
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function displayName_(p) {
  return p.name || p.partNumber || p.id || 'Unnamed part';
}

function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

function isMarked_(value) {
  if (value === true) return true;
  if (isDate_(value)) return !isNaN(value.getTime());
  if (typeof value === 'number') return value > 0;
  if (typeof value === 'string') return !!value.trim() && !/^(no|n|false|0|none|-)$/i.test(value.trim());
  return false;
}

function isBlank_(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function cellText_(value) {
  if (value === null || value === undefined) return '';
  if (isDate_(value)) return isNaN(value.getTime()) ? '' : value.toISOString().slice(0, 10);
  return String(value).trim();
}

function cleanString_(value, multiline) {
  let s = String(value == null ? '' : value);
  s = multiline ? s.replace(/\r\n?/g, '\n') : s.replace(/[\r\n\t]+/g, ' ');
  return s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
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

function toDateOrNull_(value) {
  if (isDate_(value)) return isNaN(value.getTime()) ? null : value;
  if (typeof value === 'string' && value.trim()) {
    const t = Date.parse(value.trim());
    return isNaN(t) ? null : new Date(t);
  }
  return null;
}

function toBoolean_(value) {
  return value === true || /^(true|yes|y|on|1)$/i.test(String(value).trim());
}

function round_(n, decimals) {
  const f = Math.pow(10, decimals);
  return Math.round(n * f) / f;
}

function sameValue_(a, b) {
  const blankA = a === null || a === undefined || a === '';
  const blankB = b === null || b === undefined || b === '';
  if (blankA || blankB) return blankA === blankB;
  return typeof a === 'number' || typeof b === 'number' ? Number(a) === Number(b) : String(a) === String(b);
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

function listText_(items) {
  if (items.length < 2) return items.join('');
  return items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1];
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

function partUrl_(appUrl, id) {
  return appUrl + (appUrl.indexOf('?') === -1 ? '?' : '&') + 'part=' + encodeURIComponent(id);
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

function safeEqual_(a, b) {
  const x = String(a);
  const y = String(b);
  let diff = x.length ^ y.length;
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    diff |= (x.charCodeAt(i) || 0) ^ (y.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function escapeHtml_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** JSON that is safe to drop inside a <script> tag. */
function jsonForHtml_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

/** google.script.run can't return Date objects, so turn them into ISO strings. */
function jsonSafe_(value) {
  if (isDate_(value)) return isNaN(value.getTime()) ? null : value.toISOString();
  if (Array.isArray(value)) return value.map(jsonSafe_);
  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(function (key) {
      const v = value[key];
      if (v !== undefined && typeof v !== 'function') out[key] = jsonSafe_(v);
    });
    return out;
  }
  return value;
}

function parseJson_(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function errorMessage_(err) {
  const message = err && err.message ? err.message : String(err);
  return message.replace(/^Exception:\s*/, '');
}

// ---------------------------------------------------------------------------
// The web page (src/Index.html), bundled so this is the only file to paste.
// ---------------------------------------------------------------------------

var INDEX_HTML_ = "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<style>\n  :root {\n    --bg: #f3f5f7;\n    --surface: #ffffff;\n    --surface-2: #f7f9fa;\n    --border: #e2e6ea;\n    --border-strong: #cdd4db;\n    --text: #16202a;\n    --text-2: #43505c;\n    --muted: #697684;\n    --accent: #0f766e;\n    --accent-hover: #0d625c;\n    --accent-soft: #e3f3f1;\n    --on-accent: #ffffff;\n    --ok: #177a3c;\n    --ok-soft: #e4f5ea;\n    --low: #a95306;\n    --low-soft: #fdf0d9;\n    --out: #b42318;\n    --out-soft: #fde8e6;\n    --ordered: #2952cc;\n    --ordered-soft: #e6ecfc;\n    --scrim: rgba(15, 23, 42, 0.46);\n    --shadow: 0 1px 2px rgba(16, 24, 40, 0.05), 0 1px 3px rgba(16, 24, 40, 0.07);\n    --shadow-lg: 0 18px 40px rgba(16, 24, 40, 0.2);\n    --radius: 12px;\n    --font: system-ui, -apple-system, \"Segoe UI\", Roboto, Helvetica, Arial, sans-serif;\n    color-scheme: light;\n  }\n  @media (prefers-color-scheme: dark) {\n    :root:not([data-theme=\"light\"]) {\n      --bg: #0e1318;\n      --surface: #161c23;\n      --surface-2: #1b222a;\n      --border: #28313b;\n      --border-strong: #36414d;\n      --text: #e7ecf1;\n      --text-2: #b8c2cc;\n      --muted: #8d99a6;\n      --accent: #2eb3a3;\n      --accent-hover: #43c3b3;\n      --accent-soft: #113833;\n      --on-accent: #04201c;\n      --ok: #5ad48a;\n      --ok-soft: #133021;\n      --low: #f3b04c;\n      --low-soft: #3a2a11;\n      --out: #f7837a;\n      --out-soft: #3c1a18;\n      --ordered: #86a6ff;\n      --ordered-soft: #1a2750;\n      --scrim: rgba(0, 0, 0, 0.6);\n      --shadow: 0 1px 2px rgba(0, 0, 0, 0.35);\n      --shadow-lg: 0 18px 44px rgba(0, 0, 0, 0.55);\n      color-scheme: dark;\n    }\n  }\n  /* An explicit theme choice (data-theme) wins over the system setting either way. */\n  :root[data-theme=\"dark\"] {\n    --bg: #0e1318;\n    --surface: #161c23;\n    --surface-2: #1b222a;\n    --border: #28313b;\n    --border-strong: #36414d;\n    --text: #e7ecf1;\n    --text-2: #b8c2cc;\n    --muted: #8d99a6;\n    --accent: #2eb3a3;\n    --accent-hover: #43c3b3;\n    --accent-soft: #113833;\n    --on-accent: #04201c;\n    --ok: #5ad48a;\n    --ok-soft: #133021;\n    --low: #f3b04c;\n    --low-soft: #3a2a11;\n    --out: #f7837a;\n    --out-soft: #3c1a18;\n    --ordered: #86a6ff;\n    --ordered-soft: #1a2750;\n    --scrim: rgba(0, 0, 0, 0.6);\n    --shadow: 0 1px 2px rgba(0, 0, 0, 0.35);\n    --shadow-lg: 0 18px 44px rgba(0, 0, 0, 0.55);\n    color-scheme: dark;\n  }\n\n  * { box-sizing: border-box; }\n  html { -webkit-text-size-adjust: 100%; }\n  body { margin: 0; background: var(--bg); color: var(--text); font: 15px/1.45 var(--font); }\n  button, input, select, textarea { font: inherit; color: inherit; }\n  button { -webkit-tap-highlight-color: transparent; }\n  a { color: var(--accent); }\n  [hidden] { display: none !important; }\n  :focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }\n  svg { flex: none; }\n  .sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }\n  .muted { color: var(--muted); }\n  .num { font-variant-numeric: tabular-nums; }\n\n  /* ---------- Top bar ---------- */\n  .topbar {\n    position: sticky; top: env(safe-area-inset-top, 0px); z-index: 20;\n    display: grid; grid-template-columns: minmax(0, 1fr) auto; grid-template-areas: \"brand actions\" \"tabs tabs\";\n    align-items: center; gap: 2px 12px; padding: 10px 16px 0;\n    background: var(--surface); border-bottom: 1px solid var(--border);\n  }\n  .brand { grid-area: brand; display: flex; align-items: center; gap: 10px; min-width: 0; }\n  .logo {\n    width: 32px; height: 32px; border-radius: 9px; flex: none;\n    display: grid; place-items: center; background: var(--accent); color: var(--on-accent);\n  }\n  .brand h1 { margin: 0; font-size: 17px; font-weight: 700; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }\n  .top-actions { grid-area: actions; display: flex; gap: 8px; align-items: center; }\n  .tabs { grid-area: tabs; display: flex; gap: 2px; overflow-x: auto; scrollbar-width: none; margin: 0 -4px; }\n  .tab {\n    border: 0; background: none; cursor: pointer; white-space: nowrap;\n    padding: 10px 12px 9px; border-bottom: 2px solid transparent;\n    color: var(--muted); font-weight: 600; display: inline-flex; align-items: center; gap: 7px;\n  }\n  .tab:hover { color: var(--text); }\n  .tab[aria-selected=\"true\"] { color: var(--text); border-bottom-color: var(--accent); }\n  @media (min-width: 760px) {\n    .topbar { grid-template-columns: auto 1fr auto; grid-template-areas: \"brand tabs actions\"; height: 62px; padding: 0 24px; gap: 24px; }\n    .tabs { align-self: stretch; margin: 0; }\n    .tab { padding: 0 12px; }\n  }\n\n  main { max-width: 1140px; margin: 0 auto; padding: 16px 16px 96px; }\n  @media (min-width: 760px) { main { padding: 24px 24px 80px; } }\n\n  /* ---------- Buttons ---------- */\n  .btn {\n    display: inline-flex; align-items: center; justify-content: center; gap: 7px;\n    min-height: 40px; padding: 0 14px; border-radius: 9px; cursor: pointer; white-space: nowrap;\n    border: 1px solid var(--border-strong); background: var(--surface); color: var(--text);\n    font-weight: 600; text-decoration: none; transition: background 0.12s, border-color 0.12s;\n  }\n  .btn:hover { background: var(--surface-2); }\n  .btn.primary { background: var(--accent); border-color: var(--accent); color: var(--on-accent); }\n  .btn.primary:hover { background: var(--accent-hover); border-color: var(--accent-hover); }\n  .btn.subtle { border-color: transparent; background: transparent; color: var(--text-2); }\n  .btn.subtle:hover { background: var(--surface-2); }\n  .btn.danger { color: var(--out); }\n  .btn.danger-fill { background: var(--out); border-color: var(--out); color: #fff; }\n  .btn.small { min-height: 32px; padding: 0 10px; font-size: 14px; border-radius: 8px; }\n  .btn.block { width: 100%; }\n  .btn:disabled { opacity: 0.55; cursor: default; }\n  .icon-btn {\n    width: 40px; height: 40px; flex: none; display: inline-grid; place-items: center; cursor: pointer;\n    border-radius: 9px; border: 1px solid var(--border); background: var(--surface); color: var(--text-2);\n    text-decoration: none;\n  }\n  .icon-btn:hover { background: var(--surface-2); color: var(--text); }\n  .spin svg { animation: spin 0.8s linear infinite; }\n  @keyframes spin { to { transform: rotate(360deg); } }\n  @media (max-width: 420px) { .hide-narrow { display: none; } #add-btn { padding: 0 11px; } }\n  @media (max-width: 440px) { .tab svg { display: none; } .tab { padding: 10px 10px 9px; } }\n  @media (max-width: 360px) { .brand .logo { display: none; } }\n\n  /* ---------- Banners ---------- */\n  .banner {\n    display: flex; gap: 10px; align-items: flex-start; padding: 11px 14px; margin-bottom: 12px;\n    border-radius: 10px; border: 1px solid var(--border); background: var(--surface); font-size: 14px;\n  }\n  .banner.warn { background: var(--low-soft); border-color: transparent; }\n  .banner.warn > svg { color: var(--low); }\n  .banner.error { background: var(--out-soft); border-color: transparent; }\n  .banner.error > svg { color: var(--out); }\n  .banner.info > svg { color: var(--accent); }\n  .banner .grow { flex: 1; min-width: 0; }\n  .banner .banner-actions { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 8px; }\n  .banner .close { background: none; border: 0; padding: 2px; cursor: pointer; color: var(--muted); }\n\n  /* ---------- Filters ---------- */\n  .filters { display: flex; gap: 8px; margin: 0 -1px 12px; padding: 1px; overflow-x: auto; scrollbar-width: none; }\n  .filters::-webkit-scrollbar { display: none; }\n  .filter {\n    flex: none; display: inline-flex; align-items: center; gap: 7px; height: 38px; padding: 0 8px 0 14px;\n    border-radius: 999px; border: 1px solid var(--border-strong); background: var(--surface);\n    color: var(--text-2); font-weight: 650; cursor: pointer; white-space: nowrap;\n  }\n  .filter:hover { border-color: var(--muted); color: var(--text); }\n  .filter .count {\n    min-width: 24px; height: 24px; padding: 0 7px; border-radius: 999px; display: inline-grid; place-items: center;\n    background: var(--surface-2); color: var(--text-2); font-size: 13px; font-variant-numeric: tabular-nums;\n  }\n  .filter.low.has .count { background: var(--low-soft); color: var(--low); }\n  .filter.out.has .count { background: var(--out-soft); color: var(--out); }\n  .filter.ordered.has .count { background: var(--ordered-soft); color: var(--ordered); }\n  .filter[aria-pressed=\"true\"] { background: var(--text); border-color: var(--text); color: var(--surface); }\n  .filter[aria-pressed=\"true\"] .count { background: var(--surface); color: var(--text); }\n  .filter .short { display: none; }\n  @media (max-width: 520px) {\n    .filters { gap: 6px; }\n    .filter { padding: 0 5px 0 10px; gap: 5px; font-size: 14px; height: 36px; }\n    .filter .count { min-width: 22px; height: 22px; padding: 0 6px; font-size: 12.5px; }\n    .filter .long { display: none; }\n    .filter .short { display: inline; }\n  }\n  @media (max-width: 380px) {\n    .filters { gap: 4px; }\n    .filter { padding: 0 4px 0 9px; gap: 4px; font-size: 13.5px; }\n    .filter .count { min-width: 20px; padding: 0 5px; }\n  }\n\n  .toolbar { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; }\n  .search { position: relative; flex: 1 1 260px; display: flex; align-items: center; }\n  .search > span[data-icon] { position: absolute; left: 12px; top: 0; bottom: 0; display: flex; align-items: center; color: var(--muted); pointer-events: none; }\n  .search input {\n    width: 100%; height: 42px; padding: 0 12px 0 38px; border-radius: 10px;\n    border: 1px solid var(--border-strong); background: var(--surface);\n  }\n  .toolbar select {\n    height: 42px; border-radius: 10px; border: 1px solid var(--border-strong); background: var(--surface);\n    padding: 0 10px; flex: 1 1 140px; max-width: 220px; min-width: 0;\n  }\n  @media (max-width: 520px) { .toolbar select { max-width: none; } }\n\n  /* ---------- Parts list ---------- */\n  .list { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }\n  .list-head { display: none; }\n  .row {\n    position: relative; display: grid; gap: 8px 12px; padding: 12px 14px 12px 18px;\n    grid-template-columns: minmax(0, 1fr) auto; grid-template-areas: \"main main\" \"extra stock\";\n    border-top: 1px solid var(--border);\n  }\n  .row:first-child { border-top: 0; }\n  .row::before { content: \"\"; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; }\n  .row.status-low::before { background: var(--low); }\n  .row.status-out::before { background: var(--out); }\n  .row.flash { animation: flash 1.4s ease-out; }\n  @keyframes flash { from { background: var(--accent-soft); } to { background: transparent; } }\n  .row-main { grid-area: main; min-width: 0; padding: 0; border: 0; background: none; text-align: left; cursor: pointer; }\n  .row-main:hover .name { text-decoration: underline; text-decoration-color: var(--border-strong); text-underline-offset: 3px; }\n  .row-title { display: flex; align-items: center; flex-wrap: wrap; gap: 4px 8px; }\n  .row-title .name { font-weight: 650; overflow-wrap: anywhere; }\n  .row-meta { margin-top: 2px; color: var(--muted); font-size: 13px; overflow-wrap: anywhere; }\n  .row-loc { display: none; }\n  .row-extra { grid-area: extra; display: flex; align-items: center; flex-wrap: wrap; gap: 6px 12px; min-width: 0; }\n  .row-min { display: none; }\n  .row-act:empty { display: none; }\n  .row-act .btn { min-height: 34px; padding: 0 12px; }\n  .row-act .on-order { display: inline-flex; align-items: center; gap: 5px; color: var(--ordered); font-size: 13px; font-weight: 650; }\n  .row-stock { grid-area: stock; justify-self: end; align-self: center; }\n  @media (min-width: 760px) {\n    .list-head, .row {\n      grid-template-columns: minmax(0, 1fr) minmax(90px, 150px) 64px 186px 124px;\n      grid-template-areas: \"main loc min stock act\";\n    }\n    .list-head {\n      display: grid; gap: 12px; padding: 10px 16px 10px 20px; background: var(--surface-2);\n      border-bottom: 1px solid var(--border); color: var(--muted);\n      font-size: 12px; font-weight: 650; letter-spacing: 0.04em; text-transform: uppercase;\n    }\n    .list-head .right { text-align: right; }\n    .row { align-items: center; padding: 10px 16px 10px 20px; }\n    .row-extra { display: contents; }\n    .row-min { display: block; grid-area: min; color: var(--text-2); font-size: 14px; font-variant-numeric: tabular-nums; }\n    .meta-min { display: none; }\n    .row-act, .row-act:empty { display: block; grid-area: act; justify-self: end; }\n    .row-loc { display: block; grid-area: loc; color: var(--text-2); font-size: 14px; overflow-wrap: anywhere; }\n    .meta-loc { display: none; }\n  }\n\n  .stepper {\n    display: inline-flex; align-items: stretch; overflow: hidden;\n    border: 1px solid var(--border-strong); border-radius: 10px; background: var(--surface);\n  }\n  .step {\n    width: 42px; height: 42px; border: 0; background: var(--surface-2); color: var(--text-2);\n    display: grid; place-items: center; cursor: pointer; touch-action: manipulation;\n  }\n  .step:hover { color: var(--text); background: var(--border); }\n  .step:disabled { opacity: 0.4; cursor: default; background: var(--surface-2); }\n  .qty {\n    min-width: 88px; padding: 0 6px; display: flex; align-items: baseline; justify-content: center; gap: 4px;\n    align-self: center; font-variant-numeric: tabular-nums; position: relative;\n  }\n  .qty strong { font-size: 17px; }\n  .qty small { color: var(--muted); font-size: 12px; }\n  .qty.saving strong { opacity: 0.55; }\n  .qty.saving::after {\n    content: \"\"; position: absolute; right: 5px; top: 50%; width: 6px; height: 6px; margin-top: -3px;\n    border-radius: 50%; background: var(--accent); animation: pulse 0.9s ease-in-out infinite alternate;\n  }\n  @keyframes pulse { from { opacity: 0.25; } to { opacity: 1; } }\n\n  .badge {\n    display: inline-flex; align-items: center; gap: 4px; padding: 1px 8px; border-radius: 999px;\n    font-size: 12px; font-weight: 700; line-height: 20px; white-space: nowrap;\n  }\n  .badge.low { background: var(--low-soft); color: var(--low); }\n  .badge.out { background: var(--out-soft); color: var(--out); }\n  .badge.ok { background: var(--ok-soft); color: var(--ok); }\n  .badge.ordered { background: var(--ordered-soft); color: var(--ordered); }\n\n  .more { display: flex; justify-content: center; padding: 14px; border-top: 1px solid var(--border); }\n  .empty { padding: 44px 20px; text-align: center; color: var(--muted); }\n  .empty .empty-icon {\n    width: 52px; height: 52px; margin: 0 auto 12px; border-radius: 14px; display: grid; place-items: center;\n    background: var(--accent-soft); color: var(--accent);\n  }\n  .empty h3 { margin: 0 0 4px; color: var(--text); font-size: 17px; }\n  .empty p { margin: 0 auto 16px; max-width: 420px; }\n\n  .skeleton { padding: 16px 18px; border-top: 1px solid var(--border); display: grid; gap: 8px; }\n  .skeleton:first-child { border-top: 0; }\n  .bar { height: 12px; border-radius: 6px; background: linear-gradient(90deg, var(--surface-2), var(--border), var(--surface-2)); background-size: 200% 100%; animation: shimmer 1.2s linear infinite; }\n  @keyframes shimmer { to { background-position: -200% 0; } }\n\n  /* ---------- Part drawer ---------- */\n  .drawer { position: fixed; inset: 0; z-index: 40; }\n  .drawer-backdrop { position: absolute; inset: 0; background: var(--scrim); animation: fade 0.15s ease-out; }\n  .drawer-panel {\n    position: absolute; top: 0; right: 0; bottom: 0; width: min(470px, 100%);\n    display: flex; flex-direction: column; background: var(--surface); box-shadow: var(--shadow-lg);\n    animation: slide-in 0.2s ease-out;\n  }\n  @keyframes fade { from { opacity: 0; } }\n  @keyframes slide-in { from { transform: translateX(24px); opacity: 0.4; } }\n  .drawer-head { display: flex; gap: 12px; align-items: flex-start; padding: 16px 16px 14px 20px; border-bottom: 1px solid var(--border); }\n  .drawer-head .grow { flex: 1; min-width: 0; }\n  .drawer-head h2 { margin: 2px 0 4px; font-size: 20px; line-height: 1.25; overflow-wrap: anywhere; }\n  .eyebrow { color: var(--muted); font-size: 13px; font-weight: 600; }\n  .drawer-body { flex: 1; overflow-y: auto; padding: 16px 20px 32px; display: grid; gap: 18px; align-content: start; }\n\n  .stock-card { border: 1px solid var(--border); border-radius: var(--radius); background: var(--surface-2); padding: 16px; display: grid; gap: 14px; }\n  .stock-top { display: flex; align-items: flex-end; justify-content: space-between; gap: 12px; flex-wrap: wrap; }\n  .stock-big { font-size: 36px; font-weight: 750; line-height: 1; font-variant-numeric: tabular-nums; }\n  .stock-big small { font-size: 16px; font-weight: 600; color: var(--muted); margin-left: 4px; }\n  .stock-card.status-low .stock-big { color: var(--low); }\n  .stock-card.status-out .stock-big { color: var(--out); }\n  .stock-levels { color: var(--muted); font-size: 13px; text-align: right; }\n  .segmented { display: grid; grid-template-columns: repeat(3, 1fr); border: 1px solid var(--border-strong); border-radius: 9px; overflow: hidden; }\n  .segmented button { border: 0; border-left: 1px solid var(--border-strong); background: var(--surface); padding: 8px 6px; font-weight: 600; cursor: pointer; }\n  .segmented button:first-child { border-left: 0; }\n  .segmented button[aria-pressed=\"true\"] { background: var(--accent); color: var(--on-accent); }\n  .adjust-row { display: grid; grid-template-columns: 110px 1fr; gap: 8px; }\n  .adjust-row input { height: 42px; border: 1px solid var(--border-strong); border-radius: 9px; padding: 0 12px; background: var(--surface); width: 100%; min-width: 0; }\n  .adjust-preview { font-size: 13px; color: var(--muted); min-height: 19px; }\n\n  .order-links { display: grid; gap: 8px; }\n  .order-links a.btn { justify-content: space-between; }\n  .order-links .btn span { overflow: hidden; text-overflow: ellipsis; }\n  .ordered-note { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 10px 12px; border-radius: 10px; background: var(--ordered-soft); color: var(--ordered); font-weight: 600; font-size: 14px; }\n  .ordered-note .grow { flex: 1; }\n\n  .details { display: grid; grid-template-columns: 120px minmax(0, 1fr); gap: 9px 12px; margin: 0; font-size: 14px; }\n  .details dt { color: var(--muted); }\n  .details dd { margin: 0; overflow-wrap: anywhere; white-space: pre-wrap; }\n  .section-title { font-size: 13px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.04em; margin: 0 0 8px; }\n  .drawer-actions { display: flex; flex-wrap: wrap; gap: 8px; }\n\n  /* ---------- Activity ---------- */\n  .page-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 14px; }\n  .page-head h2 { margin: 0; font-size: 20px; }\n  .timeline { list-style: none; margin: 0; padding: 0; }\n  .card-list { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; }\n  .act { display: grid; grid-template-columns: 32px minmax(0, 1fr); gap: 12px; padding: 12px 14px; border-top: 1px solid var(--border); }\n  .act:first-child { border-top: 0; }\n  .act-icon { width: 32px; height: 32px; border-radius: 50%; display: grid; place-items: center; background: var(--surface-2); color: var(--text-2); }\n  .act.used .act-icon { background: var(--low-soft); color: var(--low); }\n  .act.restocked .act-icon, .act.added .act-icon { background: var(--ok-soft); color: var(--ok); }\n  .act.alert .act-icon { background: var(--accent-soft); color: var(--accent); }\n  .act.deleted .act-icon { background: var(--out-soft); color: var(--out); }\n  .act.ordered .act-icon { background: var(--ordered-soft); color: var(--ordered); }\n  .act-line { overflow-wrap: anywhere; }\n  .act-line a, .link-btn { color: var(--text); font-weight: 650; text-decoration: none; background: none; border: 0; padding: 0; cursor: pointer; text-align: left; }\n  .act-line a:hover, .link-btn:hover { text-decoration: underline; }\n  .act-sub { color: var(--muted); font-size: 13px; margin-top: 1px; }\n  .chg { font-weight: 700; font-variant-numeric: tabular-nums; }\n  .chg.neg { color: var(--low); }\n  .chg.pos { color: var(--ok); }\n  .drawer .card-list { box-shadow: none; }\n\n  /* ---------- Settings ---------- */\n  .settings { display: grid; gap: 16px; max-width: 780px; }\n  .card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow); padding: 18px 20px; }\n  .card h2 { margin: 0 0 2px; font-size: 17px; }\n  .card .sub { margin: 0 0 16px; color: var(--muted); font-size: 14px; }\n  .card > :last-child { margin-bottom: 0; }\n  .card-actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 16px; align-items: center; }\n  .facts { margin: 14px 0 0; padding: 0; list-style: none; display: grid; gap: 6px; font-size: 14px; color: var(--text-2); }\n  .facts li { display: flex; gap: 8px; align-items: flex-start; }\n  .facts svg { margin-top: 2px; color: var(--muted); }\n\n  /* ---------- Forms ---------- */\n  .form-grid { display: grid; grid-template-columns: minmax(0, 1fr); gap: 14px; }\n  @media (min-width: 620px) { .form-grid.two { grid-template-columns: repeat(2, minmax(0, 1fr)); } .span-2 { grid-column: 1 / -1; } }\n  .field { display: flex; flex-direction: column; gap: 5px; min-width: 0; }\n  .field > label, .field > .label { font-size: 13px; font-weight: 650; color: var(--text-2); }\n  .field .req { color: var(--out); }\n  .field input, .field select, .field textarea {\n    width: 100%; height: 42px; padding: 0 12px; border-radius: 9px;\n    border: 1px solid var(--border-strong); background: var(--surface);\n  }\n  .field textarea { height: auto; min-height: 66px; padding: 10px 12px; resize: vertical; line-height: 1.4; }\n  .field input:disabled, .field select:disabled, .field textarea:disabled { background: var(--surface-2); color: var(--muted); }\n  .field input[aria-invalid=\"true\"] { border-color: var(--out); }\n  .hint { font-size: 12.5px; color: var(--muted); }\n  .hint.warn { color: var(--low); }\n  .inline { display: flex; gap: 8px; align-items: center; }\n  .inline > * { flex: 1; min-width: 0; }\n  .form-error { display: flex; gap: 8px; align-items: flex-start; padding: 10px 12px; border-radius: 9px; background: var(--out-soft); color: var(--out); font-size: 14px; font-weight: 600; margin-bottom: 14px; }\n  .switch { display: flex; align-items: flex-start; gap: 12px; cursor: pointer; }\n  .switch input { position: absolute; opacity: 0; width: 1px; height: 1px; }\n  .switch .track { flex: none; width: 42px; height: 24px; border-radius: 999px; background: var(--border-strong); position: relative; transition: background 0.15s; margin-top: 1px; }\n  .switch .track::after { content: \"\"; position: absolute; top: 3px; left: 3px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.3); transition: transform 0.15s; }\n  .switch input:checked + .track { background: var(--accent); }\n  .switch input:checked + .track::after { transform: translateX(18px); }\n  .switch input:focus-visible + .track { outline: 2px solid var(--accent); outline-offset: 2px; }\n  .switch input:disabled + .track { opacity: 0.5; }\n  .switch .switch-text { display: grid; gap: 2px; }\n  .switch .switch-text b { font-weight: 650; }\n\n  /* ---------- Dialogs ---------- */\n  dialog.modal {\n    width: min(660px, calc(100% - 24px)); max-height: calc(100% - 24px); padding: 0; border: 0; border-radius: 16px;\n    background: var(--surface); color: var(--text); box-shadow: var(--shadow-lg);\n  }\n  dialog.modal.small { width: min(420px, calc(100% - 24px)); }\n  dialog.modal::backdrop { background: var(--scrim); }\n  dialog.modal[open] { display: flex; flex-direction: column; animation: pop 0.16s ease-out; }\n  @keyframes pop { from { transform: translateY(8px) scale(0.99); opacity: 0; } }\n  .modal form { display: flex; flex-direction: column; min-height: 0; flex: 1; }\n  .modal-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 14px 16px 14px 20px; border-bottom: 1px solid var(--border); }\n  .modal-head h2 { margin: 0; font-size: 18px; }\n  .modal-body { padding: 18px 20px; overflow-y: auto; min-height: 0; flex: 1; }\n  .modal-body p { margin: 0; }\n  .modal-foot { display: flex; justify-content: flex-end; gap: 8px; padding: 12px 20px; border-top: 1px solid var(--border); flex-wrap: wrap; }\n  @media (max-width: 620px) {\n    dialog.modal:not(.small) { width: 100%; max-width: 100%; height: 100%; max-height: 100%; margin: 0; border-radius: 0; }\n  }\n\n  /* ---------- Lock screen, fatal error ---------- */\n  .screen { min-height: 100vh; display: grid; place-items: center; padding: 16px; }\n  .screen-card { width: min(400px, 100%); background: var(--surface); border: 1px solid var(--border); border-radius: 16px; box-shadow: var(--shadow-lg); padding: 28px 24px; text-align: center; }\n  .screen-card h1 { margin: 12px 0 4px; font-size: 20px; }\n  .screen-card p { margin: 0 0 18px; color: var(--muted); }\n  .screen-card .logo { width: 48px; height: 48px; border-radius: 14px; margin: 0 auto; }\n  .screen-card form { display: grid; gap: 10px; text-align: left; }\n  .screen-card input { width: 100%; height: 44px; padding: 0 12px; border-radius: 10px; border: 1px solid var(--border-strong); background: var(--surface); font-size: 16px; }\n  .screen-card .error-text { color: var(--out); font-size: 14px; min-height: 20px; }\n  .screen-card pre { text-align: left; white-space: pre-wrap; background: var(--surface-2); border-radius: 8px; padding: 10px; font-size: 13px; color: var(--text-2); }\n\n  /* ---------- Add part button on phones ---------- */\n  .fab {\n    position: fixed; right: 16px; bottom: calc(16px + env(safe-area-inset-bottom, 0px)); z-index: 30;\n    display: inline-flex; align-items: center; gap: 8px; height: 52px; padding: 0 22px 0 18px;\n    border: 0; border-radius: 999px; background: var(--accent); color: var(--on-accent);\n    font-size: 16px; font-weight: 700; box-shadow: var(--shadow-lg); cursor: pointer;\n  }\n  .fab:hover { background: var(--accent-hover); }\n  @media (min-width: 760px) { .fab { display: none; } }\n  @media (max-width: 759px) { #add-btn { display: none; } }\n\n  /* ---------- Toasts ---------- */\n  .toasts { position: fixed; left: 50%; bottom: calc(80px + env(safe-area-inset-bottom, 0px)); transform: translateX(-50%); z-index: 80; width: min(460px, calc(100% - 24px)); display: grid; gap: 8px; pointer-events: none; }\n  .toast {\n    pointer-events: auto; display: flex; gap: 10px; align-items: flex-start; padding: 11px 12px 11px 14px;\n    border-radius: 11px; background: #1d2731; color: #f3f6f9; box-shadow: var(--shadow-lg); font-size: 14px;\n    animation: rise 0.18s ease-out;\n  }\n  .toast.error { background: #9f1f16; color: #fff; }\n  .toast .grow { flex: 1; min-width: 0; padding-top: 1px; overflow-wrap: anywhere; }\n  .toast .toast-action { opacity: 1; font-weight: 700; text-decoration: underline; text-underline-offset: 3px; padding: 1px 4px; white-space: nowrap; }\n  @media (min-width: 760px) { .toasts { bottom: calc(18px + env(safe-area-inset-bottom, 0px)); } }\n  .toast button { background: none; border: 0; color: inherit; opacity: 0.75; cursor: pointer; padding: 1px; }\n  .toast.leaving { opacity: 0; transform: translateY(6px); transition: all 0.2s; }\n  @keyframes rise { from { transform: translateY(10px); opacity: 0; } }\n\n  @media (prefers-reduced-motion: reduce) {\n    *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }\n  }\n</style>\n</head>\n<body>\n\n<div id=\"app\" hidden>\n  <header class=\"topbar\">\n    <div class=\"brand\">\n      <span class=\"logo\" data-icon=\"box\" aria-hidden=\"true\"></span>\n      <h1 id=\"app-name\">Parts Inventory</h1>\n    </div>\n    <nav class=\"tabs\" role=\"tablist\" aria-label=\"Sections\">\n      <button class=\"tab\" role=\"tab\" data-view=\"inventory\" data-icon=\"list\" aria-selected=\"true\">Inventory</button>\n      <button class=\"tab\" role=\"tab\" data-view=\"activity\" data-icon=\"history\" aria-selected=\"false\">Activity</button>\n      <button class=\"tab\" role=\"tab\" data-view=\"settings\" data-icon=\"sliders\" aria-selected=\"false\">Settings</button>\n    </nav>\n    <div class=\"top-actions\">\n      <button class=\"icon-btn\" id=\"refresh-btn\" type=\"button\" title=\"Refresh\" aria-label=\"Refresh\" data-icon=\"refresh\"></button>\n      <button class=\"btn primary\" id=\"add-btn\" type=\"button\" data-icon=\"plus\"><span class=\"hide-narrow\">Add part</span></button>\n    </div>\n  </header>\n\n  <main>\n    <section id=\"view-inventory\" aria-label=\"Inventory\">\n      <div id=\"banners\"></div>\n      <div class=\"toolbar\">\n        <label class=\"search\">\n          <span data-icon=\"search\"></span>\n          <span class=\"sr-only\">Search parts</span>\n          <input id=\"search\" type=\"search\" placeholder=\"Search name, part #, location, supplier…\" autocomplete=\"off\" enterkeyhint=\"search\">\n        </label>\n        <select id=\"category-filter\" aria-label=\"Category\"></select>\n        <select id=\"sort\" aria-label=\"Sort by\">\n          <option value=\"name\">Sort: Name</option>\n          <option value=\"status\">Sort: Needs attention</option>\n          <option value=\"qty\">Sort: Fewest on hand</option>\n          <option value=\"location\">Sort: Location</option>\n          <option value=\"updated\">Sort: Recently changed</option>\n        </select>\n      </div>\n      <div class=\"filters\" id=\"filters\" role=\"group\" aria-label=\"Show\"></div>\n      <div id=\"list\" class=\"list\" aria-live=\"polite\"></div>\n    </section>\n\n    <section id=\"view-activity\" aria-label=\"Activity\" hidden>\n      <div class=\"page-head\">\n        <h2>Recent activity</h2>\n        <button class=\"btn small\" id=\"activity-refresh\" type=\"button\" data-icon=\"refresh\">Refresh</button>\n      </div>\n      <div id=\"activity\"></div>\n    </section>\n\n    <section id=\"view-settings\" aria-label=\"Settings\" hidden>\n      <div id=\"settings\" class=\"settings\"></div>\n    </section>\n  </main>\n  <button class=\"fab\" id=\"fab-add\" type=\"button\" data-icon=\"plus\">Add part</button>\n</div>\n\n<div id=\"drawer\" class=\"drawer\" hidden>\n  <div class=\"drawer-backdrop\" data-close-drawer></div>\n  <aside class=\"drawer-panel\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"drawer-title\" id=\"drawer-panel\"></aside>\n</div>\n\n<dialog id=\"part-dialog\" class=\"modal\" aria-labelledby=\"part-dialog-title\">\n  <form id=\"part-form\" novalidate>\n    <div class=\"modal-head\">\n      <h2 id=\"part-dialog-title\">Add part</h2>\n      <button class=\"icon-btn\" type=\"button\" data-close-dialog aria-label=\"Close\" data-icon=\"x\"></button>\n    </div>\n    <div class=\"modal-body\">\n      <div id=\"part-form-error\" class=\"form-error\" role=\"alert\" hidden></div>\n      <div class=\"form-grid two\">\n        <div class=\"field span-2\">\n          <label for=\"f-name\">Part name <span class=\"req\">*</span></label>\n          <input id=\"f-name\" name=\"name\" maxlength=\"200\" required autocomplete=\"off\" placeholder=\"e.g. M3 × 8 socket head screw\">\n        </div>\n        <div class=\"field\" data-field=\"partNumber\">\n          <label for=\"f-partNumber\">Part number</label>\n          <input id=\"f-partNumber\" name=\"partNumber\" maxlength=\"100\" autocomplete=\"off\" placeholder=\"SKU or manufacturer #\">\n          <span class=\"hint warn\" id=\"dup-hint\" hidden></span>\n        </div>\n        <div class=\"field\" data-field=\"category\">\n          <label for=\"f-category\">Category</label>\n          <input id=\"f-category\" name=\"category\" maxlength=\"100\" list=\"dl-category\" autocomplete=\"off\" placeholder=\"e.g. Fasteners\">\n        </div>\n        <div class=\"field\" data-field=\"location\">\n          <label for=\"f-location\">Location</label>\n          <input id=\"f-location\" name=\"location\" maxlength=\"100\" list=\"dl-location\" autocomplete=\"off\" placeholder=\"e.g. Bin A-12\">\n        </div>\n        <div class=\"field\" data-field=\"supplier\">\n          <label for=\"f-supplier\">Supplier</label>\n          <input id=\"f-supplier\" name=\"supplier\" maxlength=\"150\" list=\"dl-supplier\" autocomplete=\"off\" placeholder=\"e.g. McMaster-Carr\">\n        </div>\n        <div class=\"field\">\n          <label for=\"f-quantity\">Quantity on hand <span class=\"req\">*</span></label>\n          <div class=\"inline\">\n            <input id=\"f-quantity\" name=\"quantity\" type=\"number\" inputmode=\"decimal\" min=\"0\" step=\"any\" required>\n            <input id=\"f-unit\" name=\"unit\" maxlength=\"30\" list=\"dl-unit\" autocomplete=\"off\" placeholder=\"Unit (ea)\" aria-label=\"Unit\" data-field=\"unit\">\n          </div>\n        </div>\n        <div class=\"field\">\n          <label for=\"f-minQty\">Min qty (low-stock alert)</label>\n          <input id=\"f-minQty\" name=\"minQty\" type=\"number\" inputmode=\"decimal\" min=\"0\" step=\"any\" placeholder=\"Leave blank for no alert\">\n          <span class=\"hint\">When the quantity drops to this number or below, an email goes out with the order link.</span>\n        </div>\n        <div class=\"field\" data-field=\"reorderQty\">\n          <label for=\"f-reorderQty\">Reorder quantity</label>\n          <input id=\"f-reorderQty\" name=\"reorderQty\" type=\"number\" inputmode=\"decimal\" min=\"0\" step=\"any\" placeholder=\"How many to order\">\n        </div>\n        <div class=\"field\" data-field=\"unitCost\">\n          <label for=\"f-unitCost\">Unit cost</label>\n          <input id=\"f-unitCost\" name=\"unitCost\" type=\"number\" inputmode=\"decimal\" min=\"0\" step=\"any\" placeholder=\"Optional\">\n        </div>\n        <div class=\"field span-2\">\n          <label for=\"f-link\">Order link</label>\n          <textarea id=\"f-link\" name=\"link\" rows=\"2\" autocomplete=\"off\" spellcheck=\"false\" placeholder=\"https://www.supplier.com/product-page\"></textarea>\n          <span class=\"hint\">The page to reorder from. It's included in the low-stock email. Put a backup supplier on a second line.</span>\n        </div>\n        <div class=\"field span-2\" data-field=\"notes\">\n          <label for=\"f-notes\">Notes</label>\n          <textarea id=\"f-notes\" name=\"notes\" rows=\"2\" maxlength=\"2000\"></textarea>\n        </div>\n      </div>\n      <datalist id=\"dl-category\"></datalist>\n      <datalist id=\"dl-location\"></datalist>\n      <datalist id=\"dl-supplier\"></datalist>\n      <datalist id=\"dl-unit\"></datalist>\n    </div>\n    <div class=\"modal-foot\">\n      <button class=\"btn\" type=\"button\" data-close-dialog>Cancel</button>\n      <button class=\"btn primary\" type=\"submit\" id=\"part-save\">Save part</button>\n    </div>\n  </form>\n</dialog>\n\n<dialog id=\"confirm-dialog\" class=\"modal small\" aria-labelledby=\"confirm-title\">\n  <div class=\"modal-head\"><h2 id=\"confirm-title\">Are you sure?</h2></div>\n  <div class=\"modal-body\"><p id=\"confirm-message\"></p></div>\n  <div class=\"modal-foot\">\n    <button class=\"btn\" type=\"button\" id=\"confirm-cancel\">Cancel</button>\n    <button class=\"btn primary\" type=\"button\" id=\"confirm-ok\">OK</button>\n  </div>\n</dialog>\n\n<dialog id=\"name-dialog\" class=\"modal small\" aria-labelledby=\"name-title\">\n  <form id=\"name-form\">\n    <div class=\"modal-head\"><h2 id=\"name-title\">What's your name?</h2></div>\n    <div class=\"modal-body\">\n      <div class=\"field\">\n        <label for=\"name-input\">Your changes are logged under this name. It's saved on this device only.</label>\n        <input id=\"name-input\" maxlength=\"60\" autocomplete=\"name\" placeholder=\"e.g. Sam R.\">\n      </div>\n    </div>\n    <div class=\"modal-foot\">\n      <button class=\"btn\" type=\"button\" id=\"name-skip\">Not now</button>\n      <button class=\"btn primary\" type=\"submit\">Save</button>\n    </div>\n  </form>\n</dialog>\n\n<div id=\"screen\" class=\"screen\" hidden></div>\n<div id=\"toasts\" class=\"toasts\" aria-live=\"polite\"></div>\n\n<script>const BOOT = <?!= bootJson ?>;</script>\n<script>\n(function () {\n  'use strict';\n\n  // -------------------------------------------------------------------------\n  // Small helpers\n  // -------------------------------------------------------------------------\n\n  const $ = (sel, root) => (root || document).querySelector(sel);\n  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));\n\n  /** Builds DOM nodes. Text is always inserted as text, never as HTML. */\n  function h(tag, attrs, ...children) {\n    const el = document.createElement(tag);\n    Object.keys(attrs || {}).forEach((key) => {\n      const v = attrs[key];\n      if (v === null || v === undefined || v === false) return;\n      if (key === 'class') el.className = v;\n      else if (key.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(key.slice(2), v);\n      else if (v === true) el.setAttribute(key, '');\n      else el.setAttribute(key, String(v));\n    });\n    add(el, children);\n    return el;\n  }\n  function add(el, children) {\n    children.forEach((child) => {\n      if (child === null || child === undefined || child === false || child === '') return;\n      if (Array.isArray(child)) add(el, child);\n      else el.appendChild(child instanceof Node ? child : document.createTextNode(String(child)));\n    });\n    return el;\n  }\n  function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); return el; }\n\n  const SVG_NS = 'http://www.w3.org/2000/svg';\n  const ICONS = {\n    plus: [['path', 'M12 5v14M5 12h14']],\n    minus: [['path', 'M5 12h14']],\n    search: [['circle', 11, 11, 7], ['path', 'M20 20l-3.6-3.6']],\n    refresh: [['path', 'M20 12a8 8 0 1 1-2.34-5.66L20 8.7'], ['path', 'M20 3.5v5.2h-5.2']],\n    x: [['path', 'M18 6L6 18M6 6l12 12']],\n    box: [['path', 'M21 8l-9-5-9 5v8l9 5 9-5V8z'], ['path', 'M3.3 7.7L12 12.6l8.7-4.9M12 12.6V21']],\n    external: [['path', 'M14 4h6v6M20 4l-9 9'], ['path', 'M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5']],\n    cart: [['circle', 9, 20, 1.4], ['circle', 18, 20, 1.4], ['path', 'M2.5 3.5h2.8l2.6 11.3a1 1 0 0 0 1 .8h9.3a1 1 0 0 0 1-.8L21 7.5H6.2']],\n    pin: [['path', 'M12 21s-6.5-5.8-6.5-11A6.5 6.5 0 0 1 18.5 10C18.5 15.2 12 21 12 21z'], ['circle', 12, 10, 2.3]],\n    edit: [['path', 'M4 20h4L19 9l-4-4L4 16v4z'], ['path', 'M13.5 6.5l4 4']],\n    trash: [['path', 'M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3']],\n    truck: [['path', 'M2.5 6.5h11v9.5h-11zM13.5 10h4.2l3 3.2V16h-7.2'], ['circle', 7, 18, 1.8], ['circle', 17, 18, 1.8]],\n    mail: [['rect', 3, 5, 18, 14, 2], ['path', 'M3.5 7l8.5 6 8.5-6']],\n    clock: [['circle', 12, 12, 9], ['path', 'M12 7.5V12l3 2']],\n    alert: [['path', 'M10.3 4.2L2.6 18a2 2 0 0 0 1.7 3h15.4a2 2 0 0 0 1.7-3L13.7 4.2a2 2 0 0 0-3.4 0z'], ['path', 'M12 9.5v4M12 17h.01']],\n    check: [['path', 'M5 12.5l4.5 4.5L19 7.5']],\n    link: [['path', 'M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1'], ['path', 'M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1']],\n    lock: [['rect', 5, 11, 14, 10, 2], ['path', 'M8 11V7.5a4 4 0 0 1 8 0V11']],\n    sheet: [['rect', 3.5, 3.5, 17, 17, 2], ['path', 'M3.5 9.5h17M3.5 15h17M9.5 3.5v17']],\n    sliders: [['path', 'M4 6.5h9M17 6.5h3M4 12h3M11 12h9M4 17.5h11M19 17.5h1'], ['circle', 15, 6.5, 2], ['circle', 9, 12, 2], ['circle', 17, 17.5, 2]],\n    history: [['path', 'M3.5 12a8.5 8.5 0 1 0 2.5-6'], ['path', 'M3.5 4v4.5H8'], ['path', 'M12 8v4.3l3 1.8']],\n    list: [['path', 'M9 6.5h11M9 12h11M9 17.5h11'], ['circle', 4.5, 6.5, 1], ['circle', 4.5, 12, 1], ['circle', 4.5, 17.5, 1]],\n    copy: [['rect', 8.5, 8.5, 12, 12, 2], ['path', 'M15.5 8.5V6a2 2 0 0 0-2-2h-7.5a2 2 0 0 0-2 2v7.5a2 2 0 0 0 2 2h2.5']],\n    user: [['circle', 12, 8, 4], ['path', 'M4.5 20.5a7.5 7.5 0 0 1 15 0']],\n    info: [['circle', 12, 12, 9], ['path', 'M12 11v5.5M12 7.8v.2']],\n    send: [['path', 'M21 3L10 14'], ['path', 'M21 3l-6.5 18-4.5-7-7-4.5L21 3z']],\n    tag: [['path', 'M3 12V4a1 1 0 0 1 1-1h8l9 9-9 9-9-9z'], ['circle', 7.5, 7.5, 1.4]],\n  };\n\n  function icon(name, size) {\n    const svg = document.createElementNS(SVG_NS, 'svg');\n    const s = String(size || 18);\n    [['viewBox', '0 0 24 24'], ['width', s], ['height', s], ['fill', 'none'], ['stroke', 'currentColor'],\n      ['stroke-width', '2'], ['stroke-linecap', 'round'], ['stroke-linejoin', 'round'], ['aria-hidden', 'true'],\n      ['focusable', 'false']].forEach((a) => svg.setAttribute(a[0], a[1]));\n    (ICONS[name] || []).forEach((shape) => {\n      const el = document.createElementNS(SVG_NS, shape[0]);\n      if (shape[0] === 'path') el.setAttribute('d', shape[1]);\n      if (shape[0] === 'circle') { el.setAttribute('cx', shape[1]); el.setAttribute('cy', shape[2]); el.setAttribute('r', shape[3]); }\n      if (shape[0] === 'rect') {\n        el.setAttribute('x', shape[1]); el.setAttribute('y', shape[2]); el.setAttribute('width', shape[3]);\n        el.setAttribute('height', shape[4]); el.setAttribute('rx', shape[5] || 0);\n      }\n      svg.appendChild(el);\n    });\n    return svg;\n  }\n\n  function fmt(n) {\n    if (n === null || n === undefined || n === '' || isNaN(n)) return '—';\n    return (Math.round(Number(n) * 1000) / 1000).toLocaleString(undefined, { maximumFractionDigits: 3 });\n  }\n  function fmtMoney(n) {\n    return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });\n  }\n  function round(n) { return Math.round(n * 1000) / 1000; }\n  function toDate(iso) { const d = iso ? new Date(iso) : null; return d && !isNaN(d) ? d : null; }\n  function fmtDate(iso, withTime) {\n    const d = toDate(iso);\n    if (!d) return '';\n    const opts = { month: 'short', day: 'numeric' };\n    if (d.getFullYear() !== new Date().getFullYear()) opts.year = 'numeric';\n    if (withTime) { opts.hour = 'numeric'; opts.minute = '2-digit'; }\n    return d.toLocaleString(undefined, opts);\n  }\n  function timeAgo(iso) {\n    const d = toDate(iso);\n    if (!d) return '';\n    const s = Math.round((Date.now() - d.getTime()) / 1000);\n    if (s < 45) return 'just now';\n    if (s < 3600) return Math.round(s / 60) + ' min ago';\n    if (s < 86400) return Math.round(s / 3600) + ' hr ago';\n    if (s < 86400 * 6) return Math.round(s / 86400) + ' d ago';\n    return fmtDate(iso);\n  }\n  function hostOf(url) {\n    try { return new URL(url).hostname.replace(/^www\\./, ''); } catch (err) { return url; }\n  }\n  function safeUrl(url) { return /^https?:\\/\\//i.test(String(url || '')) ? url : null; }\n  function displayName(p) { return p.name || p.partNumber || p.id; }\n  function plural(n, one, many) { return fmt(n) + ' ' + (Number(n) === 1 ? one : many); }\n  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });\n  const byText = (a, b) => collator.compare(a || '', b || '');\n\n  const store = {\n    get(key) { try { return window.localStorage.getItem('parts-inventory.' + key); } catch (err) { return null; } },\n    set(key, value) {\n      try {\n        if (value === null || value === undefined || value === '') window.localStorage.removeItem('parts-inventory.' + key);\n        else window.localStorage.setItem('parts-inventory.' + key, value);\n      } catch (err) { /* storage can be unavailable in private windows */ }\n    },\n  };\n\n  // -------------------------------------------------------------------------\n  // State\n  // -------------------------------------------------------------------------\n\n  const state = {\n    data: null,\n    parts: [],\n    view: 'inventory',\n    filter: 'all',\n    search: '',\n    category: '',\n    sort: store.get('sort') || 'name',\n    showCount: 150,\n    selectedId: null,\n    adjustMode: 'remove',\n    activity: null,\n    activityLimit: 100,\n    activityLoadedAt: 0,\n    partActivity: {},\n    partActivityLoading: {},\n    touched: {},\n    deleted: {},\n    adjustDraft: null,\n    settingsDirty: false,\n    code: store.get('code') || '',\n    actor: store.get('actor') || '',\n    loadedAt: 0,\n    loading: false,\n    pending: {},\n    inflight: {},\n    timers: {},\n    editing: null,\n    dismissed: {},\n  };\n\n  // -------------------------------------------------------------------------\n  // Talking to the server (google.script.run)\n  // -------------------------------------------------------------------------\n\n  function call(fn, ...args) {\n    return new Promise((resolve, reject) => {\n      if (!window.google || !google.script || !google.script.run) {\n        reject(new Error('This page only works as a Google Apps Script web app.'));\n        return;\n      }\n      google.script.run\n        .withSuccessHandler(resolve)\n        .withFailureHandler(reject)[fn]({ code: state.code, actor: state.actor }, ...args);\n    });\n  }\n\n  function errorText(err) {\n    const raw = (err && err.message) || String(err || '');\n    if (/NetworkError|Failed to fetch|HTTP 0/i.test(raw)) return 'Couldn\\'t reach Google. Check the connection and try again.';\n    return raw.replace(/^(Exception|Error):\\s*/i, '').replace(/^.*?ACCESS_(DENIED|LOCKED):\\s*/, '') || 'Something went wrong.';\n  }\n  function isAccessError(err) { return /ACCESS_(DENIED|LOCKED)/.test((err && err.message) || String(err)); }\n\n  function api(fn, ...args) {\n    return call(fn, ...args).catch((err) => {\n      if (isAccessError(err)) showLock(errorText(err), /ACCESS_LOCKED/.test((err && err.message) || String(err)));\n      throw err;\n    });\n  }\n\n  function load(options) {\n    const opts = options || {};\n    if (state.loading) return Promise.resolve();\n    state.loading = true;\n    const btn = $('#refresh-btn');\n    btn.classList.add('spin');\n    const requestedAt = Date.now();\n    return api('apiGetData')\n      .then((data) => {\n        // Changes saved while this was loading are newer than what came back.\n        const newer = (id) => state.touched[id] > requestedAt;\n        const seen = {};\n        data.parts = data.parts.filter((p) => !(state.deleted[p.id] > requestedAt)).map((p) => {\n          seen[p.id] = true;\n          return (newer(p.id) && partById(p.id)) || p;\n        });\n        state.parts.forEach((p) => { if (newer(p.id) && !seen[p.id]) data.parts.push(p); });\n        applyData(data);\n        if (opts.fromLock) store.set('code', state.code);\n      })\n      .catch((err) => {\n        if (isAccessError(err)) return;\n        if (!state.data) showFatal(errorText(err));\n        else toast(errorText(err), 'error');\n      })\n      .then(() => {\n        state.loading = false;\n        btn.classList.remove('spin');\n      });\n  }\n\n  function applyData(data) {\n    state.data = data;\n    state.parts = data.parts;\n    state.loadedAt = Date.now();\n    state.activity = null;\n    state.partActivity = {};\n    $('#screen').hidden = true;\n    $('#app').hidden = false;\n    $('#app-name').textContent = data.settings.appName;\n    document.title = data.settings.appName;\n    renderAll();\n  }\n\n  // -------------------------------------------------------------------------\n  // Stock status (mirrors the server) and the numbers shown on screen\n  // -------------------------------------------------------------------------\n\n  function partById(id) { return state.parts.find((p) => p.id === id) || null; }\n  function unsaved(id) { return (state.pending[id] || 0) + (state.inflight[id] || 0); }\n  function qtyOf(p) {\n    const d = unsaved(p.id);\n    if (p.quantity === null && !d) return null;\n    return round((p.quantity || 0) + d);\n  }\n  function statusOf(p) {\n    const q = qtyOf(p);\n    if (q !== null && q <= 0) return 'out';\n    if (p.minQty !== null && q !== null && q <= p.minQty) return 'low';\n    return 'ok';\n  }\n\n  function counts() {\n    const c = { all: state.parts.length, low: 0, out: 0, ordered: 0 };\n    state.parts.forEach((p) => {\n      const s = statusOf(p);\n      if (s !== 'ok') c.low++;\n      if (s === 'out') c.out++;\n      if (p.ordered) c.ordered++;\n    });\n    return c;\n  }\n\n  function visibleParts() {\n    const terms = state.search.trim().toLowerCase().split(/\\s+/).filter(Boolean);\n    const list = state.parts.filter((p) => {\n      const s = statusOf(p);\n      if (state.filter === 'low' && s === 'ok') return false;\n      if (state.filter === 'out' && s !== 'out') return false;\n      if (state.filter === 'ordered' && !p.ordered) return false;\n      if (state.category && (p.category || '') !== state.category) return false;\n      if (!terms.length) return true;\n      const hay = [p.name, p.partNumber, p.id, p.category, p.location, p.supplier, p.notes].join(' ').toLowerCase();\n      return terms.every((t) => hay.indexOf(t) !== -1);\n    });\n    const rank = { out: 0, low: 1, ok: 2 };\n    const name = (p) => displayName(p);\n    list.sort((a, b) => {\n      if (state.sort === 'status') return rank[statusOf(a)] - rank[statusOf(b)] || byText(name(a), name(b));\n      if (state.sort === 'qty') {\n        const qa = qtyOf(a); const qb = qtyOf(b);\n        return (qa === null ? Infinity : qa) - (qb === null ? Infinity : qb) || byText(name(a), name(b));\n      }\n      if (state.sort === 'location') return byText(a.location || '￿', b.location || '￿') || byText(name(a), name(b));\n      if (state.sort === 'updated') return String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')) || byText(name(a), name(b));\n      return byText(name(a), name(b));\n    });\n    return list;\n  }\n\n  function uniqueValues(key) {\n    const seen = {};\n    state.parts.forEach((p) => { if (p[key]) seen[p[key]] = true; });\n    return Object.keys(seen).sort(byText);\n  }\n\n  // -------------------------------------------------------------------------\n  // Rendering\n  // -------------------------------------------------------------------------\n\n  function renderAll() {\n    renderBanners();\n    renderFilters();\n    renderCategoryFilter();\n    renderList();\n    if (state.selectedId) renderDrawer();\n    if (state.view === 'settings' && !state.settingsDirty) renderSettings();\n    if (state.view === 'activity') showActivity();\n  }\n\n  function renderBanners() {\n    const el = clear($('#banners'));\n    const d = state.data;\n    const owner = d.user.isOwner;\n    const banner = (kind, iconName, key, body, actions) => h('div', { class: 'banner ' + kind, role: kind === 'error' ? 'alert' : null },\n      icon(iconName, 20),\n      h('div', { class: 'grow' }, body, actions && actions.length ? h('div', { class: 'banner-actions' }, actions) : null),\n      key ? h('button', { class: 'close', type: 'button', 'aria-label': 'Dismiss', onclick: () => { state.dismissed[key] = true; renderBanners(); } }, icon('x', 16)) : null);\n\n    if (!d.settings.recipients.length) {\n      el.appendChild(banner('warn', 'mail', null,\n        [h('b', {}, 'Nobody gets low-stock emails yet. '), owner ? 'Add an email address in Settings.' : 'Ask ' + (d.info.sender || 'the app owner') + ' to add one in Settings.'],\n        owner ? [h('button', { class: 'btn small', type: 'button', onclick: () => { setView('settings'); focusLater('#s-recipients'); } }, 'Add email')] : null));\n    }\n    if (d.info.lastAlertError && !state.dismissed.alertError) {\n      el.appendChild(banner('error', 'alert', 'alertError',\n        [h('b', {}, 'The last low-stock email couldn\\'t be sent. '), d.info.lastAlertError.message, ' (', timeAgo(d.info.lastAlertError.at), '). It will be retried every hour.']));\n    }\n    if (owner && !d.info.alertsInstalled && !state.dismissed.triggers) {\n      el.appendChild(banner('warn', 'clock', 'triggers',\n        [h('b', {}, 'Automatic checks are off. '), 'Emails still go out for changes made here, but not for edits made straight in the spreadsheet. To fix that, open the spreadsheet and choose Inventory > Set up / repair.']));\n    }\n    if (!d.user.email && !state.actor && !state.dismissed.name && BOOT.mode !== 'sidebar') {\n      el.appendChild(banner('info', 'user', 'name',\n        'Add your name so your changes show up under it in the activity log.',\n        [h('button', { class: 'btn small', type: 'button', onclick: () => askName() }, 'Add my name')]));\n    }\n  }\n\n  function renderFilters() {\n    const c = counts();\n    const el = clear($('#filters'));\n    [\n      { key: 'all', label: 'All parts', short: 'All', n: c.all },\n      { key: 'low', label: 'Low stock', short: 'Low', n: c.low },\n      { key: 'out', label: 'Out of stock', short: 'Out', n: c.out },\n      { key: 'ordered', label: 'On order', short: 'Ordered', n: c.ordered },\n    ].forEach((f) => {\n      el.appendChild(h('button', {\n        class: 'filter ' + f.key + (f.n > 0 ? ' has' : ''),\n        type: 'button',\n        'data-filter': f.key,\n        'aria-pressed': String(state.filter === f.key),\n        onclick: () => setFilter(state.filter === f.key && f.key !== 'all' ? 'all' : f.key),\n      }, h('span', { class: 'long' }, f.label), h('span', { class: 'short' }, f.short), h('span', { class: 'count' }, fmt(f.n))));\n    });\n  }\n\n  function renderCategoryFilter() {\n    const select = $('#category-filter');\n    const cats = uniqueValues('category');\n    select.hidden = !state.data.fields.category || !cats.length;\n    if (state.category && cats.indexOf(state.category) === -1) state.category = '';\n    clear(select).appendChild(h('option', { value: '' }, 'All categories'));\n    cats.forEach((c) => select.appendChild(h('option', { value: c, selected: c === state.category }, c)));\n    $('#sort').value = state.sort;\n  }\n\n  function renderList() {\n    const el = clear($('#list'));\n    const parts = visibleParts();\n    if (!state.parts.length) {\n      el.appendChild(h('div', { class: 'empty' },\n        h('div', { class: 'empty-icon' }, icon('box', 26)),\n        h('h3', {}, 'No parts yet'),\n        h('p', {}, 'Add your first part, with the link to reorder it and the quantity that should trigger a low-stock email. You can also type parts straight into the Inventory tab of the spreadsheet.'),\n        h('button', { class: 'btn primary', type: 'button', onclick: () => openPartForm(null) }, icon('plus'), 'Add a part')));\n      return;\n    }\n    if (!parts.length) {\n      el.appendChild(h('div', { class: 'empty' },\n        h('div', { class: 'empty-icon' }, icon('search', 24)),\n        h('h3', {}, 'No parts match'),\n        h('p', {}, state.filter === 'low' ? 'Nothing is low on stock right now.' : 'Try a different search or filter.'),\n        h('button', { class: 'btn', type: 'button', onclick: resetFilters }, 'Show all parts')));\n      return;\n    }\n    el.appendChild(h('div', { class: 'list-head', 'aria-hidden': 'true' },\n      h('span', {}, 'Part'), h('span', {}, 'Location'), h('span', {}, 'Min'), h('span', { class: 'right' }, 'On hand'), h('span', {})));\n    parts.slice(0, state.showCount).forEach((p) => el.appendChild(partRow(p)));\n    if (parts.length > state.showCount) {\n      el.appendChild(h('div', { class: 'more' },\n        h('button', { class: 'btn', type: 'button', onclick: () => { state.showCount += 150; renderList(); } },\n          'Show more (' + (parts.length - state.showCount) + ' more)')));\n    }\n  }\n\n  function badges(p) {\n    const s = statusOf(p);\n    return [\n      s === 'low' ? h('span', { class: 'badge low' }, 'Low') : null,\n      s === 'out' ? h('span', { class: 'badge out' }, 'Out') : null,\n      p.ordered ? h('span', { class: 'badge ordered' }, icon('truck', 13), 'On order') : null,\n    ];\n  }\n\n  function partRow(p) {\n    const s = statusOf(p);\n    const q = qtyOf(p);\n    const busy = unsaved(p.id) !== 0 || p.id in state.inflight;\n    const link = safeUrl(p.links[0]);\n    // Where to find it first; on wide screens the location has its own column.\n    const meta = [\n      p.location ? h('span', { class: 'meta-loc' }, p.location, p.partNumber || p.minQty !== null ? ' · ' : '') : null,\n      p.partNumber ? h('span', {}, '#' + p.partNumber) : null,\n      p.minQty !== null ? h('span', { class: 'meta-min' }, (p.partNumber ? ' · ' : '') + 'Min ' + fmt(p.minQty)) : null,\n    ].filter(Boolean);\n    let action = null;\n    if (p.ordered) {\n      action = h('span', { class: 'on-order' }, icon('truck', 15), 'On order');\n    } else if (s !== 'ok' && link) {\n      action = h('a', {\n        class: 'btn small primary', href: link, target: '_blank', rel: 'noopener noreferrer',\n        title: 'Opens ' + hostOf(link), onclick: () => offerMarkOrdered(p.id),\n      }, icon('cart', 16), 'Reorder');\n    }\n    return h('div', { class: 'row status-' + s, 'data-id': p.id },\n      h('button', { class: 'row-main', type: 'button', onclick: () => openPart(p.id) },\n        h('span', { class: 'row-title' }, h('span', { class: 'name' }, displayName(p)), statusBadge(p)),\n        meta.length ? h('span', { class: 'row-meta' }, meta) : null),\n      h('div', { class: 'row-loc' }, p.location || h('span', { class: 'muted' }, '—')),\n      h('div', { class: 'row-extra' },\n        h('div', { class: 'row-act' }, action),\n        h('div', { class: 'row-min' }, p.minQty !== null ? fmt(p.minQty) : h('span', { class: 'muted' }, '—'))),\n      h('div', { class: 'row-stock' },\n        h('div', { class: 'stepper' },\n          h('button', {\n            class: 'step', type: 'button', 'aria-label': 'Take one ' + displayName(p) + ' out',\n            disabled: q === null || q < 1, onclick: () => quickAdjust(p.id, -1),\n          }, icon('minus')),\n          h('span', { class: 'qty' + (busy ? ' saving' : ''), 'aria-label': fmt(q) + ' on hand' },\n            h('strong', {}, fmt(q)), p.unit ? h('small', {}, p.unit) : null),\n          h('button', {\n            class: 'step', type: 'button', 'aria-label': 'Put one ' + displayName(p) + ' back', onclick: () => quickAdjust(p.id, 1),\n          }, icon('plus')))));\n  }\n\n  function statusBadge(p) {\n    const s = statusOf(p);\n    if (s === 'low') return h('span', { class: 'badge low' }, 'Low');\n    if (s === 'out') return h('span', { class: 'badge out' }, 'Out');\n    return null;\n  }\n\n  /** After Reorder opens the supplier's page, offer to mark the part as ordered. */\n  function offerMarkOrdered(id) {\n    setTimeout(() => {\n      const p = partById(id);\n      if (!p || p.ordered) return;\n      toast('Placed the order for ' + displayName(p) + '?', 'cart', { label: 'Mark as ordered', run: () => setOrdered(id, true) });\n    }, 400);\n  }\n\n  function updateRow(id) {\n    const old = $('#list .row[data-id=\"' + CSS.escape(id) + '\"]');\n    const p = partById(id);\n    if (old && p) old.replaceWith(partRow(p));\n    renderFilters();\n    if (state.selectedId === id) renderDrawer();\n  }\n\n  // -------------------------------------------------------------------------\n  // Quick +/- buttons: taps are batched and sent once you stop tapping\n  // -------------------------------------------------------------------------\n\n  function quickAdjust(id, delta) {\n    const p = partById(id);\n    if (!p) return;\n    if ((qtyOf(p) || 0) + delta < 0) return;\n    state.pending[id] = round((state.pending[id] || 0) + delta);\n    updateRow(id);\n    clearTimeout(state.timers[id]);\n    state.timers[id] = setTimeout(() => flushAdjust(id), 700);\n  }\n\n  function flushAdjust(id) {\n    clearTimeout(state.timers[id]);\n    if (id in state.inflight) {\n      state.timers[id] = setTimeout(() => flushAdjust(id), 250);\n      return;\n    }\n    const delta = state.pending[id] || 0;\n    delete state.pending[id];\n    if (!delta) { updateRow(id); return; }\n    state.inflight[id] = delta;\n    api('apiAdjustStock', { id: id, mode: delta > 0 ? 'add' : 'remove', amount: Math.abs(delta) })\n      .then((res) => {\n        delete state.inflight[id];\n        replacePart(res.part);\n        announceAlert(res.alert);\n      })\n      .catch((err) => {\n        delete state.inflight[id];\n        updateRow(id);\n        if (!isAccessError(err)) toast(errorText(err), 'error');\n      });\n  }\n\n  function flushAll() { Object.keys(state.pending).forEach(flushAdjust); }\n\n  function replacePart(part) {\n    state.touched[part.id] = Date.now();\n    const i = state.parts.findIndex((p) => p.id === part.id);\n    if (i === -1) state.parts.push(part);\n    else state.parts[i] = part;\n    delete state.partActivity[part.id];\n    state.activity = null;\n    updateRow(part.id);\n  }\n\n  function announceAlert(alert) {\n    if (!alert) return;\n    if (alert.sent) toast('Low-stock email sent to ' + alert.recipients.join(', ') + '.', 'mail');\n    else if (alert.error) toast('Couldn\\'t send the low-stock email: ' + alert.error, 'error');\n    else if (alert.skipped) toast('This part is now low. No email was sent: ' + alert.skipped, 'alert');\n  }\n\n  // -------------------------------------------------------------------------\n  // Part details drawer\n  // -------------------------------------------------------------------------\n\n  let drawerReturnFocus = null;\n\n  function openPart(id) {\n    const p = partById(id);\n    if (!p) { toast('That part wasn\\'t found. It may have been deleted.', 'error'); return; }\n    if (state.selectedId !== id) state.adjustMode = statusOf(p) === 'out' ? 'add' : 'remove';\n    drawerReturnFocus = document.activeElement;\n    if (!state.selectedId) pushBackStop('part');\n    state.selectedId = id;\n    $('#drawer').hidden = false;\n    document.body.style.overflow = 'hidden';\n    renderDrawer();\n    const close = $('#drawer-close');\n    if (close) close.focus();\n  }\n\n  function closeDrawer(fromBackButton) {\n    if (!state.selectedId) return;\n    if (!fromBackButton) popBackStop('part');\n    state.selectedId = null;\n    $('#drawer').hidden = true;\n    document.body.style.overflow = '';\n    if (drawerReturnFocus && document.contains(drawerReturnFocus)) drawerReturnFocus.focus();\n    // Refreshes are held while the panel is open; catch up now if it's been a while.\n    if (state.data && Date.now() - state.loadedAt > 60000 && !busy()) load();\n  }\n\n  function renderDrawer() {\n    const panel = clear($('#drawer-panel'));\n    const p = partById(state.selectedId);\n    if (!p) { closeDrawer(); return; }\n    const s = statusOf(p);\n    const q = qtyOf(p);\n    const fields = state.data.fields;\n\n    panel.appendChild(h('div', { class: 'drawer-head' },\n      h('div', { class: 'grow' },\n        h('div', { class: 'eyebrow' }, [p.id, p.category].filter(Boolean).join(' · ')),\n        h('h2', { id: 'drawer-title' }, displayName(p)),\n        h('div', { class: 'row-title' },\n          s === 'ok' ? h('span', { class: 'badge ok' }, icon('check', 13), 'In stock') : null, badges(p),\n          p.partNumber ? h('span', { class: 'muted' }, 'Part # ' + p.partNumber) : null)),\n      h('button', { class: 'icon-btn', id: 'drawer-close', type: 'button', 'aria-label': 'Close', onclick: closeDrawer }, icon('x'))));\n\n    const body = h('div', { class: 'drawer-body' });\n    panel.appendChild(body);\n\n    // Stock and adjusting it\n    const draft = state.adjustDraft && state.adjustDraft.id === p.id ? state.adjustDraft : null;\n    const amount = h('input', {\n      id: 'adjust-amount', type: 'number', inputmode: 'decimal', min: '0', step: 'any', 'aria-label': 'Amount',\n    });\n    amount.value = draft && draft.mode === state.adjustMode ? draft.amount\n      : state.adjustMode === 'set' ? (q === null ? '' : String(q)) : '1';\n    const note = h('input', { id: 'adjust-note', type: 'text', maxlength: '300', placeholder: 'Note (optional), e.g. job or reason', 'aria-label': 'Note' });\n    note.value = draft ? draft.note : '';\n    const saveDraft = () => { state.adjustDraft = { id: p.id, mode: state.adjustMode, amount: amount.value, note: note.value }; };\n    amount.addEventListener('input', saveDraft);\n    note.addEventListener('input', saveDraft);\n    const preview = h('div', { class: 'adjust-preview', 'aria-live': 'polite' });\n    const apply = h('button', { class: 'btn primary block', type: 'submit' }, 'Apply');\n    const updatePreview = () => {\n      const n = Number(amount.value);\n      const cur = q || 0;\n      if (amount.value === '' || !(n >= 0)) { preview.textContent = ''; return; }\n      const next = state.adjustMode === 'add' ? cur + n : state.adjustMode === 'remove' ? cur - n : n;\n      preview.textContent = next < 0 ? 'Only ' + fmt(cur) + ' on hand.' : 'New quantity: ' + fmt(next) + (p.unit ? ' ' + p.unit : '') +\n        (p.minQty !== null && next <= p.minQty && !(p.minQty !== null && cur <= p.minQty) ? ' (low: an email will go out)' : '');\n    };\n    amount.addEventListener('input', updatePreview);\n    const modeButton = (mode, label) => h('button', {\n      type: 'button', 'aria-pressed': String(state.adjustMode === mode),\n      onclick: () => {\n        if (state.adjustDraft && state.adjustDraft.id === p.id) state.adjustDraft.mode = null;\n        state.adjustMode = mode;\n        renderDrawer();\n        focusLater('#adjust-amount', true);\n      },\n    }, label);\n    const form = h('form', {\n      class: 'form-grid', onsubmit: (e) => { e.preventDefault(); submitAdjust(p.id, amount.value, note.value, apply); },\n    },\n    h('div', { class: 'segmented', role: 'group', 'aria-label': 'Change type' },\n      modeButton('remove', 'Use'), modeButton('add', 'Restock'), modeButton('set', 'Set count')),\n    h('div', { class: 'adjust-row' }, amount, note), preview, apply);\n    updatePreview();\n\n    body.appendChild(h('div', { class: 'stock-card status-' + s },\n      h('div', { class: 'stock-top' },\n        h('div', {}, h('div', { class: 'eyebrow' }, 'On hand'), h('div', { class: 'stock-big' }, fmt(q), p.unit ? h('small', {}, p.unit) : null)),\n        h('div', { class: 'stock-levels' },\n          p.minQty !== null ? h('div', {}, 'Min ', h('b', {}, fmt(p.minQty)), ' (alert at or below)') : h('div', {}, 'No min set, so no alerts'),\n          p.reorderQty !== null ? h('div', {}, 'Reorder ', h('b', {}, fmt(p.reorderQty))) : null)),\n      form));\n\n    // Ordering\n    const links = p.links.map(safeUrl).filter(Boolean);\n    const orderBox = h('div', { class: 'order-links' });\n    links.forEach((url, i) => orderBox.appendChild(h('a', {\n      class: 'btn' + (i === 0 ? ' primary' : ''), href: url, target: '_blank', rel: 'noopener noreferrer',\n    }, h('span', {}, 'Order from ' + hostOf(url)), icon('external', 16))));\n    if (!links.length) {\n      orderBox.appendChild(h('button', { class: 'btn', type: 'button', onclick: () => openPartForm(p.id, 'f-link') }, icon('link', 16), 'Add an order link'));\n    }\n    if (p.ordered) {\n      orderBox.appendChild(h('div', { class: 'ordered-note' }, icon('truck'),\n        h('span', { class: 'grow' }, 'On order' + (p.orderedAt ? ' since ' + fmtDate(p.orderedAt) : '')),\n        h('button', { class: 'btn small', type: 'button', onclick: (e) => setOrdered(p.id, false, e.currentTarget) }, 'Clear')));\n    } else if (s !== 'ok') {\n      orderBox.appendChild(h('button', { class: 'btn', type: 'button', onclick: (e) => setOrdered(p.id, true, e.currentTarget) },\n        icon('check', 16), 'Mark as ordered'));\n    }\n    body.appendChild(h('section', {}, h('h3', { class: 'section-title' }, 'Reorder'), orderBox));\n\n    // Details\n    const dl = h('dl', { class: 'details' });\n    const row = (label, value) => { if (value !== null && value !== undefined && value !== '') add(dl, [h('dt', {}, label), h('dd', {}, value)]); };\n    if (fields.location) row('Location', p.location);\n    if (fields.supplier) row('Supplier', p.supplier);\n    if (fields.unitCost && p.unitCost !== null) row('Unit cost', fmtMoney(p.unitCost));\n    if (fields.notes) row('Notes', p.notes);\n    row('Last change', p.updatedAt ? fmtDate(p.updatedAt, true) + (p.updatedBy ? ' by ' + p.updatedBy : '') : '');\n    if (p.alerted) row('Alert emailed', p.alertSentAt ? fmtDate(p.alertSentAt, true) : 'Yes');\n    body.appendChild(h('section', {}, h('h3', { class: 'section-title' }, 'Details'), dl));\n\n    body.appendChild(h('div', { class: 'drawer-actions' },\n      h('button', { class: 'btn', type: 'button', onclick: () => openPartForm(p.id) }, icon('edit', 16), 'Edit'),\n      h('button', { class: 'btn', type: 'button', onclick: () => copyPartLink(p) }, icon('copy', 16), 'Copy link'),\n      h('button', { class: 'btn danger', type: 'button', onclick: () => deletePart(p) }, icon('trash', 16), 'Delete')));\n\n    // Recent activity for this part\n    const history = h('div', { id: 'part-history' });\n    body.appendChild(h('section', {}, h('h3', { class: 'section-title' }, 'Recent activity'), history));\n    renderPartActivity(p.id);\n  }\n\n  function submitAdjust(id, amountText, note, button) {\n    const amount = Number(amountText);\n    if (amountText === '' || !(amount >= 0)) { toast('Enter an amount.', 'error'); return; }\n    if (state.adjustMode !== 'set' && amount === 0) { toast('Enter an amount greater than zero.', 'error'); return; }\n    const mode = state.adjustMode;\n    const p = partById(id);\n    const before = qtyOf(p);\n    button.disabled = true;\n    api('apiAdjustStock', { id: id, mode: mode, amount: amount, note: note })\n      .then((res) => {\n        state.adjustDraft = null;\n        replacePart(res.part);\n        const verb = mode === 'add' ? 'Restocked ' + fmt(amount) : mode === 'remove' ? 'Used ' + fmt(amount) : 'Count set to ' + fmt(amount);\n        toast(verb + ' · ' + fmt(res.part.quantity) + ' on hand' + (mode === 'set' && before !== null ? ' (was ' + fmt(before) + ')' : ''), 'check');\n        announceAlert(res.alert);\n      })\n      .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); })\n      .then(() => { button.disabled = false; });\n  }\n\n  function setOrdered(id, ordered, button) {\n    if (button) button.disabled = true;\n    api('apiSetOrdered', id, ordered)\n      .then((res) => {\n        replacePart(res.part);\n        toast(ordered ? 'Marked as ordered. Reminders skip it until it\\'s restocked.' : 'Order mark cleared.', 'truck');\n      })\n      .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); if (button) button.disabled = false; });\n  }\n\n  function loadPartActivity(id) {\n    if (state.partActivityLoading[id]) return;\n    state.partActivityLoading[id] = true;\n    api('apiGetActivity', { partId: id, limit: 12 })\n      .then((res) => { state.partActivity[id] = res.entries; }, () => { state.partActivity[id] = []; })\n      .then(() => {\n        delete state.partActivityLoading[id];\n        if (state.selectedId === id) renderPartActivity(id);\n      });\n  }\n\n  function renderPartActivity(id) {\n    const el = $('#part-history');\n    if (!el) return;\n    clear(el);\n    const entries = state.partActivity[id];\n    if (!entries) {\n      el.appendChild(h('div', { class: 'card-list' }, h('div', { class: 'skeleton' }, h('div', { class: 'bar', style: 'width:60%' }), h('div', { class: 'bar', style: 'width:35%' }))));\n      loadPartActivity(id);\n      return;\n    }\n    if (!entries.length) { el.appendChild(h('p', { class: 'muted' }, 'Nothing yet.')); return; }\n    el.appendChild(h('ul', { class: 'timeline card-list' }, entries.map((e) => activityItem(e, false))));\n  }\n\n  function copyPartLink(p) {\n    const base = state.data.settings.appUrl;\n    if (!base) { toast('Add the web app link in Settings first (it ends in /exec).', 'info'); return; }\n    const url = base + (base.indexOf('?') === -1 ? '?' : '&') + 'part=' + encodeURIComponent(p.id);\n    copyText(url).then((ok) => toast(ok ? 'Link copied. It opens this part directly (handy for QR labels on bins).' : url, ok ? 'check' : 'link'));\n  }\n\n  function copyText(text) {\n    const fallback = () => {\n      const ta = h('textarea', { style: 'position:fixed;opacity:0;top:0;left:0' });\n      ta.value = text;\n      document.body.appendChild(ta);\n      ta.select();\n      let ok = false;\n      try { ok = document.execCommand('copy'); } catch (err) { ok = false; }\n      ta.remove();\n      return ok;\n    };\n    if (navigator.clipboard && navigator.clipboard.writeText) {\n      return navigator.clipboard.writeText(text).then(() => true, () => fallback());\n    }\n    return Promise.resolve(fallback());\n  }\n\n  function deletePart(p) {\n    confirmDialog({\n      title: 'Delete ' + displayName(p) + '?',\n      message: 'This removes the part and its row from the spreadsheet. Its history stays in the Activity Log tab.',\n      ok: 'Delete part',\n      danger: true,\n    }).then((yes) => {\n      if (!yes) return;\n      api('apiDeletePart', p.id)\n        .then(() => {\n          state.deleted[p.id] = Date.now();\n          state.parts = state.parts.filter((x) => x.id !== p.id);\n          state.activity = null;\n          closeDrawer();\n          renderAll();\n          toast('Deleted ' + displayName(p) + '.', 'trash');\n        })\n        .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); });\n    });\n  }\n\n  // -------------------------------------------------------------------------\n  // Add / edit form\n  // -------------------------------------------------------------------------\n\n  const FORM_FIELDS = ['name', 'partNumber', 'category', 'location', 'supplier', 'quantity', 'unit', 'minQty', 'reorderQty', 'unitCost', 'link', 'notes'];\n  const NUMBER_FIELDS = ['quantity', 'minQty', 'reorderQty', 'unitCost'];\n\n  function formValue(p, key) {\n    if (!p) return key === 'quantity' ? '' : '';\n    if (key === 'link') return p.links.length ? p.links.join('\\n') : p.link;\n    const v = p[key];\n    return v === null || v === undefined ? '' : String(v);\n  }\n\n  function openPartForm(id, focusId) {\n    const p = id ? partById(id) : null;\n    state.editing = p ? p.id : null;\n    const dialog = $('#part-dialog');\n    const fields = state.data.fields;\n    $('#part-dialog-title').textContent = p ? 'Edit ' + displayName(p) : 'Add part';\n    $('#part-save').textContent = p ? 'Save changes' : 'Add part';\n    $('#part-form-error').hidden = true;\n    $('#dup-hint').hidden = true;\n    $$('[data-field]', dialog).forEach((el) => { el.hidden = !fields[el.getAttribute('data-field')]; });\n    FORM_FIELDS.forEach((key) => {\n      const input = $('#f-' + key);\n      input.value = formValue(p, key);\n      input.removeAttribute('aria-invalid');\n      input.dataset.original = input.value;\n    });\n    if (!p) $('#f-quantity').value = '0';\n    ['category', 'location', 'supplier'].forEach((key) => fillDatalist('#dl-' + key, uniqueValues(key)));\n    fillDatalist('#dl-unit', ['ea', 'pcs', 'box', 'pack', 'bag', 'roll', 'ft', 'm', 'lb', 'kg', 'gal', 'L'].concat(uniqueValues('unit')));\n    if (!dialog.open) openDialog(dialog);\n    focusLater('#' + (focusId || 'f-name'), true);\n  }\n\n  function fillDatalist(sel, values) {\n    const seen = {};\n    const el = clear($(sel));\n    values.forEach((v) => { if (!seen[v]) { seen[v] = true; el.appendChild(h('option', { value: v })); } });\n  }\n\n  function checkDuplicate() {\n    const value = $('#f-partNumber').value.trim().toLowerCase();\n    const hint = $('#dup-hint');\n    const other = value && state.parts.find((p) => p.id !== state.editing && (p.partNumber || '').toLowerCase() === value);\n    hint.hidden = !other;\n    if (other) hint.textContent = displayName(other) + ' (' + other.id + ') already has this part number.';\n  }\n\n  function submitPartForm(e) {\n    e.preventDefault();\n    const errorBox = $('#part-form-error');\n    errorBox.hidden = true;\n    const fields = state.data.fields;\n    const isNew = !state.editing;\n    const input = {};\n    let problem = null;\n    FORM_FIELDS.forEach((key) => {\n      const el = $('#f-' + key);\n      const container = el.closest('[data-field]');\n      if (container && container.hidden) return;\n      if (key === 'unit' && !fields.unit) return;\n      el.removeAttribute('aria-invalid');\n      const value = el.value.trim();\n      if (NUMBER_FIELDS.indexOf(key) !== -1 && value !== '' && !(Number(value) >= 0)) {\n        el.setAttribute('aria-invalid', 'true');\n        problem = problem || { el: el, message: 'Numbers must be 0 or more.' };\n      }\n      if (isNew || value !== el.dataset.original.trim()) input[key] = value;\n    });\n    if (!$('#f-name').value.trim()) problem = { el: $('#f-name'), message: 'Give the part a name.' };\n    if ($('#f-quantity').value.trim() === '') problem = problem || { el: $('#f-quantity'), message: 'Enter the quantity on hand (0 if there are none).' };\n    if (problem) {\n      problem.el.setAttribute('aria-invalid', 'true');\n      showFormError(problem.message);\n      problem.el.focus();\n      return;\n    }\n    if (!isNew && !Object.keys(input).length) { $('#part-dialog').close(); return; }\n    if (!isNew) input.id = state.editing;\n    const save = $('#part-save');\n    save.disabled = true;\n    save.textContent = 'Saving…';\n    api('apiSavePart', input)\n      .then((res) => {\n        $('#part-dialog').close();\n        replacePart(res.part);\n        if (isNew) {\n          renderAll();\n          const row = $('#list .row[data-id=\"' + CSS.escape(res.part.id) + '\"]');\n          if (row) { row.classList.add('flash'); row.scrollIntoView({ block: 'nearest' }); }\n          toast('Added ' + displayName(res.part) + ' (' + res.part.id + ').', 'check');\n        } else {\n          renderAll();\n          toast('Saved.', 'check');\n        }\n        announceAlert(res.alert);\n      })\n      .catch((err) => { if (!isAccessError(err)) showFormError(errorText(err)); })\n      .then(() => { save.disabled = false; save.textContent = isNew ? 'Add part' : 'Save changes'; });\n  }\n\n  function showFormError(message) {\n    const box = $('#part-form-error');\n    clear(box).appendChild(icon('alert', 18));\n    box.appendChild(document.createTextNode(message));\n    box.hidden = false;\n    box.scrollIntoView({ block: 'nearest' });\n  }\n\n  // -------------------------------------------------------------------------\n  // Activity\n  // -------------------------------------------------------------------------\n\n  const ACTION_STYLE = {\n    'Used': ['used', 'minus'],\n    'Restocked': ['restocked', 'plus'],\n    'Added': ['added', 'box'],\n    'Count set': ['', 'edit'],\n    'Edited': ['', 'edit'],\n    'Deleted': ['deleted', 'trash'],\n    'Alert emailed': ['alert', 'mail'],\n    'Marked ordered': ['ordered', 'truck'],\n    'Order mark cleared': ['', 'truck'],\n  };\n\n  function activityItem(e, showPart) {\n    const style = ACTION_STYLE[e.action] || ['', 'clock'];\n    const change = typeof e.change === 'number' && e.action !== 'Added'\n      ? h('span', { class: 'chg ' + (e.change < 0 ? 'neg' : 'pos') }, (e.change > 0 ? '+' : '') + fmt(e.change)) : null;\n    const partLabel = showPart\n      ? (partById(e.partId)\n        ? h('button', { class: 'link-btn', type: 'button', onclick: () => openPart(e.partId) }, e.part)\n        : h('b', {}, e.part))\n      : null;\n    return h('li', { class: 'act ' + style[0] },\n      h('span', { class: 'act-icon' }, icon(style[1], 16)),\n      h('div', {},\n        h('div', { class: 'act-line' }, e.action, change ? [' ', change] : null, partLabel ? [' · ', partLabel] : null,\n          typeof e.qtyAfter === 'number' && e.action !== 'Alert emailed' ? h('span', { class: 'muted' }, ' → ' + fmt(e.qtyAfter)) : null),\n        h('div', { class: 'act-sub' }, [e.by, timeAgo(e.at)].filter(Boolean).join(' · '), e.note ? ' · ' + e.note : '')));\n  }\n\n  function showActivity() {\n    const fresh = state.activity && Date.now() - state.activityLoadedAt < 30000;\n    if (fresh) { renderActivity(); return; }\n    if (!state.activity) renderActivity();\n    loadActivity();\n  }\n\n  function loadActivity() {\n    const btn = $('#activity-refresh');\n    btn.disabled = true;\n    return api('apiGetActivity', { limit: state.activityLimit })\n      .then((res) => { state.activity = res.entries; state.activityLoadedAt = Date.now(); renderActivity(); })\n      .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); })\n      .then(() => { btn.disabled = false; });\n  }\n\n  function renderActivity() {\n    const el = clear($('#activity'));\n    if (!state.activity) {\n      const list = h('div', { class: 'card-list' });\n      for (let i = 0; i < 5; i++) list.appendChild(h('div', { class: 'skeleton' }, h('div', { class: 'bar', style: 'width:' + (40 + i * 9) + '%' }), h('div', { class: 'bar', style: 'width:24%' })));\n      el.appendChild(list);\n      return;\n    }\n    if (!state.activity.length) {\n      el.appendChild(h('div', { class: 'card-list' }, h('div', { class: 'empty' },\n        h('div', { class: 'empty-icon' }, icon('history', 24)), h('h3', {}, 'No activity yet'),\n        h('p', {}, 'Every change made in the app shows up here, along with the low-stock emails that were sent.'))));\n      return;\n    }\n    el.appendChild(h('ul', { class: 'timeline card-list' }, state.activity.map((e) => activityItem(e, true))));\n    if (state.activity.length >= state.activityLimit && state.activityLimit < 500) {\n      el.appendChild(h('div', { class: 'more' }, h('button', {\n        class: 'btn', type: 'button',\n        onclick: () => { state.activityLimit = Math.min(state.activityLimit + 200, 500); loadActivity(); },\n      }, 'Show older')));\n    }\n  }\n\n  // -------------------------------------------------------------------------\n  // Settings\n  // -------------------------------------------------------------------------\n\n  const REMINDERS = [['Off', 'Off'], ['Daily', 'Every day'], ['Weekdays', 'Weekdays (Mon–Fri)'], ['Weekly', 'Weekly (Mondays)']];\n\n  function hourLabel(hour) {\n    const d = new Date(2000, 0, 1, hour);\n    return d.toLocaleTimeString(undefined, { hour: 'numeric' });\n  }\n\n  function renderSettings() {\n    state.settingsDirty = false;\n    const el = clear($('#settings'));\n    const d = state.data;\n    const s = d.settings;\n    const owner = d.user.isOwner;\n    const off = !owner;\n\n    const recipients = h('textarea', { id: 's-recipients', rows: '2', disabled: off, spellcheck: 'false', placeholder: 'buyer@yourcompany.com, you@gmail.com' });\n    recipients.value = s.recipients.join(', ');\n    const alerts = h('input', { id: 's-alerts', type: 'checkbox', disabled: off, checked: s.alertsEnabled });\n    const reminder = h('select', { id: 's-reminder', disabled: off }, REMINDERS.map((r) => h('option', { value: r[0], selected: r[0] === s.reminder }, r[1])));\n    const hour = h('select', { id: 's-hour', disabled: off, 'aria-label': 'Reminder time' });\n    for (let i = 0; i < 24; i++) hour.appendChild(h('option', { value: String(i), selected: i === s.reminderHour }, hourLabel(i)));\n    const syncHour = () => { hour.disabled = off || reminder.value === 'Off'; };\n    reminder.addEventListener('change', syncHour);\n    syncHour();\n    const appName = h('input', { id: 's-appname', maxlength: '60', disabled: off });\n    appName.value = s.appName;\n    const code = h('input', { id: 's-code', maxlength: '64', disabled: off, autocomplete: 'off', spellcheck: 'false', placeholder: owner ? 'Not set' : (s.accessCodeSet ? 'Set' : 'Not set') });\n    code.value = s.accessCode;\n    const appUrl = h('input', { id: 's-appurl', type: 'url', disabled: off, spellcheck: 'false', placeholder: 'Paste the web app link (ends in /exec)' });\n    appUrl.value = s.appUrl;\n\n    const saveBtn = h('button', { class: 'btn primary', type: 'submit', disabled: off }, 'Save settings');\n    const form = h('form', { class: 'card', onsubmit: (e) => {\n      e.preventDefault();\n      saveBtn.disabled = true;\n      api('apiSaveSettings', {\n        recipients: recipients.value, alertsEnabled: alerts.checked, reminder: reminder.value,\n        reminderHour: Number(hour.value), appName: appName.value, accessCode: code.value, appUrl: appUrl.value,\n      })\n        .then((res) => {\n          state.settingsDirty = false;\n          state.data.settings = res.settings;\n          $('#app-name').textContent = res.settings.appName;\n          document.title = res.settings.appName;\n          renderBanners();\n          renderSettings();\n          toast('Settings saved.', 'check');\n        })\n        .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); saveBtn.disabled = false; });\n    } },\n    h('h2', {}, 'Low-stock emails'),\n    h('p', { class: 'sub' }, owner ? 'Who gets told when parts run low, and when.' : 'Only ' + (d.info.sender || 'the person who set up the app') + ' can change these.'),\n    h('div', { class: 'form-grid two' },\n      h('div', { class: 'field span-2' }, h('label', { for: 's-recipients' }, 'Send alerts to'), recipients,\n        h('span', { class: 'hint' }, 'Separate several addresses with commas. Any email address works, including Outlook.')),\n      h('label', { class: 'switch span-2', for: 's-alerts' }, alerts, h('span', { class: 'track' }),\n        h('span', { class: 'switch-text' }, h('b', {}, 'Email as soon as a part runs low'),\n          h('span', { class: 'hint' }, 'One email per part when it drops to or below its alert level. It can email again after the part is restocked.'))),\n      h('div', { class: 'field' }, h('label', { for: 's-reminder' }, 'Reminder of parts still low'), reminder),\n      h('div', { class: 'field' }, h('label', { for: 's-hour' }, 'Reminder time'), hour),\n      h('div', { class: 'field span-2' }, h('span', { class: 'hint' }, 'The reminder lists parts that are still low and not marked as ordered. Times use the spreadsheet\\'s time zone (' + d.info.timeZone + ').'))),\n    h('h2', { style: 'margin-top:22px' }, 'App'),\n    h('p', { class: 'sub' }, 'How the app looks and who can open it.'),\n    h('div', { class: 'form-grid two' },\n      h('div', { class: 'field' }, h('label', { for: 's-appname' }, 'App name'), appName),\n      h('div', { class: 'field' }, h('label', { for: 's-code' }, 'Access code'), code,\n        h('span', { class: 'hint' }, 'Optional. People enter it once per device. Use one if anyone with the link can open the app.')),\n      h('div', { class: 'field span-2' }, h('label', { for: 's-appurl' }, 'Web app link'), appUrl,\n        h('span', { class: 'hint' }, 'Used for the \"View in the app\" links in emails. It\\'s usually filled in when the app is first opened; if it\\'s empty, paste the link from Deploy > Manage deployments.'))),\n    h('div', { class: 'card-actions' }, saveBtn));\n    const markDirty = () => { state.settingsDirty = true; };\n    form.addEventListener('input', markDirty);\n    form.addEventListener('change', markDirty);\n    el.appendChild(form);\n\n    // Emails card\n    const facts = h('ul', { class: 'facts' },\n      h('li', {}, icon('send', 16), h('span', {}, 'Emails are sent from ', h('b', {}, d.info.sender || 'the owner\\'s Google account'), '.')),\n      d.info.emailQuota !== null ? h('li', {}, icon('mail', 16), h('span', {}, d.info.emailQuota + ' emails left today (Google\\'s daily limit).')) : null,\n      h('li', {}, icon('clock', 16), h('span', {}, d.info.alertsInstalled\n        ? 'Automatic check every hour' + (d.info.lastCheck ? ', last one ' + timeAgo(d.info.lastCheck) : '') + '.'\n        : 'Automatic checks are off. In the spreadsheet, choose Inventory > Set up / repair.')),\n      d.info.lastAlertError ? h('li', {}, icon('alert', 16), h('span', {}, 'Last problem: ' + d.info.lastAlertError.message + ' (' + timeAgo(d.info.lastAlertError.at) + ')')) : null);\n    const testBtn = h('button', { class: 'btn', type: 'button', disabled: off || !s.recipients.length, onclick: () => sendEmail('apiSendTestEmail', testBtn) }, icon('send', 16), 'Send test email');\n    const remindBtn = h('button', { class: 'btn', type: 'button', disabled: off || !s.recipients.length, onclick: () => sendEmail('apiSendReminderNow', remindBtn) }, icon('mail', 16), 'Email the low-stock list now');\n    el.appendChild(h('section', { class: 'card' },\n      h('h2', {}, 'Emails'),\n      h('p', { class: 'sub' }, 'Check that alerts reach your inbox. If they don\\'t, look in the spam folder and mark them as \"not spam\".'),\n      owner ? h('div', { class: 'card-actions', style: 'margin-top:0' }, testBtn, remindBtn) : null,\n      facts));\n\n    // This device\n    const nameInput = h('input', { id: 's-name', maxlength: '60', autocomplete: 'name', placeholder: d.user.email || 'e.g. Sam R.' });\n    nameInput.value = state.actor;\n    nameInput.addEventListener('change', () => { setActor(nameInput.value); toast('Saved on this device.', 'check'); });\n    el.appendChild(h('section', { class: 'card' },\n      h('h2', {}, 'You'),\n      h('p', { class: 'sub' }, d.user.email ? 'You\\'re signed in as ' + d.user.email + '. Changes are logged under that address.' : 'Google doesn\\'t tell the app who you are, so changes are logged under this name. It\\'s saved on this device.'),\n      d.user.email ? null : h('div', { class: 'field' }, h('label', { for: 's-name' }, 'Your name'), nameInput),\n      state.code && !owner ? h('div', { class: 'card-actions' }, h('button', { class: 'btn', type: 'button', onclick: forgetCode }, icon('lock', 16), 'Forget the access code on this device')) : null));\n\n    el.appendChild(h('section', { class: 'card' },\n      h('h2', {}, 'Spreadsheet'),\n      h('p', { class: 'sub' }, 'Everything is stored in the Google Sheet \"' + d.info.spreadsheetName + '\": the Inventory, Settings and Activity Log tabs. Edits there show up here after a refresh.'),\n      h('div', { class: 'card-actions', style: 'margin-top:0' },\n        h('a', { class: 'btn', href: safeUrl(d.info.spreadsheetUrl) || '#', target: '_blank', rel: 'noopener noreferrer' }, icon('sheet', 16), 'Open the spreadsheet'),\n        h('span', { class: 'muted' }, 'Version ' + d.version))));\n  }\n\n  function sendEmail(fn, button) {\n    button.disabled = true;\n    api(fn)\n      .then((res) => {\n        if (fn === 'apiSendTestEmail') toast('Test email sent to ' + res.recipients.join(', ') + '.', 'mail');\n        else toast(res.sent ? 'Sent the list of ' + plural(res.count + res.onOrder, 'low part', 'low parts') + '.' : 'Nothing is low right now, so no email was sent.', 'mail');\n      })\n      .catch((err) => { if (!isAccessError(err)) toast(errorText(err), 'error'); })\n      .then(() => { button.disabled = false; });\n  }\n\n  function setActor(name) {\n    state.actor = String(name || '').trim().slice(0, 60);\n    store.set('actor', state.actor);\n    renderBanners();\n  }\n\n  function forgetCode() {\n    state.code = '';\n    store.set('code', '');\n    showLock();\n  }\n\n  // -------------------------------------------------------------------------\n  // Toasts and small dialogs\n  // -------------------------------------------------------------------------\n\n  /** action (optional) = { label, run } shows a button in the toast. */\n  function toast(message, kind, action) {\n    const box = $('#toasts');\n    const el = h('div', { class: 'toast' + (kind === 'error' ? ' error' : ''), role: kind === 'error' ? 'alert' : 'status' },\n      icon(kind === 'error' ? 'alert' : ICONS[kind] ? kind : 'info', 18),\n      h('div', { class: 'grow' }, message),\n      action ? h('button', { class: 'toast-action', type: 'button', onclick: () => { dismiss(); action.run(); } }, action.label) : null,\n      h('button', { type: 'button', 'aria-label': 'Dismiss', onclick: () => dismiss() }, icon('x', 16)));\n    const dismiss = () => { el.classList.add('leaving'); setTimeout(() => el.remove(), 200); };\n    box.appendChild(el);\n    while (box.children.length > 3) box.firstChild.remove();\n    setTimeout(dismiss, action ? 12000 : kind === 'error' ? 9000 : 4500);\n  }\n\n  function confirmDialog(opts) {\n    const dialog = $('#confirm-dialog');\n    $('#confirm-title').textContent = opts.title;\n    $('#confirm-message').textContent = opts.message;\n    const ok = $('#confirm-ok');\n    ok.textContent = opts.ok || 'OK';\n    ok.className = 'btn ' + (opts.danger ? 'danger-fill' : 'primary');\n    openDialog(dialog);\n    return new Promise((resolve) => {\n      const done = (value) => {\n        ok.onclick = null;\n        $('#confirm-cancel').onclick = null;\n        dialog.onclose = null;\n        if (dialog.open) dialog.close();\n        resolve(value);\n      };\n      ok.onclick = () => done(true);\n      $('#confirm-cancel').onclick = () => done(false);\n      dialog.onclose = () => done(false);\n    });\n  }\n\n  function askName() {\n    const dialog = $('#name-dialog');\n    $('#name-input').value = state.actor;\n    openDialog(dialog);\n    focusLater('#name-input', true);\n  }\n\n  // A phone's back button should close the part panel or form on top, not\n  // leave the app, so opening one adds a history entry and closing removes it.\n  let closingFromBack = false;\n  let ownBackPending = 0;\n\n  function pushBackStop(kind) {\n    try { history.pushState({ partsInventory: kind }, ''); } catch (err) { /* history not available here */ }\n  }\n\n  function popBackStop(kind) {\n    try {\n      if (history.state && history.state.partsInventory === kind) {\n        ownBackPending++;\n        history.back();\n      }\n    } catch (err) { /* history not available here */ }\n  }\n\n  function openDialog(dialog) {\n    dialog.showModal();\n    pushBackStop('dialog');\n  }\n\n  function onBack() {\n    // Our own history.back() after closing something isn't the back button.\n    if (ownBackPending > 0) { ownBackPending--; return; }\n    const dialog = $('dialog[open]');\n    closingFromBack = true;\n    if (dialog) dialog.close();\n    else if (state.selectedId) closeDrawer(true);\n    closingFromBack = false;\n  }\n\n  function focusLater(sel, select) {\n    setTimeout(() => {\n      const el = $(sel);\n      if (!el) return;\n      el.focus();\n      if (select && el.select) el.select();\n    }, 30);\n  }\n\n  // -------------------------------------------------------------------------\n  // Views, filters and full-screen states\n  // -------------------------------------------------------------------------\n\n  function setView(view) {\n    state.view = view;\n    $$('.tab').forEach((t) => t.setAttribute('aria-selected', String(t.getAttribute('data-view') === view)));\n    ['inventory', 'activity', 'settings'].forEach((v) => { $('#view-' + v).hidden = v !== view; });\n    $('#add-btn').hidden = view !== 'inventory';\n    $('#fab-add').hidden = view !== 'inventory';\n    if (view === 'settings') renderSettings();\n    if (view === 'activity') showActivity();\n    window.scrollTo(0, 0);\n  }\n\n  function setFilter(filter) {\n    state.filter = filter;\n    state.showCount = 150;\n    renderFilters();\n    renderList();\n  }\n\n  function resetFilters() {\n    state.filter = 'all';\n    state.search = '';\n    state.category = '';\n    $('#search').value = '';\n    renderCategoryFilter();\n    renderFilters();\n    renderList();\n  }\n\n  /** Shows the access code screen. message is the server's reason, shown only when a code was tried. */\n  function showLock(message, lockedOut) {\n    const screen = clear($('#screen'));\n    $('#app').hidden = true;\n    closeDrawer();\n    const input = h('input', { id: 'code-input', type: 'password', autocomplete: 'current-password', 'aria-label': 'Access code', placeholder: 'Access code', required: true });\n    const error = h('div', { class: 'error-text', role: 'alert' });\n    const btn = h('button', { class: 'btn primary block', type: 'submit' }, 'Open');\n    screen.appendChild(h('div', { class: 'screen-card' },\n      h('div', { class: 'logo' }, icon('lock', 24)),\n      h('h1', {}, (state.data && state.data.settings.appName) || BOOT.appName || 'Parts Inventory'),\n      h('p', {}, 'Enter the access code to open the inventory. You only need to do this once on this device.'),\n      h('form', { onsubmit: (e) => {\n        e.preventDefault();\n        state.code = input.value.trim();\n        if (!state.code) return;\n        btn.disabled = true;\n        error.textContent = '';\n        load({ fromLock: true }).then(() => { btn.disabled = false; openDeepLink(); });\n      } }, input, error, btn)));\n    screen.hidden = false;\n    if (message && state.code) {\n      error.textContent = message;\n      if (lockedOut) {\n        // Too many wrong tries by someone: this code may be fine, so keep it for later.\n        input.value = state.code;\n      } else {\n        // This code was turned down: forget it.\n        state.code = '';\n        store.set('code', '');\n      }\n    }\n    setTimeout(() => input.focus(), 30);\n  }\n\n  function showFatal(message) {\n    const screen = clear($('#screen'));\n    $('#app').hidden = true;\n    screen.appendChild(h('div', { class: 'screen-card' },\n      h('div', { class: 'logo' }, icon('alert', 24)),\n      h('h1', {}, 'Couldn\\'t open the inventory'),\n      h('p', {}, 'If this is a new setup, open the spreadsheet and choose Inventory > Set up / repair, then try again.'),\n      h('pre', {}, message),\n      h('button', { class: 'btn primary block', type: 'button', onclick: () => { screen.hidden = true; showLoading(); load(); } }, 'Try again')));\n    screen.hidden = false;\n  }\n\n  function showLoading() {\n    $('#app').hidden = false;\n    const list = clear($('#list'));\n    for (let i = 0; i < 6; i++) {\n      list.appendChild(h('div', { class: 'skeleton' }, h('div', { class: 'bar', style: 'width:' + (30 + ((i * 17) % 40)) + '%' }), h('div', { class: 'bar', style: 'width:18%' })));\n    }\n  }\n\n  // -------------------------------------------------------------------------\n  // Start up\n  // -------------------------------------------------------------------------\n\n  function wire() {\n    $$('[data-icon]').forEach((el) => el.insertBefore(icon(el.getAttribute('data-icon'), el.classList.contains('logo') ? 18 : 18), el.firstChild));\n    $$('.tab').forEach((t) => t.addEventListener('click', () => setView(t.getAttribute('data-view'))));\n    $('#refresh-btn').addEventListener('click', () => { flushAll(); load(); if (state.view === 'activity') loadActivity(); });\n    $('#add-btn').addEventListener('click', () => openPartForm(null));\n    $('#fab-add').addEventListener('click', () => openPartForm(null));\n    $('#activity-refresh').addEventListener('click', loadActivity);\n    let searchTimer = null;\n    $('#search').addEventListener('input', (e) => {\n      clearTimeout(searchTimer);\n      searchTimer = setTimeout(() => { state.search = e.target.value; state.showCount = 150; renderList(); }, 120);\n    });\n    $('#category-filter').addEventListener('change', (e) => { state.category = e.target.value; renderList(); });\n    $('#sort').addEventListener('change', (e) => { state.sort = e.target.value; store.set('sort', state.sort); renderList(); });\n    $('[data-close-drawer]').addEventListener('click', closeDrawer);\n    $('#part-form').addEventListener('submit', submitPartForm);\n    $('#f-partNumber').addEventListener('input', checkDuplicate);\n    $$('[data-close-dialog]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));\n    $('#name-form').addEventListener('submit', (e) => {\n      e.preventDefault();\n      setActor($('#name-input').value);\n      $('#name-dialog').close();\n      if (state.actor) toast('Thanks, ' + state.actor + '.', 'check');\n    });\n    $('#name-skip').addEventListener('click', () => { state.dismissed.name = true; $('#name-dialog').close(); renderBanners(); });\n\n    document.addEventListener('keydown', (e) => {\n      const typing = /^(INPUT|TEXTAREA|SELECT)$/.test((e.target && e.target.tagName) || '');\n      if (e.key === 'Escape' && state.selectedId && !$('dialog[open]')) { closeDrawer(); return; }\n      if (e.key === '/' && !typing && state.view === 'inventory' && !$('dialog[open]') && !state.selectedId) {\n        e.preventDefault();\n        $('#search').focus();\n      }\n    });\n    document.addEventListener('visibilitychange', () => {\n      if (document.visibilityState === 'hidden') { flushAll(); return; }\n      if (state.data && Date.now() - state.loadedAt > 60000 && !busy()) load();\n    });\n    window.addEventListener('pagehide', flushAll);\n    window.addEventListener('popstate', onBack);\n    $$('dialog').forEach((d) => d.addEventListener('close', () => { if (!closingFromBack) popBackStop('dialog'); }));\n    setInterval(() => {\n      if (document.visibilityState === 'visible' && state.data && !busy() && Date.now() - state.loadedAt > 5 * 60000) load();\n    }, 60000);\n  }\n\n  /** True while someone is in the middle of something a background refresh would disturb. */\n  function busy() {\n    return !!$('dialog[open]') || !!state.selectedId || state.settingsDirty ||\n      Object.keys(state.pending).length > 0 || Object.keys(state.inflight).length > 0;\n  }\n\n  function start() {\n    wire();\n    if (BOOT.mode === 'sidebar') document.body.classList.add('is-sidebar');\n    if (BOOT.error) { showFatal(BOOT.error); return; }\n    if (BOOT.locked && !BOOT.data) {\n      if (state.code) { showLoading(); load({ fromLock: true }).then(openDeepLink); } else showLock();\n      return;\n    }\n    if (BOOT.data) { applyData(BOOT.data); openDeepLink(); return; }\n    showLoading();\n    load().then(openDeepLink);\n  }\n\n  function openDeepLink() {\n    if (!BOOT.part || !state.data) return;\n    const wanted = String(BOOT.part).trim();\n    BOOT.part = '';\n    const part = state.parts.find((p) => p.id.toUpperCase() === wanted.toUpperCase());\n    if (part) openPart(part.id);\n    else toast('Part ' + wanted + ' wasn\\'t found. It may have been deleted.', 'error');\n  }\n\n  start();\n})();\n</script>\n</body>\n</html>\n";
