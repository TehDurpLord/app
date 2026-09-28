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
 *   managed: kept up to date by the app (grey header in the sheet)
 * The other columns are optional: the app only shows them when they exist.
 */
const FIELDS = [
  {
    key: 'id', header: 'ID', type: 'text', core: true, managed: true, width: 80,
    aliases: ['Part ID', 'Item ID'],
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
    aliases: ['Ordered', 'Order Date', 'Date Ordered', 'On Order Since'],
    note: 'Set by "Mark as ordered" in the app. Cleared automatically once the part is restocked.',
  },
  {
    key: 'alertSentAt', header: 'Alert Sent', type: 'date', core: true, managed: true, width: 140,
    aliases: ['Last Alert', 'Alerted', 'Alert Sent On'],
    note: 'When the low-stock email for this part went out. Cleared once the part is restocked, so the next time it runs low a new email goes out.',
  },
  {
    key: 'updatedAt', header: 'Last Updated', type: 'date', core: true, managed: true, width: 140,
    aliases: ['Updated', 'Updated At', 'Last Modified', 'Modified'],
  },
  {
    key: 'updatedBy', header: 'Updated By', type: 'text', core: true, managed: true, width: 180,
    aliases: ['Modified By', 'Last Updated By', 'Changed By'],
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
    help: 'Link to the web app, used in emails. Filled in automatically the first time the web app is opened.',
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
  const template = HtmlService.createTemplateFromFile('Index');
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
function hourlyCheck() {
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
  const cards = parts.map(function (p) { return partCardHtml_(p, appUrl, tz); }).join('');
  const orderedCards = onOrder.length
    ? '<h2 style="font-size:16px;margin:24px 0 10px;color:#1f2937;">Already on order</h2>' +
      onOrder.map(function (p) { return partCardHtml_(p, appUrl, tz); }).join('')
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
  parts.forEach(function (p) { lines.push(partText_(p), ''); });
  if (onOrder.length) {
    lines.push('Already on order:', '');
    onOrder.forEach(function (p) { lines.push(partText_(p), ''); });
  }
  if (appUrl) lines.push('Inventory app: ' + appUrl);
  lines.push('Spreadsheet: ' + ss.getUrl(), '', footerNote);
  return { subject: subject, html: html, text: lines.join('\n') };
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
    // Any value in these columns counts (someone may type "yes"); dates are kept when there are dates.
    ordered: !isBlank_(orderedRaw),
    orderedAt: toDateOrNull_(orderedRaw),
    alerted: !isBlank_(alertRaw),
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
    .map(function (key) { return { col: inv.cols[key], value: toCell_(key, fields[key]) }; })
    .sort(function (a, b) { return a.col - b.col; });
  // One write per run of neighbouring columns keeps this fast without
  // touching any of your own columns in between.
  let i = 0;
  while (i < cells.length) {
    let j = i;
    while (j + 1 < cells.length && cells[j + 1].col === cells[j].col + 1) j++;
    const run = cells.slice(i, j + 1);
    inv.sheet.getRange(row, run[0].col, 1, run.length).setValues([run.map(function (c) { return c.value; })]);
    i = j + 1;
  }
}

/** Writes one field for several rows with a single range write. updates = [{ row, value }] */
function writeColumnValues_(inv, key, updates) {
  const col = inv.cols[key];
  if (!col || !updates.length) return;
  if (updates.length === 1) {
    inv.sheet.getRange(updates[0].row, col).setValue(toCell_(key, updates[0].value));
    return;
  }
  let min = Infinity;
  let max = -Infinity;
  updates.forEach(function (u) {
    min = Math.min(min, u.row);
    max = Math.max(max, u.row);
  });
  const range = inv.sheet.getRange(min, col, max - min + 1, 1);
  const block = range.getValues().map(function (r) { return [typeof r[0] === 'string' ? safeCell_(r[0]) : r[0]]; });
  updates.forEach(function (u) { block[u.row - min][0] = toCell_(key, u.value); });
  range.setValues(block);
}

function toCell_(key, value) {
  const field = FIELD_BY_KEY[key];
  if (value === null || value === undefined || value === '') return '';
  if (field.type === 'number') return Number(value);
  if (field.type === 'date') return isDate_(value) ? value : new Date(value);
  return safeCell_(String(value));
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
