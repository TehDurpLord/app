'use strict';
/**
 * A small in-memory stand-in for the Google Apps Script services that
 * src/Code.js uses (SpreadsheetApp, MailApp, ScriptApp, ...), so the real
 * server code can run under Node for tests and the local preview.
 *
 * It is deliberately strict: range sizes are checked, unknown methods don't
 * exist, and text written to a cell is parsed the way Google Sheets parses
 * typing ("00123" becomes 123, "=..." becomes a formula, a leading apostrophe
 * keeps text as text). That's what catches real bugs before deploying.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC_DIR = path.join(__dirname, '..', 'src');

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

/**
 * Parses a value the way Sheets does when it is set with setValue(s).
 * In a cell formatted as Plain text ('@') text is kept exactly as written,
 * leading apostrophe included, except that "=..." is still taken as a formula
 * (the worst case, so tests catch any formula that gets through).
 */
function userEntered(value, format) {
  if (value === undefined) throw new Error('Cannot write undefined to a cell.');
  if (value === null || value === '') return null;
  if (format === '@' && typeof value === 'string') return value[0] === '=' ? { f: value } : { v: value };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('Cannot write ' + value + ' to a cell.');
    return { v: value };
  }
  if (typeof value === 'boolean') return { v: value };
  if (value instanceof Date) return { v: new Date(value.getTime()) };
  if (typeof value !== 'string') throw new Error('Cannot write a ' + typeof value + ' to a cell.');
  if (value[0] === "'") return { v: value.slice(1) };
  if (value[0] === '=') return { f: value };
  const t = value.trim();
  if (/^[+-]?\d+(\.\d+)?$/.test(t)) return { v: Number(t) };
  if (/^(true|false)$/i.test(t)) return { v: t.toLowerCase() === 'true' };
  const date = /^(\d{1,4})[-/.](\d{1,2})(?:[-/.](\d{1,4}))?$/.exec(t);
  if (date) {
    const [, a, b, c] = date;
    return { v: a.length === 4 ? new Date(+a, +b - 1, +(c || 1)) : new Date(c ? +c : 2026, +a - 1, +b) };
  }
  return { v: value };
}

function formulaValue(formula) {
  const link = /^=HYPERLINK\(\s*"([^"]*)"\s*(?:[,;]\s*"([^"]*)")?\s*\)$/i.exec(formula);
  if (link) return link[2] !== undefined ? link[2] : link[1];
  return '#ERROR!';
}

function cellValue(cell) {
  if (!cell) return '';
  if (cell.f) return formulaValue(cell.f);
  return cell.v instanceof Date ? new Date(cell.v.getTime()) : cell.v;
}

function isEmptyCell(cell) {
  return !cell || (!cell.f && (cell.v === '' || cell.v === null || cell.v === undefined));
}

class FakeRichText {
  constructor(text, url) {
    this.text = text;
    this.url = url || null;
  }
  getText() { return this.text; }
  getLinkUrl() { return this.url; }
  getRuns() { return [this]; }
}

// ---------------------------------------------------------------------------
// Builders (conditional formatting and data validation)
// ---------------------------------------------------------------------------

class ConditionalFormatRuleBuilder {
  constructor() { this.rule = { formula: null, background: null, ranges: [] }; }
  whenFormulaSatisfied(formula) { this.rule.formula = formula; return this; }
  setBackground(color) { this.rule.background = color; return this; }
  setFontColor(color) { this.rule.fontColor = color; return this; }
  setBold(bold) { this.rule.bold = bold; return this; }
  setRanges(ranges) {
    if (!Array.isArray(ranges)) throw new Error('setRanges expects an array of ranges');
    this.rule.ranges = ranges;
    return this;
  }
  build() {
    const rule = Object.assign({}, this.rule);
    return {
      _rule: rule,
      getBooleanCondition() {
        if (!rule.formula) return null;
        return { getCriteriaType: () => 'CUSTOM_FORMULA', getCriteriaValues: () => [rule.formula] };
      },
      getRanges() { return rule.ranges; },
    };
  }
}

class DataValidationBuilder {
  constructor() { this.rule = {}; }
  requireNumberGreaterThanOrEqualTo(n) { this.rule.type = 'NUMBER_GTE'; this.rule.args = [n]; return this; }
  requireNumberBetween(a, b) { this.rule.type = 'NUMBER_BETWEEN'; this.rule.args = [a, b]; return this; }
  requireValueInList(list, show) {
    if (!Array.isArray(list)) throw new Error('requireValueInList expects an array');
    this.rule.type = 'VALUE_IN_LIST'; this.rule.args = [list, show]; return this;
  }
  requireCheckbox() { this.rule.type = 'CHECKBOX'; return this; }
  requireTextIsUrl() { this.rule.type = 'TEXT_IS_URL'; return this; }
  setAllowInvalid(b) { this.rule.allowInvalid = b; return this; }
  setHelpText(t) { this.rule.helpText = t; return this; }
  build() { return Object.assign({}, this.rule); }
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------

let sheetIds = 1;

class FakeSheet {
  constructor(spreadsheet, name, rows = 1000, cols = 26) {
    this.ss = spreadsheet;
    this.name = name;
    this.maxRows = rows;
    this.maxCols = cols;
    this.grid = [];
    this.fmt = [];
    this.frozenRows = 0;
    this.rules = [];
    this.meta = {};
    this.columnWidths = {};
    this.id = sheetIds++;
  }

  cell(r, c) { return this.grid[r - 1] && this.grid[r - 1][c - 1]; }

  format(r, c) { return (this.fmt[r - 1] && this.fmt[r - 1][c - 1]) || 'General'; }

  setCell(r, c, cell) {
    if (!this.grid[r - 1]) this.grid[r - 1] = [];
    if (cell) this.grid[r - 1][c - 1] = cell;
    else delete this.grid[r - 1][c - 1];
  }

  getName() { return this.name; }
  setName(name) {
    if (this.ss.sheets.some((s) => s !== this && s.name === name)) throw new Error('A sheet with the name "' + name + '" already exists.');
    this.name = name;
    return this;
  }
  getParent() { return this.ss; }
  getSheetId() { return this.id; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }

  getLastRow() {
    for (let r = this.grid.length; r >= 1; r--) {
      const row = this.grid[r - 1];
      if (row && row.some((cell) => !isEmptyCell(cell))) return r;
    }
    return 0;
  }

  getLastColumn() {
    let last = 0;
    this.grid.forEach((row) => {
      if (!row) return;
      row.forEach((cell, i) => { if (!isEmptyCell(cell)) last = Math.max(last, i + 1); });
    });
    return last;
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') return this.getRangeA1(a);
    const numRows = c === undefined ? 1 : c;
    const numCols = d === undefined ? 1 : d;
    [a, b, numRows, numCols].forEach((n) => {
      if (!Number.isInteger(n)) throw new Error('Range coordinates must be whole numbers, got ' + n);
    });
    if (a < 1 || b < 1) throw new Error('Range starts outside the sheet: row ' + a + ', column ' + b);
    if (numRows < 1 || numCols < 1) throw new Error('The number of rows and columns in a range must be at least 1.');
    if (a + numRows - 1 > this.maxRows || b + numCols - 1 > this.maxCols) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    return new FakeRange(this, a, b, numRows, numCols);
  }

  getRangeA1(a1) {
    const m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(a1);
    if (!m) throw new Error('Bad A1 notation: ' + a1);
    const col = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const r1 = m[2] ? +m[2] : 1;
    const c1 = col(m[1]);
    const c2 = m[3] ? col(m[3]) : c1;
    const r2 = m[3] ? (m[4] ? +m[4] : this.maxRows) : r1;
    return this.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }

  getDataRange() {
    return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  appendRow(values) {
    const row = this.getLastRow() + 1;
    if (row > this.maxRows) this.maxRows = row;
    if (values.length > this.maxCols) this.maxCols = values.length;
    this.getRange(row, 1, 1, values.length).setValues([values]);
    return this;
  }

  insertRowsAfter(after, n) {
    this.grid.splice(after, 0, ...Array.from({ length: n }, () => []));
    this.fmt.splice(after, 0, ...Array.from({ length: n }, () => []));
    this.maxRows += n;
    return this;
  }
  insertRowAfter(after) { return this.insertRowsAfter(after, 1); }

  insertColumnsAfter(after, n) {
    this.grid.forEach((row) => { if (row) row.splice(after, 0, ...new Array(n)); });
    this.fmt.forEach((row) => { if (row) row.splice(after, 0, ...new Array(n)); });
    this.maxCols += n;
    return this;
  }

  deleteRow(row) {
    if (row < 1 || row > this.maxRows) throw new Error('Row ' + row + ' is out of bounds.');
    if (this.maxRows - this.frozenRows <= 1) throw new Error('Sorry, it is not possible to delete all non-frozen rows.');
    this.grid.splice(row - 1, 1);
    this.fmt.splice(row - 1, 1);
    this.maxRows -= 1;
    return this;
  }

  setFrozenRows(n) { this.frozenRows = n; return this; }
  getFrozenRows() { return this.frozenRows; }
  setColumnWidth(col, width) { this.columnWidths[col] = width; return this; }
  setTabColor(color) { this.meta.tabColor = color; return this; }
  getConditionalFormatRules() { return this.rules.slice(); }
  setConditionalFormatRules(rules) {
    if (!Array.isArray(rules)) throw new Error('setConditionalFormatRules expects an array');
    this.rules = rules.slice();
    return this;
  }

  /** Test helper: the sheet as plain values, trimmed to the used area. */
  dump() {
    return this.getLastRow() ? this.getDataRange().getValues() : [];
  }

  /** Test helper: every formula in the sheet. */
  formulas() {
    const out = [];
    this.grid.forEach((row, r) => row && row.forEach((cell, c) => { if (cell && cell.f) out.push({ row: r + 1, col: c + 1, formula: cell.f }); }));
    return out;
  }
}

const FORMAT_METHODS = [
  'setFontWeight', 'setFontColor', 'setBackground', 'setWrap', 'setWrapStrategy',
  'setHorizontalAlignment', 'setVerticalAlignment', 'setFontSize', 'setFontFamily', 'setNote',
];

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }

  eachCell(fn) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) fn(this.row + r, this.col + c, r, c);
    }
  }

  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getLastRow() { return this.row + this.numRows - 1; }
  getLastColumn() { return this.col + this.numCols - 1; }

  getValues() {
    const out = [];
    this.eachCell((r, c, i, j) => {
      if (!out[i]) out[i] = [];
      out[i][j] = cellValue(this.sheet.cell(r, c));
    });
    return out;
  }
  getValue() { return cellValue(this.sheet.cell(this.row, this.col)); }
  getDisplayValues() { return this.getValues().map((row) => row.map((v) => (v instanceof Date ? v.toISOString() : String(v)))); }
  getFormulas() {
    const out = [];
    this.eachCell((r, c, i, j) => {
      if (!out[i]) out[i] = [];
      const cell = this.sheet.cell(r, c);
      out[i][j] = cell && cell.f ? cell.f : '';
    });
    return out;
  }
  getRichTextValues() {
    const out = [];
    this.eachCell((r, c, i, j) => {
      if (!out[i]) out[i] = [];
      const cell = this.sheet.cell(r, c);
      const value = cellValue(cell);
      out[i][j] = typeof value === 'string' ? new FakeRichText(value, cell && cell.link) : null;
    });
    return out;
  }

  setValues(values) {
    if (!Array.isArray(values) || values.length !== this.numRows) {
      throw new Error('The number of rows in the data does not match the number of rows in the range. The data has ' +
        (Array.isArray(values) ? values.length : 'no') + ' but the range has ' + this.numRows + '.');
    }
    values.forEach((row) => {
      if (!Array.isArray(row) || row.length !== this.numCols) {
        throw new Error('The number of columns in the data does not match the number of columns in the range. The data has ' +
          (Array.isArray(row) ? row.length : 'no') + ' but the range has ' + this.numCols + '.');
      }
    });
    this.eachCell((r, c, i, j) => this.sheet.setCell(r, c, userEntered(values[i][j], this.sheet.format(r, c))));
    return this;
  }
  setValue(value) {
    this.eachCell((r, c) => this.sheet.setCell(r, c, userEntered(value, this.sheet.format(r, c))));
    return this;
  }
  setNumberFormat(format) {
    if (typeof format !== 'string') throw new Error('setNumberFormat needs a string');
    this.eachCell((r, c) => {
      if (!this.sheet.fmt[r - 1]) this.sheet.fmt[r - 1] = [];
      this.sheet.fmt[r - 1][c - 1] = format;
    });
    return this;
  }
  getNumberFormats() {
    const out = [];
    this.eachCell((r, c, i, j) => {
      if (!out[i]) out[i] = [];
      out[i][j] = this.sheet.format(r, c);
    });
    return out;
  }
  getNumberFormat() { return this.sheet.format(this.row, this.col); }
  clearContent() {
    this.eachCell((r, c) => this.sheet.setCell(r, c, null));
    return this;
  }
  /** Test helper: make the cell a rich-text link, like Insert > Link in Sheets. */
  setLinkForTest(text, url) {
    this.sheet.setCell(this.row, this.col, { v: text, link: url });
    return this;
  }
  insertCheckboxes() {
    this.eachCell((r, c) => {
      this.sheet.setCell(r, c, { v: false });
      this.sheet.meta['checkbox:' + r + ':' + c] = true;
    });
    return this;
  }
  setDataValidation(rule) {
    if (!rule || typeof rule !== 'object') throw new Error('setDataValidation expects a built rule');
    this.sheet.meta['validation:' + this.row + ':' + this.col] = rule;
    return this;
  }
  getA1Notation() { return 'R' + this.row + 'C' + this.col + ':R' + this.getLastRow() + 'C' + this.getLastColumn(); }
}
FORMAT_METHODS.forEach((name) => {
  FakeRange.prototype[name] = function (value) {
    if (value === undefined) throw new Error(name + ' needs a value');
    this.sheet.meta[name + ':' + this.row + ':' + this.col] = value;
    return this;
  };
});

class FakeSpreadsheet {
  constructor(opts = {}) {
    this.id = opts.id || 'spreadsheet-123';
    this.name = opts.name || 'Parts Inventory';
    this.tz = opts.timeZone || 'America/New_York';
    this.sheets = [new FakeSheet(this, 'Sheet1')];
    this.toasts = [];
    this.active = this.sheets[0];
  }
  getId() { return this.id; }
  getName() { return this.name; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getSpreadsheetTimeZone() { return this.tz; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
  insertSheet(name, index) {
    if (this.getSheetByName(name)) throw new Error('A sheet with the name "' + name + '" already exists.');
    const sheet = new FakeSheet(this, name);
    if (index === undefined) this.sheets.push(sheet);
    else this.sheets.splice(index, 0, sheet);
    return sheet;
  }
  setActiveSheet(sheet) { this.active = sheet; return sheet; }
  getActiveSheet() { return this.active; }
  toast(message, title, timeout) { this.toasts.push({ message, title, timeout }); }
}

// ---------------------------------------------------------------------------
// The environment
// ---------------------------------------------------------------------------

function formatDate(date, timeZone, pattern) {
  if (!(date instanceof Date) || isNaN(date)) throw new Error('formatDate needs a valid Date');
  const parts = {};
  new Intl.DateTimeFormat('en-US', {
    timeZone, year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric',
    second: 'numeric', hourCycle: 'h23', weekday: 'long',
  }).formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const pad = (n, len = 2) => String(n).padStart(len, '0');
  const H = Number(parts.hour) % 24;
  const M = Number(parts.month);
  const tokens = {
    yyyy: parts.year, yy: parts.year.slice(-2), MMMM: months[M - 1], MMM: months[M - 1].slice(0, 3),
    MM: pad(M), M: String(M), dd: pad(parts.day), d: String(Number(parts.day)), HH: pad(H), H: String(H),
    hh: pad(H % 12 || 12), h: String(H % 12 || 12), mm: pad(parts.minute), ss: pad(parts.second),
    a: H < 12 ? 'AM' : 'PM', EEEE: parts.weekday, EEE: parts.weekday.slice(0, 3),
  };
  return pattern.replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dd|d|HH|H|hh|h|mm|ss|a|EEEE|EEE/g, (tok, literal) =>
    (literal !== undefined ? literal : tokens[tok]));
}

/**
 * The fake Google services, without Code.js loaded. Works in Node and in a
 * browser (see dev/build-demo.js).
 * opts: { owner, active, timeZone, ui, now, serviceUrl, spreadsheetName, quota, readHtml }
 */
function createEnvironment(opts = {}) {
  const env = {
    owner: opts.owner === undefined ? 'owner@example.com' : opts.owner,
    active: opts.active === undefined ? (opts.owner === undefined ? 'owner@example.com' : opts.owner) : opts.active,
    uiAvailable: !!opts.ui,
    now: opts.now ? new Date(opts.now).getTime() : null,
    serviceUrl: opts.serviceUrl || null,
    quota: opts.quota === undefined ? 100 : opts.quota,
    mailFails: null,
    sent: [],
    alerts: [],
    alertAnswers: [],
    dialogs: [],
    sidebars: [],
    menus: [],
    triggers: [],
    props: {},
    cache: {},
    logs: [],
    lockHeld: false,
  };
  const ss = new FakeSpreadsheet({ name: opts.spreadsheetName, timeZone: opts.timeZone });
  env.ss = ss;

  const RealDate = Date;
  class ClockDate extends RealDate {
    constructor(...args) {
      if (args.length === 0 && env.now !== null) super(env.now);
      else super(...args);
    }
    static now() { return env.now !== null ? env.now : RealDate.now(); }
  }

  const Button = { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL' };
  const ui = {
    Button,
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    createMenu(name) {
      const menu = { name, items: [] };
      env.menus.push(menu);
      const api = {
        addItem(caption, fn) { menu.items.push({ caption, fn }); return api; },
        addSeparator() { menu.items.push('---'); return api; },
        addToUi() { menu.added = true; },
      };
      return api;
    },
    alert(title, prompt, buttons) {
      env.alerts.push({ title, prompt, buttons });
      return env.alertAnswers.length ? env.alertAnswers.shift() : Button.OK;
    },
    showSidebar(output) { env.sidebars.push(output); },
    showModalDialog(output, title) { env.dialogs.push({ title, html: output.getContent() }); },
    showModelessDialog(output, title) { env.dialogs.push({ title, html: output.getContent() }); },
  };

  class HtmlOutput {
    constructor(content) { this.content = content; this.title = ''; this.meta = []; }
    getContent() { return this.content; }
    setTitle(title) { this.title = title; return this; }
    getTitle() { return this.title; }
    addMetaTag(name, content) {
      if (name !== 'viewport' && !/^(apple-mobile-web-app-capable|mobile-web-app-capable)$/.test(name)) {
        throw new Error('Meta tag ' + name + ' is not allowed');
      }
      this.meta.push({ name, content });
      return this;
    }
    setWidth(w) { this.width = w; return this; }
    setHeight(h) { this.height = h; return this; }
    setXFrameOptionsMode(mode) { this.xFrame = mode; return this; }
    setFaviconUrl(url) { this.favicon = url; return this; }
  }

  const readHtml = opts.readHtml || ((name) => fs.readFileSync(path.join(SRC_DIR, name + '.html'), 'utf8'));

  const HtmlService = {
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
    SandboxMode: { IFRAME: 'IFRAME' },
    createHtmlOutput(html) { return new HtmlOutput(String(html)); },
    createHtmlOutputFromFile(name) { return new HtmlOutput(readHtml(name)); },
    createTemplateFromFile(name) {
      const source = readHtml(name);
      const template = {
        evaluate() {
          // Only <?!= name ?> is used by the app. Anything else starting with
          // "<?" would be read by Apps Script as template code, so fail loudly.
          const out = source.replace(/<\?([\s\S]*?)\?>/g, (whole, inner) => {
            const m = /^!=\s*([A-Za-z_$][\w$]*)\s*$/.exec(inner);
            if (!m) throw new Error('Unexpected template tag in ' + name + '.html: ' + whole.slice(0, 40));
            if (!(m[1] in template)) throw new Error('Template variable ' + m[1] + ' was not set');
            return String(template[m[1]]);
          });
          return new HtmlOutput(out);
        },
      };
      return template;
    },
  };

  const makeTriggerBuilder = (fn) => {
    const trigger = { fn, id: 'trigger-' + (env.triggers.length + 1) + '-' + Math.random().toString(36).slice(2, 7) };
    const create = () => {
      trigger.owner = env.owner;
      env.triggers.push(trigger);
      return {
        getHandlerFunction: () => trigger.fn,
        getUniqueId: () => trigger.id,
        getEventType: () => trigger.type,
        _trigger: trigger,
      };
    };
    return {
      forSpreadsheet(spreadsheet) {
        if (!spreadsheet || typeof spreadsheet.getId !== 'function') throw new Error('forSpreadsheet needs a Spreadsheet');
        return {
          onEdit() { trigger.type = 'ON_EDIT'; return { create }; },
          onChange() { trigger.type = 'ON_CHANGE'; return { create }; },
          onOpen() { trigger.type = 'ON_OPEN'; return { create }; },
        };
      },
      timeBased() {
        return {
          everyHours(n) {
            if (![1, 2, 4, 6, 8, 12].includes(n)) throw new Error('everyHours only accepts 1, 2, 4, 6, 8 or 12');
            trigger.type = 'CLOCK'; trigger.every = n + 'h'; return { create };
          },
          everyMinutes(n) { trigger.type = 'CLOCK'; trigger.every = n + 'm'; return { create }; },
        };
      },
    };
  };

  const globals = {
    Date: ClockDate,
    console: {
      log: (...a) => env.logs.push(['log', a.join(' ')]),
      info: (...a) => env.logs.push(['info', a.join(' ')]),
      warn: (...a) => env.logs.push(['warn', a.join(' ')]),
      error: (...a) => env.logs.push(['error', a.join(' ')]),
    },
    Logger: { log: (...a) => env.logs.push(['log', a.join(' ')]) },
    SpreadsheetApp: {
      WrapStrategy: { CLIP: 'CLIP', WRAP: 'WRAP', OVERFLOW: 'OVERFLOW' },
      BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
      getActiveSpreadsheet: () => ss,
      getActive: () => ss,
      openById: (id) => {
        if (id !== ss.id) throw new Error('Unexpected spreadsheet id ' + id);
        return ss;
      },
      getUi: () => {
        if (!env.uiAvailable) throw new Error('Exception: Cannot call SpreadsheetApp.getUi() from this context.');
        return ui;
      },
      flush: () => {},
      newConditionalFormatRule: () => new ConditionalFormatRuleBuilder(),
      newDataValidation: () => new DataValidationBuilder(),
    },
    HtmlService,
    MailApp: {
      sendEmail(options) {
        if (typeof options !== 'object' || !options) throw new Error('This app always calls sendEmail with an options object');
        ['to', 'subject', 'body'].forEach((k) => {
          if (typeof options[k] !== 'string' || !options[k]) throw new Error('sendEmail: missing ' + k);
        });
        if (env.mailFails) throw new Error(env.mailFails);
        // Gmail's limit for a message body sent from Apps Script (free accounts).
        if ((options.htmlBody || '').length + options.body.length > 200 * 1024) throw new Error('Argument too large: body');
        const recipients = options.to.split(',').length;
        if (recipients > env.quota - env.sent.length) throw new Error('Service invoked too many times for one day: email.');
        env.sent.push(Object.assign({}, options));
      },
      getRemainingDailyQuota: () => Math.max(env.quota - env.sent.length, 0),
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in env.props ? env.props[k] : null),
        setProperty(k, v) {
          if (typeof v !== 'string') throw new Error('Property values must be strings');
          env.props[k] = v;
          return this;
        },
        deleteProperty(k) { delete env.props[k]; return this; },
        getProperties: () => Object.assign({}, env.props),
      }),
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (k in env.cache ? env.cache[k] : null),
        put: (k, v, ttl) => {
          if (typeof v !== 'string') throw new Error('Cache values must be strings');
          if (ttl !== undefined && ttl > 21600) throw new Error('Cache TTL is 6 hours at most');
          env.cache[k] = v;
        },
        remove: (k) => { delete env.cache[k]; },
      }),
    },
    LockService: {
      getScriptLock: () => ({
        tryLock() {
          if (env.lockHeld) throw new Error('Lock requested while it is already held (nested withLock_ call).');
          env.lockHeld = true;
          return true;
        },
        waitLock() { this.tryLock(); },
        releaseLock() { env.lockHeld = false; },
        hasLock: () => env.lockHeld,
      }),
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => env.active || '' }),
      getEffectiveUser: () => ({ getEmail: () => env.owner || '' }),
      getScriptTimeZone: () => 'America/New_York',
    },
    ScriptApp: {
      getProjectTriggers: () => env.triggers.filter((t) => t.owner === env.owner).map((t) => ({
        getHandlerFunction: () => t.fn,
        getUniqueId: () => t.id,
        getEventType: () => t.type,
        _trigger: t,
      })),
      newTrigger: (fn) => {
        if (typeof fn !== 'string') throw new Error('newTrigger expects a function name');
        return makeTriggerBuilder(fn);
      },
      deleteTrigger: (t) => {
        const i = env.triggers.indexOf(t._trigger);
        if (i === -1) throw new Error('No such trigger');
        env.triggers.splice(i, 1);
      },
      getService: () => ({ getUrl: () => env.serviceUrl }),
    },
    Utilities: {
      formatDate,
      getUuid: () => globalThis.crypto.randomUUID(),
      sleep: () => {},
    },
  };

  return { env, ss, globals, RealDate };
}

/** Helpers around a loaded Code.js. lookup(name) returns one of its top-level functions. */
function makeApp(environment, lookup, extras) {
  const { env, ss, RealDate } = environment;
  const app = Object.assign({
    env,
    ss,
    /** The event object a trigger of this script would pass (with its real triggerUid). */
    triggerEvent(fn) {
      const t = env.triggers.find((x) => x.fn === fn);
      if (!t) throw new Error('No trigger for ' + fn);
      return { triggerUid: t.id, authMode: 'FULL' };
    },
    /** Calls a function from Code.js the way google.script.run would: arguments and result go through JSON. */
    run(name, ...args) {
      const fn = lookup(name);
      if (typeof fn !== 'function') throw new Error('No function ' + name + ' in Code.js');
      if (/_$/.test(name)) throw new Error(name + ' is private and cannot be called by google.script.run');
      const result = fn(...JSON.parse(JSON.stringify(args)));
      assertJsonSafe(result, name);
      return result === undefined ? undefined : JSON.parse(JSON.stringify(result));
    },
    sheet(name) { return ss.getSheetByName(name); },
    setTime(iso) { env.now = new RealDate(iso).getTime(); },
    as(active) { env.active = active; return app; },
  }, extras);
  return app;
}

/**
 * Creates a fresh fake Google environment and loads src/Code.js into it.
 * opts: see createEnvironment.
 */
function createApp(opts = {}) {
  const environment = createEnvironment(opts);
  const injected = Object.assign({}, environment.globals);
  const context = vm.createContext(environment.globals);
  const code = fs.readFileSync(path.join(SRC_DIR, 'Code.js'), 'utf8');
  vm.runInContext(code, context, { filename: 'Code.js' });
  return makeApp(environment, (name) => context[name], {
    context,
    /** Runs any expression inside Code.js's scope (for constants and private helpers). */
    eval(expression) { return vm.runInContext(expression, context); },
    /** Like eval, but returns plain JSON data (objects made inside the sandbox fail strict equality). */
    value(expression) { return JSON.parse(JSON.stringify(vm.runInContext(expression, context))); },
    /** Names of the functions Code.js defines at the top level (google.script.run can call the public ones). */
    functionNames() {
      return Object.keys(context).filter((k) => typeof context[k] === 'function' && !(k in injected));
    },
  });
}

/** google.script.run can't send Dates or functions back to the page; fail tests that try. */
function assertJsonSafe(value, where, trail = 'result') {
  if (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]') {
    throw new Error(where + ' returned a Date at ' + trail + ' (google.script.run would deliver null)');
  }
  if (typeof value === 'function') throw new Error(where + ' returned a function at ' + trail);
  if (value && typeof value === 'object') {
    Object.keys(value).forEach((k) => assertJsonSafe(value[k], where, trail + '.' + k));
  }
}

module.exports = { createApp, createEnvironment, makeApp, FakeSpreadsheet, FakeSheet, formatDate, userEntered };
