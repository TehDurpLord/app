#!/usr/bin/env node
'use strict';
/**
 * Builds template/parts-inventory.xlsx: the Inventory, Reorder and Settings
 * tabs, ready to upload to Google Drive, where it opens as a Google Sheet.
 * The columns, settings, colours and formulas come from src/Code.js, so the
 * file always matches what the script's setup would build.
 *
 *   node dev/build-sheet.js            write the file
 *   node dev/build-sheet.js --check    fail if the file is out of date
 *   node dev/build-sheet.js --base64   print the file base64-encoded
 *
 * Google's .xlsx import can't make tick boxes, so Ordered is a "Yes" dropdown
 * until setup runs and turns it into tick boxes. It also rejects ranges like
 * A2:A, which is why the formulas use whole columns (A:A).
 */
const fs = require('fs');
const path = require('path');
const { createApp } = require('../test/gas-fakes');
const { buildXlsx, xlsxFiles, unzip, columnLetter } = require('./xlsx');

const OUT = path.join(__dirname, '..', 'template', 'parts-inventory.xlsx');
const ROWS = 1000;

/** The workbook, described from the definitions in src/Code.js. */
function sheetSpec() {
  const app = createApp();
  const FIELDS = app.value('FIELDS');
  const SETTINGS = app.value('SETTINGS');
  const COLORS = app.value('COLORS');
  const NAMES = app.value('SHEET_NAMES');
  const STATUS_TEXT = app.value('STATUS_TEXT');
  const field = {};
  FIELDS.forEach((f) => { field[f.key] = f; });

  const layout = FIELDS.filter((f) => f.layout);
  const cols = {};
  layout.forEach((f, i) => { cols[f.key] = i + 1; });
  const column = (key) => columnLetter(cols[key]) + '2:' + columnLetter(cols[key]) + ROWS;

  const inventory = {
    name: NAMES.inventory,
    tabColor: COLORS.accent,
    frozenRows: 1,
    widths: layout.map((f) => f.width),
    rows: [layout.map((f) => ({ v: f.header, style: f.managed ? 'managedHeader' : 'header' }))],
    validations: [
      {
        type: 'decimal', operator: 'greaterThanOrEqual', formula1: '0', error: 'Enter a number (0 or more).',
        sqref: layout.filter((f) => f.type === 'number').map((f) => column(f.key)).join(' '),
      },
      // Tick boxes don't survive the import; setup turns this into one.
      { type: 'list', list: ['Yes'], sqref: column('ordered') },
    ],
    conditionalFormats: [{
      sqref: 'A2:' + columnLetter(layout.length) + ROWS,
      rules: app.context.stockRuleFormulas_(cols, 2).map((formula, i) => ({
        formula,
        fill: i === 0 ? COLORS.outRow : COLORS.lowRow,
      })),
    }],
  };

  const reorderKeys = app.value('REORDER_KEYS').filter((key) => cols[key]);
  const reorder = {
    name: NAMES.reorder,
    tabColor: COLORS.reorderTab,
    frozenRows: 2,
    widths: reorderKeys.map((key) => field[key].width),
    rows: [
      [{ v: app.value('REORDER_TITLE'), style: 'label' }],
      reorderKeys.map((key) => ({ v: field[key].header, style: 'header' })),
      [{ f: app.context.reorderFormula_(cols, reorderKeys).slice(1) }],
    ],
  };

  const settingRows = SETTINGS.map((def) => {
    let value = '';
    if (def.type === 'status') value = 'Off';
    else if (def.type === 'onoff') value = def.default ? 'On' : 'Off';
    else if (def.type !== 'emails') value = def.default;
    return [
      { v: def.label, style: 'label' },
      { v: value, style: def.type === 'status' ? 'label' : 'top' },
      { v: def.type === 'status' ? STATUS_TEXT.off : def.help, style: 'help' },
    ];
  });
  const settingValidations = [];
  SETTINGS.forEach((def, i) => {
    const ref = 'B' + (i + 2);
    if (def.type === 'onoff') settingValidations.push({ type: 'list', list: ['On', 'Off'], sqref: ref });
    if (def.type === 'choice') settingValidations.push({ type: 'list', list: def.choices, sqref: ref });
    if (def.type === 'hour') {
      settingValidations.push({ type: 'whole', operator: 'between', formula1: '0', formula2: '23', error: 'A whole number from 0 to 23.', sqref: ref });
    }
  });
  const settings = {
    name: NAMES.settings,
    tabColor: COLORS.settingsTab,
    frozenRows: 1,
    widths: [270, 280, 560],
    rows: [['Setting', 'Value', 'What it does'].map((v) => ({ v, style: 'header' }))].concat(settingRows),
    validations: settingValidations,
  };

  return { colors: COLORS, sheets: [inventory, reorder, settings] };
}

function buildSheet() {
  return buildXlsx(sheetSpec());
}

/** True when the file on disk has the same contents as a fresh build (byte-level zip details aside). */
function isUpToDate() {
  if (!fs.existsSync(OUT)) return false;
  const current = unzip(fs.readFileSync(OUT));
  const fresh = xlsxFiles(sheetSpec());
  return fresh.length === Object.keys(current).length &&
    fresh.every((file) => current[file.name] && current[file.name].toString('utf8') === file.data);
}

if (require.main === module) {
  if (process.argv.includes('--check')) {
    if (!isUpToDate()) {
      console.error('template/parts-inventory.xlsx is out of date. Run: npm run build');
      process.exit(1);
    }
    console.log('template/parts-inventory.xlsx is up to date.');
  } else if (process.argv.includes('--base64')) {
    process.stdout.write(buildSheet().toString('base64'));
  } else {
    const file = buildSheet();
    fs.mkdirSync(path.dirname(OUT), { recursive: true });
    fs.writeFileSync(OUT, file);
    console.log('Wrote template/parts-inventory.xlsx (' + file.length + ' bytes)');
  }
}

module.exports = { sheetSpec, buildSheet, isUpToDate, OUT };
