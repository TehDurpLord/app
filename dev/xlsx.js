'use strict';
/**
 * A small .xlsx writer with no dependencies, just enough for the ready-made
 * spreadsheet (dev/build-sheet.js), plus a zip reader for the tests.
 *
 * The output is kept small on purpose: it gets uploaded to Google Drive inside
 * an API call, where Drive turns it into a Google Sheet.
 */
const zlib = require('zlib');

// ---------------------------------------------------------------------------
// Zip
// ---------------------------------------------------------------------------

function crc32(buf) {
  if (typeof zlib.crc32 === 'function') return zlib.crc32(buf) >>> 0;
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    let c = (crc ^ buf[i]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** files: [{ name, data }] -> zip archive (deflated, fixed 1980-01-01 timestamps). */
function zip(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  files.forEach(({ name, data }) => {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
    const packed = zlib.deflateRawSync(raw, { level: 9 });
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0x0800, 6); // UTF-8 names
    local.writeUInt16LE(8, 8); // deflate
    local.writeUInt16LE(0, 10); // time
    local.writeUInt16LE(33, 12); // date: 1980-01-01
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    locals.push(local, nameBuf, packed);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(33, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + packed.length;
  });
  const directory = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat(locals.concat([directory, end]));
}

/** zip archive -> { name: Buffer }. Checks every file's CRC. */
function unzip(buf) {
  let end = buf.length - 22;
  while (end >= 0 && buf.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) throw new Error('Not a zip file');
  const count = buf.readUInt16LE(end + 10);
  let p = buf.readUInt32LE(end + 16);
  const out = {};
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Bad zip directory');
    const method = buf.readUInt16LE(p + 10);
    const crc = buf.readUInt32LE(p + 16);
    const size = buf.readUInt32LE(p + 20);
    const nameLength = buf.readUInt16LE(p + 28);
    const extraLength = buf.readUInt16LE(p + 30);
    const commentLength = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLength);
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const packed = buf.subarray(start, start + size);
    const data = method === 8 ? zlib.inflateRawSync(packed) : Buffer.from(packed);
    if (crc32(data) !== crc) throw new Error('CRC mismatch in ' + name);
    out[name] = data;
    p += 46 + nameLength + extraLength + commentLength;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Spreadsheet
// ---------------------------------------------------------------------------

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Cell styles (index into cellXfs). */
const STYLE = { plain: 0, header: 1, managedHeader: 2, label: 3, help: 4, top: 5 };

function esc(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function argb(hex) {
  return 'FF' + hex.replace('#', '').toUpperCase();
}

function columnLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Sheets column width in pixels -> Excel's character-based width. */
function excelWidth(px) {
  return Math.round(((px - 5) / 7) * 100) / 100;
}

function stylesXml(colors, dxfFills) {
  const font = (extra) => '<font>' + (extra || '') + '<sz val="10"/><name val="Arial"/></font>';
  const solid = (hex) => '<fill><patternFill patternType="solid"><fgColor rgb="' + argb(hex) + '"/><bgColor indexed="64"/></patternFill></fill>';
  const xf = (attrs, inner) => '<xf numFmtId="0" borderId="0" xfId="0" ' + attrs + (inner ? '>' + inner + '</xf>' : '/>');
  return XML_HEAD + '<styleSheet xmlns="' + MAIN + '">' +
    '<fonts count="4">' + font() + font('<b/><color rgb="FFFFFFFF"/>') + font('<b/>') +
    font('<color rgb="' + argb(colors.helpText) + '"/>') + '</fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    solid(colors.header) + solid(colors.managedHeader) + '</fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    xf('fontId="0" fillId="0"') +
    xf('fontId="1" fillId="2" applyFont="1" applyFill="1"') +
    xf('fontId="1" fillId="3" applyFont="1" applyFill="1"') +
    xf('fontId="2" fillId="0" applyFont="1" applyAlignment="1"', '<alignment vertical="top"/>') +
    xf('fontId="3" fillId="0" applyFont="1" applyAlignment="1"', '<alignment vertical="top" wrapText="1"/>') +
    xf('fontId="0" fillId="0" applyAlignment="1"', '<alignment vertical="top"/>') +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="' + dxfFills.length + '">' +
    dxfFills.map((hex) => '<dxf><fill><patternFill patternType="solid"><bgColor rgb="' + argb(hex) + '"/></patternFill></fill></dxf>').join('') +
    '</dxfs></styleSheet>';
}

function cellXml(ref, cell) {
  if (cell === null || cell === undefined) return '';
  const c = typeof cell === 'object' ? cell : { v: cell };
  const s = c.style ? ' s="' + STYLE[c.style] + '"' : '';
  if (c.f !== undefined) return '<c r="' + ref + '"' + s + '><f>' + esc(c.f) + '</f></c>';
  if (c.v === '' || c.v === null || c.v === undefined) return s ? '<c r="' + ref + '"' + s + '/>' : '';
  if (typeof c.v === 'number') return '<c r="' + ref + '"' + s + '><v>' + c.v + '</v></c>';
  if (typeof c.v === 'boolean') return '<c r="' + ref + '"' + s + ' t="b"><v>' + (c.v ? 1 : 0) + '</v></c>';
  return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(c.v) + '</t></is></c>';
}

function sheetXml(sheet, index, dxfIndex) {
  const parts = [XML_HEAD + '<worksheet xmlns="' + MAIN + '">'];
  if (sheet.tabColor) parts.push('<sheetPr><tabColor rgb="' + argb(sheet.tabColor) + '"/></sheetPr>');
  parts.push('<sheetViews><sheetView workbookViewId="0"' + (index === 0 ? ' tabSelected="1"' : '') + '>');
  if (sheet.frozenRows) {
    parts.push('<pane ySplit="' + sheet.frozenRows + '" topLeftCell="A' + (sheet.frozenRows + 1) +
      '" activePane="bottomLeft" state="frozen"/>');
  }
  parts.push('</sheetView></sheetViews>');
  if (sheet.widths && sheet.widths.length) {
    parts.push('<cols>' + sheet.widths.map((px, i) =>
      '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + excelWidth(px) + '" customWidth="1"/>').join('') + '</cols>');
  }
  parts.push('<sheetData>' + sheet.rows.map((row, r) =>
    '<row r="' + (r + 1) + '">' + row.map((cell, c) => cellXml(columnLetter(c + 1) + (r + 1), cell)).join('') + '</row>').join('') +
    '</sheetData>');
  (sheet.conditionalFormats || []).forEach((cf) => {
    parts.push('<conditionalFormatting sqref="' + cf.sqref + '">' + cf.rules.map((rule, i) =>
      '<cfRule type="expression" dxfId="' + dxfIndex(rule.fill) + '" priority="' + (i + 1) + '"><formula>' +
      esc(rule.formula) + '</formula></cfRule>').join('') + '</conditionalFormatting>');
  });
  const validations = sheet.validations || [];
  if (validations.length) {
    parts.push('<dataValidations count="' + validations.length + '">' + validations.map((v) => {
      const attrs = v.type === 'list'
        ? 'type="list"'
        : 'type="' + v.type + '" operator="' + v.operator + '"';
      const formulas = v.type === 'list'
        ? '<formula1>' + esc('"' + v.list.join(',') + '"') + '</formula1>'
        : '<formula1>' + v.formula1 + '</formula1>' + (v.formula2 !== undefined ? '<formula2>' + v.formula2 + '</formula2>' : '');
      return '<dataValidation ' + attrs + ' allowBlank="1" showErrorMessage="1"' +
        (v.error ? ' error="' + esc(v.error) + '"' : '') + ' sqref="' + v.sqref + '">' + formulas + '</dataValidation>';
    }).join('') + '</dataValidations>');
  }
  parts.push('</worksheet>');
  return parts.join('');
}

/** spec: { colors, sheets: [{ name, tabColor, frozenRows, widths, rows, validations, conditionalFormats }] } -> files */
function xlsxFiles(spec) {
  const dxfFills = [];
  const dxfIndex = (hex) => {
    if (dxfFills.indexOf(hex) === -1) dxfFills.push(hex);
    return dxfFills.indexOf(hex);
  };
  const sheets = spec.sheets.map((sheet, i) => ({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(sheet, i, dxfIndex) }));
  const n = spec.sheets.length;
  return [
    {
      name: '[Content_Types].xml',
      data: XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        spec.sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) +
          '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
        '</Types>',
    },
    {
      name: '_rels/.rels',
      data: XML_HEAD + '<Relationships xmlns="' + PKG + '"><Relationship Id="rId1" Type="' + REL +
        '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    },
    {
      name: 'xl/workbook.xml',
      data: XML_HEAD + '<workbook xmlns="' + MAIN + '" xmlns:r="' + REL + '"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
        spec.sheets.map((s, i) => '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') +
        '</sheets></workbook>',
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: XML_HEAD + '<Relationships xmlns="' + PKG + '">' +
        spec.sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + REL + '/worksheet" Target="worksheets/sheet' +
          (i + 1) + '.xml"/>').join('') +
        '<Relationship Id="rId' + (n + 1) + '" Type="' + REL + '/styles" Target="styles.xml"/></Relationships>',
    },
    // Styles last: the conditional-format colours are collected while the sheets are written.
  ].concat(sheets, [{ name: 'xl/styles.xml', data: stylesXml(spec.colors, dxfFills) }]);
}

function buildXlsx(spec) {
  return zip(xlsxFiles(spec));
}

module.exports = { buildXlsx, xlsxFiles, zip, unzip, columnLetter };
