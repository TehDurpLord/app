#!/usr/bin/env node
'use strict';
/**
 * Builds install/Code.gs: src/Code.js with src/Index.html bundled into it, so
 * installing the app means pasting one file into the Apps Script editor.
 *
 *   node dev/build-bundle.js           write install/Code.gs
 *   node dev/build-bundle.js --check   exit with an error if it's out of date
 *
 * Run it after changing anything in src/ (the tests check it's up to date).
 */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const OUT = path.join(root, 'install', 'Code.gs');

function buildBundle() {
  const code = fs.readFileSync(path.join(root, 'src', 'Code.js'), 'utf8');
  const html = fs.readFileSync(path.join(root, 'src', 'Index.html'), 'utf8');
  return '// Parts Inventory: the whole app in one file. Paste all of it into Code.gs in the Apps Script editor.\n' +
    '// Built from src/Code.js and src/Index.html by dev/build-bundle.js; edit those, not this.\n\n' +
    code.trimEnd() + '\n\n' +
    '// ---------------------------------------------------------------------------\n' +
    '// The web page (src/Index.html), bundled so this is the only file to paste.\n' +
    '// ---------------------------------------------------------------------------\n\n' +
    'var INDEX_HTML_ = ' + JSON.stringify(html) + ';\n';
}

if (require.main === module) {
  const bundle = buildBundle();
  if (process.argv.includes('--check')) {
    const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
    if (current !== bundle) {
      console.error('install/Code.gs is out of date. Run: npm run build');
      process.exit(1);
    }
    console.log('install/Code.gs is up to date.');
  } else {
    fs.mkdirSync(path.dirname(OUT), { recursive: true });
    fs.writeFileSync(OUT, bundle);
    console.log('Wrote install/Code.gs (' + Math.round(bundle.length / 1024) + ' KB)');
  }
}

module.exports = { buildBundle, OUT };
