#!/usr/bin/env node
'use strict';
/**
 * Builds a one-file demo of the app that runs in any browser, with no Google
 * account: the real src/Index.html and src/Code.js on the fake Google services
 * from test/gas-fakes.js, filled with the example parts from dev/sample-data.js.
 * Nothing is saved, and emails are shown in the page instead of being sent.
 *
 *   node dev/build-demo.js [output.html]     (default: dist/parts-inventory-demo.html)
 *
 * The output is page content without <html>/<head>/<body> (it's published as a
 * Claude artifact, which adds those). Use --standalone for a complete HTML file.
 */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const read = (file) => fs.readFileSync(path.join(root, file), 'utf8');
const args = process.argv.slice(2);
const standalone = args.includes('--standalone');
const out = path.resolve(args.find((a) => !a.startsWith('--')) || path.join(root, 'dist', 'parts-inventory-demo.html'));

/** Makes JS source safe inside an inline <script>. These strings only occur inside JS strings, where "<\/" and "<\!" read the same. */
function inline(source) {
  return source.replace(/<\/script/gi, '<\\/script').replace(/<!--/g, '<\\!--');
}

function replaceOnce(text, find, replacement) {
  const i = text.indexOf(find);
  if (i === -1 || text.indexOf(find, i + find.length) !== -1) throw new Error('Expected exactly one: ' + find.slice(0, 60));
  return text.slice(0, i) + replacement + text.slice(i + find.length);
}

const code = read('src/Code.js');
const fakes = read('test/gas-fakes.js');
const sample = read('dev/sample-data.js');
let page = read('src/Index.html');

const serverFunctions = ['setup', 'hourlyCheck'].concat((code.match(/^function (api\w+)\(/gm) || []).map((m) => m.slice(9, -1)));
const googleServices = ['SpreadsheetApp', 'HtmlService', 'MailApp', 'PropertiesService', 'CacheService', 'LockService',
  'Session', 'ScriptApp', 'Utilities', 'Logger', 'console', 'Date'];

const demoCss = `<style>
  .demo-bar {
    display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 8px 12px;
    padding: 8px 16px; background: var(--accent-soft); color: var(--text); font-size: 13.5px;
    border-bottom: 1px solid var(--border);
  }
  .demo-bar .demo-text { min-width: 0; flex: 1 1 260px; }
  .demo-bar .demo-actions { display: flex; gap: 6px; flex-wrap: wrap; }
  .demo-count {
    display: inline-grid; place-items: center; min-width: 20px; height: 20px; padding: 0 6px; border-radius: 999px;
    background: var(--accent); color: var(--on-accent); font-size: 12px; font-weight: 700;
  }
  #demo-emails.pulse { animation: demo-pulse 1s ease-out 2; }
  @keyframes demo-pulse { 50% { box-shadow: 0 0 0 4px var(--accent-soft), 0 0 0 6px var(--accent); } }
  .demo-mail-list { display: grid; gap: 12px; }
  .demo-mail { border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }
  .demo-mail summary { cursor: pointer; padding: 10px 14px; background: var(--surface-2); font-weight: 650; overflow-wrap: anywhere; }
  .demo-mail .demo-meta { font-weight: 400; color: var(--muted); font-size: 13px; }
  .demo-mail .demo-body { overflow-x: auto; }
  .demo-empty { color: var(--muted); }
</style>`;

const demoMarkup = `<div class="demo-bar" id="demo-bar">
  <span class="demo-text" id="demo-text"><b>Demo</b> with example parts. Nothing is saved, and emails show up here instead of being sent.</span>
  <span class="demo-actions">
    <button class="btn small" type="button" id="demo-emails">Emails sent <span class="demo-count" id="demo-count">0</span></button>
    <button class="btn small subtle" type="button" id="demo-reset">Start over</button>
  </span>
</div>
<dialog id="demo-mail" class="modal" aria-labelledby="demo-mail-title">
  <div class="modal-head">
    <h2 id="demo-mail-title">Emails the app sent</h2>
    <button class="icon-btn" type="button" id="demo-mail-close" aria-label="Close">&times;</button>
  </div>
  <div class="modal-body"><div class="demo-mail-list" id="demo-mail-list"></div></div>
</dialog>
`;

const runtime = `<script>
(function () {
  var shims = {
    fs: { readFileSync: function () { throw new Error('There are no files in the demo.'); } },
    path: { join: function () { return Array.prototype.slice.call(arguments).filter(Boolean).join('/'); } },
    vm: {},
  };
  var fakesModule = { exports: {} };
  (function (require, module, __dirname) {
${inline(fakes)}
  })(function (name) { return shims[name]; }, fakesModule, '');
  var sampleModule = { exports: {} };
  (function (require, module) {
${inline(sample)}
  })(function () { return {}; }, sampleModule);
  window.__demo = {
    fakes: fakesModule.exports,
    seed: sampleModule.exports.seed,
    environment: fakesModule.exports.createEnvironment({
      owner: 'you@example.com',
      spreadsheetName: 'Parts Inventory (demo)',
      timeZone: (Intl.DateTimeFormat().resolvedOptions().timeZone) || 'America/New_York',
    }),
  };
})();
</script>
<script>
(function (${googleServices.join(', ')}) {
${inline(code)}
window.__demo.server = { ${serverFunctions.map((f) => f + ': ' + f).join(', ')} };
})(${googleServices.map((g) => 'window.__demo.environment.globals.' + g).join(', ')});
</script>
<script>
(function () {
  var demo = window.__demo;
  var app = demo.fakes.makeApp(demo.environment, function (name) { return demo.server[name]; });
  app.run('setup');
  demo.seed(app);
  app.env.sent.length = 0;

  // google.script.run, answered by the Code.js loaded above.
  function runner(ok, fail) {
    return new Proxy({}, { get: function (_, name) {
      if (name === 'withSuccessHandler') return function (fn) { return runner(fn, fail); };
      if (name === 'withFailureHandler') return function (fn) { return runner(ok, fn); };
      if (name === 'withUserObject') return function () { return runner(ok, fail); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        setTimeout(function () {
          var result;
          var error = null;
          try { result = app.run.apply(app, [name].concat(args)); } catch (err) { error = err; }
          setTimeout(function () {
            if (error) { if (fail) fail(new Error(String(error.message || error))); }
            else if (ok) ok(result === undefined ? null : result);
            showCount();
          }, 120);
        }, 30);
      };
    } });
  }
  window.google = { script: { run: runner(null, null), host: { close: function () {} } } };

  var shown = 0;
  function showCount() {
    var count = app.env.sent.length;
    document.getElementById('demo-count').textContent = String(count);
    var button = document.getElementById('demo-emails');
    if (count > shown) {
      button.classList.remove('pulse');
      void button.offsetWidth;
      button.classList.add('pulse');
    }
    shown = count;
  }

  function note(text) {
    var el = document.getElementById('demo-text');
    var before = el.innerHTML;
    el.textContent = text;
    setTimeout(function () { el.innerHTML = before; }, 4000);
  }

  function openEmails() {
    var list = document.getElementById('demo-mail-list');
    while (list.firstChild) list.removeChild(list.firstChild);
    var mails = app.env.sent.slice().reverse();
    if (!mails.length) {
      var empty = document.createElement('p');
      empty.className = 'demo-empty';
      empty.textContent = 'No emails yet. Tap \\u2212 on a part until it drops to its alert level, or use Settings > Send test email.';
      list.appendChild(empty);
    }
    mails.forEach(function (mail, i) {
      var item = document.createElement('details');
      item.className = 'demo-mail';
      item.open = i === 0;
      var summary = document.createElement('summary');
      summary.textContent = mail.subject;
      var meta = document.createElement('div');
      meta.className = 'demo-meta';
      meta.textContent = 'To ' + mail.to.split(',').join(', ') + ', from ' + mail.name + ' (' + demo.environment.env.owner + ')';
      summary.appendChild(meta);
      var body = document.createElement('div');
      body.className = 'demo-body';
      body.innerHTML = mail.htmlBody; // built by Code.js, which escapes everything typed into the app
      item.appendChild(summary);
      item.appendChild(body);
      list.appendChild(item);
    });
    document.getElementById('demo-mail').showModal();
  }

  document.getElementById('demo-emails').addEventListener('click', openEmails);
  document.getElementById('demo-mail-close').addEventListener('click', function () { document.getElementById('demo-mail').close(); });
  document.getElementById('demo-reset').addEventListener('click', function () { location.reload(); });
  document.addEventListener('click', function (e) {
    var link = e.target.closest && e.target.closest('a[href*="docs.google.com/spreadsheets"]');
    if (!link) return;
    e.preventDefault();
    note('In the real app this opens your Google Sheet. The demo has no spreadsheet.');
  }, true);
})();
</script>
`;

// The artifact page supplies <html>, <head> and <body>; drop the ones the Apps Script page carries.
page = replaceOnce(page, '<!DOCTYPE html>\n<html lang="en">\n<head>\n<base target="_top">\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n',
  '<title>Parts Inventory Demo</title>\n');
page = replaceOnce(page, '</style>\n</head>\n<body>\n', '</style>\n' + demoCss + '\n' + demoMarkup);
page = replaceOnce(page, '<script>const BOOT = <?!= bootJson ?>;</script>', runtime + '<script>const BOOT = {"mode":"demo","part":""};</script>');
page = replaceOnce(page, '</body>\n</html>\n', '');
if (page.indexOf('<?!= bootJson') !== -1) throw new Error('The page template tag was not replaced.');

if (standalone) {
  page = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n</head>\n<body>\n' +
    page + '</body>\n</html>\n';
}
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, page);
console.log('Wrote ' + path.relative(process.cwd(), out) + ' (' + Math.round(page.length / 1024) + ' KB)');
