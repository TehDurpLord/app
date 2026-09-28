#!/usr/bin/env node
'use strict';
/**
 * Local preview of the web app, with no Google account needed.
 *
 *   npm run preview           then open http://localhost:8080
 *
 * The real src/Code.js runs on an in-memory copy of the Google services
 * (see test/gas-fakes.js) and the real src/Index.html is served with a
 * stand-in for google.script.run, so what you click here runs the same code
 * as the deployed app. Emails aren't sent; they're listed at /emails.
 *
 * Extra URLs:
 *   /?part=P-0004     open a part directly (like the links in emails)
 *   /?as=visitor      see the app as a coworker (not the owner)
 *   /?code=bolt-7731  turn on an access code (then open /?as=visitor)
 *   /?empty=1         start with no parts
 *   /emails           the emails the app "sent"
 *   /reset            start over with the example data
 */
const http = require('http');
const { URL } = require('url');
const { createApp } = require('../test/gas-fakes');
const { seed } = require('./sample-data');

const PORT = Number(process.env.PORT) || 8080;
const DELAY = Number(process.env.DELAY || 250); // google.script.run is never instant
const OWNER = 'shop@example.com';

let app;
function reset(options) {
  const opts = options || {};
  app = createApp({
    owner: OWNER,
    spreadsheetName: 'Maintenance Parts',
    serviceUrl: 'https://script.google.com/macros/s/PREVIEW/exec',
  });
  app.run('setup');
  if (!opts.empty) seed(app);
  app.env.sent.length = 0;
  if (opts.code) app.run('apiSaveSettings', {}, { accessCode: opts.code });
  return app;
}
reset();

function shim(as) {
  return '<script>(function () {\n' +
    '  var AS = ' + JSON.stringify(as) + ', DELAY = ' + DELAY + ';\n' +
    '  function runner(ok, fail) {\n' +
    '    return new Proxy({}, { get: function (_, name) {\n' +
    '      if (name === "withSuccessHandler") return function (fn) { return runner(fn, fail); };\n' +
    '      if (name === "withFailureHandler") return function (fn) { return runner(ok, fn); };\n' +
    '      if (name === "withUserObject") return function () { return runner(ok, fail); };\n' +
    '      return function () {\n' +
    '        var args = Array.prototype.slice.call(arguments);\n' +
    '        fetch("/rpc", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fn: name, args: args, as: AS }) })\n' +
    '          .then(function (r) { return r.json(); })\n' +
    '          .then(function (res) { setTimeout(function () {\n' +
    '            if (res.error) { if (fail) fail(new Error(res.error)); } else if (ok) ok(res.result);\n' +
    '          }, DELAY); })\n' +
    '          .catch(function (err) { if (fail) fail(new Error("NetworkError: " + err.message)); });\n' +
    '      };\n' +
    '    } });\n' +
    '  }\n' +
    '  window.google = { script: { run: runner(null, null), host: { close: function () {} } } };\n' +
    '})();</script>\n';
}

function asUser(as, fn) {
  app.env.active = as === 'visitor' ? '' : OWNER;
  try {
    return fn();
  } finally {
    app.env.active = OWNER;
  }
}

function send(res, status, type, body) {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(body);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function emailsPage() {
  const items = app.env.sent.map((m, i) => '<section><h2>' + escapeHtml(m.subject) + '</h2><p>To: ' + escapeHtml(m.to) +
    ' &middot; From: ' + escapeHtml(m.name) + '</p><iframe src="/emails/' + i + '"></iframe></section>').reverse().join('');
  return '<!doctype html><meta charset="utf-8"><title>Sent emails</title><style>body{font:15px system-ui;margin:24px;background:#eee}' +
    'section{background:#fff;padding:16px;margin:0 0 20px;border-radius:10px}h2{font-size:17px;margin:0}p{color:#666;margin:4px 0 12px}' +
    'iframe{width:100%;height:640px;border:1px solid #ddd;border-radius:6px}</style>' +
    '<h1>Emails the app sent (' + app.env.sent.length + ')</h1>' + (items || '<p>None yet. Take some parts below their alert level.</p>');
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (req.method === 'POST' && url.pathname === '/rpc') {
    let body = '';
    req.on('data', (chunk) => { body += chunk; });
    req.on('end', () => {
      try {
        const { fn, args, as } = JSON.parse(body);
        const result = asUser(as, () => app.run(fn, ...args));
        send(res, 200, 'application/json', JSON.stringify({ result: result === undefined ? null : result }));
      } catch (err) {
        send(res, 200, 'application/json', JSON.stringify({ error: String(err.message || err).replace(/^Error: /, '') }));
      }
    });
    return;
  }
  if (url.pathname === '/' || url.pathname === '/exec') {
    const q = Object.fromEntries(url.searchParams);
    if (q.empty || q.code) reset({ empty: !!q.empty, code: q.code });
    const as = q.as === 'visitor' ? 'visitor' : 'owner';
    const output = asUser(as, () => app.context.doGet({ parameter: q }));
    const html = output.getContent().replace('<script>const BOOT', shim(as) + '<script>const BOOT');
    send(res, 200, 'text/html; charset=utf-8', html);
    return;
  }
  if (url.pathname === '/sidebar') {
    const output = asUser('owner', () => {
      app.env.uiAvailable = true;
      app.run('menuOpenSidebar');
      app.env.uiAvailable = false;
      return app.env.sidebars.pop();
    });
    const html = output.getContent().replace('<script>const BOOT', shim('owner') + '<script>const BOOT');
    send(res, 200, 'text/html; charset=utf-8', '<!doctype html><title>Sidebar</title><body style="margin:0;background:#888">' +
      '<iframe style="width:300px;height:100vh;border:0;background:#fff" srcdoc="' + escapeHtml(html) + '"></iframe>');
    return;
  }
  if (url.pathname === '/emails') { send(res, 200, 'text/html; charset=utf-8', emailsPage()); return; }
  const email = /^\/emails\/(\d+)$/.exec(url.pathname);
  if (email && app.env.sent[+email[1]]) {
    send(res, 200, 'text/html; charset=utf-8', '<!doctype html><meta charset="utf-8"><body style="margin:0">' + app.env.sent[+email[1]].htmlBody);
    return;
  }
  if (url.pathname === '/reset') { reset(); res.writeHead(302, { Location: '/' }); res.end(); return; }
  if (url.pathname === '/state.json') { send(res, 200, 'application/json', JSON.stringify({ sent: app.env.sent, sheets: app.ss.getSheets().map((s) => ({ name: s.getName(), values: s.dump() })) })); return; }
  send(res, 404, 'text/plain', 'Not found');
});

if (require.main === module) {
  server.listen(PORT, () => console.log('Parts Inventory preview: http://localhost:' + PORT + '  (emails: /emails)'));
}

module.exports = { server, reset, getApp: () => app };
