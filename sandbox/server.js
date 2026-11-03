#!/usr/bin/env node
'use strict';
/*
 * The Ops Hub sandbox: the whole platform on this computer, with fake data.
 *
 *   node sandbox/server.js              start it (http://localhost:8080/sandbox/)
 *   node sandbox/server.js --port 9000  another port
 *   node sandbox/server.js --fresh      ignore saved data and start from the fixtures
 *   node sandbox/server.js --host 0.0.0.0   let other devices on the network open it (a phone)
 *
 * It serves the repo's own pages and runs the repo's own Apps Script files, so every
 * edit shows up on refresh. Addresses that would reach Google, Fabman or real people are
 * swapped for stand-ins (lib/rewrite.js, public/guard.js); emails go to an outbox.
 * No npm install needed: plain Node 18 or newer.
 */

process.env.TZ = process.env.SANDBOX_TZ || 'America/New_York';   // the projects run on New York time

const http = require('http');
const fs = require('fs');
const path = require('path');

const { Account } = require('./lib/gas');
const { Runner, defineProjects } = require('./lib/projects');
const site = require('./lib/site');
const fixtures = require('./lib/fixtures');
const store = require('./lib/store');
const forms = require('./lib/forms');
const pages = require('./lib/pages');
const { makeRewriter, headTags, inject } = require('./lib/rewrite');

const REPO = path.resolve(__dirname, '..');
const PUBLIC = path.join(__dirname, 'public');

// ---------------------------------------------------------------- options

const args = process.argv.slice(2);
const opt = (name, fallback) => { const i = args.findIndex((a) => a === '--' + name || a.indexOf('--' + name + '=') === 0); if (i < 0) return fallback; const a = args[i]; return a.indexOf('=') > 0 ? a.slice(a.indexOf('=') + 1) : (args[i + 1] && args[i + 1].indexOf('--') !== 0 ? args[i + 1] : true); };
const PORT = Number(opt('port', process.env.PORT || 8080));
const HOST = String(opt('host', '127.0.0.1'));
const FRESH = !!opt('fresh', false);
const STATE_FILE = path.resolve(String(opt('state', path.join(__dirname, '.state', 'state.json'))));   // tests use their own

// ---------------------------------------------------------------- the fake world

const cfg = site.readConfig(REPO);
const peekRunner = new Runner(REPO, new Account(), defineProjects({}, {}));
const ids = site.readIds(cfg, (k, e) => peekRunner.peek(k, e));
const projects = defineProjects(ids, { space: cfg.spaceUrl, inventory: cfg.inventoryUrl, command: cfg.statsUrl });
const rewriter = makeRewriter(cfg);
// Which spreadsheets the Command Center reads, for the Sheets list.
const SOURCES = {};
[[ids.space, 'Command Center'], [ids.ell, 'Command Center'], [ids.fabman, 'Command Center'], [ids.master, 'Command Center (switched off)'], [ids.drivers, 'Command Center'], [ids.oldInventory, 'Command Center'], [ids.people, 'Staff page (people, schedule)']]
  .forEach(([id, who]) => { (SOURCES[id] = SOURCES[id] || []).push(who); });
SOURCES[ids.space].unshift('Space Status web app');

let account, runner;
function fresh() {
  const built = fixtures.build(ids, (acc) => new Runner(REPO, acc, projects));
  account = built.account; runner = built.runner;
  saveNow();
}
function boot() {
  const saved = FRESH ? null : store.load(STATE_FILE);
  if (saved) { account = saved; runner = new Runner(REPO, account, projects); return 'saved data from ' + (account.builtAt ? new Date(account.builtAt).toLocaleString() : 'before'); }
  fresh();
  return 'fresh fake data';
}

let saveTimer = null;
function saveNow() { try { store.save(STATE_FILE, account); } catch (err) { console.warn('  Could not save the sandbox: ' + err.message); } }
function saveSoon() { if (!account.dirty) return; clearTimeout(saveTimer); saveTimer = setTimeout(saveNow, 400); }

const ctx = () => ({ account, runner, projects, ids, sources: SOURCES, status });

function status() {
  const out = { open: 0, pending: 0, projects: 0, proposals: 0, items: 0, mail: account.outbox.length };
  const ss = account.ss(ids.space);
  const rowsOf = (sheet) => { const sh = ss && ss.getSheetByName(sheet); if (!sh) return []; const v = sh.rows(); const H = (v[0] || []).map((x) => String(x).trim().toLowerCase()); return v.slice(1).map((r) => (k) => r[H.indexOf(k)]); };
  rowsOf('Form Responses').forEach((get) => { if (!get('issue token') || get('addressed at')) return; if (get('completed at')) out.pending++; else out.open++; });
  rowsOf('Projects').forEach((get) => { if (/assigned|in progress|pending/i.test(String(get('status')))) out.projects++; });
  rowsOf('Proposals').forEach((get) => { if (/provisional/i.test(String(get('status')))) out.proposals++; });
  const inv = account.ss(ids.inventory), master = inv && inv.getSheetByName('Inventory Master');
  if (master) out.items = master.rows().slice(1).filter((r) => String(r[0] || '').trim()).length;
  return out;
}

// ---------------------------------------------------------------- HTTP helpers

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.ico': 'image/x-icon', '.woff2': 'font/woff2', '.woff': 'font/woff', '.otf': 'font/otf', '.ttf': 'font/ttf', '.md': 'text/plain; charset=utf-8', '.txt': 'text/plain; charset=utf-8', '.gs': 'text/plain; charset=utf-8', '.pdf': 'application/pdf' };
const TEXT = /^(text\/|application\/json|image\/svg)/;

function send(res, code, body, type, headers) {
  res.writeHead(code, Object.assign({ 'Content-Type': type || 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' }, headers || {}));
  res.end(body);
}
const json = (res, obj, code) => send(res, code || 200, JSON.stringify(obj), 'application/json; charset=utf-8');
const redirect = (res, to) => send(res, 302, '', 'text/plain', { Location: to });

function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    const chunks = []; let size = 0;
    req.on('data', (c) => { size += c.length; if (size > (limit || 30e6)) { reject(new Error('Upload too large')); req.destroy(); } else chunks.push(c); });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}
const parseJson = (s) => { try { return JSON.parse(s || '{}'); } catch (e) { return {}; } };

// What a run sends back to the page: what happened, without the raw objects.
function runSummary(r) {
  const run = r.run || {};
  return {
    ok: r.ok, error: r.error, project: run.project, fn: run.fn, ms: run.ms,
    value: r.value && r.value.isHtmlOutput ? { html: true } : (r.value && typeof r.value.getContent === 'function' ? { text: r.value.getContent().slice(0, 2000) } : safeValue(r.value)),
    logs: run.logs, ui: run.ui, menus: run.menus,
    mails: (run.mails || []).map((id) => { const m = account.outbox.find((x) => x.id === id); return m ? { id: m.id, subject: m.subject, to: m.to } : { id }; }),
  };
}
function safeValue(v) { try { return v === undefined ? undefined : JSON.parse(JSON.stringify(v)); } catch (e) { return String(v); } }

// ---------------------------------------------------------------- the site's own files

// Real exports, the script sources (05 holds live keys) and dotfiles are never served.
// The site does not need any of them; the scripts run from /sandbox/code.
const BLOCKED = /^(downloads|email-notifications|node_modules|apps-script|sandbox)(\/|$)|^analysis_raw\.json$|^analyze_files\.py$|(^|\/)\./;

function serveFile(req, res, urlPath) {
  let asked;
  try { asked = decodeURIComponent(urlPath).replace(/^\/+/, ''); } catch (e) { return send(res, 400, 'Bad path', 'text/plain'); }
  const full = path.resolve(REPO, asked || 'index.html');
  if (full !== REPO && full.indexOf(REPO + path.sep) !== 0) return send(res, 403, 'Forbidden', 'text/plain');
  // Check the path as it resolves, so ../ tricks cannot reach a blocked folder.
  const rel = path.relative(REPO, full).split(path.sep).join('/') || 'index.html';
  if (BLOCKED.test(rel)) return send(res, 404, /^apps-script/.test(rel) ? 'The sandbox runs the scripts instead of serving them. See /sandbox/code' : 'Not served by the sandbox (real data, private files or sandbox internals).', 'text/plain; charset=utf-8');
  let stat;
  try { stat = fs.statSync(full); } catch (e) { return send(res, 404, notFound(rel), 'text/html; charset=utf-8'); }
  if (stat.isDirectory()) return fs.existsSync(path.join(full, 'index.html')) ? redirect(res, '/' + (rel === 'index.html' ? '' : rel.replace(/\/?$/, '/')) + 'index.html') : send(res, 404, notFound(rel));
  const ext = path.extname(full).toLowerCase();
  const type = TYPES[ext] || 'application/octet-stream';
  if (!TEXT.test(type)) return send(res, 200, fs.readFileSync(full), type);
  let text = rewriter.files(fs.readFileSync(full, 'utf8'));
  if (rel === 'config.js') {
    text += '\n// ---- added by the sandbox (the file on disk is unchanged) ----\nwindow.SANDBOX = true;\nif (window.INVENTORY) window.INVENTORY.appUrl = location.origin + \'/inventory.html\';\n';
  }
  if (ext === '.html') text = inject(text, headTags(rewriter));
  send(res, 200, text, type);
}
function notFound(rel) { return '<!DOCTYPE html><meta charset="utf-8"><title>Not found</title><body style="font:15px/1.6 system-ui;margin:40px"><h1>Not found</h1><p><code>' + pages.h(rel) + '</code> is not in the repo.</p><p><a href="/sandbox/">Sandbox control panel</a></p>'; }

// ---------------------------------------------------------------- the web apps

function serveWebApp(req, res, key, method, query, body) {
  const project = projects[key];
  const fn = method === 'POST' ? 'doPost' : 'doGet';
  const r = runner.exec(key, fn, [Runner.webEvent(query, body, req.headers['content-type'])]);
  saveSoon();
  if (!r.ok) return send(res, 500, pages.gasError(r.error));
  const out = r.value;
  if (out && out.isHtmlOutput) {
    let html = out.html;
    if (!/<html[\s>]/i.test(html)) {
      html = '<!DOCTYPE html><html><head><meta charset="utf-8">' + (out.title ? '<title>' + pages.h(out.title) + '</title>' : '')
        + out.metas.map(([n, c]) => '<meta name="' + pages.h(n) + '" content="' + pages.h(c) + '">').join('') + '</head><body>' + html + '</body></html>';
    }
    html = inject(rewriter.gas(html), headTags(rewriter, { gas: project.local.replace(/\/exec$/, '/run') }));
    return send(res, 200, html);
  }
  if (out && typeof out.getContent === 'function') {
    const mime = { JAVASCRIPT: 'text/javascript; charset=utf-8', JSON: 'application/json; charset=utf-8', CSV: 'text/csv; charset=utf-8', XML: 'application/xml' }[out.mime] || 'text/plain; charset=utf-8';
    return send(res, 200, rewriter.data(out.getContent()), mime);
  }
  send(res, 200, '<!DOCTYPE html><meta charset="utf-8"><body style="font:15px Arial;margin:40px">The script completed but did not return anything.</body>');
}

async function serveRun(req, res, key) {
  const body = parseJson(await readBody(req));
  const r = runner.exec(key, String(body.fn || ''), Array.isArray(body.args) ? body.args : [], { client: true });
  saveSoon();
  if (!r.ok) return json(res, { ok: false, error: r.error });
  let value;
  try { value = r.value === undefined ? undefined : JSON.parse(rewriter.gas(JSON.stringify(r.value))); } catch (e) { value = null; }
  json(res, { ok: true, value });
}

// ---------------------------------------------------------------- sheets

// "2026-10-21", "10/21/2026 9:30", "42", "TRUE": typed into a cell the way Sheets reads them.
function typed(value, kind) {
  if (kind === 'bool') return value === true || value === 'true';
  const s = String(value == null ? '' : value);
  if (s === '') return '';
  if (/^(true|false)$/i.test(s)) return /^true$/i.test(s);
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  return s;
}

function menusFor(ssId) {
  const out = [];
  Object.values(projects).forEach((p) => {
    if (p.boundId !== ssId || !runner.functionsOf(p.key).some((f) => f.name === 'onOpen')) return;
    const r = runner.exec(p.key, 'onOpen', [{ source: account.ss(ssId), authMode: 'LIMITED' }]);
    (r.run.menus || []).forEach((m) => out.push(Object.assign({ project: p.key }, m)));
  });
  return out;
}

function gviz(res, ss, query) {
  const tqx = String(query.get('tqx') || '');
  const handler = (/responseHandler:([A-Za-z_$][\w$]*)/.exec(tqx) || [])[1] || 'google.visualization.Query.setResponse';
  const sh = ss.getSheetById(query.get('gid')) || ss.getSheets()[0];
  const v = sh.rows();
  const sel = /select\s+(.+)/i.exec(String(query.get('tq') || ''));
  const cols = sel ? sel[1].split(',').map((c) => c.trim().toUpperCase()).map((c) => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1) : (v[0] || []).map((_, i) => i);
  const cell = (x) => (x instanceof Date ? { v: 'Date(' + x.getFullYear() + ',' + x.getMonth() + ',' + x.getDate() + ')', f: x.toLocaleDateString('en-US') } : x === '' ? null : { v: x });
  const table = {
    cols: cols.map((c) => ({ id: String.fromCharCode(65 + c), label: v[0] && v[0][c] !== undefined ? String(v[0][c]) : '', type: 'string' })),
    rows: v.slice(1).map((r) => ({ c: cols.map((c) => cell(r[c] === undefined ? '' : r[c])) })),
  };
  send(res, 200, '/*O_o*/\n' + handler + '(' + JSON.stringify({ version: '0.6', reqId: '0', status: 'ok', sig: 'sandbox', table }) + ');', 'text/javascript; charset=utf-8');
}

// ---------------------------------------------------------------- time passing

function ageOneDay() {
  const DAY = 86400000;
  Object.values(account.spreadsheets).forEach((ss) => ss.sheets.forEach((sh) => {
    sh.cells.forEach((cl) => { if (cl.v instanceof Date) cl.v = new Date(cl.v.getTime() - DAY); });
  }));
  Object.values(account.drive.files).forEach((f) => { if (f.created instanceof Date) f.created = new Date(f.created.getTime() - DAY); });
  Object.values(account.state).forEach((s) => { s.cache = {}; });
  account.agedDays = (account.agedDays || 0) + 1;
  account.dirty = true;
}

// ---------------------------------------------------------------- README as a page

function markdown(md) {
  const esc = pages.h;
  const inline = (s) => esc(s).replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>').replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  const out = []; let list = null, code = null;
  const close = () => { if (list) { out.push('</' + list + '>'); list = null; } };
  md.split('\n').forEach((line) => {
    if (code !== null) { if (/^```/.test(line)) { out.push('<pre class="sbx-log">' + esc(code.join('\n')) + '</pre>'); code = null; } else code.push(line); return; }
    if (/^```/.test(line)) { close(); code = []; return; }
    let m;
    if ((m = /^(#{1,3})\s+(.*)$/.exec(line))) { close(); out.push('<h' + (m[1].length) + '>' + inline(m[2]) + '</h' + m[1].length + '>'); return; }
    if ((m = /^\s*[-*]\s+(.*)$/.exec(line))) { if (list !== 'ul') { close(); out.push('<ul>'); list = 'ul'; } out.push('<li>' + inline(m[1]) + '</li>'); return; }
    if ((m = /^\s*\d+\.\s+(.*)$/.exec(line))) { if (list !== 'ol') { close(); out.push('<ol>'); list = 'ol'; } out.push('<li>' + inline(m[1]) + '</li>'); return; }
    if (!line.trim()) { close(); return; }
    if (list && /^\s{2,}/.test(line)) { out[out.length - 1] = out[out.length - 1].replace(/<\/li>$/, ' ' + inline(line.trim()) + '</li>'); return; }
    close(); out.push('<p>' + inline(line) + '</p>');
  });
  close();
  return '<article class="sbx-doc">' + out.join('\n').replace(/<\/p>\n<p>/g, '</p>\n<p>') + '</article>';
}

// ---------------------------------------------------------------- routes

async function route(req, res) {
  const url = new URL(req.url, 'http://sandbox.local');
  const p = url.pathname;
  const q = url.searchParams;
  const method = req.method;

  if (p === '/' || p === '/sandbox') return redirect(res, '/sandbox/');
  if (!p.startsWith('/sandbox/')) return serveFile(req, res, p);

  if (p === '/sandbox/') return send(res, 200, pages.panel(ctx()));
  if (p === '/sandbox/readme') {
    const file = path.join(__dirname, 'README.md');
    const md = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '# How the sandbox works\n\nsandbox/README.md is missing.';
    return send(res, 200, pages.layout('How the sandbox works', markdown(md), ctx()));
  }
  if (p.startsWith('/sandbox/static/')) {
    const f = path.resolve(PUBLIC, p.slice('/sandbox/static/'.length));
    if (f.indexOf(PUBLIC + path.sep) !== 0 || !fs.existsSync(f)) return send(res, 404, 'Not found', 'text/plain');
    return send(res, 200, fs.readFileSync(f), TYPES[path.extname(f)] || 'application/octet-stream');
  }

  // The sandbox's own POSTs are JSON. Insisting on it means a page on another site
  // cannot press the sandbox's buttons (a JSON POST from elsewhere needs permission
  // the sandbox never gives).
  if (method === 'POST' && /^\/sandbox\/(api\/|space\/run$|forms\/)/.test(p) && !/^application\/json/i.test(String(req.headers['content-type'] || ''))) {
    return json(res, { ok: false, error: 'The sandbox only accepts JSON here.' }, 415);
  }

  // Web apps: /sandbox/space/exec, /sandbox/inventory/exec, /sandbox/command/exec
  let m = /^\/sandbox\/(space|inventory|command)\/(exec|dev)$/.exec(p);
  if (m) return serveWebApp(req, res, m[1], method, q, method === 'POST' ? await readBody(req) : undefined);
  if ((m = /^\/sandbox\/(space)\/run$/.exec(p)) && method === 'POST') return serveRun(req, res, m[1]);

  // JSON API used by the sandbox's own pages
  if (p.startsWith('/sandbox/api/')) {
    if (p === '/sandbox/api/status') return json(res, Object.assign(status(), { problems: runner.problems.slice(-10) }));
    if (p === '/sandbox/api/ids') return json(res, ids);
    // A tab's values as JSON (dates as ISO text), for scripts and tests:
    //   /sandbox/api/sheet?id=<spreadsheet id>&sheet=<tab name>
    if (p === '/sandbox/api/sheet' && method === 'GET') {
      const ss = account.ss(q.get('id') || ''), sh = ss && (q.get('sheet') ? ss.getSheetByName(q.get('sheet')) : ss.getSheets()[0]);
      if (!sh) return json(res, { ok: false, error: 'No such sheet' }, 404);
      return json(res, { ok: true, spreadsheet: ss.getName(), sheet: sh.getName(), gid: sh.getSheetId(), values: sh.rows().map((r) => r.map((v) => (v instanceof Date ? v.toISOString() : v))) });
    }
    if (method !== 'POST') return json(res, { ok: false, error: 'POST only' }, 405);
    const body = parseJson(await readBody(req));
    if (p === '/sandbox/api/run') {
      if (!projects[body.project]) return json(res, { ok: false, error: 'No project ' + body.project });
      // Run from a sheet page: that tab is the active sheet, as on Google.
      const onSs = body.id && account.ss(body.id), onSheet = onSs && onSs.getSheetByName(body.sheet || '');
      if (onSheet) onSs.setActiveSheet(onSheet);
      const r = runner.exec(body.project, String(body.fn || ''), [], { answer: body.answer, promptText: body.promptText });
      saveSoon();
      return json(res, Object.assign(runSummary(r), { html: pages.resultBlock(runSummary(r)) }));
    }
    if (p === '/sandbox/api/cell') {
      const ss = account.ss(body.id), sh = ss && ss.getSheetByName(body.sheet);
      if (!sh) return json(res, { ok: false, error: 'No such sheet' });
      const row = Number(body.row), col = Number(body.col);
      if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
      if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
      const range = sh.getRange(row, col);
      const oldValue = range.getValue();
      const value = typed(body.value, body.kind);
      if (typeof value === 'string' && value.charAt(0) === '=') range.setFormula(value);   // shown, not calculated
      else range.setValue(value);
      const event = { range, value: value instanceof Date ? value.toISOString() : value, oldValue: oldValue instanceof Date ? oldValue.toISOString() : oldValue, source: ss, user: { getEmail: () => account.user }, authMode: 'FULL', triggerUid: 'sandbox-edit' };
      const results = runner.fire(ss.id, 'edit', event).concat(runner.fire(ss.id, 'change', { changeType: 'EDIT', source: ss, authMode: 'FULL' }));
      saveSoon();
      return json(res, { ok: true, results: results.map(runSummary), html: results.map((x) => pages.resultBlock(runSummary(x))).join('') });
    }
    if (p === '/sandbox/api/row') {
      const ss = account.ss(body.id), sh = ss && ss.getSheetByName(body.sheet);
      if (!sh || !(Number(body.row) > 1)) return json(res, { ok: false, error: 'No such row' });
      sh.deleteRow(Number(body.row));
      const results = runner.fire(ss.id, 'change', { changeType: 'REMOVE_ROW', source: ss, authMode: 'FULL' });
      saveSoon();
      return json(res, { ok: true, html: results.map((x) => pages.resultBlock(runSummary(x))).join('') });
    }
    if (p === '/sandbox/api/reminders' || p === '/sandbox/api/day') {
      if (p === '/sandbox/api/day') ageOneDay();
      const before = account.outbox.length;
      const r = runner.exec('space', 'sendReminders', []);
      saveSoon();
      const sent = account.outbox.length - before;
      const lead = '<p class="sbx-ui"><b>' + (p === '/sandbox/api/day' ? 'A day went by.' : 'Reminder job ran.') + '</b> ' + (sent ? sent + ' reminder email' + (sent === 1 ? '' : 's') + ' sent.' : 'No reminders were due.') + '</p>';
      return json(res, Object.assign(runSummary(r), { html: lead + pages.resultBlock(runSummary(r)) }));
    }
    if (p === '/sandbox/api/reset') { fresh(); return json(res, { ok: true }); }
    if (p === '/sandbox/api/fabman') {
      if (typeof body.down === 'boolean') account.fabman.down = body.down;
      if (typeof body.refuseKey === 'boolean') account.fabman.refuseKey = body.refuseKey;
      account.dirty = true; saveSoon();
      return json(res, { ok: true, down: account.fabman.down, refuseKey: account.fabman.refuseKey });
    }
    if (p === '/sandbox/api/outbox/clear') { account.outbox = []; account.dirty = true; saveSoon(); return json(res, { ok: true }); }
    return json(res, { ok: false, error: 'Unknown request' }, 404);
  }

  // Sheets: /sandbox/sheets/<id>/edit, /sandbox/sheets/<id>/gviz/tq
  if (p === '/sandbox/sheets/' || p === '/sandbox/sheets') return send(res, 200, pages.sheetsList(ctx()));
  if ((m = /^\/sandbox\/sheets\/([-\w]+)(\/.*)?$/.exec(p))) {
    const ss = account.ss(m[1]);
    if (m[2] && /^\/gviz\/tq/.test(m[2])) return ss ? gviz(res, ss, q) : send(res, 404, '', 'text/javascript');
    if (!ss) return send(res, 404, pages.notInSandbox(ctx(), m[1]));
    return send(res, 200, pages.sheetPage(ctx(), ss, q.get('gid'), menusFor(ss.id)));
  }

  // Outbox
  if (p === '/sandbox/outbox') return send(res, 200, pages.outbox(ctx()));
  if ((m = /^\/sandbox\/outbox\/([\w]+)(\/body|\/cid\/(.+))?$/.exec(p))) {
    const mail = account.outbox.find((x) => x.id === m[1]);
    if (!mail) return send(res, 404, pages.layout('Not found', '<h1>That email is not in the outbox</h1><p><a href="/sandbox/outbox">Outbox</a></p>', ctx()));
    if (m[3]) { const img = mail.images[decodeURIComponent(m[3])]; return img ? send(res, 200, Buffer.from(img.data, 'base64'), img.mime) : send(res, 404, '', 'text/plain'); }
    if (m[2] === '/body') {
      const html = rewriter.gas(mail.html).replace(/(["'(])cid:([^"')\s]+)/g, (all, q1, cid) => q1 + '/sandbox/outbox/' + mail.id + '/cid/' + encodeURIComponent(cid));
      return send(res, 200, inject('<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_blank"></head><body style="margin:0">' + html + '</body></html>', headTags(rewriter)));
    }
    return send(res, 200, pages.mailPage(ctx(), mail));
  }

  // Drive
  if (p === '/sandbox/drive/thumbnail' || p === '/sandbox/drive/uc' || (m = /^\/sandbox\/drive\/raw\/([-\w]+)$/.exec(p))) {
    const f = account.fileById(m && m[1] ? m[1] : q.get('id'));
    if (!f || f.trashed) return send(res, 404, 'Not found', 'text/plain');
    return send(res, 200, Buffer.from(f.bytes.map((b) => b & 255)), f.mime);
  }
  if ((m = /^\/sandbox\/drive\/file\/([-\w]+)/.exec(p)) || p === '/sandbox/drive/open') {
    const f = account.fileById(m ? m[1] : q.get('id'));
    if (!f) return send(res, 404, pages.layout('Not found', '<h1>No such file in the sandbox Drive</h1>', ctx()));
    return send(res, 200, pages.drivePage(ctx(), f));
  }

  // Forms
  if ((m = /^\/sandbox\/forms\/([\w-]+)$/.exec(p))) {
    const form = forms.definitions(account, ids)[m[1]];
    if (!form) return send(res, 404, pages.layout('Not found', '<h1>No such form</h1>', ctx()));
    if (method === 'GET') return send(res, 200, pages.formPage(ctx(), form, q.get('editor') === '1'));
    const body = parseJson(await readBody(req));
    const files = {};
    (body.files || []).forEach((f) => { const file = forms.fileFromDataUrl(f.name, f.dataUrl); if (file) (files[f.question] = files[f.question] || []).push(file); });
    let out;
    try { out = forms.submit(runner, account, form, body.answers || {}, files); } catch (err) { return json(res, { ok: false, error: err.message }); }
    saveSoon();
    const t = out.trigger ? runSummary(out.trigger) : null;
    return json(res, { ok: true, row: out.row, html: '<p class="sbx-ui"><b>Response recorded</b> on row ' + out.row + ' of ' + pages.h(form.tab) + '. <a href="/sandbox/sheets/' + pages.h(form.sheet) + '/edit">See the sheet</a></p>' + (t ? pages.resultBlock(t) : '') });
  }

  if (p === '/sandbox/fabman') return send(res, 200, pages.fabmanPage(ctx()));
  if (p === '/sandbox/code') return send(res, 200, pages.codeList(ctx()));
  if ((m = /^\/sandbox\/code\/(\w+)$/.exec(p)) && projects[m[1]]) return send(res, 200, pages.codePage(ctx(), m[1]));
  if (p === '/sandbox/leaving') return send(res, 200, pages.leavingPage(ctx(), q.get('to') || ''));

  send(res, 404, pages.layout('Not found', '<h1>Not found</h1><p>The sandbox has no page at <code>' + pages.h(p) + '</code>.</p><p><a href="/sandbox/">Control panel</a></p>', ctx()));
}

// ---------------------------------------------------------------- start

const how = boot();
const server = http.createServer((req, res) => {
  route(req, res).catch((err) => {
    console.error(err);
    if (!res.headersSent) send(res, 500, pages.gasError('The sandbox itself hit an error: ' + err.message + '\n\n' + err.stack), 'text/html; charset=utf-8');
  });
});
server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.log('\n  Port ' + PORT + ' is already in use. The sandbox may already be running: http://localhost:' + PORT + '/sandbox/');
    console.log('  To run a second copy, pick another port: node sandbox/server.js --port ' + (PORT + 1) + '\n');
    process.exit(0);
  }
  throw err;
});
server.listen(PORT, HOST, () => {
  console.log('\n  Ops Hub sandbox is running with ' + how + '.');
  console.log('  Control panel:  http://localhost:' + PORT + '/sandbox/');
  console.log('  Staff page:     http://localhost:' + PORT + '/index.html');
  console.log('  Admin page:     http://localhost:' + PORT + '/admin.html\n');
  if (HOST === '0.0.0.0' || HOST === '::') {
    const lan = [].concat(...Object.values(require('os').networkInterfaces())).filter((n) => n && n.family === 'IPv4' && !n.internal).map((n) => n.address);
    if (lan.length) console.log('  From a phone on the same network: ' + lan.map((a) => 'http://' + a + ':' + PORT + '/sandbox/').join('  or  ') + '\n');
  }
  console.log('  Edit any file and refresh. Stop with Ctrl+C. Data is kept in sandbox/.state/.\n');
});
process.on('SIGINT', () => { if (account && account.dirty) saveNow(); process.exit(0); });
process.on('SIGTERM', () => { if (account && account.dirty) saveNow(); process.exit(0); });
