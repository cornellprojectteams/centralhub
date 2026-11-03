#!/usr/bin/env node
'use strict';
/*
 * Self-test for the sandbox. Builds the fake data, runs the main paths of every Apps
 * Script project, then starts a throwaway server and requests the main pages.
 * Nothing it does touches your saved sandbox data (sandbox/.state).
 *
 *   node sandbox/check.js
 *
 * Run it after changing the sandbox itself, or when the sandbox misbehaves and you want
 * to know whether the problem is in your change or in the sandbox.
 */

process.env.TZ = process.env.SANDBOX_TZ || 'America/New_York';

const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const { Account } = require('./lib/gas');
const { Runner, defineProjects } = require('./lib/projects');
const site = require('./lib/site');
const fixtures = require('./lib/fixtures');
const store = require('./lib/store');
const forms = require('./lib/forms');
const { makeRewriter } = require('./lib/rewrite');

const REPO = path.resolve(__dirname, '..');
let passed = 0, failed = 0;
function check(name, fn) {
  try {
    const note = fn();
    passed++;
    console.log('  ok    ' + name + (note ? '  (' + note + ')' : ''));
  } catch (err) {
    failed++;
    console.log('  FAIL  ' + name + '\n        ' + String(err && err.message || err).split('\n').join('\n        '));
  }
}
function must(cond, msg) { if (!cond) throw new Error(msg); }
function ran(r, what) { must(r.ok, (what || r.run.fn) + ' failed: ' + r.error); return r.value; }

console.log('\nOps Hub sandbox self-test\n');

// ---------------------------------------------------------------- the scripts and the fake data

const cfg = site.readConfig(REPO);
const peek = new Runner(REPO, new Account(), defineProjects({}, {}));
const ids = site.readIds(cfg, (k, e) => peek.peek(k, e));
const projects = defineProjects(ids, { space: cfg.spaceUrl, inventory: cfg.inventoryUrl, command: cfg.statsUrl });
let account, runner;

check('config.js and app.js are readable', () => {
  must(cfg.spaceUrl && cfg.inventoryUrl && cfg.statsUrl, 'missing a web app URL (SPACE_STATUS_URL, INVENTORY.apiUrl or STATS_ENDPOINT)');
  return 'three web app URLs found';
});

check('every Apps Script file compiles', () => {
  const files = new Set();
  Object.values(projects).forEach((p) => p.files.forEach((f) => files.add(f)));
  files.forEach((f) => peek.script(f));
  return files.size + ' files';
});

check('the scripts point at the sheets the sandbox has', () => {
  must(ids.space === peek.peek('space', 'CONFIG.spreadsheetId'), 'CONFIG.spreadsheetId could not be read from 02_notify_on_submit.gs');
  return 'Space Status ' + ids.space.slice(0, 8) + '...';
});

check('fake data builds', () => {
  const t = Date.now();
  const built = fixtures.build(ids, (acc) => new Runner(REPO, acc, projects));
  account = built.account; runner = built.runner;
  must(Object.keys(account.spreadsheets).length >= 9, 'expected at least 9 spreadsheets');
  must(account.outbox.length === 3, 'expected the 3 fresh reports to send 3 emails, got ' + account.outbox.length);
  return Object.keys(account.spreadsheets).length + ' spreadsheets, ' + Object.keys(account.drive.files).length + ' Drive files, ' + (Date.now() - t) + ' ms';
});

if (!account) { console.log('\nThe fake data did not build, so the rest cannot run.\n'); process.exit(1); }

const web = (key, params) => {
  const q = new URLSearchParams(params || {});
  return ran(runner.exec(key, 'doGet', [Runner.webEvent(q)]), key + ' doGet ' + q.toString());
};
const team = () => account.ss(ids.space).getSheetByName('Team Contacts').rows()[2];

check('Space Status pages render', () => {
  const pages = [{}, { view: 'all', admin: '1', embed: '1' }, { view: 'Baja', act: '1' }, { view: 'Operations Team', act: '1', embed: '1' },
    { module: 'projects' }, { module: 'projects', admin: '1' }, { module: 'projects-dash' }, { module: 'proposals' }, { module: 'proposals', admin: '1' }, { team: team()[7] }];
  pages.forEach((p) => { const out = web('space', p); must(out && out.isHtmlOutput && out.html.length > 500, 'empty page for ' + JSON.stringify(p)); });
  return pages.length + ' pages';
});

check('a report sends the team an email', () => {
  const form = forms.definitions(account, ids)['space-status'];
  const before = account.outbox.length;
  const out = forms.submit(runner, account, form, { NetID: 'chk01', 'Responsible Team': 'Baja', 'Respond within': '7 to 10 days', 'Issue Type': 'Housekeeping', 'Required action': 'Clean up and organize the area', 'Details / comments': 'Self-test report' }, {});
  must(out.trigger && out.trigger.ok, 'sendEmails failed: ' + (out.trigger && out.trigger.error));
  must(account.outbox.length === before + 1, 'no email in the outbox');
  const mail = account.outbox[account.outbox.length - 1];
  must(/baja/i.test(mail.to), 'the email did not go to Baja: ' + mail.to);
  return mail.subject;
});

check('the email link opens the mark-complete page', () => {
  const resp = account.ss(ids.space).getSheetByName('Form Responses').rows();
  const H = resp[0].map((h) => String(h).toLowerCase());
  const token = resp[resp.length - 1][H.indexOf('issue token')];
  must(token, 'the script did not stamp an issue token');
  const out = web('space', { id: token });
  must(/Mark complete/.test(out.html), 'no Mark complete page');
  const done = ran(runner.exec('space', 'submitIssueCompletion', [token, 'data:image/png;base64,iVBORw0KGgo=', 'done.png', false, 'Self-test'], { client: true }));
  must(done && done.ok, 'submitIssueCompletion: ' + JSON.stringify(done));
  const ok = ran(runner.exec('space', 'approveIssueCompletion', [token, ''], { client: true }));
  must(ok && ok.ok, 'approveIssueCompletion: ' + JSON.stringify(ok));
  return 'submitted and approved';
});

check('the daily reminder job runs', () => { ran(runner.exec('space', 'sendReminders', [])); return account.outbox.length + ' emails in the outbox'; });

check('projects and proposals work end to end', () => {
  const c = ran(runner.exec('space', 'tpCreateProject', ['Self-test project', 'Check', '', '', '', '', '', '', ''], { client: true }));
  must(c.ok, 'create: ' + c.error);
  must(ran(runner.exec('space', 'tpJoinProject', [c.project.id, 'Tester'], { client: true })).ok, 'join failed');
  const props = account.ss(ids.space).getSheetByName('Proposals').rows().filter((r) => r[6] === 'Provisional');
  must(props.length >= 2, 'expected provisional proposals');
  const v = ran(runner.exec('space', 'prVote', [props[0][0], 'chk99', 4, 2, ''], { client: true }));
  must(v.ok, 'vote failed: ' + v.error);
  return 'created, joined, voted';
});

check('the Command Center API answers', () => {
  const out = web('command', { fresh: '1' });
  const data = JSON.parse(out.getContent());
  must(data && typeof data === 'object', 'not JSON');
  return Object.keys(data).slice(0, 4).join(', ') + '...';
});

check('the inventory API answers', () => {
  const out = web('inventory', { action: 'list' });
  const data = JSON.parse(out.getContent());
  must(data.ok !== false && (data.items || []).length > 10, 'no items: ' + out.getContent().slice(0, 200));
  return data.items.length + ' items';
});

check('the Fabman script reaches the fake Fabman', () => {
  ran(runner.exec('fabman', 'fetchFabmanMembers', []));
  must(account.fabman.calls.length > 0, 'no calls reached the fake Fabman');
  const rows = account.ss(ids.fabman).getSheetByName('Dashboard').getLastRow() - 1;
  must(rows === account.fabman.members.length, 'Dashboard has ' + rows + ' rows for ' + account.fabman.members.length + ' members');
  return account.fabman.calls.length + ' requests, ' + rows + ' members listed';
});

check('scripts cannot reach the internet', () => {
  const svc = require('./lib/gas').makeServices(account, projects.command, { logs: [], ui: [], menus: [], mails: [] });
  let blocked = false;
  try { svc.UrlFetchApp.fetch('https://example.com/'); } catch (e) { blocked = /does not reach the internet/.test(e.message); }
  must(blocked, 'UrlFetchApp reached an outside address');
  return 'outside requests are stopped';
});

check('the fleet sheet menu works', () => {
  const r = runner.fire(ids.drivers, 'open', { source: account.ss(ids.drivers) });
  const menus = r.flatMap((x) => x.run.menus || []);
  must(menus.length, 'onOpen added no menu');
  const first = menus[0].items.find((i) => i.fn);
  ran(runner.exec('fleet', first.fn, []));
  return menus[0].name + ' > ' + first.label;
});

check('no script errors along the way', () => {
  must(!runner.problems.length, runner.problems.map((p) => p.project + '.' + p.fn + ': ' + p.error).join('\n'));
});

check('saved data loads back the same', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sbx-check-')), 'state.json');
  store.save(file, account);
  const back = store.load(file);
  must(back, 'could not load the saved file');
  ['space', 'inventory', 'fabman', 'drivers'].forEach((k) => {
    const a = account.ss(ids[k]), b = back.ss(ids[k]);
    a.getSheets().forEach((sh) => must(JSON.stringify(sh.rows()) === JSON.stringify(b.getSheetByName(sh.getName()).rows()), 'tab ' + sh.getName() + ' changed after save and load'));
  });
  must(back.outbox.length === account.outbox.length, 'outbox changed');
  fs.rmSync(path.dirname(file), { recursive: true, force: true });
  return 'sheets, Drive and outbox match';
});

check('real addresses are swapped in the pages', () => {
  const rw = makeRewriter(cfg);
  ['index.html', 'admin.html', 'config.js', 'app.js', 'team.html', 'command-center.html'].forEach((f) => {
    const out = rw.files(fs.readFileSync(path.join(REPO, f), 'utf8'));
    [cfg.spaceUrl, cfg.inventoryUrl, cfg.statsUrl].forEach((u) => must(out.indexOf(u) < 0, f + ' still has ' + u));
  });
  return 'web app URLs replaced';
});

check('the Google copy kit makes safe copies that run', () => {
  const { makeCopies } = require('./google/make-copies');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sbx-gcopy-'));
  try {
    const me = 'new.hire@sandbox.test';
    const made = makeCopies({ me, out: dir });
    // A Google account with none of the real sheets in it: if a copy reached for one, it would fail.
    const acc = new Account();
    acc.user = me;
    const gp = { gcopy: { key: 'gcopy', name: 'Google copy', files: made.files.map((f) => path.join(dir, f)), prodUrl: 'https://script.google.com/macros/s/YOURCOPY/exec' } };
    const gr = new Runner(REPO, acc, gp);
    ran(gr.exec('gcopy', 'sandboxSetup', []), 'sandboxSetup');
    const sheetId = acc.projectState('gcopy').props.SANDBOX_SHEET_ID;
    must(sheetId && acc.ss(sheetId), 'sandboxSetup made no sheet');
    const tabs = acc.ss(sheetId).getSheets().map((s) => s.getName());
    ['Form Responses', 'Team Contacts', 'Projects', 'Proposals'].forEach((t) => must(tabs.indexOf(t) >= 0, 'no ' + t + ' tab (got ' + tabs.join(', ') + ')'));
    must(acc.projectState('gcopy').triggers.some((t) => t.fn === 'sendEmails' && t.type === 'submit'), 'no form-submit trigger');
    ran(gr.exec('gcopy', 'sandboxSubmitTestReport', []), 'sandboxSubmitTestReport');
    must(acc.outbox.length === 1, 'expected one email, got ' + acc.outbox.length);
    acc.outbox.forEach((m) => {
      must(m.to === me && !m.cc && !m.bcc, 'an email went to ' + [m.to, m.cc, m.bcc].join(' '));
      must(/^\[SANDBOX\]/.test(m.subject), 'subject not marked: ' + m.subject);
      must(/YOURCOPY\/exec\?id=/.test(m.html), 'the email button does not open the copy');
    });
    [{ view: 'all', admin: '1' }, { module: 'projects', admin: '1' }, { module: 'proposals' }].forEach((p) => {
      const out = ran(gr.exec('gcopy', 'doGet', [Runner.webEvent(new URLSearchParams(p))]), 'doGet ' + JSON.stringify(p));
      must(out && out.isHtmlOutput, 'no page for ' + JSON.stringify(p));
    });
    must(!gr.problems.length, gr.problems.map((p) => p.fn + ': ' + p.error).join('\n'));
    return made.files.length + ' files; setup, test report and pages work; mail only to the owner';
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

// ---------------------------------------------------------------- the server, over HTTP

function freePort() {
  return new Promise((resolve) => { const s = require('net').createServer().listen(0, '127.0.0.1', () => { const p = s.address().port; s.close(() => resolve(p)); }); });
}

(async () => {
  const port = await freePort();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sbx-check-'));
  const child = spawn(process.execPath, [path.join(__dirname, 'server.js'), '--fresh', '--port', String(port), '--state', path.join(dir, 'state.json')], { stdio: ['ignore', 'pipe', 'pipe'] });
  let log = '';
  child.stdout.on('data', (d) => { log += d; });
  child.stderr.on('data', (d) => { log += d; });
  const up = await new Promise((resolve) => {
    const end = Date.now() + 15000;
    (function poll() {
      if (/is running/.test(log)) return resolve(true);
      if (Date.now() > end || child.exitCode !== null) return resolve(false);
      setTimeout(poll, 100);
    })();
  });
  const results = [];
  if (up) {
    const pages = [['/sandbox/', 200], ['/index.html', 200], ['/admin.html', 200], ['/developer.html', 200], ['/inventory.html', 200], ['/command-center.html', 200],
      ['/sandbox/sheets/', 200], ['/sandbox/sheets/' + ids.space + '/edit', 200], ['/sandbox/outbox', 200], ['/sandbox/code', 200], ['/sandbox/code/space', 200],
      ['/sandbox/fabman', 200], ['/sandbox/forms/space-status', 200], ['/sandbox/readme', 200], ['/sandbox/space/exec?view=all&admin=1', 200],
      ['/sandbox/inventory/exec?action=list', 200], ['/sandbox/command/exec', 200], ['/apps-script/05_fabman_members.gs', 404], ['/downloads/', 404], ['/sandbox/.state/state.json', 404]];
    for (const [p, want] of pages) {
      try { const r = await fetch('http://127.0.0.1:' + port + p); results.push([p, r.status, want]); } catch (e) { results.push([p, 'no answer', want]); }
    }
  }
  child.kill('SIGTERM');
  fs.rmSync(dir, { recursive: true, force: true });
  check('the server starts', () => { must(up, 'the server did not start:\n' + log.trim()); return 'port ' + port; });
  if (up) check('the main pages answer', () => {
    const bad = results.filter((r) => r[1] !== r[2]);
    must(!bad.length, bad.map((r) => r[0] + ' gave ' + r[1] + ', expected ' + r[2]).join('\n'));
    return results.length + ' addresses, private files refused';
  });
  console.log('\n' + (failed ? failed + ' failed, ' + passed + ' passed.' : 'All ' + passed + ' checks passed.') + '\n');
  process.exit(failed ? 1 : 0);
})();
