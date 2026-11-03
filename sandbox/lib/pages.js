'use strict';
// The sandbox's own pages: the control panel, the sheet viewer, the outbox, the script
// runner, the forms and the fake Fabman. Plain HTML built on the server; the buttons on
// them are wired up in public/sandbox.js.

const { colLetters, displayValue } = require('./gas');

const h = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// Links to the real thing go through the "leaving the sandbox" page first.
const leave = (url) => '/sandbox/leaving?to=' + encodeURIComponent(url);
const when = (iso) => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }); };

function layout(title, body, ctx, opts) {
  opts = opts || {};
  const n = ctx.account.outbox.length;
  const nav = [
    ['/sandbox/', 'Control panel', 'panel'], ['/index.html', 'Staff page'], ['/admin.html', 'Admin'],
    ['/sandbox/outbox', 'Outbox' + (n ? ' <b>' + n + '</b>' : ''), 'outbox'], ['/sandbox/sheets/', 'Sheets', 'sheets'], ['/sandbox/code', 'Scripts', 'code'],
  ].map(([href, label, key]) => '<a href="' + href + '"' + (key && key === opts.nav ? ' class="is-on" aria-current="page"' : '') + '>' + label + '</a>').join('');
  return '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<title>' + h(title) + ' | Ops Hub sandbox</title><link rel="icon" href="/assets/favicon.svg" type="image/svg+xml">'
    + '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
    + '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap">'
    + '<link rel="stylesheet" href="/sandbox/static/sandbox.css"></head><body class="sbx">'
    + '<header class="sbx-head"><div class="sbx-head-in">'
    + '<a class="sbx-brand" href="/sandbox/"><img src="/assets/cornell-seal.svg" alt="" width="40" height="40"><span><strong>Ops Hub sandbox</strong><small>Fake data. Nothing leaves this computer.</small></span></a>'
    + '<nav class="sbx-nav" aria-label="Sandbox">' + nav + '</nav></div></header>'
    + '<main class="sbx-main' + (opts.wide ? ' is-wide' : '') + '">' + body + '</main>'
    + '<footer class="sbx-foot"><span>Sandbox data was ' + (ctx.account.builtAt ? 'set up ' + h(when(ctx.account.builtAt)) : 'loaded') + (ctx.account.agedDays ? ', then aged ' + ctx.account.agedDays + ' day' + (ctx.account.agedDays === 1 ? '' : 's') : '') + '.</span>'
    + '<span>Edit any file in the repo and refresh to see it. <a href="/sandbox/readme">How the sandbox works</a></span></footer>'
    + '<script src="/sandbox/static/sandbox.js"></script></body></html>';
}

// ---------------------------------------------------------------- control panel

const TRY = [
  ['report', 'File a test report as a staff member, then find the team’s email in the <a href="/sandbox/outbox">Outbox</a>.', '/sandbox/forms/space-status', 'Open the form'],
  ['complete', 'Press the button in that email and finish the job with a photo (any picture will do). Then approve it on the Admin page, under Open issues dashboard.', '/admin.html', 'Admin page'],
  ['sendback', 'Send a pending job back with a reason, and read the email the team gets.', '/admin.html', 'Admin page'],
  ['day', 'Skip ahead a day (button below) and see which reminder emails go out.', '', ''],
  ['project', 'Create a project on the Admin page, join it from the staff page (Your work, Projects tab), finish it with an after photo, then approve it.', '/admin.html', 'Admin page'],
  ['proposal', 'Rate three proposals from the staff page (Propose an improvement) with any made-up NetID. Give the cord reels a 3 or more for impact: your third rating counts and pushes it into Projects.', '/index.html', 'Staff page'],
  ['inventory', 'Add a tool in the inventory, print its QR label, and log maintenance on it.', '/inventory.html?admin=1', 'Inventory (admin)'],
  ['command', 'Open the Command Center and watch your changes show up (use its refresh button).', '/command-center.html', 'Command Center'],
  ['sheet', 'Open the Space Status sheet, clear a closed job’s Addressed at cell, and see the job come back on the dashboard.', '/sandbox/sheets/{space}', 'Space Status sheet'],
  ['scripts', 'Run Driver cleanup on the drivers sheet, then Flag graduates and Process graduates in the Fabman sheet. Check the fake Fabman afterwards.', '/sandbox/sheets/{drivers}', 'Drivers sheet'],
  ['signups', 'Seven FabMan Agreement sign-ups are waiting without accounts. In the Fabman sheet, use the Fabman sign-ups menu: Check sign-ups, then Process sign-ups since a date. Then sign the agreement yourself (Make something happen) and watch the row and its Fabman status appear on the FabMan Agreement tab.', '/sandbox/sheets/{fabman}', 'Fabman sheet'],
  ['outage', 'On the Fake Fabman page, make Fabman fail or refuse the key, sign the agreement again, and see the sign-up wait and then finish once Fabman is back.', '/sandbox/fabman', 'Fake Fabman'],
  ['change', 'Change something: a color in styles.css, an email subject in SEVERITY (apps-script/02_notify_on_submit.gs) or a tile on admin.html. Refresh to see it.', '', ''],
  ['break', 'Break a script on purpose (delete a bracket in a .gs file), see how the sandbox reports it, then put it back.', '/sandbox/code', 'Scripts'],
  ['reset', 'Reset the sandbox whenever you want a clean start.', '', ''],
];

function panel(ctx) {
  const s = ctx.status();
  const ids = ctx.ids;
  const stat = (n, label, href) => '<a class="sbx-stat" href="' + href + '"><b>' + n + '</b><span>' + label + '</span></a>';
  const card = (href, title, sub) => '<a class="sbx-card" href="' + href + '"><strong>' + title + '</strong><span>' + sub + '</span></a>';
  const problems = ctx.runner.problems.slice(-6).reverse();
  return layout('Control panel', ''
    + '<section class="sbx-hero"><div><p class="sbx-kicker">Sandbox</p><h1>The whole hub, with nothing real in it</h1>'
    + '<p class="sbx-lede">Everything here runs on this computer: the site, the Space Status web app, the inventory, the Command Center and the sheet scripts, all on made-up data. '
    + 'Emails land in the <a href="/sandbox/outbox">outbox</a> instead of being sent, and nothing reaches Google, Fabman or a real person. Click around, break things, reset.</p></div>'
    + '<div class="sbx-stats">' + stat(s.open, 'open issues', '/admin.html') + stat(s.pending, 'waiting for approval', '/admin.html') + stat(s.projects, 'active projects', '/admin.html')
    + stat(s.proposals, 'proposals up for rating', '/index.html') + stat(s.items, 'inventory items', '/inventory.html?admin=1') + stat(s.mail, 'emails in the outbox', '/sandbox/outbox') + '</div></section>'

    + '<section class="sbx-sec"><h2>Open the hub</h2><div class="sbx-cards">'
    + card('/index.html', 'Staff page', 'Your work, proposals, shift log, inventory, schedule')
    + card('/admin.html', 'Admin page', 'Dashboards and links, admin controls on')
    + card('/developer.html', 'Developer page', 'Script projects and the repo')
    + card('/command-center.html', 'Command Center', 'What needs attention, from every sheet')
    + card('/inventory.html', 'Inventory', 'The staff view of the shop board')
    + card('/inventory.html?admin=1', 'Inventory (admin)', 'Add, edit, photos, maintenance, counts')
    + card('/team.html?team=Baja', 'A team’s issue page', 'What Baja sees')
    + card('/docs.html', 'Admin guide', 'How the admin side is used')
    + card('/dev-guide.html', 'Developer guide', 'Editing, publishing, how the scripts work')
    + '</div></section>'

    + '<section class="sbx-sec"><h2>Make something happen</h2><div class="sbx-cards">'
    + card('/sandbox/forms/space-status', 'File a space report', 'The report form; the team gets an email')
    + card('/sandbox/forms/ell', 'Log an ELL shift', 'Feeds the Command Center’s ELL tile')
    + card('/sandbox/forms/fabman-agreement', 'Sign the FabMan Agreement', 'The Qualtrics survey: a Fabman account within seconds')
    + card('/sandbox/forms/fabman-signup', 'Old FabMan Agreement form', 'The Google Form it replaces, still handled')
    + '</div><div class="sbx-actions">'
    + '<button type="button" class="sbx-btn" data-api="/sandbox/api/reminders" data-out="#act-out">Run today’s reminder job</button>'
    + '<button type="button" class="sbx-btn" data-api="/sandbox/api/day" data-out="#act-out">Skip ahead a day</button>'
    + '<span class="sbx-hint">Skipping ahead moves every date in the sandbox back a day, then runs the daily reminder job, as if a day had passed.</span>'
    + '</div><div id="act-out" class="sbx-result" hidden></div></section>'

    + '<section class="sbx-sec"><h2>Behind the scenes</h2><div class="sbx-cards">'
    + card('/sandbox/outbox', 'Outbox', 'Every email the scripts tried to send')
    + card('/sandbox/sheets/', 'Sheets', 'The fake spreadsheets; edit cells, run menus')
    + card('/sandbox/code', 'Scripts', 'Run any Apps Script function, see logs and triggers')
    + card('/sandbox/fabman', 'Fake Fabman', 'Members and trainings the Fabman script changes')
    + '</div></section>'

    + '<section class="sbx-sec"><h2>Things to try</h2><p class="sbx-note">A tour of the platform. Ticks are kept in this browser.</p><ol class="sbx-try">'
    + TRY.map(([key, text, href, label]) => '<li><label><input type="checkbox" data-try="' + key + '"><span>' + text + '</span></label>'
      + (href ? ' <a class="sbx-try-go" href="' + href.replace('{space}', ids.space).replace('{drivers}', ids.drivers).replace('{fabman}', ids.fabman) + '">' + label + ' &rarr;</a>' : '') + '</li>').join('')
    + '</ol></section>'

    + (problems.length ? '<section class="sbx-sec"><h2>Recent script errors</h2><ul class="sbx-problems">'
      + problems.map((p) => '<li><span>' + h(when(p.at)) + ' &middot; ' + h(p.project) + ' &middot; ' + h(p.fn) + '</span><code>' + h(p.error) + '</code></li>').join('') + '</ul></section>' : '')

    + '<section class="sbx-sec sbx-danger"><h2>Start over</h2><p>Put every sheet, photo, email and fake Fabman member back the way the sandbox started. Your code is not touched.</p>'
    + '<button type="button" class="sbx-btn sbx-btn--danger" data-api="/sandbox/api/reset" data-confirm="Press again to reset everything" data-reload="1">Reset the sandbox</button></section>',
  ctx, { nav: 'panel' });
}

// ---------------------------------------------------------------- run results (shared by several pages)

function resultBlock(r) {
  if (!r) return '';
  const out = [];
  out.push('<p class="sbx-run-head ' + (r.ok ? 'is-ok' : 'is-bad') + '">' + (r.ok ? 'Ran' : 'Failed') + ' <code>' + h(r.fn || '') + '</code>' + (r.project ? ' in ' + h(r.project) : '') + (r.ms !== undefined ? ' <span>' + r.ms + ' ms</span>' : '') + '</p>');
  if (!r.ok && r.error) out.push('<pre class="sbx-err">' + h(r.error) + '</pre>');
  (r.ui || []).forEach((u) => {
    if (u.kind === 'toast') out.push('<p class="sbx-ui"><b>Toast</b> ' + h(u.title ? u.title + ': ' : '') + h(u.text) + '</p>');
    if (u.kind === 'alert') out.push('<div class="sbx-ui"><b>Dialog' + (u.title ? ': ' + h(u.title) : '') + '</b><p>' + h(u.text).replace(/\n/g, '<br>') + '</p><span>Answered ' + h(u.answer) + '</span></div>');
    if (u.kind === 'prompt') out.push('<div class="sbx-ui"><b>Prompt</b><p>' + h(u.text) + '</p><span>Answered ' + h(u.answer) + '</span></div>');
    if (u.kind === 'dialog') out.push('<div class="sbx-ui"><b>' + h(u.title || 'Dialog') + '</b><iframe class="sbx-dialog" sandbox="" srcdoc="' + h(u.html) + '"></iframe></div>');
  });
  if (r.mails && r.mails.length) out.push('<p class="sbx-ui"><b>Emails</b> ' + r.mails.map((m) => '<a href="/sandbox/outbox/' + h(m.id) + '">' + h(m.subject || '(no subject)') + '</a>').join(', ') + '</p>');
  if (r.logs && r.logs.length) out.push('<pre class="sbx-log">' + h(r.logs.join('\n')) + '</pre>');
  return out.join('');
}

// ---------------------------------------------------------------- sheets

function sheetsList(ctx) {
  const rows = Object.values(ctx.account.spreadsheets).map((ss) => {
    const bound = Object.values(ctx.projects).filter((p) => p.boundId === ss.id).map((p) => p.name);
    const users = Object.values(ctx.projects).filter((p) => p.dataSheet === ss.id || (ctx.sources[ss.id] || []).length).length;
    return '<tr><td><a href="/sandbox/sheets/' + h(ss.id) + '/edit">' + h(ss.getName()) + '</a></td><td>' + ss.getSheets().length + '</td><td>' + h(bound.join(', ') || '') + '</td><td>' + h((ctx.sources[ss.id] || []).join(', ')) + '</td></tr>';
  }).join('');
  return layout('Sheets', '<h1>Sheets</h1><p class="sbx-lede">Fake copies of every spreadsheet the platform uses, under the same ids as the real ones, so the scripts open them without a code change. Click a cell to change it; the sheet’s triggers run the way they would on Google.</p>'
    + '<div class="sbx-scroll"><table class="sbx-table"><thead><tr><th>Spreadsheet</th><th>Tabs</th><th>Script bound to it</th><th>Read by</th></tr></thead><tbody>' + rows + '</tbody></table></div>', ctx, { nav: 'sheets' });
}

function sheetPage(ctx, ss, gid, menus) {
  const sheets = ss.getSheets();
  const sh = (gid && ss.getSheetById(gid)) || ss.getActiveSheet() || sheets[0];
  const lr = Math.max(sh.getLastRow(), 1), lc = Math.max(sh.getLastColumn(), 1);
  const rows = Math.min(lr + 3, sh.getMaxRows()), cols = Math.min(lc + 1, sh.getMaxColumns());
  const project = Object.values(ctx.projects).find((p) => p.boundId === ss.id);
  const buttons = project && project.buttons && project.buttons[sh.getName()];
  const realUrl = 'https://docs.google.com/spreadsheets/d/' + ss.id + '/edit';
  let grid = '<table class="sbx-grid" data-ss="' + h(ss.id) + '" data-sheet="' + h(sh.getName()) + '"><thead><tr><th></th>';
  for (let c = 1; c <= cols; c++) grid += '<th>' + colLetters(c) + '</th>';
  grid += '</tr></thead><tbody>';
  for (let r = 1; r <= rows; r++) {
    grid += '<tr><th>' + r + (r > 1 && r <= lr ? '<button type="button" class="sbx-del" data-row="' + r + '" title="Delete row ' + r + '" aria-label="Delete row ' + r + '">&times;</button>' : '') + '</th>';
    for (let c = 1; c <= cols; c++) {
      const cl = sh.cells.get(r + ',' + c) || {};
      const v = sh.valueAt(r, c);
      const check = cl.dv && cl.dv.calls && cl.dv.calls.some((x) => x[0] === 'requireCheckbox');
      const style = (cl.setBackground ? 'background:' + cl.setBackground + ';' : '') + (cl.setFontColor ? 'color:' + cl.setFontColor + ';' : '') + (cl.setFontWeight === 'bold' ? 'font-weight:700;' : '');
      const attrs = ' data-r="' + r + '" data-c="' + c + '"' + (style ? ' style="' + h(style) + '"' : '');
      if (check) { grid += '<td class="is-check"' + attrs + '><input type="checkbox"' + (v === true ? ' checked' : '') + ' aria-label="' + colLetters(c) + r + '"></td>'; continue; }
      if (cl.f) { grid += '<td class="is-formula"' + attrs + ' title="' + h(cl.f) + '">' + h(v !== '' ? displayValue(v) : cl.f) + '</td>'; continue; }
      const text = displayValue(v);
      grid += '<td' + attrs + (v instanceof Date ? ' data-type="date"' : typeof v === 'number' ? ' data-type="number"' : typeof v === 'boolean' ? ' data-type="bool"' : '') + (cl.note ? ' data-note="' + h(cl.note) + '"' : '') + '>' + h(text) + '</td>';
    }
    grid += '</tr>';
  }
  grid += '</tbody></table>';
  const tabs = sheets.map((s) => '<a href="?gid=' + s.getSheetId() + '"' + (s === sh ? ' class="is-on"' : '') + '>' + h(s.getName()) + (s.isSheetHidden() ? ' <i>hidden</i>' : '') + '</a>').join('');
  const menuHtml = (menus || []).map((m) => '<details class="sbx-menu"><summary>' + h(m.name) + '</summary><div>'
    + m.items.map((it) => it.sep ? '<hr>' : '<button type="button" data-run="' + h(m.project) + '" data-fn="' + h(it.fn) + '">' + h(it.label) + '</button>').join('') + '</div></details>').join('');
  // The real sheet runs these from drawings on a tab; here they are chips above the grid.
  const elsewhere = project && project.buttons ? Object.keys(project.buttons).filter((t) => t !== sh.getName() && ss.getSheetByName(t)) : [];
  const buttonHtml = (buttons || elsewhere.length) ? '<div class="sbx-sheet-buttons">'
    + (buttons ? '<span>Buttons on this tab:</span>' + buttons.map((fn) => '<button type="button" class="sbx-chip" data-run="' + h(project.key) + '" data-fn="' + h(fn) + '">' + h(fn) + '</button>').join('') : '')
    + (elsewhere.length ? '<span class="sbx-hint">' + (buttons ? 'More on ' : 'Buttons are on the ') + elsewhere.map((t) => '<a href="?gid=' + ss.getSheetByName(t).getSheetId() + '">' + h(t) + '</a>').join(', ') + (buttons ? '' : ' tab' + (elsewhere.length > 1 ? 's' : '')) + '.</span>' : '')
    + '</div>' : '';
  const users = Object.values(ctx.projects).filter((p) => p.dataSheet === ss.id && p !== project);
  const readers = (ctx.sources[ss.id] || []).filter((w) => !users.some((p) => p.name === w));
  const fakeOnly = /^sandbox-/.test(ss.id);
  return layout(ss.getName(), '<div class="sbx-sheet-head"><div><p class="sbx-kicker">Sandbox copy of a Google Sheet</p><h1>' + h(ss.getName()) + '</h1>'
    + '<p class="sbx-note">' + (project ? 'Bound script: <a href="/sandbox/code/' + h(project.key) + '">' + h(project.name) + '</a>. ' : '')
    + (users.length ? 'Used by ' + users.map((p) => '<a href="/sandbox/code/' + h(p.key) + '">' + h(p.name) + '</a>').join(', ') + '. ' : '')
    + (readers.length ? 'Read by ' + h(readers.join(', ')) + '. ' : '')
    + (fakeOnly ? 'Made up for the sandbox; there is no real one.' : 'The real one is <a href="' + h(leave(realUrl)) + '">here</a> (outside the sandbox).') + '</p></div>'
    + '<div class="sbx-sheet-tools">' + menuHtml + (menuHtml || buttons ? answerPicker() : '') + '</div></div>'
    + buttonHtml
    + '<nav class="sbx-tabs" aria-label="Tabs">' + tabs + '</nav>'
    + '<div class="sbx-grid-wrap">' + grid + '</div>'
    + '<p class="sbx-hint">Click a cell to change it (Enter saves, Esc cancels). A date is typed as 2026-10-21 or 10/21/2026 9:30. Formulas are shown, not calculated, except the inventory’s Status and Next Maintenance.</p>'
    + '<div id="run-out" class="sbx-result" hidden></div>', ctx, { nav: 'sheets', wide: true });
}

function answerPicker() {
  return '<div class="sbx-answer" role="group" aria-label="How to answer the script’s questions"><span>If a script asks, answer</span>'
    + '<label><input type="radio" name="sbx-answer" value="yes" checked> Yes / OK</label><label><input type="radio" name="sbx-answer" value="no"> No / Cancel</label>'
    + '<label>and type <input type="text" name="sbx-prompt" placeholder="into a prompt" autocomplete="off"></label></div>';
}

function notInSandbox(ctx, id) {
  const real = 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
  return layout('Not in the sandbox', '<h1>This sheet is not in the sandbox</h1><p class="sbx-lede">The page asked for a spreadsheet the sandbox has no copy of. '
    + 'Nothing in the platform’s code reads it, so there was no need for a fake one.</p>'
    + '<p><a class="sbx-btn" href="/sandbox/sheets/">See the sandbox’s sheets</a> <a class="sbx-btn sbx-btn--ghost" href="' + h(leave(real)) + '">Open the real one</a></p>', ctx, { nav: 'sheets' });
}

// ---------------------------------------------------------------- outbox

function outbox(ctx) {
  const mails = ctx.account.outbox.slice().reverse();
  const rows = mails.map((m) => '<tr><td>' + h(when(m.at)) + '</td><td><a href="/sandbox/outbox/' + h(m.id) + '">' + h(m.subject || '(no subject)') + '</a></td><td>' + h(m.to) + '</td><td>' + h(m.project) + '</td></tr>').join('');
  return layout('Outbox', '<div class="sbx-row"><div><h1>Outbox</h1><p class="sbx-lede">Every email the scripts sent, newest first. Nothing here was delivered. Real addresses that the scripts copy on every email (the always-CC in 02_notify_on_submit.gs) show up as they would in production.</p></div>'
    + (mails.length ? '<button type="button" class="sbx-btn sbx-btn--ghost" data-api="/sandbox/api/outbox/clear" data-confirm="Press again to empty the outbox" data-reload="1">Empty the outbox</button>' : '') + '</div>'
    + (mails.length ? '<div class="sbx-scroll"><table class="sbx-table"><thead><tr><th>Sent</th><th>Subject</th><th>To</th><th>From script</th></tr></thead><tbody>' + rows + '</tbody></table></div>' : '<p class="sbx-empty">No emails yet. <a href="/sandbox/forms/space-status">File a report</a> to send one.</p>'), ctx, { nav: 'outbox' });
}

function mailPage(ctx, m) {
  const meta = [['From', m.from], ['To', m.to], ['Cc', m.cc], ['Bcc', m.bcc], ['Reply-To', m.replyTo], ['Sent', when(m.at) + ' by ' + m.project + ' (' + m.kind + 'App)']].filter((x) => x[1])
    .map(([k, v]) => '<dt>' + k + '</dt><dd>' + h(v) + '</dd>').join('');
  const body = m.html ? '<iframe class="sbx-mail-body" title="Email body" src="/sandbox/outbox/' + h(m.id) + '/body"></iframe>' : '<pre class="sbx-mail-text">' + h(m.body) + '</pre>';
  return layout(m.subject || 'Email', '<p><a href="/sandbox/outbox">&larr; Outbox</a></p><h1 class="sbx-mail-subject">' + h(m.subject || '(no subject)') + '</h1><dl class="sbx-meta">' + meta + '</dl>'
    + (m.attachments && m.attachments.length ? '<p class="sbx-note">Attachments: ' + h(m.attachments.join(', ')) + '</p>' : '') + body
    + '<p class="sbx-hint">Buttons and links in the email open the sandbox’s pages.</p>', ctx, { nav: 'outbox' });
}

// ---------------------------------------------------------------- Drive

function drivePage(ctx, f) {
  const isImg = /^image\//.test(f.mime);
  const folder = ctx.account.drive.folders[f.parentId];
  return layout(f.name, '<p class="sbx-kicker">Sandbox Drive file</p><h1>' + h(f.name) + '</h1>'
    + '<p class="sbx-note">' + h(f.mime) + ' &middot; ' + f.bytes.length.toLocaleString() + ' bytes' + (folder ? ' &middot; in ' + h(folder.name) : '') + (f.trashed ? ' &middot; <b>in the trash</b>' : '') + (f.access ? ' &middot; shared: ' + h(f.access) : '') + '</p>'
    + (isImg ? '<p><img class="sbx-drive-img" src="/sandbox/drive/raw/' + h(f.id) + '" alt="' + h(f.name) + '"></p>' : '<p><a class="sbx-btn" href="/sandbox/drive/raw/' + h(f.id) + '">Download</a></p>'), ctx);
}

// ---------------------------------------------------------------- forms

function formPage(ctx, form, editor) {
  const q = form.questions.map((qq, i) => {
    const name = 'q' + i, req = qq.required ? ' required' : '';
    let input = '';
    if (qq.kind === 'text') input = '<input type="text" name="' + name + '"' + req + '>';
    if (qq.kind === 'paragraph') input = '<textarea name="' + name + '" rows="3"' + req + '></textarea>';
    if (qq.kind === 'select') input = '<select name="' + name + '"' + req + '><option value="">Choose</option>' + qq.choices.map((c) => '<option>' + h(c) + '</option>').join('') + '</select>';
    if (qq.kind === 'radio') input = qq.choices.map((c, j) => '<label class="sbx-opt"><input type="radio" name="' + name + '" value="' + h(c) + '"' + (j === 0 ? req : '') + '> ' + h(c) + '</label>').join('');
    if (qq.kind === 'checks') input = qq.choices.map((c) => '<label class="sbx-opt"><input type="checkbox" name="' + name + '" value="' + h(c) + '"> ' + h(c) + '</label>').join('');
    if (qq.kind === 'file') input = '<input type="file" name="' + name + '" accept="image/*" multiple>';
    return '<fieldset class="sbx-q" data-title="' + h(qq.title) + '" data-kind="' + qq.kind + '"' + (qq.required ? ' data-required="1"' : '') + '><legend>' + h(qq.title) + (qq.required ? ' <i>*</i>' : '') + '</legend>' + input + '</fieldset>';
  }).join('');
  const editorNote = '<div class="sbx-callout"><b>Form editor stand-in.</b> The questions for this sandbox form are defined in <code>sandbox/lib/forms.js</code>. '
    + 'The scripts read answers by question title, so a renamed question has to be renamed in the script too.'
    + (form.key === 'space-status' ? ' Here the team list comes from the Team Contacts tab; the real form keeps its own list, which must match Team Contacts.' : '') + '</div>';
  return layout(form.title, '<form class="sbx-form" data-form="' + h(form.key) + '"><p class="sbx-kicker">Sandbox copy of a Google Form</p><h1>' + h(form.title) + '</h1><p class="sbx-lede">' + h(form.intro) + '</p>'
    + (editor ? editorNote : '')
    + (form.collectEmail ? '<fieldset class="sbx-q" data-title="Email Address" data-kind="text" data-required="1"><legend>Email <i>*</i></legend><input type="email" name="email" value="' + h(ctx.account.user) + '" required></fieldset>' : '')
    + q + '<div class="sbx-form-bar"><button type="submit" class="sbx-btn">Submit</button><span class="sbx-hint">Writes a row to ' + h(form.tab) + (form.trigger ? ' and runs ' + h(form.trigger.fn) + ' in the ' + h(form.trigger.project) + ' script' : '') + '.</span></div>'
    + '<div id="form-out" class="sbx-result" hidden></div></form>', ctx);
}

// ---------------------------------------------------------------- fake Fabman

function fabmanPage(ctx) {
  const f = ctx.account.fabman;
  const names = { 2946: 'Spaces', 2947: 'Lab', 2948: 'Machines', 2949: 'Mills', 2950: 'Supervisor' };
  const rows = f.members.map((m) => '<tr' + (m.state === 'locked' ? ' class="is-muted"' : '') + '><td>' + m.id + '</td><td>' + h(m.firstName + ' ' + m.lastName) + '</td><td>' + h(m.company) + '</td><td>' + h((m.metadata || {}).graduationYear) + '</td><td>' + h((m.metadata || {}).apron || '') + '</td><td>' + h(m.state) + '</td><td>'
    + h((f.trainings[m.id] || []).map((t) => names[t.trainingCourse] || t.trainingCourse).join(', ')) + '</td><td>' + h(m.emailAddress) + '</td></tr>').join('');
  const calls = f.calls.slice(-25).reverse().map((c) => '<li><span>' + h(when(c.at)) + '</span> <code>' + h(c.method.toUpperCase() + ' ' + c.path) + '</code></li>').join('');
  const sw = (key, on, what) => '<button type="button" class="sbx-btn' + (on ? '' : ' sbx-btn--ghost') + '" data-api="/sandbox/api/fabman" data-body=\'' + JSON.stringify({ [key]: !on }) + '\' data-reload="1">' + (on ? 'Stop: ' : '') + what + '</button>';
  return layout('Fake Fabman', '<p class="sbx-kicker">Stand-in for fabman.io</p><h1>Fake Fabman</h1><p class="sbx-lede">The Fabman script talks to this instead of the real Fabman, with the same requests. '
    + f.members.length + ' members, ' + f.invitations.length + ' invitations sent. The real admin site is <a href="' + h(leave('https://fabman.io/manage/2366')) + '">fabman.io</a>.</p>'
    + '<div class="sbx-actions">' + sw('down', f.down, 'Fail every request (an outage)') + sw('refuseKey', f.refuseKey, 'Refuse the API key') + '<span class="sbx-hint">'
    + (f.down || f.refuseKey ? '<b>Fabman is ' + (f.down ? 'down' : 'refusing the key') + ' right now.</b> Sign-ups wait with "Will retry" and finish on the next run after you stop it.' : 'To see how the scripts cope when Fabman is out or the key stops working.') + '</span></div>'
    + '<div class="sbx-grid-wrap"><table class="sbx-table"><thead><tr><th>ID</th><th>Name</th><th>Team</th><th>Grad</th><th>Apron</th><th>State</th><th>Trainings</th><th>Email</th></tr></thead><tbody>' + rows + '</tbody></table></div>'
    + '<h2>Latest requests</h2>' + (calls ? '<ul class="sbx-calls">' + calls + '</ul>' : '<p class="sbx-empty">None yet. Run a button in the <a href="/sandbox/sheets/' + h(ctx.ids.fabman) + '/edit">Fabman members sheet</a>.</p>'), ctx, { wide: true });
}

// ---------------------------------------------------------------- scripts

function codeList(ctx) {
  const cards = Object.values(ctx.projects).map((p) => '<a class="sbx-card" href="/sandbox/code/' + p.key + '"><strong>' + h(p.name) + '</strong><span>' + h(p.blurb) + '</span><code>' + p.files.map((f) => h(f.replace('apps-script/', ''))).join(' &middot; ') + '</code></a>').join('');
  return layout('Scripts', '<h1>Scripts</h1><p class="sbx-lede">The Apps Script projects, run straight from the repo’s <code>apps-script</code> folder. On Google you paste these files into a project; here, saving a file is enough.</p><div class="sbx-cards">' + cards + '</div>', ctx, { nav: 'code' });
}

function codePage(ctx, key) {
  const p = ctx.projects[key];
  const st = ctx.account.projectState(key);
  const fns = ctx.runner.functionsOf(key);
  let load = '';
  try { p.files.forEach((f) => ctx.runner.script(f)); } catch (err) { load = '<pre class="sbx-err">' + h(err.message) + '</pre>'; }
  const list = (arr) => arr.map((f) => '<li><button type="button" class="sbx-chip" data-run="' + key + '" data-fn="' + h(f.name) + '">' + h(f.name) + '</button><small>' + h(f.file.replace('apps-script/', '')) + '</small></li>').join('');
  const pub = fns.filter((f) => !f.private && !/^(doGet|doPost)$/.test(f.name)), priv = fns.filter((f) => f.private);
  const triggers = st.triggers.map((t) => '<li><code>' + h(t.fn) + '</code> on ' + h(t.type) + (t.every ? ' every ' + h(t.every) : '') + '</li>').join('');
  const props = Object.keys(st.props).map((k) => '<tr><td><code>' + h(k) + '</code></td><td>' + h(st.props[k]) + '</td></tr>').join('');
  const logs = st.logs.slice(-40).map((l) => h(when(l.at)) + '  ' + h(l.text)).join('\n');
  const sheet = p.boundId && ctx.account.ss(p.boundId);
  return layout(p.name, '<p><a href="/sandbox/code">&larr; Scripts</a></p><p class="sbx-kicker">Apps Script project</p><h1>' + h(p.name) + '</h1><p class="sbx-lede">' + h(p.blurb) + '</p>'
    + '<dl class="sbx-meta"><dt>Files</dt><dd>' + p.files.map((f) => '<code>' + h(f) + '</code>').join(' ') + '</dd>'
    + (sheet ? '<dt>Bound to</dt><dd><a href="/sandbox/sheets/' + h(sheet.id) + '/edit">' + h(sheet.getName()) + '</a></dd>' : '')
    + (p.webapp ? '<dt>Web app</dt><dd><a href="' + h(p.local) + '">' + h(p.local) + '</a> (on Google: the /exec URL in config.js or app.js)</dd>' : '')
    + (p.editor ? '<dt>On Google</dt><dd><a href="' + h(leave(p.editor)) + '">The real Apps Script project</a></dd>' : '') + '</dl>'
    + load + answerPicker()
    + '<h2>Run a function</h2><p class="sbx-note">Like the Run button in the Apps Script editor. Functions ending in _ are helpers; they run here, but a page cannot call them.</p>'
    + '<ul class="sbx-fns">' + list(pub) + '</ul>'
    + (priv.length ? '<details class="sbx-more"><summary>' + priv.length + ' helper functions</summary><ul class="sbx-fns">' + list(priv) + '</ul></details>' : '')
    + '<div id="run-out" class="sbx-result" hidden></div>'
    + '<h2>Triggers</h2>' + (triggers ? '<ul class="sbx-list">' + triggers + '</ul>' : '<p class="sbx-empty">None installed.</p>')
    + (key === 'space' ? '<p class="sbx-note">On Google this project also has a form-submit trigger (sendEmails, owned by another account) and a daily trigger for sendReminders. In the sandbox the report form calls sendEmails, and the control panel runs sendReminders.</p>' : '')
    + '<h2>Script properties</h2>' + (props ? '<div class="sbx-scroll"><table class="sbx-table"><tbody>' + props + '</tbody></table></div>' : '<p class="sbx-empty">None set.</p>')
    + '<h2>Recent log</h2>' + (logs ? '<pre class="sbx-log">' + logs + '</pre>' : '<p class="sbx-empty">Nothing logged yet.</p>'), ctx, { nav: 'code' });
}

// ---------------------------------------------------------------- leaving

function leavingPage(ctx, to) {
  let kind = 'a real site';
  if (/docs\.google\.com\/document/.test(to)) kind = 'a real Google Doc. Reading it is fine; please don’t edit it from here by mistake';
  else if (/docs\.google\.com\/spreadsheets/.test(to)) kind = 'a real Google Sheet. Changes there are real';
  else if (/forms\.gle|docs\.google\.com\/forms/.test(to)) kind = 'a real Google Form. A response there reaches real teams';
  else if (/script\.google\.com/.test(to)) kind = 'real Apps Script';
  else if (/drive\.google\.com/.test(to)) kind = 'a real Drive file';
  else if (/fabman\.io/.test(to)) kind = 'the real Fabman. Changes there are real';
  return layout('Leaving the sandbox', '<section class="sbx-leave"><p class="sbx-kicker">Leaving the sandbox</p><h1>This link goes outside the sandbox</h1>'
    + '<p class="sbx-lede">It opens ' + kind + '.</p><p class="sbx-url"><code>' + h(to) + '</code></p>'
    + '<p><a class="sbx-btn sbx-btn--ghost" href="javascript:history.back()">Go back</a> <a class="sbx-btn" href="' + h(to) + '" rel="noopener" data-leave="1">Open it anyway</a></p></section>', ctx);
}

// ---------------------------------------------------------------- Apps Script error page

function gasError(message) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Error</title><style>body{font:15px/1.6 Arial,sans-serif;margin:40px;color:#222}h1{font-size:20px;color:#b31b1b}pre{white-space:pre-wrap;background:#f6f5f1;border:1px solid #e7e7e3;border-radius:8px;padding:12px 14px}</style></head>'
    + '<body><h1>Error</h1><pre>' + h(message) + '</pre><p>This is the sandbox’s copy of the Apps Script error page. Fix the script and refresh. <a href="/sandbox/code" target="_top">Scripts</a></p></body></html>';
}

module.exports = { layout, panel, sheetsList, sheetPage, notInSandbox, outbox, mailPage, drivePage, formPage, fabmanPage, codeList, codePage, leavingPage, gasError, resultBlock, h };
