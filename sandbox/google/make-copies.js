#!/usr/bin/env node
'use strict';
/*
 * Makes safe copies of the Space Status scripts for trying a change on real Google,
 * in your own account, before it goes into the live project.
 *
 *   node sandbox/google/make-copies.js you@cornell.edu
 *
 * The copies land in sandbox/google/out/space-status/ (git ignores that folder). They:
 *   - use a spreadsheet and form of your own, made by sandboxSetup(), never the real sheet
 *   - send every email to you only, with [SANDBOX] in the subject (no CC, no alias)
 *   - work out their own web app address, so email buttons open your copy
 * The script refuses to write anything if a copy still mentions a real sheet id or a real
 * address, so a change to the original files cannot quietly make the copies unsafe.
 *
 * See sandbox/google/README.md for the steps on Google.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const REPO = path.resolve(__dirname, '..', '..');
const FILES = ['apps-script/02_notify_on_submit.gs', 'apps-script/04_tasks_projects.gs', 'apps-script/07_proposals.gs'];

// Ids and addresses that must never appear in a copy.
function realThings(repo) {
  const ids = new Set();
  const grab = (text) => { (text.match(/\b1[A-Za-z0-9_-]{40,}/g) || []).forEach((m) => ids.add(m)); };
  ['apps-script/02_notify_on_submit.gs', 'apps-script/command-center/Code.gs', 'apps-script/fleet_driver_cleanup.gs', 'config.js']
    .forEach((f) => { try { grab(fs.readFileSync(path.join(repo, f), 'utf8')); } catch (e) { /* file moved: the checks below still run */ } });
  return { ids: Array.from(ids), addresses: /[\w.+-]+@cornell\.edu/i };
}

// The team list from 01_setup_team_contacts.gs, so the form matches the real one.
function teams(repo) {
  try {
    const box = {};
    vm.runInNewContext(fs.readFileSync(path.join(repo, 'apps-script/01_setup_team_contacts.gs'), 'utf8') + '\n;this.__t = TEAMS;', box, { timeout: 1000 });
    if (Array.isArray(box.__t) && box.__t.length) return box.__t.slice();
  } catch (e) { /* fall through */ }
  return ['Baja', 'Formula', 'Rocketry', 'Operations Team', 'Unknown/unsure'];
}

// Swap one CONFIG line, and insist that it was there exactly once.
function swapLine(src, key, replacement, file) {
  const re = new RegExp('^(\\s*)' + key + ':.*$', 'm');
  const hits = src.match(new RegExp('^\\s*' + key + ':', 'gm')) || [];
  if (hits.length !== 1) throw new Error(file + ': expected one "' + key + ':" line in CONFIG, found ' + hits.length + '. The original changed shape; update make-copies.js.');
  return src.replace(re, (all, indent) => indent + key + ': ' + replacement + ',   // SANDBOX COPY');
}

function banner(me, stamp, file) {
  return '// SANDBOX COPY of ' + file + ', made by sandbox/google/make-copies.js on ' + stamp + '.\n'
    + '// It uses your own sheet and sends every email to ' + me + ' only.\n'
    + '// Never paste this file into the live Space Status project.\n\n';
}

function setupFile(me, stamp, teamList, forms) {
  const J = (v) => JSON.stringify(v);
  return banner(me, stamp, 'the sandbox helpers') + `/**
 * Helpers for this sandbox copy. Run them from the editor (pick the function in the
 * toolbar, then Run):
 *   sandboxSetup()             once: makes your sheet and form, links them, adds the trigger
 *   sandboxSubmitTestReport()  files a report as if someone used the form
 *   sandboxDailyReminders()    optional: run sendReminders every morning, like the real one
 *   sandboxStatus()            logs your sheet, form and web app links
 *   sandboxRemoveTriggers()    stops the triggers when you are done
 */

var SANDBOX = {
  me: ${J(me)},
  teams: ${J(teamList)},
  respondWithin: ${J(forms.WITHIN)},
  issueTypes: ${J(forms.ISSUE_TYPES)},
  actions: ${J(forms.ACTIONS)},
};

// Every email the copied scripts send comes through here and goes to you alone.
var SANDBOX_MAIL = {
  sendEmail: function (to, subject, body, opts) {
    if (to && typeof to === 'object') { opts = to; to = opts.to; subject = opts.subject; body = opts.body; }
    opts = Object.assign({}, opts || {});
    var was = [String(to || ''), opts.cc ? 'cc ' + opts.cc : '', opts.bcc ? 'bcc ' + opts.bcc : ''].filter(String).join('; ');
    delete opts.cc; delete opts.bcc; delete opts.from;
    var note = 'Sandbox copy. In production this email would go to: ' + was;
    if (opts.htmlBody) {
      opts.htmlBody = '<div style="font:13px/1.5 Arial,sans-serif;background:#fff4d6;border:1px solid #f0c050;border-radius:6px;padding:8px 12px;margin:0 0 14px">'
        + note.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</div>' + opts.htmlBody;
    }
    GmailApp.sendEmail(SANDBOX.me, '[SANDBOX] ' + subject, (body || '') + '\\n\\n' + note, opts);
  },
  getRemainingDailyQuota: function () { return MailApp.getRemainingDailyQuota(); },
};

function sandboxSheetId_() {
  var id = PropertiesService.getScriptProperties().getProperty('SANDBOX_SHEET_ID');
  if (!id) throw new Error('Run sandboxSetup first.');
  return id;
}

function sandboxSetup() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SANDBOX_SHEET_ID')) {
    Logger.log('Already set up. sandboxStatus() shows the links. To start over, delete the SANDBOX_SHEET_ID and SANDBOX_FORM_ID script properties (Project Settings) and run this again.');
    return sandboxStatus();
  }
  var ss = SpreadsheetApp.create('Space Status (sandbox copy)');
  props.setProperty('SANDBOX_SHEET_ID', ss.getId());
  CONFIG.spreadsheetId = ss.getId();   // CONFIG was read before the property existed

  // The report form, linked to the sheet. Apps Script cannot add a file-upload
  // question; add "Photo Upload" by hand in the form editor if you want photos.
  var form = FormApp.create('Space Status Report (sandbox copy)');
  form.setDescription('Sandbox copy. Reports here email only ' + SANDBOX.me + '.');
  form.setCollectEmail(true);
  form.addTextItem().setTitle(CONFIG.headers.netid).setRequired(true);
  form.addListItem().setTitle(CONFIG.headers.team).setChoiceValues(SANDBOX.teams).setRequired(true);
  form.addMultipleChoiceItem().setTitle(CONFIG.headers.status).setChoiceValues(SANDBOX.respondWithin).setRequired(true);
  form.addMultipleChoiceItem().setTitle(CONFIG.headers.issueType).setChoiceValues(SANDBOX.issueTypes).setRequired(true);
  form.addMultipleChoiceItem().setTitle(CONFIG.headers.action).setChoiceValues(SANDBOX.actions).setRequired(true);
  form.addParagraphTextItem().setTitle(CONFIG.headers.details).setRequired(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty('SANDBOX_FORM_ID', form.getId());
  SpreadsheetApp.flush();

  ss = SpreadsheetApp.openById(ss.getId());
  var linked = ss.getSheets().filter(function (s) { return s.getFormUrl(); })[0];
  if (linked) linked.setName(CONFIG.responsesSheet);
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  // Team Contacts: every team's addresses are yours.
  var H = ['Team', 'Liaison name', 'Liaison email', 'Team general email', 'Safety lead name', 'Safety lead email', 'Extra CC', 'Team token'];
  var tc = ss.getSheetByName(CONFIG.contactsSheet) || ss.insertSheet(CONFIG.contactsSheet);
  var rows = SANDBOX.teams.map(function (t) {
    return [t, 'Sandbox liaison', SANDBOX.me, SANDBOX.me, 'Sandbox safety lead', SANDBOX.me, '', Utilities.getUuid().replace(/-/g, '').slice(0, 16)];
  });
  tc.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold');
  tc.getRange(2, 1, rows.length, H.length).setValues(rows);
  tc.setFrozenRows(1);

  // The tabs and columns the scripts set up for themselves.
  setupTasksProjects();
  setupProposals();
  sandboxInstallTriggers();
  sandboxStatus();
  Logger.log('Next: Deploy > New deployment > Web app (Execute as: Me). Then submit the form, or run sandboxSubmitTestReport.');
}

function sandboxInstallTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendEmails') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendEmails').forSpreadsheet(sandboxSheetId_()).onFormSubmit().create();
}

function sandboxDailyReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(8).create();
  Logger.log('sendReminders will run every morning around 8. sandboxRemoveTriggers() stops it.');
}

function sandboxRemoveTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('All triggers removed. Reports on the sandbox form no longer send email.');
}

// Files a report the way the form would: a new row, then the form-submit handler.
function sandboxSubmitTestReport(team, within, details) {
  var ss = SpreadsheetApp.openById(sandboxSheetId_());
  var sh = ss.getSheetByName(CONFIG.responsesSheet);
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var answers = {};
  answers[CONFIG.headers.timestamp] = new Date();
  answers[CONFIG.headers.email] = SANDBOX.me;
  answers[CONFIG.headers.netid] = 'test1';
  answers[CONFIG.headers.team] = team || 'Baja';
  answers[CONFIG.headers.status] = within || SANDBOX.respondWithin[1];
  answers[CONFIG.headers.issueType] = SANDBOX.issueTypes[5] || SANDBOX.issueTypes[0];
  answers[CONFIG.headers.action] = SANDBOX.actions[0];
  answers[CONFIG.headers.details] = details || 'Test report from sandboxSubmitTestReport';
  var row = H.map(function (h) { return Object.prototype.hasOwnProperty.call(answers, h) ? answers[h] : ''; });
  sh.appendRow(row);
  var r = sh.getLastRow();
  var namedValues = {};
  H.forEach(function (h, i) { namedValues[h] = [row[i] instanceof Date ? row[i].toLocaleString() : String(row[i])]; });
  sendEmails({ namedValues: namedValues, range: sh.getRange(r, 1, 1, H.length), values: row.map(String) });
  Logger.log('Filed a test report on row ' + r + '. Check your inbox for [SANDBOX].');
}

function sandboxStatus() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SANDBOX_SHEET_ID'), formId = props.getProperty('SANDBOX_FORM_ID');
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  var lines = [
    'Sheet:   ' + (sheetId ? 'https://docs.google.com/spreadsheets/d/' + sheetId + '/edit' : 'not made yet (run sandboxSetup)'),
    'Form:    ' + (formId ? FormApp.openById(formId).getPublishedUrl() : 'not made yet'),
    'Web app: ' + (url || 'not deployed yet (Deploy > New deployment > Web app)'),
    'Emails:  all go to ' + SANDBOX.me,
    'Triggers: ' + ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join(', '),
  ];
  if (url) lines.push('Pages:   ' + url + '?view=all&admin=1  (admin dashboard), ' + url + '?module=projects&admin=1, ' + url + '?module=proposals');
  Logger.log(lines.join('\\n'));
  return lines;
}
`;
}

function makeCopies(opts) {
  const me = String(opts.me || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(me)) throw new Error('Give your own email address: node sandbox/google/make-copies.js you@cornell.edu');
  const repo = opts.repo || REPO;
  const out = opts.out || path.join(__dirname, 'out', 'space-status');
  const stamp = new Date().toISOString().slice(0, 10);
  const forms = require('../lib/forms');
  const real = realThings(repo);

  const files = {};
  FILES.forEach((f) => {
    let src = fs.readFileSync(path.join(repo, f), 'utf8');
    if (/02_notify_on_submit/.test(f)) {
      src = swapLine(src, 'spreadsheetId', "PropertiesService.getScriptProperties().getProperty('SANDBOX_SHEET_ID') || 'Run sandboxSetup first'", f);
      src = swapLine(src, 'fleetSpreadsheetId', "''", f);
      src = swapLine(src, 'sender', "''", f);
      src = swapLine(src, 'alwaysCc', '[]', f);
      src = swapLine(src, 'fallbackEmail', JSON.stringify(me).replace(/"/g, "'"), f);
      src = swapLine(src, 'webAppUrl', "PropertiesService.getScriptProperties().getProperty('SANDBOX_WEBAPP_URL') || (function () { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })()", f);
    }
    src = src.replace(/\b(GmailApp|MailApp)\.sendEmail\(/g, 'SANDBOX_MAIL.sendEmail(');
    files[path.basename(f)] = banner(me, stamp, f) + src;
  });
  files['00_sandbox.gs'] = setupFile(me, stamp, teams(repo), forms);

  // Refuse to write a copy that could reach the real platform.
  Object.keys(files).forEach((name) => {
    const text = files[name];
    real.ids.forEach((id) => { if (text.indexOf(id) >= 0) throw new Error(name + ' still contains the real id ' + id + '. Nothing was written.'); });
    const addr = text.split(me).join('').replace(/\/\/[^\n]*/g, '').match(real.addresses);
    if (addr) throw new Error(name + ' still contains the real address ' + addr[0] + '. Nothing was written.');
    if (/\b(GmailApp|MailApp)\.sendEmail\(/.test(text.replace(/var SANDBOX_MAIL[\s\S]*?\n};/, ''))) throw new Error(name + ' sends email outside SANDBOX_MAIL. Nothing was written.');
  });

  fs.mkdirSync(out, { recursive: true });
  Object.keys(files).forEach((name) => fs.writeFileSync(path.join(out, name), files[name]));
  return { out, files: Object.keys(files).sort() };
}

module.exports = { makeCopies };

if (require.main === module) {
  try {
    const r = makeCopies({ me: process.argv[2] });
    console.log('\nSafe copies written to ' + path.relative(process.cwd(), r.out) + ':');
    r.files.forEach((f) => console.log('  ' + f));
    console.log('\nNext: follow sandbox/google/README.md (make an Apps Script project, add these files, run sandboxSetup).\n');
  } catch (err) {
    console.error('\n' + err.message + '\n');
    process.exit(1);
  }
}
