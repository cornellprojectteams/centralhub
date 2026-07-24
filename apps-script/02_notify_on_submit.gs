/**
 * Space Status: notifier + acknowledgement web app + reminders + team portal.
 * Emails are clean Cornell letters (Georgia) with a red action button.
 * Web pages are Swiss minimal (white, thin red top line, bold sans).
 * All inline-styled so they render reliably.
 *
 * Any change to web-app logic needs a NEW deployment version (Deploy > Manage
 * deployments > pencil > Version: New version). Same URL after.
 */

const CONFIG = {
  spreadsheetId: '1mZrlnA-GiVKB4_Um21aMH9jsSkf6TUJhauDWivnf7-I',
  // Equipment / Inventory registry lives in its OWN spreadsheet, not the space-issues one.
  // Set this to your inventory spreadsheet id (this default is "Tools, purchasing & locks").
  registrySpreadsheetId: '1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU',
  // Fleet driver-cleanup tool (?tool=drivers). Runs on the reservation sheet.
  fleetSpreadsheetId: '1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M',
  fleetGradYearThrough: 0,   // remove this graduation year and earlier; 0 = current year.
  toolPassHash: '993b97603778a944ff00c86775aeb4c852a6deae031dbee58206dc3ae1e3242c',   // SHA-256 of the tool passcode (default "bigred")
  responsesSheet: 'Form Responses',
  contactsSheet: 'Team Contacts',
  sender: 'eng_projectteams@cornell.edu',                          // go-live: 'eng_projectteams@cornell.edu' (send-as alias on this account)
  alwaysCc: ['nhh5@cornell.edu'],                        // go-live: ['nhh5@cornell.edu']
  fallbackEmail: 'nhh5@cornell.edu',
  webAppUrl: 'https://script.google.com/macros/s/AKfycbwNbGjVcBrcsMZiOl2nXzpqZHz04nvKLm9D_aC0VJDz7Xxxf_4kLKlNSOHubPXj1X74/exec',  // MUST be the published /exec URL; email acknowledge links use this
  notifiedHeader: 'Notified at',
  issueTokenHeader: 'Issue token',
  addressedHeader: 'Addressed at',            // set on ADMIN APPROVAL = Completed
  lastReminderHeader: 'Last reminder',
  completionPhotoHeader: 'Completion photo',  // evidence uploaded by the doer
  completedAtHeader: 'Completed at',           // doer submitted -> Pending approval
  sentBackHeader: 'Sent back reason',          // admin's note when sending a submission back (photo is kept)
  headers: {
    timestamp: 'Timestamp',
    email: 'Email Address',
    netid: 'NetID',
    team: 'Responsible Team',
    issueType: 'Issue Type',
    action: 'Required action',
    details: 'Details / comments',
    status: 'Current Status',
    photo: 'Photo Upload',
  },
};

const SLA_DAYS = { red: 0, orange: 3, yellow: 10, ivory: 21, purple: null };
const COLOR_HEX = { red: '#b31b1b', orange: '#e08a1e', yellow: '#caa12a', ivory: '#b8a06a', purple: '#7c3aed' };

const SEVERITY = {
  red:    { subject: 'Urgent: same-day action needed',     open: "A space safety issue requiring immediate, same-day attention has been reported in your team's space.", ask: 'Please resolve this today' },
  orange: { subject: 'Action needed within 2 to 3 days',   open: "A space safety issue needing prompt attention has been reported in your team's space.",                ask: 'Please resolve this within 2 to 3 business days' },
  yellow: { subject: 'Please address within 7 to 10 days', open: "A space stewardship concern has been reported in your team's space.",                                  ask: 'Please resolve this within 7 to 10 days' },
  ivory:  { subject: 'For attention within 14 to 21 days', open: "A space issue needing attention has been reported in your team's space.",                              ask: 'Please address this within the next 2 to 3 weeks' },
  purple: { subject: 'Improvement opportunity',            open: "An opportunity to improve your team's space has been reported.",                                       ask: 'Please address this when time allows' },
};
const SEVERITY_DEFAULT = { subject: 'Space issue reported', open: "A space issue has been reported in your team's space.", ask: 'Please address this at your earliest convenience' };

// The form-submit trigger (created by another account) calls sendEmails(); route it here.
function sendEmails(e) { onFormSubmit(e); }

function onFormSubmit(e) {
  if (!e || !e.namedValues) { Logger.log('Run from the form-submit trigger.'); return; }
  const data = dataFromNamedValues_(e.namedValues);
  const token = newToken_();
  const result = sendNotification_(data, token);
  if (result.sent && e.range) stampRow_(e.range.getSheet(), e.range.getRow(), token);
}

function dataFromNamedValues_(nv) {
  const map = {};
  Object.keys(nv).forEach(function (k) { map[norm_(k)] = nv[k]; });
  const get = function (h) { const v = map[norm_(h)]; return v && v[0] ? String(v[0]).trim() : ''; };
  return readData_(get);
}

function readData_(get) {
  const h = CONFIG.headers;
  return {
    timestamp: get(h.timestamp) || new Date().toLocaleString(),
    email: get(h.email), netid: get(h.netid), team: get(h.team),
    issueType: get(h.issueType), action: get(h.action), details: get(h.details),
    status: get(h.status), photo: get(h.photo),
  };
}

function recipientsFor_(team) {
  const c = lookupTeam_(team);
  const to = uniqEmails_([c && c.generalEmail, c && c.liaisonEmail]);
  const cc = uniqEmails_([c && c.safetyEmail, c && c.extraCc].concat(CONFIG.alwaysCc));
  if (!to.length) to.push(CONFIG.fallbackEmail);
  return { to: to, cc: cc, contact: c };
}

function sendNotification_(data, token) {
  const team = data.team || 'Unknown / unsure';
  const rec = recipientsFor_(team);
  const color = parseColor_(data.status);
  const sev = SEVERITY[color] || SEVERITY_DEFAULT;
  const subject = sev.subject + ' - ' + team + (data.issueType ? ' (' + phrase_(data.issueType) + ')' : '');
  let inlineImages = {}, photoCid = '';
  if (data.photo) {
    const ids = extractFileIds_(data.photo);
    ids.forEach(function (id) {
      try { DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (err) { Logger.log('Photo share failed for ' + id + ': ' + err); }
    });
    if (ids.length) {
      try {
        inlineImages = { spaceStatusPhoto: DriveApp.getFileById(ids[0]).getBlob() };
        photoCid = 'spaceStatusPhoto';
      } catch (err) { Logger.log('Photo not embedded: ' + err); }
    }
  }
  const opts = {
    htmlBody: buildEmail_(data, color, photoCid, token),
    name: 'Engineering Student Project Teams',
    cc: rec.cc.join(','), replyTo: data.email || CONFIG.fallbackEmail,
  };
  if (CONFIG.sender) opts.from = CONFIG.sender;
  if (photoCid) opts.inlineImages = inlineImages;
  GmailApp.sendEmail(rec.to.join(','), subject, '', opts);
  Logger.log('Sent "' + team + '" to: ' + rec.to.join(',') + ' | cc: ' + rec.cc.join(','));
  return { sent: true };
}

function sendReminders() {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) { Logger.log('No rows.'); return; }
  const H = v[0].map(norm_);
  const ci = function (name) { return H.indexOf(norm_(name)); };
  const cTs = ci(CONFIG.headers.timestamp), cTeam = ci(CONFIG.headers.team), cStatus = ci(CONFIG.headers.status);
  const cIssue = ci(CONFIG.headers.issueType), cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details);
  const cEmail = ci(CONFIG.headers.email), cNotified = ci(CONFIG.notifiedHeader), cToken = ci(CONFIG.issueTokenHeader);
  const cAddr = ensureColumn_(sh, CONFIG.addressedHeader) - 1;
  const cLast = ensureColumn_(sh, CONFIG.lastReminderHeader) - 1;
  const cCompleted = ci(CONFIG.completedAtHeader);   // -1 until first completion is submitted
  const DAY = 86400000;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let sent = 0;

  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    if (cNotified < 0 || !row[cNotified]) continue;            // never notified
    if (row[cAddr]) continue;                                   // addressed / closed
    if (cCompleted >= 0 && row[cCompleted]) continue;          // completion submitted, awaiting approval -> pause reminders
    if (cToken < 0 || !row[cToken]) continue;
    const color = parseColor_(row[cStatus]);
    if (!color || SLA_DAYS[color] == null) continue;            // purple / unknown -> no deadline, no reminders
    const reportDate = new Date(row[cTs]);
    if (isNaN(reportDate.getTime())) continue;
    reportDate.setHours(0, 0, 0, 0);
    if (today <= reportDate) continue;                          // day 0 is covered by the initial notice
    const windowDays = SLA_DAYS[color];
    const deadline = new Date(reportDate); deadline.setDate(deadline.getDate() + windowDays); deadline.setHours(0, 0, 0, 0);
    const interval = Math.max(1, Math.round(windowDays / 2));   // recurring cadence = (due - report) / 2
    let dueToday;
    if (row[cLast]) {
      const last = new Date(row[cLast]); last.setHours(0, 0, 0, 0);
      dueToday = (today - last) / DAY >= interval;              // every `interval` days after the last reminder
    } else {
      dueToday = today >= deadline;                             // first reminder on (or after) the due date
    }
    if (!dueToday) continue;

    const data = {
      timestamp: row[cTs], team: cTeam >= 0 ? String(row[cTeam]).trim() : '',
      issueType: cIssue >= 0 ? String(row[cIssue]).trim() : '', action: cAction >= 0 ? String(row[cAction]).trim() : '',
      details: cDetails >= 0 ? String(row[cDetails]).trim() : '', status: cStatus >= 0 ? String(row[cStatus]).trim() : '',
      email: cEmail >= 0 ? String(row[cEmail]).trim() : '',
    };
    const token = String(row[cToken]);
    const rec = recipientsFor_(data.team || 'Unknown / unsure');
    const subject = 'Reminder - ' + (data.team || '') + ' space issue needs attention' + (data.issueType ? ' (' + phrase_(data.issueType) + ')' : '');
    const opts = { htmlBody: buildReminder_(data, color, token, deadline), name: 'Engineering Student Project Teams', cc: rec.cc.join(','), replyTo: data.email || CONFIG.fallbackEmail };
    if (CONFIG.sender) opts.from = CONFIG.sender;
    GmailApp.sendEmail(rec.to.join(','), subject, '', opts);
    sh.getRange(i + 1, cLast + 1).setValue(new Date());
    sent++;
  }
  Logger.log('Reminders sent: ' + sent);
}

function lookupTeam_(team) {
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  if (!sh) { Logger.log('No "' + CONFIG.contactsSheet + '" tab.'); return null; }
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return null;
  const H = v[0].map(norm_);
  const idx = function (name) { return H.indexOf(norm_(name)); };
  const cTeam = idx('Team'), cGen = idx('Team general email'), cLia = idx('Liaison email'),
        cLiaN = idx('Liaison name'), cSafe = idx('Safety lead email'), cExtra = idx('Extra CC');
  if (cTeam < 0) { Logger.log('Team Contacts has no "Team" header.'); return null; }
  for (let i = 1; i < v.length; i++) {
    if (norm_(v[i][cTeam]) === norm_(team)) {
      return {
        team: v[i][cTeam],
        generalEmail: cGen >= 0 ? v[i][cGen] : '', liaisonEmail: cLia >= 0 ? v[i][cLia] : '',
        liaisonName: cLiaN >= 0 ? v[i][cLiaN] : '', safetyEmail: cSafe >= 0 ? v[i][cSafe] : '',
        extraCc: cExtra >= 0 ? v[i][cExtra] : '',
      };
    }
  }
  return null;
}

function lookupTeamByToken_(token) {
  if (!token) return '';
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  if (!sh) return '';
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const cTeam = H.indexOf(norm_('Team')), cTok = H.indexOf(norm_('Team token'));
  if (cTeam < 0 || cTok < 0) return '';
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][cTok]) === String(token)) return String(v[i][cTeam]).trim();
  }
  return '';
}

// All known team names from the contacts sheet (for the admin reassign dropdown).
function icTeamNames_() {
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const cTeam = v[0].map(norm_).indexOf(norm_('Team'));
  if (cTeam < 0) return [];
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const t = String(v[i][cTeam] || '').trim();
    if (t && out.indexOf(t) < 0) out.push(t);
  }
  return out;
}

// <option> list for the team select, keeping the current value even if unlisted.
function icTeamOptions_(teams, current) {
  const list = teams.slice();
  const cur = String(current || '').trim();
  if (cur && list.indexOf(cur) < 0) list.unshift(cur);
  let out = '<option value=""' + (cur ? '' : ' selected') + '>Unassigned</option>';
  list.forEach(function (t) {
    out += '<option value="' + escapeHtml_(t) + '"' + (t === cur ? ' selected' : '') + '>' + escapeHtml_(t) + '</option>';
  });
  return out;
}

// Distinct existing values from the sheet, so the editor's pick-lists only offer
// values that already exist rather than free text (no typos, no invented values).
function icFieldOptions_() {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  if (!sh) return { actions: [], issueTypes: [] };
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { actions: [], issueTypes: [] };
  const H = v[0].map(norm_);
  const cAct = H.indexOf(norm_(CONFIG.headers.action));
  const cType = H.indexOf(norm_(CONFIG.headers.issueType));
  const actions = [], types = [];
  for (let i = 1; i < v.length; i++) {
    if (cAct >= 0) { const a = String(v[i][cAct] || '').trim(); if (a && actions.indexOf(a) < 0) actions.push(a); }
    if (cType >= 0) { const t = String(v[i][cType] || '').trim(); if (t && types.indexOf(t) < 0) types.push(t); }
  }
  actions.sort(); types.sort();
  return { actions: actions, issueTypes: types };
}

// <option> list keeping the current value even if unlisted. The value is the raw
// stored text (what downstream logic reads); the label is the tidied phrase_ form.
function icPickOptions_(values, current) {
  const list = values.slice();
  const cur = String(current || '').trim();
  if (cur && list.indexOf(cur) < 0) list.unshift(cur);
  let out = '<option value=""' + (cur ? '' : ' selected') + '>Not set</option>';
  list.forEach(function (val) {
    out += '<option value="' + escapeHtml_(val) + '"' + (val === cur ? ' selected' : '') + '>' + escapeHtml_(phrase_(val)) + '</option>';
  });
  return out;
}

// Admin-only edit form for one issue card (hidden until Edit is tapped).
function icEditForm_(rid, it, teams, issueTypes) {
  return '<div id="' + rid + '-edit" class="ic-edit" hidden>'
    + '<label class="ic-edit-lbl">Team<select id="' + rid + '-eteam" class="ic-edit-in">' + icTeamOptions_(teams, it.team) + '</select></label>'
    + '<label class="ic-edit-lbl">Issue type<select id="' + rid + '-etype" class="ic-edit-in">' + icPickOptions_(issueTypes, it.issueType) + '</select></label>'
    + '<label class="ic-edit-lbl">Details<textarea id="' + rid + '-edetails" class="ic-edit-in" rows="3">' + escapeHtml_(it.details || '') + '</textarea></label>'
    + '<div class="ic-edit-btns"><button type="button" class="btn btn-primary" onclick="icEditSave(\'' + rid + '\',\'' + it.token + '\')">Save changes</button>'
    + '<button type="button" class="btn btn-ghost" onclick="icEditCancel(\'' + rid + '\')">Cancel</button>'
    + '<span id="' + rid + '-emsg" class="ic-edit-msg"></span></div>'
    + '</div>';
}

function parseColor_(status) {
  if (!status) return '';
  const m = String(status).trim().match(/^[A-Za-z]+/);
  const c = m ? m[0].toLowerCase() : '';
  return SLA_DAYS.hasOwnProperty(c) ? c : '';
}

function deadlineDate_(color) {
  const days = SLA_DAYS[color];
  if (days === null || days === undefined) return null;
  const d = new Date(); d.setDate(d.getDate() + days);
  return d;
}
function fmtDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy'); }
function fmtShort_(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMMM d, yyyy'); }
function daysLeftLabel_(d) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const dd = new Date(d); dd.setHours(0, 0, 0, 0);
  const days = Math.round((dd - today) / 86400000);
  if (days <= 0) return 'Due today';
  if (days === 1) return '1 day left';
  return days + ' days left';
}

function addressedButton_(token) {
  // Must be the published /exec URL from CONFIG.webAppUrl.
  // Do NOT use ScriptApp.getService().getUrl() here: from a trigger it returns the
  // editor-only /dev URL, which recipients cannot open ("unable to open the file").
  if (!CONFIG.webAppUrl || !token) return '';
  const url = CONFIG.webAppUrl + '?id=' + encodeURIComponent(token);
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:20px 0 4px"><tr>'
    + '<td style="background:#8f1515;border-radius:6px"><a href="' + url + '" style="display:inline-block;padding:12px 24px;color:#fff;font:bold 14px Arial,sans-serif;text-decoration:none;border-radius:6px">Mark complete</a></td>'
    + '</tr></table>';
}

// Clean Cornell letter. No header wordmark or urgency line (already in the subject).
function buildEmail_(data, color, photoCid, token) {
  const sev = SEVERITY[color] || SEVERITY_DEFAULT;
  const team = data.team || '';
  const issue = data.issueType ? phrase_(data.issueType) : '';
  const action = data.action ? phrase_(data.action) : '';
  const dl = deadlineDate_(color);
  const L = function (html) { return '<p style="margin:12px 0 0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + html + '</p>'; };

  let out = '<div style="max-width:600px;margin:0;padding:8px 6px;font-family:Georgia,serif;color:#1a1a1a">';
  out += '<p style="margin:0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + (team ? 'Dear ' + escapeHtml_(team) + ' team,' : 'Hello,') + '</p>';
  out += L(escapeHtml_(sev.open));
  let pb = '';
  if (issue) pb += 'The concern relates to ' + escapeHtml_(lcFirst_(issue));
  if (action) pb += (pb ? ', and we ask that you ' : 'We ask that you ') + escapeHtml_(lcFirst_(action));
  if (pb) out += L(pb + '.');
  if (data.details) out += L('The reporter noted: &ldquo;' + escapeHtml_(data.details) + '&rdquo;');
  out += L(escapeHtml_(sev.ask) + '.' + (dl ? ' The deadline is <b>' + escapeHtml_(fmtDate_(dl)) + '</b>.' : ''));
  out += addressedButton_(token);
  out += L('With thanks,<br>Engineering Student Project Teams');
  if (photoCid) out += '<p style="margin:18px 0 6px;color:#777;font:13px Arial">Reference photo:</p><img src="cid:' + photoCid + '" width="240" style="border-radius:6px">';
  out += '</div>';
  return out;
}

function buildReminder_(data, color, token, deadline) {
  const team = data.team || '';
  const issue = data.issueType ? phrase_(data.issueType) : '';
  const action = data.action ? phrase_(data.action) : '';
  const L = function (html) { return '<p style="margin:12px 0 0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + html + '</p>'; };

  let out = '<div style="max-width:600px;margin:0;padding:8px 6px;font-family:Georgia,serif;color:#1a1a1a">';
  out += '<p style="margin:0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + (team ? 'Dear ' + escapeHtml_(team) + ' team,' : 'Hello,') + '</p>';
  out += L('This is a follow-up on a space issue in your team\'s space that has not yet been marked as addressed. It was due on <b>' + escapeHtml_(fmtDate_(deadline)) + '</b>.');
  let pb = '';
  if (issue) pb += 'The concern relates to ' + escapeHtml_(lcFirst_(issue));
  if (action) pb += (pb ? ', and we ask that you ' : 'We ask that you ') + escapeHtml_(lcFirst_(action));
  if (pb) out += L(pb + '.');
  out += L('Please resolve it as soon as possible.');
  out += addressedButton_(token);
  out += L('With thanks,<br>Engineering Student Project Teams');
  out += '</div>';
  return out;
}

// ---- Web app pages (Swiss minimal) ----

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const embed = p.embed === '1' || p.embed === 'true';
  // Admin mode: set only by links from the (unlisted) admin page. It unlocks the
  // admin controls without a passcode - the admin page's obscurity is the gate.
  const admin = p.admin === '1' || p.admin === 'true';
  if (p.module === 'projects') return projectsPage_(embed, admin);   // Multi-user projects (04_tasks_projects.gs)
  if (p.module === 'projects-dash') return projectsDashboardPage_(embed);   // Projects dashboard (04_tasks_projects.gs)
  if (p.team) { const t = lookupTeamByToken_(p.team); return t ? teamPortal_(t, false) : htmlPage_('Invalid link', 'This team link is not recognized.'); }
  if (p.view === 'all') return allIssuesPage_(embed, admin);
  if (p.registry) return registryPage_(String(p.registry), p.embed === '1' || p.embed === 'true');   // read-only Equipment / Inventory tables
  if (p.view) return teamPortal_(String(p.view), !(p.act === '1' || p.act === 'true'), p.embed === '1' || p.embed === 'true');   // ?view=<team> read-only; &act=1 markable; &embed=1 no top bar
  if (p.id) return confirmPage_(p.id);
  return pickerPage_();
}

function portalStyles_() {
  return ''
    + '@import url("https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap");'
    + '*,*::before,*::after{box-sizing:border-box}'
    + 'body{margin:0;font-family:Inter,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased}'
    + '.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;line-height:1;padding:11px 20px;border-radius:10px;border:none;cursor:pointer;text-decoration:none;white-space:nowrap;transition:transform .16s ease,box-shadow .16s ease,background .16s ease,border-color .16s ease}'
    + '.btn-primary{color:#fff;background:linear-gradient(180deg,#d62b2b 0%,#b31b1b 55%,#8f1515 100%);box-shadow:0 4px 14px rgba(179,27,27,.28)}'
    + '.btn-primary:hover{transform:translateY(-1px);box-shadow:0 8px 20px rgba(179,27,27,.36)}'
    + '.btn-primary:active{transform:translateY(0)}'
    + '.btn-confirm{color:#fff;background:linear-gradient(180deg,#1d9d5b 0%,#157a47 100%);box-shadow:0 4px 12px rgba(21,122,71,.25)}'
    + '.btn-confirm:hover{transform:translateY(-1px);box-shadow:0 8px 18px rgba(21,122,71,.32)}'
    + '.btn-ghost{color:#666;background:#fff;border:1.5px solid #e0e0dc;padding:10px 16px;border-radius:10px}'
    + '.btn-ghost:hover{color:#333;border-color:#ccc;background:#fafaf8}'
    + '.btn-row{display:inline-flex;align-items:center;gap:8px;flex-wrap:wrap}'
    + '.stats{display:flex;flex-wrap:wrap;gap:10px;margin-top:14px}'
    + '.stat{flex:1;min-width:120px;padding:16px 18px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.06);position:relative;overflow:hidden}'
    + '.stat::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#8f1515,#b31b1b,#f0c050)}'
    + '.stat-label{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c}'
    + '.stat-val{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:28px;font-weight:800;letter-spacing:-.04em;line-height:1.1;margin-top:6px;color:#111}'
    + '.stat-val--danger{color:#b31b1b}'
    + '.filters{display:flex;flex-wrap:wrap;gap:10px;margin:18px 0 10px;padding:14px 16px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.search-wrap{flex:1;min-width:180px;position:relative}'
    + '.search-wrap input{width:100%;font:inherit;font-size:14px;padding:12px 14px 12px 40px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fafaf8;outline:none;transition:border-color .15s,box-shadow .15s}'
    + '.search-wrap input:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12);background:#fff}'
    + '.search-wrap::before{content:"";position:absolute;left:14px;top:50%;transform:translateY(-50%);width:16px;height:16px;background:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'16\' height=\'16\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%239a9a96\' stroke-width=\'2.2\' stroke-linecap=\'round\'%3E%3Ccircle cx=\'11\' cy=\'11\' r=\'7\'/%3E%3Cline x1=\'21\' y1=\'21\' x2=\'16.65\' y2=\'16.65\'/%3E%3C/svg%3E") center/contain no-repeat;pointer-events:none}'
    + '.filters select{font:inherit;font-size:14px;font-weight:600;padding:12px 14px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;outline:none;cursor:pointer;min-width:150px;color:#333}'
    + '.filters select:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12)}'
    + '.toggle{display:inline-flex;align-items:center;gap:8px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#444;padding:10px 14px;border-radius:10px;border:1.5px solid #e0e0dc;background:#fff;cursor:pointer;user-select:none;transition:border-color .15s,background .15s,box-shadow .15s}'
    + '.toggle:hover{border-color:#ccc;box-shadow:0 2px 6px rgba(20,20,30,.06)}'
    + '.toggle.is-on,.toggle:has(input:checked){border-color:#b31b1b;background:#fff8f8;color:#8f1515;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.toggle input{width:16px;height:16px;accent-color:#b31b1b;cursor:pointer;margin:0}'
    + '.section-head{display:flex;align-items:center;gap:11px;margin:26px 0 14px}'
    + '.section-label{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.14em;text-transform:uppercase}'
    + '.section-label--late{color:#b31b1b}'
    + '.section-label--open{color:#8a857c}'
    + '.section-count{font-size:12px;color:#bbb;font-weight:700}'
    + '.section-rule{flex:1;height:1.5px;background:linear-gradient(90deg,#e8e8e8,transparent)}'
    + '.card{background:#fff;border-radius:14px;border:1.5px solid #ececea;box-shadow:0 2px 8px rgba(20,20,30,.06),0 1px 3px rgba(20,20,30,.04);overflow:hidden;margin-bottom:12px;transition:opacity .25s ease,transform .25s ease,box-shadow .25s ease}'
    + '@keyframes cardApprove{0%{box-shadow:0 0 0 0 rgba(21,122,71,0)}22%{box-shadow:0 0 0 3px rgba(21,122,71,.5),0 6px 18px rgba(21,122,71,.22)}100%{box-shadow:0 0 0 0 rgba(21,122,71,0)}}'
    + '.card--approved{animation:cardApprove .85s ease-out}'
    + '.card:hover{box-shadow:0 8px 24px rgba(20,20,30,.08);transform:translateY(-1px)}'
    + '.card-body{padding:18px 20px 16px}'
    + '.card-foot{display:flex;align-items:center;gap:8px;flex-wrap:wrap;border-top:1.5px solid #f1f1f1;padding:13px 20px;background:linear-gradient(180deg,#fcfcfb 0%,#f8f8f6 100%)}'
    + '.card-foot>:first-child{margin-right:auto}'
    + '.card-foot .btn-row{gap:8px}'
    + '.card-team{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b}'
    + '.card-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;letter-spacing:-.02em;line-height:1.25;margin-top:3px;color:#111}'
    + '.card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}'
    + '.card-field{margin-top:13px}'
    + '.card-flabel{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#a8a29e;margin-bottom:4px}'
    + '.card-action{font-size:14.5px;color:#26231f;font-weight:600;line-height:1.5}'
    + '.card-details{font-size:13.5px;color:#57534e;line-height:1.62;white-space:pre-line;padding-left:12px;border-left:2px solid #ece9e2}'
    + '.chip{flex:0 0 auto;white-space:nowrap;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;padding:5px 11px;border-radius:999px}'
    + '.chip--overdue{color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0}'
    + '.chip--due{color:#6b665e;background:#f0efe9;border:1px solid #e5e4de}'
    + '.due{font-size:13px;font-weight:600;color:#777}'
    + '.due--late{color:#b31b1b}'
    + '.due--done{color:#157a47;font-weight:700}'
    + '.empty{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;color:#8a857c;margin-top:12px;padding:24px 20px;text-align:center;background:#fff;border:1.5px dashed #ddd;border-radius:14px}'
    + '.page-head{margin-bottom:4px}'
    + '.page-kicker{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#9a958c}'
    + '.page-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:30px;font-weight:800;letter-spacing:-.035em;line-height:1.05;margin-top:8px;color:#111}'
    + '.page-rule{width:46px;height:3px;background:linear-gradient(90deg,#8f1515,#b31b1b,#f0c050);margin-top:12px;border-radius:99px}'
    + '@media(max-width:600px){.card-foot{flex-direction:column;align-items:stretch}.btn-row{width:100%;justify-content:stretch}.btn{flex:1;justify-content:center}.filters select{width:100%}.stat-val{font-size:24px}}';
}

function swissShell_(innerHtml, pageTitle, wide, embedded) {
  const maxW = wide ? '960px' : '600px';
  const pad = embedded ? '18px 18px 28px' : '40px 28px 64px';
  const topBar = embedded ? '' : '<div style="height:5px;background:linear-gradient(90deg,#8f1515,#b31b1b,#8f1515)"></div>';
  const enh = '<style>' + portalStyles_()
    + '@media(max-width:600px){.page-title{font-size:24px !important}}'
    + '</style>';
  const html = enh
    + '<div style="margin:0;background:#f5f4f0;min-height:100vh;font-family:Helvetica,Arial,sans-serif;color:#111">'
    + topBar
    + '<div style="max-width:' + maxW + ';margin:0 auto;padding:' + pad + '">' + innerHtml + '</div>'
    + '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(pageTitle || 'Space Status')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function htmlPage_(title, bodyHtml) {
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">' + escapeHtml_(title) + '</div>'
    + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:14px">' + bodyHtml + '</div>';
  return swissShell_(inner, 'Space Status');
}

// Reached from the "Mark complete" button in the notification email (?id=token).
// Most tasks need an evidence photo; the photo-optional action types (see
// icPhotoOptional_) can also be finished with one tap. Either way -> Pending approval.
function confirmPage_(id) {
  const info = findIssue_(id);
  if (!info) return htmlPage_('Not found', 'We could not find that item. It may have been removed.');
  if (info.addressed) return htmlPage_('Already completed', 'This was approved as complete on ' + escapeHtml_(fmtShort_(info.addressed)) + '.');
  if (info.completedAt) return htmlPage_('Pending approval', 'This was already submitted on ' + escapeHtml_(fmtShort_(info.completedAt)) + '. It is waiting for an admin to review.');
  const photoOptional = icPhotoOptional_(info.action);
  const sentBackBanner = info.sentBackReason
    ? '<div style="margin-top:14px;font:600 14px/1.6 Arial,sans-serif;color:#8a4b00;background:#fdf2df;border:1px solid #f4dfb0;border-radius:10px;padding:11px 14px"><b>Sent back:</b> ' + escapeHtml_(info.sentBackReason) + '</div>'
    : '';
  const lead = photoOptional
    ? 'Tap <b>Mark done</b> and an admin gives it a quick review. A photo is optional here, so add one only if it helps.'
    : 'Add a photo of the completed work and an admin gives it a quick review.';
  const controls = photoOptional
    ? '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'cf-file\').click()">Add a photo</button>'
      + '<button type="button" class="btn btn-primary" onclick="cfDone()">Mark done</button>'
    : '<button type="button" class="btn btn-primary" onclick="document.getElementById(\'cf-file\').click()">Complete with a photo</button>';
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">Mark complete</div>'
    + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:12px">' + lead + '</div>'
    + sentBackBanner
    + '<div id="act" style="margin-top:24px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
    +   '<input type="file" accept="image/*" id="cf-file" style="display:none" onchange="cfFile(this)">'
    +   controls
    + '</div>'
    + '<div id="cf-photo"></div>'
    + '<div id="done" style="display:none;font-size:22px;font-weight:800;letter-spacing:-.02em;color:#1d7a46;margin-top:14px"></div>'
    + '<script>'
    + 'function cfDoneUi(photoId){tpConfetti();var a=document.getElementById("act");a.style.display="none";if(photoId){document.getElementById("cf-photo").innerHTML="<div class=\\"tp-photos\\"><div class=\\"tp-photo\\"><figure><figcaption>Completion photo</figcaption><img src=\\"https://drive.google.com/thumbnail?id="+photoId+"&sz=w600\\" style=\\"max-width:100%;max-height:220px;border-radius:10px;border:1px solid #ececec\\"></figure></div></div>";}var d=document.getElementById("done");d.style.display="block";d.innerHTML="\\u2713 Submitted, pending approval";}'
    + 'function cfFile(input){tpReadFile(input,function(res){if(!res)return;var a=document.getElementById("act");a.innerHTML="<span class=\\"tp-hint\\">Uploading photo\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Upload failed")+"</span>";return;}cfDoneUi(r.photoId);'
    + '}).withFailureHandler(function(){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Upload failed. Please retry.</span>";}).submitIssueCompletion(' + JSON.stringify(id) + ',res.dataUrl,res.name);});}'
    + 'function cfDone(){var a=document.getElementById("act");a.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Could not save")+"</span>";return;}cfDoneUi("");'
    + '}).withFailureHandler(function(){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Could not save. Please retry.</span>";}).submitIssueCompletion(' + JSON.stringify(id) + ',"","");}'
    + '</script>';
  return swissShell_(tpStyles_() + inner + tpSharedJs_(), 'Space Status');
}

function teamPortal_(team, readOnly, embedded) {
  if (!team) return htmlPage_('Invalid link', 'This team link is not recognized.');
  const data = listTeamIssues_(team);
  const issues = data.open;
  const pendingCount = issues.filter(function (x) { return x.pending; }).length;
  const activeCount = issues.length - pendingCount;
  const overdue = issues.filter(function (x) { return x.overdue; }).length;

  let inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#9a958c">Project Teams Ops Hub</div>'
    + '<div class="swh" style="font-size:34px;font-weight:800;letter-spacing:-.035em;line-height:1;margin-top:10px">' + escapeHtml_(team) + '</div>'
    + '<div style="width:46px;height:3px;background:#b31b1b;margin-top:12px"></div>';

  inner += '<div class="ic-summary" style="margin-top:18px">'
    + '<span class="ic-sum"><b id="sum-open">' + activeCount + '</b> open</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum"><b id="sum-pending">' + pendingCount + '</b> pending</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum ic-sum--danger"><b id="sum-over">' + overdue + '</b> overdue</span>'
    + (data.resolved ? '<span class="ic-dot"></span><span class="ic-sum"><b>' + data.resolved + '</b> done</span>' : '')
    + '</div>';

  if (!issues.length) {
    inner += '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:30px">No open issues. Everything is in good shape.</div>';
    return swissShell_(tpStyles_() + inner + tpSharedJs_(), 'Space Status - ' + team, false, embedded);
  }

  let idx = 0;
  const renderCard = function (it) {
    const rid = 'iss' + (idx++);
    const isPending = it.pending;
    const dl = it.deadline ? fmtShort_(it.deadline) : 'No set deadline';
    const dueCls = it.overdue ? 'due due--late' : 'due';
    const isSentBack = !isPending && !!it.sentBackReason;
    const chip = isPending
      ? '<span class="tp-pill tp-pill--pending">Pending approval</span>'
      : (isSentBack
        ? '<span class="tp-pill tp-pill--sent">Sent back</span>'
        : (it.overdue
          ? '<span class="chip chip--overdue">Overdue</span>'
          : (it.deadline ? '<span class="chip chip--due">' + daysLeftLabel_(it.deadline) + '</span>' : '')));
    const statusTxt = isPending ? 'Submitted, awaiting approval' : (isSentBack ? 'Sent back' : ('Due ' + escapeHtml_(dl)));
    const accent = COLOR_HEX[it.color] || '#d6d3ce';
    let foot = '';
    if (!readOnly) foot = isPending ? icPendingFoot_(rid, it.token) : icOpenFoot_(rid, it.token, it.photoOptional);
    return '<div class="card" id="' + rid + '" style="border-left:4px solid ' + accent + ';border-right:4px solid ' + accent + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-po="' + (it.photoOptional ? '1' : '0') + '">'
      + '<div class="card-body">'
      +   '<div class="card-head"><div>'
      +     '<div class="card-title">' + escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue') + '</div></div>'
      +     '<span id="' + rid + '-pill">' + chip + '</span>'
      +   '</div>'
      +   (it.action ? '<div class="card-field"><span class="card-flabel">Required action</span><div class="card-action">' + escapeHtml_(phrase_(it.action)) + '</div></div>' : '')
      +   (it.details ? '<div class="card-field"><span class="card-flabel">Details</span><div class="card-details">' + escapeHtml_(it.details) + '</div></div>' : '')
      +   (it.photos && it.photos.length ? photoStrip_(it.photos) : '')
      +   icPhotoBlock_(rid, it.completionPhoto)
      +   icNoteContainer_(rid, isPending, it.sentBackReason, !readOnly, it.photoOptional)
      + '</div>'
      + '<div class="card-foot">'
      +   '<span id="' + rid + '-status" class="' + (isPending ? 'due' : dueCls) + '">' + statusTxt + '</span>'
      +   (readOnly ? '' : '<span id="' + rid + '-act" class="btn-row">' + foot + '</span>')
      + '</div>'
      + '</div>';
  };

  const pendingList = issues.filter(function (x) { return x.pending; });
  const activeList = issues.filter(function (x) { return !x.pending; });
  const overdueList = activeList.filter(function (x) { return x.overdue; });
  const openList = activeList.filter(function (x) { return !x.overdue; });
  const sectionHeader = function (label, count, color) {
    return '<div style="display:flex;align-items:center;gap:11px;margin:28px 0 14px">'
      + '<span style="font-size:11px;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:' + color + '">' + label + '</span>'
      + '<span style="font-size:12px;color:#bbb;font-weight:600">' + count + '</span>'
      + '<span style="flex:1;height:1px;background:#e8e8e8"></span></div>';
  };

  if (overdueList.length) {
    inner += sectionHeader('Overdue', overdueList.length, '#b31b1b');
    overdueList.forEach(function (it) { inner += renderCard(it); });
  }
  if (openList.length) {
    inner += sectionHeader('Open', openList.length, '#8a857c');
    openList.forEach(function (it) { inner += renderCard(it); });
  }
  if (pendingList.length) {
    inner += sectionHeader('Pending approval', pendingList.length, '#b06a00');
    pendingList.forEach(function (it) { inner += renderCard(it); });
  }

  const scripts = tpSharedJs_() + (readOnly ? '' : icClientJs_());
  return swissShell_(tpStyles_() + inner + scripts, 'Space Status - ' + team, false, embedded);
}

function pickerPage_() {
  const names = teamNames_();
  const url = ScriptApp.getService().getUrl();
  let opts = '<option value="" disabled selected>Choose your team…</option>';
  names.forEach(function (n) { opts += '<option value="' + escapeHtml_(n) + '">' + escapeHtml_(n) + '</option>'; });
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#9a958c">Project Teams Ops Hub</div>'
    + '<div class="swh" style="font-size:34px;font-weight:800;letter-spacing:-.035em;line-height:1.05;margin-top:10px">Your team&rsquo;s space status</div>'
    + '<div style="width:46px;height:3px;background:#b31b1b;margin-top:12px"></div>'
    + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:18px">Select your team to see its open space issues and deadlines.</div>'
    + '<form action="' + escapeHtml_(url) + '" method="get" target="_top" style="margin-top:22px;display:flex;flex-wrap:wrap;gap:12px">'
    +   '<select name="view" required style="flex:1;min-width:240px;font-size:15px;padding:12px 14px;border:1.5px solid #ddd;border-radius:8px;background:#fff;color:#111">' + opts + '</select>'
    +   '<button type="submit" class="b" style="font-size:14px;font-weight:700;color:#fff;background:#b31b1b;border:none;padding:0 24px;border-radius:8px;cursor:pointer">View open issues</button>'
    + '</form>';
  return swissShell_(inner, 'Space Status');
}

function teamNames_() {
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const c = v[0].map(norm_).indexOf(norm_('Team'));
  if (c < 0) return [];
  const seen = {}, out = [];
  for (let i = 1; i < v.length; i++) {
    const t = String(v[i][c]).trim();
    if (t && !seen[norm_(t)]) { seen[norm_(t)] = 1; out.push(t); }
  }
  out.sort(function (a, b) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : (x > y ? 1 : 0); });
  return out;
}

function listTeamIssues_(teamName) {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const ci = function (n) { return H.indexOf(norm_(n)); };
  const cTeam = ci(CONFIG.headers.team), cStatus = ci(CONFIG.headers.status), cIssue = ci(CONFIG.headers.issueType),
        cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details), cTs = ci(CONFIG.headers.timestamp),
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader), cPhoto = ci(CONFIG.headers.photo),
        cCompletedAt = ci(CONFIG.completedAtHeader), cCompletionPhoto = ci(CONFIG.completionPhotoHeader), cSentBack = ci(CONFIG.sentBackHeader);
  const open = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTeam < 0 || norm_(v[i][cTeam]) !== norm_(teamName)) continue;
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) { resolved++; continue; }
    const pending = cCompletedAt >= 0 && v[i][cCompletedAt] ? true : false;
    const color = parseColor_(v[i][cStatus]);
    let deadline = null, overdue = false;
    if (color && color !== 'purple') {
      const rd = new Date(v[i][cTs]);
      if (!isNaN(rd.getTime())) { deadline = new Date(rd); deadline.setDate(deadline.getDate() + SLA_DAYS[color]); deadline.setHours(0, 0, 0, 0); overdue = today > deadline; }
    }
    if (pending) overdue = false;   // waiting on review, not late
    open.push({
      token: String(v[i][cTok]),
      issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
      action: cAction >= 0 ? String(v[i][cAction]).trim() : '',
      photoOptional: icPhotoOptional_(cAction >= 0 ? v[i][cAction] : ''),
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      photos: cPhoto >= 0 ? extractFileIds_(v[i][cPhoto]) : [],
      color: color, deadline: deadline, overdue: overdue,
      pending: pending, completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
      sentBackReason: (!pending && cSentBack >= 0) ? String(v[i][cSentBack] || '').trim() : '',
    });
  }
  open.sort(function (a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (a.deadline && b.deadline) return a.deadline - b.deadline;
    if (a.deadline) return -1; if (b.deadline) return 1; return 0;
  });
  return { open: open, resolved: resolved };
}

function listAllIssues_() {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const ci = function (n) { return H.indexOf(norm_(n)); };
  const cTeam = ci(CONFIG.headers.team), cStatus = ci(CONFIG.headers.status), cIssue = ci(CONFIG.headers.issueType),
        cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details), cTs = ci(CONFIG.headers.timestamp),
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader), cPhoto = ci(CONFIG.headers.photo),
        cCompletedAt = ci(CONFIG.completedAtHeader), cCompletionPhoto = ci(CONFIG.completionPhotoHeader), cSentBack = ci(CONFIG.sentBackHeader);
  const open = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) { resolved++; continue; }
    const pending = cCompletedAt >= 0 && v[i][cCompletedAt] ? true : false;
    const color = parseColor_(v[i][cStatus]);
    let deadline = null, overdue = false;
    if (color && color !== 'purple') {
      const rd = new Date(v[i][cTs]);
      if (!isNaN(rd.getTime())) { deadline = new Date(rd); deadline.setDate(deadline.getDate() + SLA_DAYS[color]); deadline.setHours(0, 0, 0, 0); overdue = today > deadline; }
    }
    if (pending) overdue = false;   // the team already did their part; it is waiting on review, not late
    open.push({
      team: cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
      token: String(v[i][cTok]),
      issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
      action: cAction >= 0 ? String(v[i][cAction]).trim() : '',
      photoOptional: icPhotoOptional_(cAction >= 0 ? v[i][cAction] : ''),
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      photos: cPhoto >= 0 ? extractFileIds_(v[i][cPhoto]) : [],
      color: color, deadline: deadline, overdue: overdue,
      pending: pending, completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
      sentBackReason: (!pending && cSentBack >= 0) ? String(v[i][cSentBack] || '').trim() : '',
    });
  }
  open.sort(function (a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (a.deadline && b.deadline) return a.deadline - b.deadline;
    if (a.deadline) return -1; if (b.deadline) return 1; return 0;
  });
  return { open: open, resolved: resolved };
}

// Every open issue across all teams. Completing one now requires an evidence
// photo -> confetti -> Pending approval, then an admin approves/sends back.
// Embedded in the hub panel when embed=1.
function allIssuesPage_(embedded, admin) {
  const data = listAllIssues_();
  const issues = data.open;
  const pendingList = issues.filter(function (x) { return x.pending; });
  const active = issues.filter(function (x) { return !x.pending; });   // not yet submitted
  const overdueList = active.filter(function (x) { return x.overdue; });
  const openList = active.filter(function (x) { return !x.overdue; });
  const overdue = overdueList.length;
  const teamSet = {};
  issues.forEach(function (it) { if (it.team) teamSet[it.team] = 1; });
  const teamOpts = Object.keys(teamSet).sort().map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('');

  let inner = '';
  if (!embedded) {
    inner += '<div class="page-head">'
      + '<div class="page-kicker">Project Teams Ops Hub</div>'
      + '<div class="page-title">Open space issues</div>'
      + '<div class="page-rule"></div>'
      + '</div>';
  }


  inner += '<div class="ic-summary">'
    + '<span class="ic-sum"><b id="sum-open">' + active.length + '</b> open</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum"><b id="sum-pending">' + pendingList.length + '</b> pending</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum ic-sum--danger"><b id="sum-over">' + overdue + '</b> overdue</span>'
    + '</div>';

  inner += '<div class="filters">'
    + '<div class="search-wrap"><input id="q" type="search" placeholder="Search team, issue, details…" oninput="flt()"></div>'
    + '<select id="team" onchange="flt()"><option value="">All teams</option>' + teamOpts + '</select>'
    + '<label class="toggle"><input id="odue" type="checkbox" onchange="this.closest(\'.toggle\').classList.toggle(\'is-on\', this.checked); flt()"> Overdue only</label>'
    + '</div>';

  if (!issues.length) {
    inner += '<div class="empty">No open issues. Everything is in good shape.</div>';
    return swissShell_(tpStyles_() + inner + tpSharedJs_() + tpAdminRevealJs_(admin), 'Open issues', true, embedded);
  }

  const sectionHead = function (label, count, cls) {
    return '<div class="section-head">'
      + '<span class="section-label ' + cls + '">' + label + '</span>'
      + '<span class="section-count">' + count + '</span>'
      + '<span class="section-rule"></span></div>';
  };

  let idx = 0;
  const teams = icTeamNames_();
  const fieldOpts = icFieldOptions_();
  const renderCard = function (it) {
    const rid = 'iss' + (idx++);
    const isPending = it.pending;
    const dl = it.deadline ? fmtShort_(it.deadline) : 'No set deadline';
    const dueCls = it.overdue ? 'due due--late' : 'due';
    const isSentBack = !isPending && !!it.sentBackReason;
    const chip = isPending
      ? '<span class="tp-pill tp-pill--pending">Pending approval</span>'
      : (isSentBack
        ? '<span class="tp-pill tp-pill--sent">Sent back</span>'
        : (it.overdue
          ? '<span class="chip chip--overdue">Overdue</span>'
          : (it.deadline ? '<span class="chip chip--due">' + daysLeftLabel_(it.deadline) + '</span>' : '')));
    const hay = ((it.team || '') + ' ' + (it.issueType || '') + ' ' + (it.action || '') + ' ' + (it.details || '')).toLowerCase();
    const accent = COLOR_HEX[it.color] || '#d6d3ce';
    const statusTxt = isPending ? 'Submitted, awaiting approval' : (isSentBack ? 'Sent back' : ('Due ' + escapeHtml_(dl)));
    const foot = isPending ? icPendingFoot_(rid, it.token) : icOpenFoot_(rid, it.token, it.photoOptional);
    return '<div class="card" id="' + rid + '" style="border-left:4px solid ' + accent + ';border-right:4px solid ' + accent + '" data-team="' + escapeHtml_(it.team) + '" data-over="' + (it.overdue ? '1' : '0') + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-po="' + (it.photoOptional ? '1' : '0') + '" data-hay="' + escapeHtml_(hay) + '">'
      + '<div class="card-body">'
      +   '<div id="' + rid + '-view">'
      +     '<div class="card-head">'
      +       '<div><div class="card-team" id="' + rid + '-vteam">' + escapeHtml_(it.team || 'Unassigned') + '</div>'
      +       '<div class="card-title" id="' + rid + '-vtype">' + escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue') + '</div></div>'
      +       '<span id="' + rid + '-pill">' + chip + '</span>'
      +     '</div>'
      +     '<div class="card-field" id="' + rid + '-vaction-wrap"' + (it.action ? '' : ' hidden') + '><span class="card-flabel">Required action</span><div class="card-action" id="' + rid + '-vaction">' + escapeHtml_(phrase_(it.action)) + '</div></div>'
      +     '<div class="card-field" id="' + rid + '-vdetails-wrap"' + (it.details ? '' : ' hidden') + '><span class="card-flabel">Details</span><div class="card-details" id="' + rid + '-vdetails">' + escapeHtml_(it.details) + '</div></div>'
      +     (it.photos && it.photos.length ? photoStrip_(it.photos) : '')
      +     icPhotoBlock_(rid, it.completionPhoto)
      +     icNoteContainer_(rid, isPending, it.sentBackReason, true, it.photoOptional)
      +   '</div>'
      +   icEditForm_(rid, it, teams, fieldOpts.issueTypes)
      + '</div>'
      + '<div class="card-foot">'
      +   '<span id="' + rid + '-status" class="' + (isPending ? 'due' : dueCls) + '">' + statusTxt + '</span>'
      +   '<span id="' + rid + '-act" class="btn-row">' + foot + '</span>'
      +   '<span class="tp-admin btn-row" hidden><button type="button" class="btn btn-ghost" onclick="icEditOpen(\'' + rid + '\')">Edit</button>'
      +     '<span id="' + rid + '-delwrap"><button type="button" class="btn btn-ghost tp-del" onclick="icDelOpen(\'' + rid + '\',\'' + it.token + '\')">Delete</button></span></span>'
      + '</div>'
      + '</div>';
  };

  // Pending approval first: these are the items waiting on the admin's action.
  if (pendingList.length) {
    inner += sectionHead('Pending approval', pendingList.length, 'section-label--late');
    pendingList.forEach(function (it) { inner += renderCard(it); });
  }
  if (overdueList.length) {
    inner += sectionHead('Overdue', overdueList.length, 'section-label--late');
    overdueList.forEach(function (it) { inner += renderCard(it); });
  }
  if (openList.length) {
    inner += sectionHead('Open', openList.length, 'section-label--open');
    openList.forEach(function (it) { inner += renderCard(it); });
  }

  inner += '<div id="empty" class="empty" style="display:none">No issues match those filters.</div>';

  inner += '<script>'
    + 'function flt(){var q=document.getElementById("q").value.toLowerCase().trim();var tm=document.getElementById("team").value;var od=document.getElementById("odue").checked;var n=0;'
    +   'document.querySelectorAll(".card").forEach(function(c){var ok=(!q||c.dataset.hay.indexOf(q)>=0)&&(!tm||c.dataset.team===tm)&&(!od||c.dataset.over==="1");c.style.display=ok?"":"none";if(ok)n++;});'
    +   'document.getElementById("empty").style.display=n?"none":"block";}'
    + '</script>';
  return swissShell_(tpStyles_() + inner + tpSharedJs_() + icClientJs_() + tpAdminRevealJs_(admin), 'Open issues', true, embedded);
}

// Read-only registry table (Equipment or Inventory) for the hub. Phase 1 CMMS.
function registryPage_(which, embedded) {
  const inv = String(which).toLowerCase() === 'inventory';
  const tab = inv ? 'Inventory' : 'Equipment';
  const title = inv ? 'Inventory' : 'Equipment registry';
  const data = readTab_(tab);
  const H = data.headers, rows = data.rows;

  let head = '';
  if (!embedded) {
    head = '<div class="page-head"><div class="page-kicker">Ops registry</div>'
      + '<div class="page-title">' + escapeHtml_(title) + '</div><div class="page-rule"></div></div>';
  }

  if (!rows.length) {
    return swissShell_(head + '<div class="empty">Nothing here yet. Add rows in the "' + escapeHtml_(tab) + '" tab.</div>', title, true, embedded);
  }

  const idx = function (name) {
    for (let i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(name)) return i; }
    return -1;
  };
  const cOnHand = idx('On hand'), cTeam = idx('Owning team');

  let controls = '';
  if (!inv && cTeam >= 0) {
    const seen = {}, opts = [];
    rows.forEach(function (r) { const t = String(r[cTeam]).trim(); if (t && !seen[t.toLowerCase()]) { seen[t.toLowerCase()] = 1; opts.push(t); } });
    opts.sort();
    controls += '<select id="team" onchange="flt()"><option value="">All teams</option>'
      + opts.map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('') + '</select>';
  }

  let inner = head
    + '<div class="filters">'
    +   '<div class="search-wrap"><input id="q" type="search" placeholder="Search ' + escapeHtml_(tab.toLowerCase()) + '" oninput="flt()"></div>'
    +   controls
    + '</div>'
    + '<div class="reg-scroll"><table class="reg-table"><thead><tr>';
  if (inv) inner += '<th></th>';
  H.forEach(function (h) { inner += '<th>' + escapeHtml_(h) + '</th>'; });
  inner += '</tr></thead><tbody>';

  rows.forEach(function (r) {
    const out = inv && cOnHand >= 0 && isNum_(r[cOnHand]) && Number(r[cOnHand]) === 0;
    const hay = r.map(function (c) { return String(c); }).join(' ').toLowerCase();
    const team = (!inv && cTeam >= 0) ? String(r[cTeam]).trim() : '';
    inner += '<tr class="reg-row" data-hay="' + escapeHtml_(hay) + '" data-team="' + escapeHtml_(team) + '">';
    if (inv) inner += '<td>' + (out ? '<span class="reg-chip">Out</span>' : '') + '</td>';
    r.forEach(function (c, i) {
      const cls = (inv && i === cOnHand && out) ? ' class="reg-lowval"' : '';
      inner += '<td' + cls + '>' + regCell_(c) + '</td>';
    });
    inner += '</tr>';
  });
  inner += '</tbody></table></div>'
    + '<div id="empty" class="empty" style="display:none">Nothing matches those filters.</div>';

  inner += '<style>'
    + '.reg-scroll{overflow-x:auto;margin-top:4px}'
    + '.reg-table{width:100%;border-collapse:collapse;font-size:13.5px;white-space:nowrap}'
    + '.reg-table th{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#8a857c;text-align:left;padding:0 14px 9px 0;border-bottom:2px solid #1c1a17}'
    + '.reg-table td{padding:11px 14px 11px 0;border-bottom:1px solid #ececea;color:#26231f;vertical-align:top}'
    + '.reg-table tbody tr:hover td{background:#fcfcfb}'
    + '.reg-table a{color:#b31b1b;font-weight:600;text-decoration:none;border-bottom:1px solid #f0c050}'
    + '.reg-lowval{color:#b31b1b;font-weight:800}'
    + '.reg-chip{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0;padding:3px 8px;border-radius:999px}'
    + '</style>';

  inner += '<script>'
    + 'function flt(){var q=document.getElementById("q").value.toLowerCase().trim();'
    +   'var tmEl=document.getElementById("team");var tm=tmEl?tmEl.value:"";var n=0;'
    +   'document.querySelectorAll(".reg-row").forEach(function(r){var ok=(!q||r.dataset.hay.indexOf(q)>=0)&&(!tm||r.dataset.team===tm);r.style.display=ok?"":"none";if(ok)n++;});'
    +   'document.getElementById("empty").style.display=n?"none":"block";}'
    + '</script>';

  return swissShell_(inner, title, true, embedded);
}

function regCell_(v) {
  if (v instanceof Date) return escapeHtml_(fmtShort_(v));
  const s = String(v == null ? '' : v).trim();
  if (/^https?:\/\//i.test(s)) return '<a href="' + escapeHtml_(s) + '" target="_blank" rel="noopener">link</a>';
  return escapeHtml_(s);
}

function isNum_(v) { return v !== '' && v != null && !isNaN(Number(v)); }

function readTab_(name) {
  const sh = registrySs_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  const v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [] };
  const headers = v[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return String(c).trim() === ''; })) continue;
    rows.push(v[i]);
  }
  return { headers: headers, rows: rows };
}

// Drive file IDs from a "Photo Upload" cell (one or more comma-separated URLs).
function extractFileIds_(s) {
  const out = [], seen = {};
  const str = String(s == null ? '' : s);
  const re = /[-\w]{25,}/g;
  let m;
  while ((m = re.exec(str)) !== null) { if (!seen[m[0]]) { seen[m[0]] = 1; out.push(m[0]); } }
  return out;
}

// Thumbnails for the portal cards. One photo -> larger; several -> a strip.
function photoStrip_(ids) {
  if (ids.length === 1) {
    const id = ids[0];
    return '<a href="https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view" target="_blank" rel="noopener" style="display:inline-block;line-height:0;margin-top:12px">'
      + '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w800" alt="Reported photo" loading="lazy" style="max-width:100%;max-height:240px;border-radius:10px;border:1px solid #ececec">'
      + '</a>';
  }
  let s = '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px">';
  ids.forEach(function (id) {
    s += '<a href="https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view" target="_blank" rel="noopener" style="display:block;line-height:0">'
      + '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w400" alt="Reported photo" loading="lazy" style="height:96px;width:auto;border-radius:8px;border:1px solid #ececec">'
      + '</a>';
  });
  return s + '</div>';
}

// One-time backfill: make existing reported photos viewable by link so the portal can show them.
function sharePhotos() {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  const cPhoto = v[0].map(norm_).indexOf(norm_(CONFIG.headers.photo));
  if (cPhoto < 0) { Logger.log('No "' + CONFIG.headers.photo + '" column.'); return; }
  let shared = 0, failed = 0;
  for (let i = 1; i < v.length; i++) {
    extractFileIds_(v[i][cPhoto]).forEach(function (id) {
      try { DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); shared++; }
      catch (err) { failed++; Logger.log('Share failed ' + id + ': ' + err); }
    });
  }
  Logger.log('Photos shared: ' + shared + ' | failed: ' + failed);
}

function confirmAddressed(id) {
  const info = findIssue_(id);
  if (!info) return 'We could not find that issue.';
  if (info.addressed) return 'Already addressed on ' + fmtShort_(info.addressed) + '.';
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  sh.getRange(info.rowIndex, ensureColumn_(sh, CONFIG.addressedHeader)).setValue(new Date());
  return 'Marked as addressed.';
}

function findIssue_(id) {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const cTok = H.indexOf(norm_(CONFIG.issueTokenHeader));
  if (cTok < 0) return null;
  const cTeam = H.indexOf(norm_(CONFIG.headers.team));
  const cIssue = H.indexOf(norm_(CONFIG.headers.issueType));
  const cAction = H.indexOf(norm_(CONFIG.headers.action));
  const cAddr = H.indexOf(norm_(CONFIG.addressedHeader));
  const cComp = H.indexOf(norm_(CONFIG.completedAtHeader));
  const cSent = H.indexOf(norm_(CONFIG.sentBackHeader));
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][cTok]) === String(id)) {
      return {
        rowIndex: i + 1,
        team: cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
        issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
        action: cAction >= 0 ? String(v[i][cAction]).trim() : '',
        addressed: cAddr >= 0 ? v[i][cAddr] : '',
        completedAt: cComp >= 0 ? v[i][cComp] : '',
        sentBackReason: cSent >= 0 ? String(v[i][cSent] || '').trim() : '',
      };
    }
  }
  return null;
}

function stampRow_(sheet, rowNum, token) {
  sheet.getRange(rowNum, ensureColumn_(sheet, CONFIG.issueTokenHeader)).setValue(token);
  sheet.getRange(rowNum, ensureColumn_(sheet, CONFIG.notifiedHeader)).setValue(new Date());
}

function ensureColumn_(sheet, header) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_);
  let i = headers.indexOf(norm_(header));
  if (i >= 0) return i + 1;
  sheet.getRange(1, lastCol + 1).setValue(header);
  return lastCol + 1;
}

function fillTeamTokens() {
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  const tokCol = ensureColumn_(sh, 'Team token');
  const teamCol = ensureColumn_(sh, 'Team');
  let added = 0;
  for (let r = 2; r <= sh.getLastRow(); r++) {
    const team = String(sh.getRange(r, teamCol).getValue() || '').trim();
    const tok = String(sh.getRange(r, tokCol).getValue() || '').trim();
    if (team && !tok) { sh.getRange(r, tokCol).setValue(newToken_()); added++; }
  }
  Logger.log('Tokens added: ' + added);
}

function ss_() { return SpreadsheetApp.openById(CONFIG.spreadsheetId); }
function registrySs_() { return SpreadsheetApp.openById(CONFIG.registrySpreadsheetId); }
function newToken_() { return Utilities.getUuid().replace(/-/g, ''); }
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
function lcFirst_(s) { return s ? s.charAt(0).toLowerCase() + s.slice(1) : s; }
function phrase_(s) { return String(s == null ? '' : s).replace(/\s*\/\s*/g, ' or ').trim(); }
function escapeHtml_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function uniqEmails_(arr) {
  const seen = {}, out = [];
  (arr || []).forEach(function (e) {
    const v = String(e || '').trim();
    if (v && v.indexOf('@') > 0 && !seen[v.toLowerCase()]) { seen[v.toLowerCase()] = true; out.push(v); }
  });
  return out;
}

function listResponseHeaders() {
  const ss = ss_();
  const sh = ss.getSheetByName(CONFIG.responsesSheet);
  if (!sh) { Logger.log('Tabs: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(' | ')); return; }
  Logger.log('Headers in "' + CONFIG.responsesSheet + '":');
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) { Logger.log('  [' + (i + 1) + '] ' + h); });
}