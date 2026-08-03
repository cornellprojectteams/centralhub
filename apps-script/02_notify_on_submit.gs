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
  const reportedOn = data.timestamp ? fmtDate_(new Date(data.timestamp)) : '';
  out += L('This is a follow-up on a space issue in your team\'s space'
    + (reportedOn ? ', reported on <b>' + escapeHtml_(reportedOn) + '</b>,' : '')
    + ' that has not yet been marked as addressed. It was due on <b>' + escapeHtml_(fmtDate_(deadline)) + '</b>.');
  let pb = '';
  if (issue) pb += 'The concern relates to ' + escapeHtml_(lcFirst_(issue));
  if (action) pb += (pb ? ', and we ask that you ' : 'We ask that you ') + escapeHtml_(lcFirst_(action));
  if (pb) out += L(pb + '.');
  if (data.details) out += L('The reporter noted: &ldquo;' + escapeHtml_(data.details) + '&rdquo;');
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
  if (p.module === 'registry-dash') return registryDashboardPage_(embed);   // CMMS dashboard: equipment, inventory, maintenance, action items
  if (p.team) { const t = lookupTeamByToken_(p.team); return t ? teamPortal_(t, false) : htmlPage_('Invalid link', 'This team link is not recognized.'); }
  if (p.view === 'all') return allIssuesPage_(embed, admin);
  if (p.registry === 'labels') return regLabelsPage_(String(p.which || 'equipment'));   // printable QR labels
  if (p.registry === 'item') return regItemPage_(String(p.which || 'equipment'), String(p.id || ''), admin);   // single item (scan target)
  if (p.registry) return registryPage_(String(p.registry), embed, admin);   // Equipment / Inventory tables (editable in admin mode)
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
    + '.btn-confirm{color:#fff;background:#157a47;box-shadow:0 1px 2px rgba(21,122,71,.2)}'
    + '.btn-confirm:hover{background:#12693c}'
    + '.btn-confirm:active{transform:translateY(1px)}'
    + '.btn-ghost{color:#57534e;background:#faf9f6;border:1.5px solid #e2ddd6;padding:9px 15px;border-radius:9px;font-weight:700}'
    + '.btn-ghost:hover{color:#292524;border-color:#b5b0a8;background:#fff}'
    + '.btn-ghost:active{background:#f0efe9}'
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
    + '.card-foot .btn{padding:8px 14px;font-size:12.5px;border-radius:9px}'
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
    return '<div class="card" id="' + rid + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-po="' + (it.photoOptional ? '1' : '0') + '">'
      + '<div class="card-accent" id="' + rid + '-accent" style="background:' + accent + '"></div>'
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
    return '<div class="card" id="' + rid + '" data-team="' + escapeHtml_(it.team) + '" data-over="' + (it.overdue ? '1' : '0') + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-po="' + (it.photoOptional ? '1' : '0') + '" data-hay="' + escapeHtml_(hay) + '">'
      + '<div class="card-accent" id="' + rid + '-accent" style="background:' + accent + '"></div>'
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
// CMMS dashboard: one read-only glance at the shop - open work, stock to reorder,
// maintenance coming due, and the live action-item feed. Reads the registry spreadsheet.
function registryDashboardPage_(embedded) {
  var eq = readTab_('Equipment');
  var invd = readTab_('Inventory');
  var pm = readTab_('Maintenance');
  var act = readTab_('Action items');

  var col = function (H, name) { for (var i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(name)) return i; } return -1; };
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var DAY = 86400000;
  var toDate = function (v) { if (v instanceof Date) return v; if (v === '' || v == null) return null; var d = new Date(v); return isNaN(d.getTime()) ? null : d; };
  var startOf = function (d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
  var truthy = function (v) { if (v === true) return true; var s = String(v == null ? '' : v).trim().toLowerCase(); return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'active' || s === 'x'; };
  var dueLabel = function (d, diff) {
    if (!d || diff === null) return { txt: 'No date set', cls: 'is-ok' };
    if (diff < 0) return { txt: 'Overdue ' + (-diff) + 'd', cls: 'is-over' };
    if (diff === 0) return { txt: 'Due today', cls: 'is-over' };
    if (diff <= 7) return { txt: 'Due in ' + diff + 'd', cls: 'is-soon' };
    return { txt: fmtShort_(d), cls: 'is-ok' };
  };
  var statusPill = function (s) {
    var t = String(s || '').toLowerCase();
    if (t.indexOf('block') >= 0 || t.indexOf('urgent') >= 0) return 'pill-out';
    if (t.indexOf('progress') >= 0 || t.indexOf('order') >= 0 || t.indexOf('wait') >= 0) return 'pill-prog';
    return 'pill-open';
  };

  // Inventory: healthy vs low vs out, and a reorder list.
  var iOn = col(invd.headers, 'On hand'), iRe = col(invd.headers, 'Reorder point'),
      iItem = col(invd.headers, 'Item'), iUnit = col(invd.headers, 'Unit'), iImg = col(invd.headers, 'Image');
  var invCounted = 0, outCount = 0, lowCount = 0, healthy = 0, reorder = [];
  invd.rows.forEach(function (r) {
    if (iOn < 0 || !isNum_(r[iOn])) return;
    var on = Number(r[iOn]);
    var re = (iRe >= 0 && isNum_(r[iRe])) ? Number(r[iRe]) : 0;
    var name = iItem >= 0 ? String(r[iItem]).trim() : '';
    var unit = iUnit >= 0 ? String(r[iUnit]).trim() : '';
    invCounted++;
    var out = on <= 0;
    var low = on > 0 && re > 0 && on <= re;
    if (out) outCount++; else if (low) lowCount++; else healthy++;
    if (out || low) reorder.push({ name: name, on: on, re: re, unit: unit, out: out, img: iImg >= 0 ? r[iImg] : '' });
  });
  reorder.sort(function (a, b) { if (a.out !== b.out) return a.out ? -1 : 1; return (b.re - b.on) - (a.re - a.on); });

  // Maintenance: active tasks, soonest due first.
  var mAct = col(pm.headers, 'Active'), mNext = col(pm.headers, 'Next due'),
      mTask = col(pm.headers, 'Task'), mArea = col(pm.headers, 'Equipment / area');
  var pmDue = 0, pmList = [];
  pm.rows.forEach(function (r) {
    var active = mAct < 0 ? true : truthy(r[mAct]);
    if (!active) return;
    var next = mNext >= 0 ? toDate(r[mNext]) : null;
    var diff = next ? Math.round((startOf(next) - today) / DAY) : null;
    if (diff !== null && diff <= 0) pmDue++;
    pmList.push({ task: mTask >= 0 ? String(r[mTask]).trim() : '', area: mArea >= 0 ? String(r[mArea]).trim() : '', next: next, diff: diff });
  });
  pmList.sort(function (a, b) { var av = a.diff === null ? 1e9 : a.diff, bv = b.diff === null ? 1e9 : b.diff; return av - bv; });

  // Action items: everything not yet done, soonest due first.
  var aStatus = col(act.headers, 'Status'), aDone = col(act.headers, 'Completed'), aTitle = col(act.headers, 'Title'),
      aLoc = col(act.headers, 'Location'), aEquip = col(act.headers, 'Equipment / part'), aWho = col(act.headers, 'Assigned to'),
      aDue = col(act.headers, 'Due'), aCreated = col(act.headers, 'Created');
  var openActions = [];
  act.rows.forEach(function (r) {
    var st = aStatus >= 0 ? String(r[aStatus] || '').trim() : '';
    var stl = st.toLowerCase();
    var done = (aDone >= 0 && String(r[aDone]).trim() !== '') || stl === 'done' || stl === 'completed' || stl === 'closed';
    if (done) return;
    openActions.push({
      title: aTitle >= 0 ? String(r[aTitle]).trim() : '',
      loc: aLoc >= 0 ? String(r[aLoc]).trim() : '',
      equip: aEquip >= 0 ? String(r[aEquip]).trim() : '',
      who: aWho >= 0 ? String(r[aWho]).trim() : '',
      status: st || 'Open',
      due: aDue >= 0 ? toDate(r[aDue]) : null,
      created: aCreated >= 0 ? toDate(r[aCreated]) : null
    });
  });
  openActions.sort(function (a, b) {
    var ad = a.due ? startOf(a.due).getTime() : 8.64e15, bd = b.due ? startOf(b.due).getTime() : 8.64e15;
    if (ad !== bd) return ad - bd;
    return (b.created ? b.created.getTime() : 0) - (a.created ? a.created.getTime() : 0);
  });

  var assets = eq.rows.length, invItems = invd.rows.length, reorderCount = outCount + lowCount, openCount = openActions.length;

  var cmms = '<style>'
    + '.cmms-list{display:flex;flex-direction:column}'
    + '.cmms-row{display:grid;grid-template-columns:1fr auto;gap:2px 14px;align-items:center;padding:13px 0;border-bottom:1px solid #f0efe9}'
    + '.cmms-row:first-child{padding-top:2px}.cmms-row:last-child{border-bottom:0;padding-bottom:2px}'
    + '.cmms-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:700;color:#26231f;line-height:1.3}'
    + '.cmms-sub{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;color:#8a857c;margin-top:2px}'
    + '.cmms-when{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;white-space:nowrap}'
    + '.cmms-when-wrap{display:flex;align-items:center;gap:10px;justify-self:end;white-space:nowrap}'
    + '.cmms-when.is-over{color:#b31b1b}.cmms-when.is-soon{color:#b06a00}.cmms-when.is-ok{color:#8a857c}'
    + '.pill{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;border-radius:99px;padding:3px 9px;white-space:nowrap}'
    + '.pill-open{color:#b06a00;background:#fdf3df;border:1px solid #f0dca6}'
    + '.pill-out{color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0}'
    + '.pill-prog{color:#2563c9;background:#eaf1fe;border:1px solid #cfe0fb}'
    + '.dash-bar-row{grid-template-columns:26px 1fr auto}'
    + '.dash-rthumb .reg-thumb{width:26px;height:26px;border-radius:7px;overflow:hidden;background:#f6f4ef;border:1px solid #ececea;display:flex;align-items:center;justify-content:center}'
    + '.dash-rthumb .reg-thumb img{width:100%;height:100%;object-fit:cover;display:block}'
    + '.dash-rthumb .reg-thumb-icon{color:#b0a99e}.dash-rthumb .reg-thumb-icon svg{width:15px;height:15px}'
    + '@media(max-width:520px){.cmms-row{grid-template-columns:1fr}.cmms-when-wrap{justify-self:start;margin-top:6px}}'
    + '</style>';

  var inner = tpDashStyles_() + cmms + '<div class="dash">';

  inner += '<div class="dash-hero"><div class="dash-hero-glow"></div>'
    + '<div class="dash-kicker">Ops registry</div>'
    + '<h1 class="dash-h1">Equipment &amp; Inventory</h1>'
    + '<div class="dash-hero-row">'
    +   '<div><div class="dash-hero-num">' + openCount + '</div><div class="dash-hero-lbl">open action item' + (openCount === 1 ? '' : 's') + ' to fix, restock, or maintain</div></div>'
    +   '<div class="dash-hero-mini">'
    +     '<div><b>' + reorderCount + '</b>to reorder</div>'
    +     '<div><b>' + pmDue + '</b>maintenance due</div>'
    +   '</div>'
    + '</div></div>';

  var tile = function (n, lbl, c) { return '<div class="dash-tile" style="--c:' + c + '"><div class="dash-tile-num">' + n + '</div><div class="dash-tile-lbl">' + lbl + '</div></div>'; };
  inner += '<div class="dash-tiles">'
    + tile(assets, 'Assets tracked', '#2563c9')
    + tile(invItems, 'Inventory items', '#0d9488')
    + tile(reorderCount, 'Low / out of stock', '#b31b1b')
    + tile(pmDue, 'Maintenance due', '#e08a1e')
    + '</div>';

  if (!assets && !invItems && !pm.rows.length && !act.rows.length) {
    inner += '<div class="dash-empty">The registry is empty. Run <b>Set up / update tabs</b> in the registry sheet, then add equipment and inventory.</div></div>';
    return swissShell_(inner, 'Equipment & inventory', true, embedded);
  }

  var ring;
  if (!invCounted) {
    ring = '<div class="dash-card"><div class="dash-card-h">Inventory in stock</div><div class="dash-ringwrap"><div class="dash-ring-cap">No inventory counts yet.</div></div></div>';
  } else {
    var inStock = Math.round(healthy / invCounted * 100);
    var offset = 100 - inStock;
    ring = '<div class="dash-card"><div class="dash-card-h">Inventory in stock</div><div class="dash-ringwrap">'
      + '<svg class="dash-ring" viewBox="0 0 42 42" aria-hidden="true">'
      +   '<defs><linearGradient id="dashgrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8f1515"/><stop offset="0.6" stop-color="#d62b2b"/><stop offset="1" stop-color="#f0c050"/></linearGradient></defs>'
      +   '<circle class="dash-ring-bg" cx="21" cy="21" r="15.915"/>'
      +   '<circle class="dash-ring-fg" cx="21" cy="21" r="15.915" stroke-dasharray="100" stroke-dashoffset="100" data-off="' + offset + '" transform="rotate(-90 21 21)"/>'
      +   '<text class="dash-ring-num" x="21" y="20.4" text-anchor="middle">' + inStock + '%</text>'
      +   '<text class="dash-ring-sub" x="21" y="26" text-anchor="middle">IN STOCK</text>'
      + '</svg>'
      + '<div class="dash-ring-cap">' + healthy + ' of ' + invCounted + ' items healthy</div>'
      + '</div></div>';
  }

  var bars = '<div class="dash-card"><div class="dash-card-h">Stock to reorder</div>';
  if (!reorder.length) {
    bars += '<div class="dash-gal-who">Everything is above its reorder point.</div>';
  } else {
    bars += '<div class="dash-bars">';
    reorder.slice(0, 7).forEach(function (it) {
      var frac = it.re > 0 ? Math.max(0, Math.min(1, it.on / it.re)) : (it.on > 0 ? 1 : 0);
      var val = it.out ? '<span class="pill pill-out">Out</span>' : escapeHtml_(it.on + (it.unit ? ' ' + it.unit : '') + ' / ' + it.re);
      bars += '<div><div class="dash-bar-row"><span class="dash-rthumb">' + regThumb_(it.img, it.name, 'inventory') + '</span><span class="dash-bar-name">' + escapeHtml_(it.name || '(unnamed)') + '</span><span class="dash-bar-val">' + val + '</span></div>'
        + '<div class="dash-bar-track"><div class="dash-bar-fill" style="--w:' + frac.toFixed(3) + '"></div></div></div>';
    });
    bars += '</div>';
  }
  bars += '</div>';
  inner += '<div class="dash-grid">' + ring + bars + '</div>';

  if (pmList.length) {
    inner += '<div class="dash-sec">Maintenance</div><div class="dash-card"><div class="cmms-list">';
    pmList.slice(0, 8).forEach(function (m) {
      var w = dueLabel(m.next, m.diff);
      inner += '<div class="cmms-row">'
        + '<div><div class="cmms-title">' + escapeHtml_(m.task || '(untitled)') + '</div>'
        + (m.area ? '<div class="cmms-sub">' + escapeHtml_(m.area) + '</div>' : '') + '</div>'
        + '<div class="cmms-when ' + w.cls + '">' + escapeHtml_(w.txt) + '</div>'
        + '</div>';
    });
    inner += '</div></div>';
  }

  inner += '<div class="dash-sec">Open action items</div>';
  if (!openActions.length) {
    inner += '<div class="dash-empty">No open action items. The shop is caught up.</div>';
  } else {
    inner += '<div class="dash-card"><div class="cmms-list">';
    openActions.slice(0, 12).forEach(function (a) {
      var subBits = [];
      if (a.loc) subBits.push(a.loc);
      if (a.equip) subBits.push(a.equip);
      if (a.who) subBits.push(a.who);
      var due = a.due ? dueLabel(a.due, Math.round((startOf(a.due) - today) / DAY)) : null;
      inner += '<div class="cmms-row">'
        + '<div><div class="cmms-title">' + escapeHtml_(a.title || '(untitled)') + '</div>'
        + (subBits.length ? '<div class="cmms-sub">' + escapeHtml_(subBits.join(' · ')) + '</div>' : '') + '</div>'
        + '<div class="cmms-when-wrap"><span class="pill ' + statusPill(a.status) + '">' + escapeHtml_(a.status) + '</span>'
        + (due ? '<span class="cmms-when ' + due.cls + '">' + escapeHtml_(due.txt) + '</span>' : '') + '</div>'
        + '</div>';
    });
    inner += '</div></div>';
    if (openActions.length > 12) inner += '<div class="dash-gal-who" style="margin-top:10px">Showing 12 of ' + openActions.length + ' open items. Open the action items sheet for the rest.</div>';
  }

  inner += '</div>';
  inner += '<script>requestAnimationFrame(function(){var c=document.querySelector(".dash-ring-fg");if(c)c.style.strokeDashoffset=c.getAttribute("data-off");});</script>';
  return swissShell_(inner, 'Equipment & inventory', true, embedded);
}

// Field spec drives the add/edit forms and identifies each tab's natural key.
function regFields_(which) {
  if (String(which).toLowerCase() === 'inventory') {
    return { tab: 'Inventory', key: 'Item', idPrefix: '', which: 'inventory', noun: 'item',
      fields: [
        { h: 'Item', label: 'Item', req: true }, { h: 'Location', label: 'Location' },
        { h: 'On hand', label: 'On hand', type: 'number' }, { h: 'Unit', label: 'Unit' },
        { h: 'Reorder point', label: 'Reorder point', type: 'number' }, { h: 'Reorder qty', label: 'Reorder qty', type: 'number' },
        { h: 'Supplier', label: 'Supplier' }, { h: 'Product link', label: 'Product link', type: 'url' },
        { h: 'eShop info', label: 'eShop info' }, { h: 'Assign to', label: 'Assign to' },
        { h: 'Last restocked', label: 'Last restocked' }, { h: 'Image', label: 'Image', type: 'image' } ] };
  }
  return { tab: 'Equipment', key: 'Asset ID', idPrefix: 'EQ-', which: 'equipment', noun: 'equipment',
    fields: [
      { h: 'Asset ID', label: 'Asset ID', auto: true }, { h: 'Name', label: 'Name', req: true },
      { h: 'Category', label: 'Category' }, { h: 'Location', label: 'Location' },
      { h: 'Owning team', label: 'Owning team' }, { h: 'Owner', label: 'Owner' },
      { h: 'Status', label: 'Status' }, { h: 'Installed', label: 'Installed' }, { h: 'Notes', label: 'Notes' },
      { h: 'Image', label: 'Image', type: 'image' } ] };
}

// A thumbnail for a row: the stored photo if there is one, else a category icon so
// every item still reads at a glance. Icons are inline SVG, nothing to fetch.
function regThumb_(url, name, kind) {
  const u = String(url == null ? '' : url).trim();
  if (/^https?:\/\//i.test(u) || u.indexOf('data:') === 0) {
    return '<span class="reg-thumb"><img src="' + escapeHtml_(u) + '" loading="lazy" alt="" onerror="this.parentNode.classList.add(\'reg-thumb-broken\')"></span>';
  }
  return '<span class="reg-thumb reg-thumb-icon">' + regCatIcon_(name, kind) + '</span>';
}

function regCatIcon_(name, kind) {
  const t = (String(name || '') + ' ' + String(kind || '')).toLowerCase();
  const has = function (arr) { return arr.some(function (w) { return t.indexOf(w) >= 0; }); };
  const svg = function (paths) { return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">' + paths + '</svg>'; };
  if (has(['drill'])) return svg('<path d="M3 8h9v5H5a2 2 0 0 1-2-2V8z"/><path d="M12 9h4l2-2v6l-2-2h-4"/><path d="M7 13v4"/>');
  if (has(['saw', 'blade'])) return svg('<circle cx="9" cy="12" r="6"/><path d="M9 9v3l2 1"/><path d="M15 12h6"/>');
  if (has(['mill', 'lathe', 'press', 'grinder', 'sander', 'machine', 'pump', 'compressor'])) return svg('<rect x="4" y="4" width="16" height="12" rx="1"/><path d="M8 20h8M12 16v4M8 8h8v4H8z"/>');
  if (has(['wrench', 'socket', 'ratchet', 'allen', 'plier', 'cutter', 'adapter', 'screw', 'punch', 'vice', 'shear', 'hack', 'caliper', 'tool', 'tape'])) return svg('<path d="M14 7a4 4 0 0 0-5 5l-6 6 2 2 6-6a4 4 0 0 0 5-5l-2 2-2-2 2-2z"/>');
  if (has(['cabinet', 'storage', 'shelf', 'bench', 'tote', 'container', 'crate', 'bin', 'reel', 'cart'])) return svg('<rect x="4" y="3" width="16" height="18" rx="1"/><path d="M4 9h16M4 15h16M9 6h.01M9 12h.01M9 18h.01"/>');
  if (has(['filter'])) return svg('<path d="M4 4h16l-6 8v6l-4 2v-8L4 4z"/>');
  if (has(['glove', 'respirat', 'ppe', 'safety', 'flammable', 'shield'])) return svg('<path d="M12 3l7 3v5c0 4-3 7-7 9-4-2-7-5-7-9V6l7-3z"/>');
  return svg('<path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/>');
}

function readTabRows_(name) {
  const sh = registrySs_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  const v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [] };
  const headers = v[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return String(c).trim() === ''; })) continue;
    rows.push({ row: i + 1, cells: v[i] });
  }
  return { headers: headers, rows: rows };
}

function regRawVal_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (v === 0) return '0';
  return String(v == null ? '' : v);
}

// Build the table body + a {rowNumber: {header: value}} map, reused by the page and refreshes.
function regBuildTbody_(which, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const data = readTabRows_(spec.tab);
  const H = data.headers, rows = data.rows;
  const idxOf = function (n) { for (let i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(n)) return i; } return -1; };
  const cOnHand = idxOf('On hand'), cTeam = idxOf('Owning team'), cImg = idxOf('Image');
  const cName = idxOf(inv ? 'Item' : 'Name'), cCat = idxOf('Category');
  let html = ''; const mapParts = [], teams = {};
  rows.forEach(function (rr) {
    const r = rr.cells;
    const out = inv && cOnHand >= 0 && isNum_(r[cOnHand]) && Number(r[cOnHand]) <= 0;
    const hay = r.map(function (c) { return String(c); }).join(' ').toLowerCase();
    const team = (!inv && cTeam >= 0) ? String(r[cTeam]).trim() : '';
    if (team) teams[team] = 1;
    const nm = cName >= 0 ? r[cName] : '', kind = (cCat >= 0 ? r[cCat] : '') + ' ' + spec.which;
    html += '<tr class="reg-row" data-hay="' + escapeHtml_(hay) + '" data-team="' + escapeHtml_(team) + '">';
    if (admin) html += '<td class="reg-actcell"><button type="button" class="reg-mini" onclick="regOpenEdit(' + rr.row + ')">Edit</button><button type="button" class="reg-mini reg-del" onclick="regDeleteRow(' + rr.row + ',this)">Delete</button></td>';
    html += '<td class="reg-thumbcell">' + regThumb_(cImg >= 0 ? r[cImg] : '', nm, kind) + '</td>';
    if (inv) html += '<td>' + (out ? '<span class="reg-chip">Out</span>' : '') + '</td>';
    r.forEach(function (c, i) { if (i === cImg) return; const cls = (inv && i === cOnHand && out) ? ' class="reg-lowval"' : ''; html += '<td' + cls + '>' + regCell_(c) + '</td>'; });
    html += '</tr>';
    const obj = {}; H.forEach(function (h, i) { obj[h] = regRawVal_(r[i]); });
    mapParts.push(JSON.stringify(String(rr.row)) + ':' + JSON.stringify(obj));
  });
  return { html: html, mapJson: '{' + mapParts.join(',') + '}', headers: H, inv: inv, teams: Object.keys(teams).sort(), key: spec.key, imgIdx: cImg };
}

// Client-callable: re-render the table body after a change (no page reload, which blanks the sandbox).
function regRowsHtml(which, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized.' };
  const b = regBuildTbody_(which, true);
  return { ok: true, html: b.html, mapJson: b.mapJson };
}

function registryPage_(which, embedded, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const title = inv ? 'Inventory' : 'Equipment registry';
  const b = regBuildTbody_(which, admin);
  const H = b.headers;

  let head = '';
  if (!embedded) {
    head = '<div class="page-head"><div class="page-kicker">Ops registry</div>'
      + '<div class="page-title">' + escapeHtml_(title) + '</div><div class="page-rule"></div></div>';
  }

  let toolbar = '';
  if (admin) {
    toolbar = regUnlockBar_()
      + '<div class="reg-tools" id="reg-tools">'
      +   '<button type="button" class="btn btn-primary" onclick="regOpenAdd()">+ Add ' + escapeHtml_(spec.noun) + '</button>'
      +   '<button type="button" class="btn btn-ghost" onclick="regScan()">Scan</button>'
      +   (inv ? '<button type="button" class="btn btn-ghost" id="reg-backfill" onclick="regBackfill()">Fetch images</button>' : '')
      +   '<a class="btn btn-ghost" href="?registry=labels&which=' + spec.which + '&admin=1" target="_blank" rel="noopener">Print labels</a>'
      + '</div>';
  }

  let controls = '';
  if (!inv && b.teams.length) {
    controls = '<select id="team" onchange="flt()"><option value="">All teams</option>'
      + b.teams.map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('') + '</select>';
  }

  let inner = '<div id="reg-root">' + head + toolbar
    + '<div class="filters"><div class="search-wrap"><input id="q" type="search" placeholder="Search ' + escapeHtml_(spec.tab.toLowerCase()) + '" oninput="flt()"></div>' + controls + '</div>'
    + '<div class="reg-scroll"><table class="reg-table"><thead><tr>'
    + (admin ? '<th>Actions</th>' : '') + '<th></th>' + (inv ? '<th></th>' : '')
    + H.map(function (h, i) { return i === b.imgIdx ? '' : '<th>' + escapeHtml_(h) + '</th>'; }).join('')
    + '</tr></thead><tbody id="reg-tbody">' + b.html + '</tbody></table></div>'
    + '<div id="empty" class="empty" style="display:none">Nothing matches those filters.</div>';

  if (!b.html) inner += '<div class="empty">Nothing here yet.' + (admin ? ' Unlock, then add the first ' + escapeHtml_(spec.noun) + '.' : '') + '</div>';

  inner += regStyles_() + regFilterJs_();
  if (admin) {
    inner += regFormOverlay_(which)
      + '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_KEY=' + JSON.stringify(b.key)
      + ';var REG_ROWS=' + b.mapJson.replace(/<\//g, '<\\/') + ';var ADMIN_PASS="";</script>'
      + regEditJs_();
  }
  inner += '</div>';
  return swissShell_(inner, title, true, embedded);
}

// Single-item view: the QR-label scan target. Focused card with quick actions.
function regItemPage_(which, id, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  let found = null;
  data.rows.forEach(function (rr) { if (ki >= 0 && String(rr.cells[ki]).trim() === String(id).trim()) found = rr; });

  const head = '<div class="page-head"><div class="page-kicker">' + (inv ? 'Inventory item' : 'Equipment') + '</div>';
  if (!found) {
    return swissShell_('<div id="reg-root">' + head + '<div class="page-title">Not found</div><div class="page-rule"></div></div>'
      + '<div class="empty">No ' + escapeHtml_(spec.noun) + ' matches "' + escapeHtml_(id) + '".</div>'
      + '<p style="margin-top:16px"><a class="btn btn-ghost" href="?registry=' + spec.which + '&admin=1">Back to ' + escapeHtml_(spec.tab.toLowerCase()) + '</a></p></div>', 'Not found', false, false);
  }
  const r = found.cells;
  const hi = function (n) { return H.map(function (h) { return norm_(h); }).indexOf(norm_(n)); };
  const nameI = hi(inv ? 'Item' : 'Name'), imgI = hi('Image'), catI = hi('Category');
  const name = nameI >= 0 ? String(r[nameI]) : String(id);

  const imgUrl = imgI >= 0 ? String(r[imgI]).trim() : '';
  const hero = (/^https?:\/\//i.test(imgUrl) || imgUrl.indexOf('data:') === 0)
    ? '<img class="reg-hero" src="' + escapeHtml_(imgUrl) + '" alt="" onerror="this.style.display=\'none\'">'
    : '<span class="reg-thumb reg-thumb-icon" style="width:72px;height:72px;border-radius:14px;margin-bottom:14px">' + regCatIcon_(name, (catI >= 0 ? r[catI] : '') + ' ' + spec.which) + '</span>';

  let fieldsHtml = '';
  H.forEach(function (h, i) {
    if (i === imgI) return;
    if (!String(r[i]).trim()) return;
    fieldsHtml += '<div class="card-field"><span class="card-flabel">' + escapeHtml_(h) + '</span><div class="card-action">' + regCell_(r[i]) + '</div></div>';
  });

  let inner = '<div id="reg-root">' + head + '<div class="page-title">' + escapeHtml_(name) + '</div><div class="page-rule"></div></div>'
    + '<div class="card"><div class="card-body">' + hero + fieldsHtml + '</div>';
  if (admin) {
    inner += regUnlockBar_()
      + '<div class="card-foot" id="reg-tools"><span class="reg-mini-note">' + escapeHtml_(spec.key) + ': ' + escapeHtml_(String(id)) + '</span>'
      + '<span class="btn-row">'
      + (inv ? '<button type="button" class="btn btn-confirm" onclick="regRestockOne(' + found.row + ')">Mark restocked</button>' : '')
      + '<button type="button" class="btn btn-ghost" onclick="regOpenEdit(' + found.row + ')">Edit</button>'
      + '<button type="button" class="btn btn-ghost reg-del" onclick="regDeleteRow(' + found.row + ',this)">Delete</button>'
      + '</span></div>';
  }
  inner += '</div><p style="margin-top:16px"><a class="btn btn-ghost" href="?registry=' + spec.which + '&admin=1">Back to ' + escapeHtml_(spec.tab.toLowerCase()) + '</a></p>';

  inner += regStyles_();
  if (admin) {
    const one = {}; one[String(found.row)] = (function () { const o = {}; H.forEach(function (h, i) { o[h] = regRawVal_(r[i]); }); return o; })();
    inner += regFormOverlay_(which)
      + '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_KEY=' + JSON.stringify(spec.key)
      + ';var REG_ROWS=' + JSON.stringify(one).replace(/<\//g, '<\\/') + ';var ADMIN_PASS="";var REG_ITEM=1;</script>'
      + regEditJs_();
  }
  inner += '</div>';
  return swissShell_(inner, name, false, false);
}

// Printable QR labels. Each label deep-links to its item page; scan with any phone camera.
function regLabelsPage_(which) {
  const spec = regFields_(which);
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  const ni = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.tab === 'Inventory' ? 'Item' : 'Name'));
  let base = '';
  try { base = ScriptApp.getService().getUrl(); } catch (e) { base = ''; }

  let cards = '';
  data.rows.forEach(function (rr) {
    const key = ki >= 0 ? String(rr.cells[ki]).trim() : '';
    if (!key) return;
    const name = ni >= 0 ? String(rr.cells[ni]).trim() : key;
    const url = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'registry=item&which=' + spec.which + '&id=' + encodeURIComponent(key) + '&admin=1';
    cards += '<div class="lab"><div class="lab-qr" data-url="' + escapeHtml_(url) + '"></div>'
      + '<div class="lab-name">' + escapeHtml_(name) + '</div><div class="lab-id">' + escapeHtml_(key) + '</div></div>';
  });

  const inner = '<div class="lab-head"><div><div class="page-kicker">Ops registry</div><div class="page-title">' + escapeHtml_(spec.tab) + ' labels</div></div>'
    + '<button class="btn btn-primary lab-print" onclick="window.print()">Print</button></div>'
    + '<p class="lab-hint">Print, cut, and stick one on each ' + escapeHtml_(spec.noun) + '. Scanning a label with any phone camera opens its record.</p>'
    + '<div class="lab-grid">' + (cards || '<div class="empty">Nothing to label yet.</div>') + '</div>'
    + '<style>'
    + '.lab-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap}'
    + '.lab-hint{color:#8a857c;font-size:13.5px;margin:6px 0 16px}'
    + '.lab-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px}'
    + '.lab{border:1.5px solid #ddd;border-radius:12px;padding:14px;text-align:center;background:#fff;break-inside:avoid}'
    + '.lab-qr{width:150px;height:150px;margin:0 auto 10px;display:flex;align-items:center;justify-content:center}'
    + '.lab-qr img,.lab-qr canvas{width:150px !important;height:150px !important}'
    + '.lab-qr .lab-fallback{font-size:9px;word-break:break-all;color:#666;line-height:1.3}'
    + '.lab-name{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-weight:800;font-size:14px;color:#111;line-height:1.2}'
    + '.lab-id{font-size:11px;color:#8a857c;margin-top:3px;font-weight:700;letter-spacing:.04em}'
    + '@media print{.lab-print{display:none}.lab-head{margin-bottom:8px}}'
    + '</style>'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>'
    + '<script>(function(){function draw(){document.querySelectorAll(".lab-qr").forEach(function(el){var u=el.getAttribute("data-url");'
    + 'if(window.QRCode){try{new QRCode(el,{text:u,width:150,height:150,correctLevel:QRCode.CorrectLevel.M});return;}catch(e){}}'
    + 'el.innerHTML="<span class=\\"lab-fallback\\">"+u+"</span>";});}'
    + 'if(window.QRCode){draw();}else{var t=0,iv=setInterval(function(){t+=200;if(window.QRCode||t>4000){clearInterval(iv);draw();}},200);}})();</script>';

  return swissShell_(inner, spec.tab + ' labels', true, false);
}

// ---- registry admin: shared styles, filter, unlock, form overlay, edit JS ----

function regStyles_() {
  return '<style>'
    + '.reg-scroll{overflow-x:auto;margin-top:4px}'
    + '.reg-table{width:100%;border-collapse:collapse;font-size:13.5px;white-space:nowrap}'
    + '.reg-table th{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#8a857c;text-align:left;padding:0 14px 9px 0;border-bottom:2px solid #1c1a17}'
    + '.reg-table td{padding:11px 14px 11px 0;border-bottom:1px solid #ececea;color:#26231f;vertical-align:top}'
    + '.reg-table tbody tr:hover td{background:#fcfcfb}'
    + '.reg-table a{color:#b31b1b;font-weight:600;text-decoration:none;border-bottom:1px solid #f0c050}'
    + '.reg-lowval{color:#b31b1b;font-weight:800}'
    + '.reg-chip{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0;padding:3px 8px;border-radius:999px}'
    + '.reg-thumbcell{width:52px;padding-right:10px !important}'
    + '.reg-thumb{display:flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:9px;overflow:hidden;background:#f6f4ef;border:1px solid #ececea}'
    + '.reg-thumb img{width:100%;height:100%;object-fit:cover;display:block}'
    + '.reg-thumb-icon{color:#b0a99e}.reg-thumb-icon svg{width:22px;height:22px}'
    + '.reg-thumb-broken img{display:none}.reg-thumb-broken::after{content:"";width:22px;height:22px;background:#e6e2da;border-radius:5px}'
    + '.reg-tools{display:none;gap:8px;flex-wrap:wrap;margin:14px 0 4px}.reg-unlocked .reg-tools{display:flex}'
    + '.reg-hero{width:100%;max-width:280px;border-radius:14px;border:1px solid #ececea;margin-bottom:14px;display:block}'
    + '.reg-imgrow{grid-column:1 / -1;display:flex;gap:12px;align-items:center}'
    + '.reg-imgrow .reg-preview{flex:0 0 auto;width:56px;height:56px;border-radius:10px;background:#f6f4ef;border:1px solid #ececea;overflow:hidden;display:flex;align-items:center;justify-content:center;color:#b0a99e}'
    + '.reg-imgrow .reg-preview img{width:100%;height:100%;object-fit:cover}.reg-imgrow .reg-imgfields{flex:1;min-width:0;display:flex;flex-direction:column;gap:5px}'
    + '.reg-imgrow .reg-imgfields .reg-imgtop{display:flex;gap:8px;align-items:center}.reg-imgrow input{flex:1}'
    + '.reg-fetch{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#57534e;background:#faf9f6;border:1.5px solid #e2ddd6;border-radius:8px;padding:8px 11px;cursor:pointer;white-space:nowrap}.reg-fetch:hover{border-color:#b5b0a8;color:#292524}'
    + '.reg-actcell{white-space:nowrap}.reg-actcell .reg-mini{display:none}.reg-unlocked .reg-actcell .reg-mini{display:inline-block}'
    + '.reg-mini{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#57534e;background:#faf9f6;border:1.5px solid #e2ddd6;border-radius:8px;padding:5px 10px;margin-right:6px;cursor:pointer}'
    + '.reg-mini:hover{border-color:#b5b0a8;color:#292524}.reg-mini.reg-del:hover{border-color:#e6b3b3;color:#b31b1b;background:#fdf5f5}'
    + '.reg-mini-note{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:.04em;color:#a8a29e}'
    + '.reg-lock{display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:#fff;border:1.5px solid #ececea;border-radius:12px;padding:12px 14px;margin:14px 0 4px}.reg-unlocked .reg-lock{display:none}'
    + '.reg-lock span.lab-t{font-size:13px;color:#57534e;font-weight:600}'
    + '.reg-lock input{font:inherit;font-size:14px;padding:9px 12px;border:1.5px solid #e0e0dc;border-radius:9px;outline:none}'
    + '.reg-msg{font-size:12.5px;color:#8a857c;font-weight:600}'
    + '.reg-ov{position:fixed;inset:0;z-index:80;display:flex;align-items:flex-start;justify-content:center;padding:5vh 16px}'
    + '.reg-ov[hidden]{display:none}'
    + '.reg-ov-bd{position:absolute;inset:0;background:rgba(20,17,14,.5)}'
    + '.reg-ov-card{position:relative;background:#fff;border-radius:16px;max-width:520px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 24px 60px rgba(0,0,0,.3)}'
    + '.reg-ov-h{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;padding:20px 22px 4px}'
    + '.reg-form{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:14px 22px}'
    + '.reg-f{display:flex;flex-direction:column;gap:5px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif}'
    + '.reg-f.wide{grid-column:1 / -1}'
    + '.reg-f span{font-size:10px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#8a857c}'
    + '.reg-f span i{color:#b31b1b;font-style:normal}'
    + '.reg-f input{font:inherit;font-size:14px;padding:10px 12px;border:1.5px solid #e0e0dc;border-radius:9px;outline:none}'
    + '.reg-f input:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.reg-ov-foot{display:flex;align-items:center;gap:10px;padding:14px 22px 20px;border-top:1.5px solid #f1f1f1;margin-top:6px}'
    + '.reg-ov-foot>:first-child{margin-right:auto}'
    + '@media(max-width:520px){.reg-form{grid-template-columns:1fr}}'
    + '</style>';
}

function regFilterJs_() {
  return '<script>'
    + 'function flt(){var qEl=document.getElementById("q");if(!qEl)return;var q=qEl.value.toLowerCase().trim();'
    + 'var tmEl=document.getElementById("team");var tm=tmEl?tmEl.value:"";var n=0;'
    + 'document.querySelectorAll(".reg-row").forEach(function(r){var ok=(!q||r.dataset.hay.indexOf(q)>=0)&&(!tm||r.dataset.team===tm);r.style.display=ok?"":"none";if(ok)n++;});'
    + 'var e=document.getElementById("empty");if(e)e.style.display=n?"none":"block";}'
    + '</script>';
}

function regUnlockBar_() {
  return '<div class="reg-lock" id="reg-lock"><span class="lab-t">Enter the admin passcode to add, edit, or delete.</span>'
    + '<input type="password" id="reg-pass" placeholder="Passcode" onkeydown="if(event.key===\'Enter\')regUnlock()">'
    + '<button type="button" class="btn btn-ghost" onclick="regUnlock()">Unlock</button>'
    + '<span id="reg-lock-msg" class="reg-msg"></span></div>';
}

function regFormOverlay_(which) {
  const spec = regFields_(which);
  const hasLink = spec.fields.some(function (f) { return f.h === 'Product link'; });
  const fields = spec.fields.map(function (f) {
    if (f.auto) return '';
    if (f.type === 'image') {
      return '<div class="reg-f wide reg-imgrow"><span class="reg-preview" id="reg-preview"></span>'
        + '<div class="reg-imgfields"><span>' + escapeHtml_(f.label) + '</span><div class="reg-imgtop">'
        + '<input data-h="Image" id="reg-imgurl" type="url" placeholder="Paste an image URL' + (hasLink ? ', or fetch from the product link' : '') + '" oninput="regPreview()">'
        + (hasLink ? '<button type="button" class="reg-fetch" id="reg-fetchbtn" onclick="regFetchImg()">Fetch from link</button>' : '')
        + '</div></div></div>';
    }
    const type = f.type === 'number' ? 'number' : (f.type === 'url' ? 'url' : 'text');
    const wide = (f.h === 'Notes' || f.h === 'eShop info' || f.h === 'Product link') ? ' wide' : '';
    return '<label class="reg-f' + wide + '"><span>' + escapeHtml_(f.label) + (f.req ? ' <i>required</i>' : '') + '</span>'
      + '<input data-h="' + escapeHtml_(f.h) + '" type="' + type + '"' + (type === 'number' ? ' step="any"' : '') + '></label>';
  }).join('');
  return '<div class="reg-ov" id="reg-ov" hidden><div class="reg-ov-bd" onclick="regClose()"></div>'
    + '<div class="reg-ov-card"><div class="reg-ov-h" id="reg-ov-title">Add</div>'
    + '<div class="reg-form">' + fields + '</div>'
    + '<div class="reg-ov-foot"><span id="reg-msg" class="reg-msg"></span>'
    + '<button type="button" class="btn btn-ghost" onclick="regClose()">Cancel</button>'
    + '<button type="button" class="btn btn-confirm" id="reg-save" onclick="regSave()">Save</button></div>'
    + '</div></div>';
}

function regEditJs_() {
  return '<script>'
    + 'var REG_EROW=null,REG_EKEY=null;'
    + 'function regInputs(){return [].slice.call(document.querySelectorAll("#reg-ov .reg-form input"));}'
    + 'function regFill(v){regInputs().forEach(function(i){i.value=(v&&v[i.dataset.h]!=null)?v[i.dataset.h]:"";});regPreview();}'
    + 'function regPreview(){var el=document.getElementById("reg-imgurl");var p=document.getElementById("reg-preview");if(!p)return;var u=el?el.value:"";p.innerHTML=/^https?:|^data:/.test(u)?("<img src=\\""+u.replace(/"/g,"%22")+"\\" onerror=\\"this.parentNode.textContent=\'\\u2715\'\\">"):"";}'
    + 'function regFetchImg(){var link="";regInputs().forEach(function(i){if(i.dataset.h==="Product link")link=i.value;});if(!link){regMsg("Add a product link first.",true);return;}var b=document.getElementById("reg-fetchbtn");if(b){b.disabled=true;b.textContent="Fetching...";}'
    + 'google.script.run.withSuccessHandler(function(res){if(b){b.disabled=false;b.textContent="Fetch from link";}if(res&&res.ok&&res.image){document.getElementById("reg-imgurl").value=res.image;regPreview();regMsg("Found an image.");}else{regMsg("No image at that link. Paste one instead.",true);}}).withFailureHandler(function(e){if(b){b.disabled=false;b.textContent="Fetch from link";}regMsg(String(e),true);}).regFetchImage(link,ADMIN_PASS);}'
    + 'function regBackfill(){if(!confirm("Fetch a photo for every item that has a product link but no image? This can take a minute."))return;var el=document.getElementById("reg-backfill");if(el){el.disabled=true;el.textContent="Fetching...";}'
    + 'google.script.run.withSuccessHandler(function(res){if(el){el.disabled=false;el.textContent="Fetch images";}if(res&&res.ok){alert("Added "+res.added+" image(s) of "+res.checked+" item(s) with links.");regRefresh();}else{alert((res&&res.error)||"Could not run.");}}).withFailureHandler(function(e){if(el){el.disabled=false;el.textContent="Fetch images";}alert(String(e));}).regBackfillImages(REG_WHICH,ADMIN_PASS);}'
    + 'function regCollect(){var o={};regInputs().forEach(function(i){o[i.dataset.h]=i.value;});return o;}'
    + 'function regShow(){document.getElementById("reg-ov").hidden=false;}'
    + 'function regClose(){document.getElementById("reg-ov").hidden=true;}'
    + 'function regMsg(m,err){var e=document.getElementById("reg-msg");if(e){e.textContent=m||"";e.style.color=err?"#b31b1b":"#8a857c";}}'
    + 'function regUnlock(){var p=(document.getElementById("reg-pass")||{}).value||"";var m=document.getElementById("reg-lock-msg");if(m){m.textContent="Checking...";}'
    + 'google.script.run.withSuccessHandler(function(ok){if(ok){ADMIN_PASS=p;try{sessionStorage.setItem("regPass",p);}catch(e){}document.getElementById("reg-root").classList.add("reg-unlocked");}else if(m){m.textContent="Wrong passcode.";m.style.color="#b31b1b";}}).withFailureHandler(function(){if(m){m.textContent="Could not verify. Retry.";}}).tpCheckPass(p);}'
    + 'function regOpenAdd(){REG_EROW=null;REG_EKEY=null;document.getElementById("reg-ov-title").textContent="Add";regFill({});regMsg("");regShow();var f=document.querySelector("#reg-ov .reg-form input");if(f)f.focus();}'
    + 'function regOpenEdit(row){REG_EROW=row;var v=REG_ROWS[row]||{};REG_EKEY=v[REG_KEY];document.getElementById("reg-ov-title").textContent="Edit";regFill(v);regMsg("");regShow();}'
    + 'function regAfter(res){var s=document.getElementById("reg-save");if(s)s.disabled=false;if(res&&res.ok){regClose();regRefresh();}else{regMsg((res&&res.error)||"Could not save.",true);}}'
    + 'function regSave(){var vals=regCollect();var s=document.getElementById("reg-save");if(s)s.disabled=true;regMsg("Saving...");'
    + 'var fail=function(e){if(s)s.disabled=false;regMsg(String(e&&e.message||e),true);};'
    + 'if(REG_EROW==null){google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regAdd(REG_WHICH,vals,ADMIN_PASS);}'
    + 'else{google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regUpdateRow(REG_WHICH,REG_EROW,REG_EKEY,vals,ADMIN_PASS);}}'
    + 'function regDeleteRow(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";if(!confirm("Delete \\""+key+"\\"? This cannot be undone."))return;if(btn)btn.disabled=true;'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){if(typeof REG_ITEM!=="undefined"){location.href="?registry="+REG_WHICH+"&admin=1";}else{var tr=btn&&btn.closest?btn.closest(".reg-row"):null;if(tr)tr.remove();delete REG_ROWS[row];}}else{if(btn)btn.disabled=false;alert((res&&res.error)||"Could not delete.");}}).withFailureHandler(function(e){if(btn)btn.disabled=false;alert(String(e));}).regDelete(REG_WHICH,row,key,ADMIN_PASS);}'
    + 'function regRestockOne(row){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";google.script.run.withSuccessHandler(function(res){if(res&&res.ok){location.href="?registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(key)+"&admin=1";}else{alert((res&&res.error)||"Could not update.");}}).withFailureHandler(function(e){alert(String(e));}).regRestock(row,key,ADMIN_PASS);}'
    + 'function regRefresh(){if(typeof REG_ITEM!=="undefined"){location.href="?registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(REG_EKEY||"")+"&admin=1";return;}'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){document.getElementById("reg-tbody").innerHTML=res.html;REG_ROWS=JSON.parse(res.mapJson);if(typeof flt==="function")flt();}}).regRowsHtml(REG_WHICH,ADMIN_PASS);}'
    + 'function regScan(){if(!("BarcodeDetector" in window)){alert("Live scanning is not available in this browser. Use Print labels and scan them with your phone camera instead.");return;}'
    + 'navigator.mediaDevices.getUserMedia({video:{facingMode:"environment"}}).then(function(stream){regRunScan(stream);}).catch(function(){alert("Camera is blocked here (common in this app). Use Print labels and scan with your phone camera instead.");});}'
    + 'function regRunScan(stream){var ov=document.createElement("div");ov.style.cssText="position:fixed;inset:0;z-index:90;background:#000;display:flex;flex-direction:column;align-items:center;justify-content:center";'
    + 'var vid=document.createElement("video");vid.setAttribute("playsinline","");vid.style.cssText="max-width:100%;max-height:80vh";vid.srcObject=stream;vid.play();'
    + 'var btn=document.createElement("button");btn.textContent="Close";btn.className="btn btn-ghost";btn.style.margin="16px";btn.onclick=function(){stream.getTracks().forEach(function(t){t.stop();});ov.remove();};'
    + 'ov.appendChild(vid);ov.appendChild(btn);document.body.appendChild(ov);'
    + 'var det=new BarcodeDetector();var loop=function(){if(!ov.isConnected)return;det.detect(vid).then(function(codes){if(codes&&codes.length){var val=codes[0].rawValue||"";stream.getTracks().forEach(function(t){t.stop();});ov.remove();var q=document.getElementById("q");if(q){q.value=val;if(typeof flt==="function")flt();}}else{requestAnimationFrame(loop);}}).catch(function(){requestAnimationFrame(loop);});};requestAnimationFrame(loop);}'
    + '(function(){try{var p=sessionStorage.getItem("regPass");if(p){google.script.run.withSuccessHandler(function(ok){if(ok){ADMIN_PASS=p;document.getElementById("reg-root").classList.add("reg-unlocked");}}).tpCheckPass(p);}}catch(e){}})();'
    + '</script>';
}

// ---- registry server writes (all gated by the shared passcode) ----

function regSheet_(which) {
  const name = String(which).toLowerCase() === 'inventory' ? 'Inventory' : 'Equipment';
  const sh = registrySs_().getSheetByName(name);
  if (!sh) throw new Error('No "' + name + '" tab. Run the registry setup first.');
  return sh;
}

function regHeaderMap_(sh) {
  const hs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const map = {}; hs.forEach(function (h, i) { map[h] = i; });
  return { headers: hs, map: map };
}

function regNextId_(sh, hm, key, prefix) {
  const ci = hm.map[key];
  let max = 0;
  if (ci == null) return prefix + '001';
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][ci]).trim();
    if (id.indexOf(prefix) === 0) { const n = parseInt(id.slice(prefix.length), 10); if (!isNaN(n) && n > max) max = n; }
  }
  let s = String(max + 1); while (s.length < 3) s = '0' + s;
  return prefix + s;
}

function regAdd(which, vals, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  vals = vals || {};
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  const reqF = spec.fields.filter(function (f) { return f.req; })[0];
  if (reqF && !String(vals[reqF.h] || '').trim()) return { ok: false, error: reqF.label + ' is required.' };
  const row = hm.headers.map(function (h) { return vals[h] != null ? vals[h] : ''; });
  if (spec.idPrefix) { const ki = hm.map[spec.key]; if (ki != null && !String(row[ki]).trim()) row[ki] = regNextId_(sh, hm, spec.key, spec.idPrefix); }
  sh.appendRow(row);
  return { ok: true };
}

function regUpdateRow(which, rowNum, expectedKey, vals, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map[spec.key];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedKey).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  hm.headers.forEach(function (h, i) { if (vals[h] != null && !(spec.idPrefix && i === ki)) cur[i] = vals[h]; });
  sh.getRange(rowNum, 1, 1, hm.headers.length).setValues([cur]);
  return { ok: true };
}

function regDelete(which, rowNum, expectedKey, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map[spec.key];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedKey).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  sh.deleteRow(rowNum);
  return { ok: true };
}

function regRestock(rowNum, expectedItem, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  const sh = regSheet_('inventory');
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map['Item'];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedItem).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  const oni = hm.map['On hand'], rqi = hm.map['Reorder qty'], lri = hm.map['Last restocked'];
  const qty = (rqi != null && Number(cur[rqi]) > 0) ? Number(cur[rqi]) : 0;
  if (oni != null) cur[oni] = (Number(cur[oni]) || 0) + qty;
  if (lri != null) cur[lri] = new Date();
  sh.getRange(rowNum, 1, 1, hm.headers.length).setValues([cur]);
  return { ok: true };
}

// Best-effort product image: fetch the page and read its og:image / twitter:image
// meta tag. Many suppliers (MSC, Amazon) block server requests; those just return ''.
function regFetchOgImage_(url) {
  url = String(url || '').trim();
  if (!/^https?:\/\//i.test(url)) return '';
  try {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36', 'Accept': 'text/html' }
    });
    if (resp.getResponseCode() >= 400) return '';
    const html = resp.getContentText();
    if (!html || html.length < 500) return '';
    const m = html.match(/<meta[^>]+(?:property|name)=["'](?:og:image(?::secure_url)?|twitter:image(?::src)?)["'][^>]*>/i);
    if (!m) return '';
    const c = m[0].match(/content=["']([^"']+)["']/i);
    let img = c ? c[1].trim() : '';
    if (img.indexOf('//') === 0) img = 'https:' + img;
    return /^https?:\/\//i.test(img) ? img : '';
  } catch (e) { return ''; }
}

function regFetchImage(link, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  return { ok: true, image: regFetchOgImage_(link) };
}

// Prepopulate: for every row with a product link but no image, try to pull one.
function regBackfillImages(which, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Not authorized. Unlock first.' };
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  const li = hm.map['Product link'], ii = hm.map['Image'];
  if (li == null || ii == null) return { ok: true, checked: 0, added: 0 };
  const vals = sh.getDataRange().getValues();
  let checked = 0, added = 0;
  for (let i = 1; i < vals.length; i++) {
    const link = String(vals[i][li] || '').trim();
    const img = String(vals[i][ii] || '').trim();
    if (!link || img) continue;
    checked++;
    const found = regFetchOgImage_(link);
    if (found) { sh.getRange(i + 1, ii + 1).setValue(found); added++; }
    Utilities.sleep(150);
    if (checked >= 150) break;
  }
  return { ok: true, checked: checked, added: added };
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
function escapeHtml_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
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