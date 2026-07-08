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
  responsesSheet: 'Form Responses',
  contactsSheet: 'Team Contacts',
  sender: '',                          // go-live: 'eng_projectteams@cornell.edu' (send-as alias on this account)
  alwaysCc: [],                        // go-live: ['nhh5@cornell.edu']
  fallbackEmail: 'com34@cornell.edu',
  webAppUrl: 'https://script.google.com/macros/s/AKfycbwOnNmpSXc3biH14Fm9iLcUQ2X0UK-Gx5kQpNmrBsHd3K-l2u0GjsMblOumiY73drM_/exec',
  notifiedHeader: 'Notified at',
  issueTokenHeader: 'Issue token',
  addressedHeader: 'Addressed at',
  lastReminderHeader: 'Last reminder',
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
  const DAY = 86400000;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let sent = 0;

  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    if (cNotified < 0 || !row[cNotified]) continue;            // never notified
    if (row[cAddr]) continue;                                   // addressed / closed
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
  if (!CONFIG.webAppUrl || !token) return '';
  const url = CONFIG.webAppUrl + '?id=' + encodeURIComponent(token);
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:20px 0 4px"><tr>'
    + '<td style="background:#8f1515;border-radius:6px"><a href="' + url + '" style="display:inline-block;padding:12px 24px;color:#fff;font:bold 14px Arial,sans-serif;text-decoration:none;border-radius:6px">Mark this issue as addressed</a></td>'
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
  if (p.team) { const t = lookupTeamByToken_(p.team); return t ? teamPortal_(t, false) : htmlPage_('Invalid link', 'This team link is not recognized.'); }
  if (p.view === 'all') return allIssuesPage_();          // filterable all-open-issues dashboard (Issues tab)
  if (p.view) return teamPortal_(String(p.view), true);   // read-only, reached from the Ops Hub team picker
  if (p.id) return confirmPage_(p.id);
  return pickerPage_();
}

function swissShell_(innerHtml, pageTitle) {
  const enh = '<style>'
    + '.b:hover{background:#7a1212 !important}'
    + '@media(max-width:600px){.swh{font-size:26px !important}.swfoot{flex-direction:column !important;align-items:flex-start !important;gap:12px !important}}'
    + '</style>';
  const html = enh
    + '<div style="margin:0;background:#fafafa;min-height:100vh;font-family:Helvetica,Arial,sans-serif;color:#111">'
    +   '<div style="height:5px;background:#b31b1b"></div>'
    +   '<div style="max-width:600px;margin:0 auto;padding:40px 28px 64px">' + innerHtml + '</div>'
    + '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(pageTitle || 'Space Status')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // allow embedding in the Ops Hub Issues tab
}

function htmlPage_(title, bodyHtml) {
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">' + escapeHtml_(title) + '</div>'
    + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:14px">' + bodyHtml + '</div>';
  return swissShell_(inner, 'Space Status');
}

function confirmPage_(id) {
  const info = findIssue_(id);
  if (!info) return htmlPage_('Not found', 'We could not find that issue. It may have been removed.');
  if (info.addressed) return htmlPage_('Already addressed', 'This was marked as addressed on ' + escapeHtml_(fmtShort_(info.addressed)) + '.');
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">Mark this issue as addressed?</div>'
    + '<div id="act" style="margin-top:28px"><a class="b" href="javascript:void(0)" onclick="go()" style="display:inline-block;font-size:14px;font-weight:700;color:#fff;background:#b31b1b;padding:13px 28px;border-radius:4px;text-decoration:none">Yes, mark as addressed</a></div>'
    + '<div id="done" style="display:none;font-size:22px;font-weight:800;letter-spacing:-.02em;color:#1d7a46;margin-top:6px"></div>'
    + '<script>'
    + 'function go(){document.getElementById("act").innerHTML="Saving...";'
    + 'google.script.run.withSuccessHandler(ok).withFailureHandler(err).confirmAddressed(' + JSON.stringify(id) + ');}'
    + 'function ok(m){document.getElementById("act").style.display="none";var d=document.getElementById("done");d.style.display="block";d.innerHTML="\\u2713 Marked as addressed";}'
    + 'function err(x){document.getElementById("act").innerHTML="Something went wrong. Please reply to the email instead.";}'
    + '</script>';
  return swissShell_(inner, 'Space Status');
}

function teamPortal_(team, readOnly) {
  if (!team) return htmlPage_('Invalid link', 'This team link is not recognized.');
  const data = listTeamIssues_(team);
  const issues = data.open;
  const overdue = issues.filter(function (x) { return x.overdue; }).length;
  const total = issues.length + data.resolved;
  const pct = total > 0 ? Math.round(data.resolved / total * 100) : 0;

  let inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#9a958c">Project Teams Ops Hub</div>'
    + '<div class="swh" style="font-size:34px;font-weight:800;letter-spacing:-.035em;line-height:1;margin-top:10px">' + escapeHtml_(team) + '</div>'
    + '<div style="width:46px;height:3px;background:#b31b1b;margin-top:12px"></div>';

  inner += '<div style="display:flex;align-items:center;gap:26px;margin-top:26px">'
    + '<svg width="84" height="84" viewBox="0 0 42 42" aria-hidden="true">'
    +   '<circle cx="21" cy="21" r="15.915" fill="none" stroke="#e9e9e9" stroke-width="3.4"></circle>'
    +   '<circle id="ring-fg" cx="21" cy="21" r="15.915" fill="none" stroke="#1d7a46" stroke-width="3.4" stroke-linecap="round" stroke-dasharray="' + pct + ' ' + (100 - pct) + '" transform="rotate(-90 21 21)"></circle>'
    +   '<text id="ring-pct" x="21" y="20.5" text-anchor="middle" font-size="9" font-weight="800" fill="#111">' + pct + '%</text>'
    +   '<text x="21" y="27" text-anchor="middle" font-size="3.3" letter-spacing="0.4" fill="#9a958c">RESOLVED</text>'
    + '</svg>'
    + '<div style="font-size:15px;line-height:2;color:#555">'
    +   '<div><b id="res-n" style="color:#111;font-size:17px">' + data.resolved + '</b> &nbsp;resolved</div>'
    +   '<div><b id="open-n" style="color:#111;font-size:17px">' + issues.length + '</b> &nbsp;open'
    +     '<span id="over-wrap" style="' + (overdue ? '' : 'display:none;') + 'color:#b31b1b;font-weight:700;margin-left:6px"><span id="over-n">' + overdue + '</span> overdue</span>'
    +   '</div>'
    + '</div></div>';

  if (!issues.length) {
    inner += '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:30px">No open issues. Everything is in good shape.</div>';
    return swissShell_(inner, 'Space Status - ' + team);
  }

  let idx = 0;
  const renderCard = function (it) {
    const rid = 'iss' + (idx++);
    const dl = it.deadline ? fmtShort_(it.deadline) : 'No set deadline';
    const dueColor = it.overdue ? '#b31b1b' : '#777';
    const chip = it.overdue
      ? '<span style="flex:0 0 auto;font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#b31b1b;background:#fbeaea;padding:4px 10px;border-radius:999px">Overdue</span>'
      : (it.deadline ? '<span style="flex:0 0 auto;font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#8a857c;background:#f0efe9;padding:4px 10px;border-radius:999px">' + daysLeftLabel_(it.deadline) + '</span>' : '');
    return '<div id="' + rid + '" style="background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.06),0 1px 3px rgba(20,20,30,.05);overflow:hidden;margin-bottom:12px">'
      + '<div style="padding:18px 22px 16px">'
      +   '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:14px">'
      +     '<div style="font-size:18px;font-weight:800;letter-spacing:-.02em">' + escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue') + '</div>'
      +     chip
      +   '</div>'
      +   (it.action ? '<div style="font-size:14px;color:#555;margin-top:6px">' + escapeHtml_(phrase_(it.action)) + '</div>' : '')
      +   (it.details ? '<div style="font-size:13px;color:#8a857c;margin-top:4px">' + escapeHtml_(it.details) + '</div>' : '')
      +   (it.photos && it.photos.length ? photoStrip_(it.photos) : '')
      + '</div>'
      + '<div class="swfoot" style="display:flex;justify-content:space-between;align-items:center;border-top:1px solid #f1f1f1;padding:13px 22px;background:#fcfcfb">'
      +   '<span id="' + rid + '-status" style="font-size:13px;color:' + dueColor + ';font-weight:600">Due ' + escapeHtml_(dl) + '</span>'
      +   (readOnly ? '' : '<span id="' + rid + '-act">'
      +     '<span id="' + rid + '-btn"><a class="b" href="javascript:void(0)" onclick="ask(\'' + rid + '\')" style="font-size:13px;font-weight:700;color:#fff;background:#b31b1b;padding:10px 18px;border-radius:6px;text-decoration:none;white-space:nowrap">Mark addressed</a></span>'
      +     '<span id="' + rid + '-confirm" style="display:none;white-space:nowrap">'
      +       '<a class="b" href="javascript:void(0)" onclick="doMark(\'' + rid + '\',\'' + it.token + '\',' + (it.overdue ? 'true' : 'false') + ')" style="font-size:13px;font-weight:700;color:#fff;background:#b31b1b;padding:10px 16px;border-radius:6px;text-decoration:none">Confirm</a>'
      +       '<a href="javascript:void(0)" onclick="cancelMark(\'' + rid + '\')" style="font-size:13px;color:#8a857c;text-decoration:none;margin-left:12px">Cancel</a>'
      +     '</span>'
      +   '</span>')
      + '</div>'
      + '</div>';
  };

  const overdueList = issues.filter(function (x) { return x.overdue; });
  const openList = issues.filter(function (x) { return !x.overdue; });
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

  if (!readOnly) inner += '<script>'
    + 'var ssRes=' + data.resolved + ',ssOpen=' + issues.length + ',ssOver=' + overdue + ',ssTot=' + total + ';'
    + 'function setT(id,t){var e=document.getElementById(id);if(e)e.textContent=t;}'
    + 'function ask(r){document.getElementById(r+"-btn").style.display="none";document.getElementById(r+"-confirm").style.display="inline";}'
    + 'function cancelMark(r){document.getElementById(r+"-confirm").style.display="none";document.getElementById(r+"-btn").style.display="inline";}'
    + 'function doMark(r,t,od){var a=document.getElementById(r+"-act");a.innerHTML="Saving...";'
    + 'google.script.run.withSuccessHandler(function(){'
    +   'a.innerHTML="";var s=document.getElementById(r+"-status");if(s){s.textContent="\\u2713 Addressed";s.style.color="#1d7a46";}'
    +   'document.getElementById(r).style.opacity="0.55";'
    +   'ssRes++;ssOpen--;if(od)ssOver--;'
    +   'var pct=ssTot>0?Math.round(ssRes/ssTot*100):0;'
    +   'setT("res-n",ssRes);setT("open-n",ssOpen);setT("ring-pct",pct+"%");setT("over-n",ssOver);'
    +   'var ow=document.getElementById("over-wrap");if(ow)ow.style.display=ssOver>0?"inline":"none";'
    +   'var fg=document.getElementById("ring-fg");if(fg)fg.setAttribute("stroke-dasharray",pct+" "+(100-pct));'
    + '}).withFailureHandler(function(){var a=document.getElementById(r+"-act");a.innerHTML="Please retry.";}).confirmAddressed(t);}'
    + '</script>';
  return swissShell_(inner, 'Space Status - ' + team);
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
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader), cPhoto = ci(CONFIG.headers.photo);
  const open = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTeam < 0 || norm_(v[i][cTeam]) !== norm_(teamName)) continue;
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) { resolved++; continue; }
    const color = parseColor_(v[i][cStatus]);
    let deadline = null, overdue = false;
    if (color && color !== 'purple') {
      const rd = new Date(v[i][cTs]);
      if (!isNaN(rd.getTime())) { deadline = new Date(rd); deadline.setDate(deadline.getDate() + SLA_DAYS[color]); deadline.setHours(0, 0, 0, 0); overdue = today > deadline; }
    }
    open.push({
      token: String(v[i][cTok]),
      issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
      action: cAction >= 0 ? String(v[i][cAction]).trim() : '',
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      photos: cPhoto >= 0 ? extractFileIds_(v[i][cPhoto]) : [],
      color: color, deadline: deadline, overdue: overdue,
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
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader);
  const open = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) { resolved++; continue; }
    const color = parseColor_(v[i][cStatus]);
    let deadline = null, overdue = false;
    if (color && color !== 'purple') {
      const rd = new Date(v[i][cTs]);
      if (!isNaN(rd.getTime())) { deadline = new Date(rd); deadline.setDate(deadline.getDate() + SLA_DAYS[color]); deadline.setHours(0, 0, 0, 0); overdue = today > deadline; }
    }
    open.push({
      team: cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
      token: String(v[i][cTok]),
      issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
      action: cAction >= 0 ? String(v[i][cAction]).trim() : '',
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      deadline: deadline, overdue: overdue,
    });
  }
  open.sort(function (a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (a.deadline && b.deadline) return a.deadline - b.deadline;
    if (a.deadline) return -1; if (b.deadline) return 1; return 0;
  });
  return { open: open, resolved: resolved };
}

// Every open issue across all teams, with filters and Mark complete. Embedded in the hub Issues tab.
function allIssuesPage_() {
  const data = listAllIssues_();
  const issues = data.open;
  const overdue = issues.filter(function (x) { return x.overdue; }).length;
  const teamSet = {};
  issues.forEach(function (it) { if (it.team) teamSet[it.team] = 1; });
  const teamOpts = Object.keys(teamSet).sort().map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('');

  let inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#9a958c">Project Teams Ops Hub</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.035em;line-height:1.05;margin-top:8px">Open space issues</div>'
    + '<div style="width:46px;height:3px;background:#b31b1b;margin-top:12px"></div>'
    + '<div style="font-size:14px;color:#555;margin-top:14px"><b id="sum-open" style="color:#111;font-size:16px">' + issues.length + '</b> open &nbsp;&middot;&nbsp; <span style="color:#b31b1b;font-weight:700"><span id="sum-over">' + overdue + '</span> overdue</span></div>';

  inner += '<div style="display:flex;flex-wrap:wrap;gap:10px;margin:18px 0 8px">'
    + '<input id="q" type="search" placeholder="Search issues" oninput="flt()" style="flex:1;min-width:160px;font-size:14px;padding:9px 12px;border:1.5px solid #ddd;border-radius:8px">'
    + '<select id="team" onchange="flt()" style="font-size:14px;padding:9px 12px;border:1.5px solid #ddd;border-radius:8px;background:#fff"><option value="">All teams</option>' + teamOpts + '</select>'
    + '<label style="display:inline-flex;align-items:center;gap:7px;font-size:14px;color:#333;padding:0 4px;white-space:nowrap"><input id="odue" type="checkbox" onchange="flt()"> Overdue only</label>'
    + '</div>';

  if (!issues.length) {
    inner += '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:24px">No open issues. Everything is in good shape.</div>';
    return swissShell_(inner, 'Open issues');
  }

  let idx = 0;
  issues.forEach(function (it) {
    const rid = 'iss' + (idx++);
    const dl = it.deadline ? fmtShort_(it.deadline) : 'No set deadline';
    const dueColor = it.overdue ? '#b31b1b' : '#777';
    const chip = it.overdue
      ? '<span style="flex:0 0 auto;font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#b31b1b;background:#fbeaea;padding:4px 10px;border-radius:999px">Overdue</span>'
      : (it.deadline ? '<span style="flex:0 0 auto;font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#8a857c;background:#f0efe9;padding:4px 10px;border-radius:999px">' + daysLeftLabel_(it.deadline) + '</span>' : '');
    const hay = ((it.team || '') + ' ' + (it.issueType || '') + ' ' + (it.action || '') + ' ' + (it.details || '')).toLowerCase();
    inner += '<div class="card" id="' + rid + '" data-team="' + escapeHtml_(it.team) + '" data-over="' + (it.overdue ? '1' : '0') + '" data-hay="' + escapeHtml_(hay) + '" style="background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.06),0 1px 3px rgba(20,20,30,.05);overflow:hidden;margin-bottom:12px">'
      + '<div style="padding:16px 20px 14px">'
      +   '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">'
      +     '<div><div style="font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:#b31b1b">' + escapeHtml_(it.team || 'Unassigned') + '</div>'
      +       '<div style="font-size:17px;font-weight:800;letter-spacing:-.02em;margin-top:2px">' + escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue') + '</div></div>'
      +     chip
      +   '</div>'
      +   (it.action ? '<div style="font-size:13.5px;color:#555;margin-top:6px">' + escapeHtml_(phrase_(it.action)) + '</div>' : '')
      +   (it.details ? '<div style="font-size:13px;color:#8a857c;margin-top:4px">' + escapeHtml_(it.details) + '</div>' : '')
      + '</div>'
      + '<div class="swfoot" style="display:flex;justify-content:space-between;align-items:center;border-top:1px solid #f1f1f1;padding:12px 20px;background:#fcfcfb">'
      +   '<span id="' + rid + '-status" style="font-size:13px;color:' + dueColor + ';font-weight:600">Due ' + escapeHtml_(dl) + '</span>'
      +   '<span id="' + rid + '-act">'
      +     '<span id="' + rid + '-btn"><a class="b" href="javascript:void(0)" onclick="ask(\'' + rid + '\')" style="font-size:13px;font-weight:700;color:#fff;background:#b31b1b;padding:9px 16px;border-radius:6px;text-decoration:none;white-space:nowrap">Mark complete</a></span>'
      +     '<span id="' + rid + '-confirm" style="display:none;white-space:nowrap">'
      +       '<a class="b" href="javascript:void(0)" onclick="doMark(\'' + rid + '\',\'' + it.token + '\',' + (it.overdue ? 'true' : 'false') + ')" style="font-size:13px;font-weight:700;color:#fff;background:#b31b1b;padding:9px 14px;border-radius:6px;text-decoration:none">Confirm</a>'
      +       '<a href="javascript:void(0)" onclick="cancelMark(\'' + rid + '\')" style="font-size:13px;color:#8a857c;text-decoration:none;margin-left:12px">Cancel</a>'
      +     '</span>'
      +   '</span>'
      + '</div>'
      + '</div>';
  });

  inner += '<div id="empty" style="display:none;font-size:15px;color:#8a857c;margin-top:8px">No issues match those filters.</div>';

  inner += '<script>'
    + 'var ssOpen=' + issues.length + ',ssOver=' + overdue + ';'
    + 'function flt(){var q=document.getElementById("q").value.toLowerCase().trim();var tm=document.getElementById("team").value;var od=document.getElementById("odue").checked;var n=0;'
    +   'document.querySelectorAll(".card").forEach(function(c){var ok=c.dataset.done!=="1"&&(!q||c.dataset.hay.indexOf(q)>=0)&&(!tm||c.dataset.team===tm)&&(!od||c.dataset.over==="1");c.style.display=ok?"":"none";if(ok)n++;});'
    +   'document.getElementById("empty").style.display=n?"none":"block";}'
    + 'function ask(r){document.getElementById(r+"-btn").style.display="none";document.getElementById(r+"-confirm").style.display="inline";}'
    + 'function cancelMark(r){document.getElementById(r+"-confirm").style.display="none";document.getElementById(r+"-btn").style.display="inline";}'
    + 'function doMark(r,t,od){var a=document.getElementById(r+"-act");a.innerHTML="Saving...";'
    +   'google.script.run.withSuccessHandler(function(){'
    +     'var c=document.getElementById(r);c.dataset.done="1";a.innerHTML="";var s=document.getElementById(r+"-status");if(s){s.textContent="\\u2713 Completed";s.style.color="#1d7a46";}c.style.opacity="0.55";'
    +     'ssOpen--;if(od)ssOver--;document.getElementById("sum-open").textContent=ssOpen;document.getElementById("sum-over").textContent=ssOver;'
    +   '}).withFailureHandler(function(){document.getElementById(r+"-act").innerHTML="Please retry.";}).confirmAddressed(t);}'
    + '</script>';
  return swissShell_(inner, 'Open issues');
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
  const cAddr = H.indexOf(norm_(CONFIG.addressedHeader));
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][cTok]) === String(id)) {
      return {
        rowIndex: i + 1,
        team: cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
        issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
        addressed: cAddr >= 0 ? v[i][cAddr] : '',
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
