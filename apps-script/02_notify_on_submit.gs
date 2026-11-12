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
  // Fleet driver-cleanup tool (?tool=drivers). Runs on the reservation sheet.
  fleetSpreadsheetId: '1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M',
  fleetGradYearThrough: 0,   // remove this graduation year and earlier; 0 = current year.
  toolPassHash: '993b97603778a944ff00c86775aeb4c852a6deae031dbee58206dc3ae1e3242c',   // SHA-256 of the tool passcode (default "bigred")
  responsesSheet: 'Form Responses',
  contactsSheet: 'Team Contacts',
  emailLogSheet: 'Email log',                 // one row per email sent; made on first use
  sender: 'eng_projectteams@cornell.edu',                          // go-live: 'eng_projectteams@cornell.edu' (send-as alias on this account)
  alwaysCc: ['nhh5@cornell.edu'],                        // go-live: ['nhh5@cornell.edu']
  fallbackEmail: 'nhh5@cornell.edu',
  webAppUrl: 'https://script.google.com/macros/s/AKfycbwNbGjVcBrcsMZiOl2nXzpqZHz04nvKLm9D_aC0VJDz7Xxxf_4kLKlNSOHubPXj1X74/exec',  // MUST be the published /exec URL; email acknowledge links use this
  notifiedHeader: 'Notified at',
  issueTokenHeader: 'Issue token',
  addressedHeader: 'Addressed at',            // set on ADMIN APPROVAL = Completed
  lastReminderHeader: 'Last reminder',
  reminderHoldHeader: 'Reminder hold',        // admin cancelled reminders: "Off" = none for this job, a date = none before it
  completionPhotoHeader: 'Completion photo',  // evidence uploaded by the doer
  completionNoteHeader: 'Completion note',     // optional text the doer sends with the photos
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
    // Form question title. Older rows and the sheet column may still say
    // "Current Status". statusHeaderNames_() is what the email code reads.
    status: 'Respond within',
    photo: 'Photo Upload',
  },
};

// Accepted titles for the response-time question, new name first.
function statusHeaderNames_() { return ['Respond within', 'Current Status']; }

// Index in an already-normalized header row, or -1. Does not create a column.
function findHeader_(H, names) {
  for (var i = 0; i < names.length; i++) {
    var ix = H.indexOf(norm_(names[i]));
    if (ix >= 0) return ix;
  }
  return -1;
}

function firstNamed_(get, names) {
  var fallback = '';
  for (var i = 0; i < names.length; i++) {
    var v = get(names[i]);
    if (!v) continue;
    if (!fallback) fallback = v;
    if (parseColor_(v)) return v;
  }
  return fallback;
}

// The form rename left two columns titled "Respond within". Old rows keep the
// color in the first; new answers land in the later one. Use a value this
// parser understands, preferring the later column when both match.
function statusFromRow_(row, H) {
  var names = statusHeaderNames_();
  var parsed = '';
  var fallback = '';
  for (var i = 0; i < H.length; i++) {
    var hit = false;
    for (var n = 0; n < names.length; n++) {
      if (H[i] === norm_(names[n])) { hit = true; break; }
    }
    if (!hit) continue;
    var val = String(row[i] == null ? '' : row[i]).trim();
    if (!val) continue;
    fallback = val;
    if (parseColor_(val)) parsed = val;
  }
  return parsed || fallback;
}

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
    status: firstNamed_(get, statusHeaderNames_()), photo: get(h.photo),
  };
}

// contacts is optional: pass teamContacts_() when resolving many teams in one run,
// so the Team Contacts tab is read once.
function recipientsFor_(team, contacts) {
  const c = lookupTeam_(team, contacts);
  const to = uniqEmails_([c && c.generalEmail, c && c.liaisonEmail]);
  const cc = uniqEmails_([c && c.safetyEmail, c && c.extraCc].concat(CONFIG.alwaysCc));
  if (!to.length) to.push(CONFIG.fallbackEmail);
  return { to: to, cc: cc, contact: c };
}

// why is the Email log note: "New report" by default, or the reason for a resend.
function sendNotification_(data, token, why) {
  const team = data.team || 'Unknown / unsure';
  const rec = recipientsFor_(team);
  const color = parseColor_(data.status);
  const sev = SEVERITY[color] || SEVERITY_DEFAULT;
  const infoOnly = norm_(data.issueType) === 'informational';
  const subject = infoOnly
    ? 'For your awareness - ' + team
    : sev.subject + ' - ' + team + (data.issueType ? ' (' + phrase_(data.issueType) + ')' : '');
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
  const entry = { type: 'Notice', team: team, to: rec.to.join(','), cc: rec.cc.join(','), subject: subject, token: token, why: why || 'New report', body: opts.htmlBody };
  try {
    GmailApp.sendEmail(rec.to.join(','), subject, '', opts);
  } catch (err) {
    entry.result = 'Failed: ' + (err && err.message ? err.message : err);
    logEmail_(entry);
    throw err;
  }
  logEmail_(entry);
  Logger.log('Sent "' + team + '" to: ' + rec.to.join(',') + ' | cc: ' + rec.cc.join(','));
  return { sent: true };
}

// Sent-back notice. Same recipients as the assignment email. The completion
// photo (the after photo) is embedded in the letter.
function sendSentBack_(data, token, reason) {
  const team = data.team || 'Unknown / unsure';
  const rec = recipientsFor_(team);
  const issue = data.issueType ? phrase_(data.issueType) : '';
  const subject = 'Sent back — ' + team + (issue ? ': ' + issue : '');
  const blobs = photoBlobs_(data.completionPhoto);
  const cids = blobs.map(function (_, i) { return 'afterPhoto' + i; });
  const inlineImages = {};
  cids.forEach(function (cid, i) { inlineImages[cid] = blobs[i]; });
  const opts = {
    htmlBody: buildSentBack_(data, token, reason, cids),
    name: 'Engineering Student Project Teams',
    cc: rec.cc.join(','),
    replyTo: data.email || CONFIG.fallbackEmail,
  };
  if (CONFIG.sender) opts.from = CONFIG.sender;
  if (!opts.cc) delete opts.cc;
  const why = String(reason || '').trim();
  const entry = { type: 'Sent back', team: team, to: rec.to.join(','), cc: rec.cc.join(','), subject: subject, token: token, why: why ? 'Reason: ' + why : 'No reason given' };
  // Same pattern as the assignment email: inline photo only. If that throws,
  // send the letter with a Drive link so the notice still goes out.
  try {
    if (cids.length) opts.inlineImages = inlineImages;
    GmailApp.sendEmail(rec.to.join(','), subject, '', opts);
  } catch (err) {
    Logger.log('Send-back with photo failed, retrying without it: ' + err);
    delete opts.inlineImages;
    delete opts.from;
    opts.htmlBody = buildSentBack_(data, token, reason, []);
    try {
      GmailApp.sendEmail(rec.to.join(','), subject, '', opts);
    } catch (err2) {
      entry.result = 'Failed: ' + (err2 && err2.message ? err2.message : err2);
      entry.body = opts.htmlBody;
      logEmail_(entry);
      throw err2;
    }
  }
  entry.body = opts.htmlBody;
  logEmail_(entry);
  Logger.log('Sent back "' + team + '" to: ' + rec.to.join(',') + ' | photos: ' + cids.length);
  return { sent: true };
}

function photoBlobs_(cell) {
  const blobs = [];
  extractFileIds_(cell).forEach(function (id, i) {
    try {
      const file = DriveApp.getFileById(id);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      const blob = file.getBlob();
      const type = String(blob.getContentType() || '');
      if (type.indexOf('image/') !== 0) return;
      blob.setName('after-' + (i + 1) + '.jpg');
      blobs.push(blob);
    } catch (err) { Logger.log('Photo not attached ' + id + ': ' + err); }
  });
  return blobs;
}

function buildSentBack_(data, token, reason, cids) {
  const team = data.team || '';
  const issue = data.issueType ? phrase_(data.issueType) : '';
  const L = function (html) { return '<p style="margin:12px 0 0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + html + '</p>'; };
  let out = '<div style="max-width:600px;margin:0;padding:8px 6px;font-family:Georgia,serif;color:#1a1a1a">';
  out += '<p style="margin:0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + (team ? 'Dear ' + escapeHtml_(team) + ' team,' : 'Hello,') + '</p>';
  out += L('Your completion' + (issue ? ' of the ' + escapeHtml_(lcFirst_(issue)) + ' item' : '') + ' was sent back, so this is open again.');
  if (data.details) out += L('When this was posted, the reporter noted: &ldquo;' + escapeHtml_(data.details) + '&rdquo;');
  const why = String(reason || '').trim();
  if (why) out += L('<b>Reason:</b> ' + escapeHtml_(why));
  if (data.completionNote) out += L('Your note: &ldquo;' + escapeHtml_(data.completionNote) + '&rdquo;');
  out += L('Please make the correction, then mark it complete again.');
  out += addressedButton_(token, team);
  out += L('With thanks,<br>Engineering Student Project Teams');
  const ids = extractFileIds_(data.completionPhoto);
  if (cids && cids.length) {
    out += '<p style="margin:18px 0 6px;color:#777;font:13px Arial">The photo you submitted:</p>';
    cids.forEach(function (cid) {
      out += '<img src="cid:' + cid + '" width="240" style="border-radius:6px;margin:0 8px 8px 0">';
    });
  } else if (ids.length) {
    out += '<p style="margin:18px 0 6px;color:#777;font:13px Arial">The photo you submitted:</p>';
    ids.forEach(function (id) {
      out += '<p style="margin:0 0 8px"><a href="https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view">Open the after photo</a></p>';
    });
  }
  out += '</div>';
  return out;
}

// ---- Reminders ----

function dayStart_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function dateOrNull_(v) { if (!v) return null; const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
// Whole days from b to a. Rounded, so the 23- and 25-hour days at a clock change count as one.
function dayDiff_(a, b) { return Math.round((dayStart_(a) - dayStart_(b)) / 86400000); }

// Every issue the script has emailed about, and what the reminder job does with it.
// sendReminders sends the open rows marked send. The Reminder emails page
// (?view=emails) shows this same list, so the page and the job cannot disagree.
// Read-only: it never writes to the sheet.
//   open   not yet approved. state is 'open', 'sentback' or 'pending'.
//   closed approved (Addressed at is set). Never reminded.
// An admin can cancel reminders from the page (Reminder hold column, emCancelReminder).
function reminderPlan_() {
  const out = { open: [], closed: [], cols: { token: -1, addressed: -1, completed: -1, hold: -1 } };
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  if (!sh) return out;
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return out;
  const H = v[0].map(norm_);
  const ci = function (name) { return H.indexOf(norm_(name)); };
  const cTs = ci(CONFIG.headers.timestamp), cTeam = ci(CONFIG.headers.team);
  const cIssue = ci(CONFIG.headers.issueType), cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details);
  const cEmail = ci(CONFIG.headers.email), cNotified = ci(CONFIG.notifiedHeader), cToken = ci(CONFIG.issueTokenHeader);
  const cAddr = ci(CONFIG.addressedHeader), cLast = ci(CONFIG.lastReminderHeader), cCompleted = ci(CONFIG.completedAtHeader);
  const cSent = ci(CONFIG.sentBackHeader), cPhoto = ci(CONFIG.completionPhotoHeader), cHold = ci(CONFIG.reminderHoldHeader);
  out.cols = { token: cToken, addressed: cAddr, completed: cCompleted, hold: cHold };
  const cell = function (row, c) { return c >= 0 ? row[c] : ''; };
  const text = function (row, c) { const x = cell(row, c); return String(x == null ? '' : x).trim(); };
  const today = dayStart_(new Date());

  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    if (!cell(row, cNotified) || !cell(row, cToken)) continue;   // never emailed: not tracked
    const status = statusFromRow_(row, H);
    const p = {
      rowNum: i + 1, token: String(row[cToken]),
      team: text(row, cTeam), issueType: text(row, cIssue), action: text(row, cAction), details: text(row, cDetails),
      status: status, color: parseColor_(status), email: text(row, cEmail), timestamp: cell(row, cTs),
      reported: dateOrNull_(cell(row, cTs)), lastReminder: dateOrNull_(cell(row, cLast)),
      completedAt: dateOrNull_(cell(row, cCompleted)), addressed: dateOrNull_(cell(row, cAddr)),
      sentBackReason: text(row, cSent),
      state: 'open', send: false, next: null, deadline: null, interval: 0, why: '', hints: [],
      hold: '', skipped: null,   // hold: 'off' = reminders cancelled for the job, 'skip' = one reminder cancelled
    };
    if (cell(row, cAddr)) { p.state = 'done'; out.closed.push(p); continue; }   // approved: closed
    out.open.push(p);
    const holdRaw = cell(row, cHold);
    const holdOff = !!holdRaw && norm_(holdRaw) === 'off';
    if (holdOff) p.hold = 'off';
    if (cell(row, cCompleted)) {                                                // submitted, awaiting approval: paused
      p.state = 'pending';
      p.why = 'Paused. The team submitted this' + (p.completedAt ? ' on ' + fmtShort_(p.completedAt) : '') + ' and it is waiting for approval.';
      continue;
    }
    // A completion photo or a send-back reason with no Completed at: the team
    // submitted it and it was reopened.
    if (p.sentBackReason || text(row, cPhoto)) {
      p.state = 'sentback';
      p.hints.push('The team marked this complete and it was sent back' + (p.sentBackReason ? ' (' + p.sentBackReason + ')' : '') + '. Reminders run until they submit it again.');
    }
    const windowDays = p.color ? SLA_DAYS[p.color] : null;
    if (windowDays == null) {                                                   // purple / unknown -> no deadline, no reminders
      p.why = p.color === 'purple' ? 'No deadline (when time allows), so no reminders.' : 'No deadline on this report, so no reminders.';
      continue;
    }
    if (!p.reported) { p.why = 'The report has no date, so no reminders.'; continue; }
    const reportDay = dayStart_(p.reported);
    p.deadline = addDays_(reportDay, windowDays);
    p.interval = Math.max(1, Math.round(windowDays / 2));                       // recurring cadence = (due - report) / 2
    if (cell(row, cLast)) {
      if (!p.lastReminder) { p.why = 'The Last reminder cell is not a date, so no reminders. Clear the cell to restart them.'; continue; }
      p.next = addDays_(dayStart_(p.lastReminder), p.interval);                 // every `interval` days after the last reminder
    } else {
      p.next = p.deadline > reportDay ? p.deadline : addDays_(reportDay, 1);    // first reminder on the due date; day 0 is covered by the notice
    }
    if (holdOff) { p.next = null; p.why = 'Reminders are cancelled for this job.'; continue; }
    const holdUntil = dateOrNull_(holdRaw);
    if (holdUntil && dayStart_(holdUntil) > p.next) {                           // one reminder cancelled: wait for the hold date
      p.hold = 'skip'; p.skipped = p.next; p.next = dayStart_(holdUntil);
    }
    p.send = today > reportDay && today >= p.next;
    const every = 'every ' + (p.interval === 1 ? 'day' : p.interval + ' days');
    if (p.hold === 'skip' && !p.send) {
      p.why = 'The reminder for ' + fmtShort_(p.skipped) + ' was cancelled. The next one goes out ' + fmtShort_(p.next) + '.';
    } else if (p.send) {
      p.why = p.lastReminder
        ? 'Still open. Last reminder ' + fmtShort_(p.lastReminder) + '. Repeats ' + every + '.'
        : 'Still open. Due ' + fmtShort_(p.deadline) + '. First reminder.';
    } else {
      p.why = p.lastReminder
        ? 'Last reminder ' + fmtShort_(p.lastReminder) + '. Repeats ' + every + '.'
        : 'Due ' + fmtShort_(p.deadline) + '. The first reminder goes out then.';
    }
  }

  // The same problem filed twice (same team and issue type, reported within two
  // weeks). A team that finished one report still gets reminders for the other.
  const all = out.open.concat(out.closed);
  out.open.forEach(function (p) {
    if (p.state === 'pending' || !p.reported) return;
    const twins = all.filter(function (q) {
      return q !== p && q.reported && norm_(q.team) === norm_(p.team) && norm_(q.issueType) === norm_(p.issueType)
        && Math.abs(q.reported - p.reported) <= 14 * 86400000;
    });
    const finished = twins.filter(function (q) { return q.state === 'done' || q.state === 'pending'; });
    const stillOpen = twins.filter(function (q) { return q.state === 'open' || q.state === 'sentback'; });
    if (finished.length) {
      const f = finished[0], when = f.addressed || f.completedAt;
      p.hints.push('The team already finished a similar report' + (when ? ' on ' + fmtShort_(when) : '')
        + ' (reported ' + fmtShort_(f.reported) + '). This is a separate report and it is still open.');
    }
    if (stillOpen.length) p.hints.push('The team has ' + (stillOpen.length + 1) + ' open reports of this type. Each one sends its own reminder.');
  });
  return out;
}

// The reminder email for one plan row, exactly as sendReminders sends it.
function reminderMail_(p, contacts) {
  const rec = recipientsFor_(p.team || 'Unknown / unsure', contacts);
  return { to: rec.to.join(','), cc: rec.cc.join(','), subject: reminderSubject_(p), body: buildReminder_(p, p.color, p.token, p.deadline) };
}

function reminderSubject_(p) {
  return norm_(p.issueType) === 'informational'
    ? 'Reminder - ' + (p.team || '') + ' note is still open'
    : 'Reminder - ' + (p.team || '') + ' space issue needs attention' + (p.issueType ? ' (' + phrase_(p.issueType) + ')' : '');
}

function sendReminders() {
  // One run at a time. Two at once (a second trigger, or a manual run during the
  // daily one) would both read the sheet before either stamped Last reminder, and
  // every team would get the reminder twice.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('Another reminder run is in progress. Skipped.'); return; }
  let sent = 0, failed = 0, skipped = 0;
  try {
    const plan = reminderPlan_();
    const due = plan.open.filter(function (p) { return p.send; });
    if (due.length) {
      const sh = ss_().getSheetByName(CONFIG.responsesSheet);
      const lastCol = ensureColumn_(sh, CONFIG.lastReminderHeader);
      const width = sh.getLastColumn();
      const contacts = teamContacts_();
      due.forEach(function (p) {
        // Read the row again just before sending. Skip it if the team completed it, or
        // an admin closed, deleted or cancelled it, since the plan was read.
        const now = sh.getRange(p.rowNum, 1, 1, width).getValues()[0];
        const c = plan.cols;
        const holdNow = c.hold >= 0 ? now[c.hold] : '';
        const holdDate = dateOrNull_(holdNow);
        const held = !!holdNow && (norm_(holdNow) === 'off' || (!!holdDate && dayStart_(holdDate) > dayStart_(new Date())));
        if (String(now[c.token]) !== p.token || (c.addressed >= 0 && now[c.addressed]) || (c.completed >= 0 && now[c.completed]) || held) { skipped++; return; }
        const mail = reminderMail_(p, contacts);
        const opts = { htmlBody: mail.body, name: 'Engineering Student Project Teams', cc: mail.cc, replyTo: p.email || CONFIG.fallbackEmail };
        if (CONFIG.sender) opts.from = CONFIG.sender;
        const entry = { type: 'Reminder', team: p.team, to: mail.to, cc: mail.cc, subject: mail.subject, token: p.token, why: p.why, body: mail.body };
        try {
          GmailApp.sendEmail(mail.to, mail.subject, '', opts);
        } catch (err) {
          // One bad address must not stop the reminders after it.
          failed++;
          entry.result = 'Failed: ' + (err && err.message ? err.message : err);
          logEmail_(entry);
          Logger.log('Reminder to "' + p.team + '" failed: ' + err);
          return;
        }
        sh.getRange(p.rowNum, lastCol).setValue(new Date());
        sent++;
        logEmail_(entry);
      });
    }
  } finally {
    lock.releaseLock();
  }
  Logger.log('Reminders sent: ' + sent + (skipped ? ' | skipped (completed or cancelled during the run): ' + skipped : '') + (failed ? ' | failed: ' + failed : ''));
  // Fail the run so the trigger's failure email reports it.
  if (failed) throw new Error(failed + ' reminder email' + (failed === 1 ? '' : 's') + ' did not send. See the Email log tab.');
}

// ---- Email log ----

const EMAIL_LOG_HEADERS = ['Sent at', 'Type', 'Team', 'To', 'Cc', 'Subject', 'Issue token', 'Why', 'Result', 'Body'];

// The Email log tab: one row per email this script sends. Made on first use.
function emailLogSheet_(create) {
  const ss = ss_();
  let sh = ss.getSheetByName(CONFIG.emailLogSheet);
  if (sh || !create) return sh;
  try {
    sh = ss.insertSheet(CONFIG.emailLogSheet, ss.getSheets().length);
  } catch (err) {
    return ss.getSheetByName(CONFIG.emailLogSheet);   // another run made it first
  }
  const n = EMAIL_LOG_HEADERS.length;
  sh.getRange(1, 1, 1, n).setValues([EMAIL_LOG_HEADERS]);
  // Looks only. The Body column holds the email's HTML for the Reminder emails page.
  try { sh.getRange(1, 1, 1, n).setFontWeight('bold'); sh.setFrozenRows(1); sh.hideColumns(n); } catch (err) { /* cosmetic */ }
  return sh;
}

// Record one email. Never throws: a logging problem must not block or repeat a send.
//   e: { type, team, to, cc, subject, token, why, result, body }. result defaults to 'Sent'.
function logEmail_(e) {
  try {
    const sh = emailLogSheet_(true);
    sh.appendRow([new Date(), e.type || '', e.team || '', e.to || '', e.cc || '', e.subject || '', e.token || '',
      e.why || '', e.result || 'Sent', String(e.body || '').slice(0, 45000)]);
  } catch (err) {
    Logger.log('Email log write failed: ' + err);
  }
}

// The newest `limit` log rows, oldest first.
function emailLogRead_(limit) {
  const sh = emailLogSheet_(false);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit || 120, last - 1);
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  const v = sh.getRange(last - n + 1, 1, n, H.length).getValues();
  const col = {};
  ['Sent at', 'Type', 'Team', 'To', 'Cc', 'Subject', 'Issue token', 'Why', 'Result', 'Body'].forEach(function (h) { col[h] = H.indexOf(norm_(h)); });
  const get = function (row, h) { return col[h] >= 0 && row[col[h]] != null ? row[col[h]] : ''; };
  const out = [];
  v.forEach(function (row) {
    const at = dateOrNull_(get(row, 'Sent at'));
    if (!at) return;
    out.push({
      at: at, type: String(get(row, 'Type')).trim(), team: String(get(row, 'Team')).trim(),
      to: String(get(row, 'To')).trim(), cc: String(get(row, 'Cc')).trim(), subject: String(get(row, 'Subject')).trim(),
      token: String(get(row, 'Issue token')).trim(), why: String(get(row, 'Why')).trim(),
      result: String(get(row, 'Result')).trim() || 'Sent', body: String(get(row, 'Body')), logged: true,
    });
  });
  return out;
}

// Team Contacts, read once and keyed by normalized team name.
function teamContacts_() {
  const map = {};
  const sh = ss_().getSheetByName(CONFIG.contactsSheet);
  if (!sh) { Logger.log('No "' + CONFIG.contactsSheet + '" tab.'); return map; }
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return map;
  const H = v[0].map(norm_);
  const idx = function (name) { return H.indexOf(norm_(name)); };
  const cTeam = idx('Team'), cGen = idx('Team general email'), cLia = idx('Liaison email'),
        cLiaN = idx('Liaison name'), cSafe = idx('Safety lead email'), cExtra = idx('Extra CC');
  if (cTeam < 0) { Logger.log('Team Contacts has no "Team" header.'); return map; }
  for (let i = 1; i < v.length; i++) {
    const key = norm_(v[i][cTeam]);
    if (!key || map[key]) continue;   // first row for a team wins
    map[key] = {
      team: v[i][cTeam],
      generalEmail: cGen >= 0 ? v[i][cGen] : '', liaisonEmail: cLia >= 0 ? v[i][cLia] : '',
      liaisonName: cLiaN >= 0 ? v[i][cLiaN] : '', safetyEmail: cSafe >= 0 ? v[i][cSafe] : '',
      extraCc: cExtra >= 0 ? v[i][cExtra] : '',
    };
  }
  return map;
}

function lookupTeam_(team, contacts) {
  return (contacts || teamContacts_())[norm_(team)] || null;
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

// Current Status drives the email subject, the deadline, and reminders.
// Older rows start with a color ("Red - Code Compliance..."). The form now
// stores a timeframe only. Both wordings must map to the same SLA key.
// Keep this in sync with parseColor_ in command-center/Code.gs.
function parseColor_(status) {
  const t = String(status || '').trim().toLowerCase().replace(/-/g, ' ').replace(/\s+/g, ' ');
  if (!t) return '';
  const word = t.match(/\b(red|orange|yellow|ivory|purple)\b/);
  if (word) return word[1];
  if (t.indexOf('same day') >= 0) return 'red';
  if (t.indexOf('14 to 21') >= 0) return 'ivory';
  if (t.indexOf('7 to 10') >= 0) return 'yellow';
  if (t.indexOf('2 to 3') >= 0) return 'orange';
  if (t.indexOf('when time allows') >= 0 || t.indexOf('as time permits') >= 0) return 'purple';
  return '';
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

function addressedButton_(token, team) {
  // Must be the published /exec URL from CONFIG.webAppUrl.
  // Do NOT use ScriptApp.getService().getUrl() here: from a trigger it returns the
  // editor-only /dev URL, which recipients cannot open ("unable to open the file").
  if (!CONFIG.webAppUrl || !token) return '';
  const url = CONFIG.webAppUrl + '?id=' + encodeURIComponent(token);
  const label = 'Mark complete';
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:20px 0 4px"><tr>'
    + '<td style="background:#8f1515;border-radius:6px"><a href="' + url + '" style="display:inline-block;padding:12px 24px;color:#fff;font:bold 14px Arial,sans-serif;text-decoration:none;border-radius:6px">' + label + '</a></td>'
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
  const infoOnly = norm_(data.issueType) === 'informational';
  out += '<p style="margin:0;color:#1a1a1a;font:15px/1.7 Georgia,serif">' + (team ? 'Dear ' + escapeHtml_(team) + ' team,' : 'Hello,') + '</p>';
  if (infoOnly) {
    out += L('This is for your awareness. No action is requested.');
  } else {
    out += L(escapeHtml_(sev.open));
    let pb = '';
    if (issue) pb += 'The concern relates to ' + escapeHtml_(lcFirst_(issue));
    if (action) pb += (pb ? ', and we ask that you ' : 'We ask that you ') + escapeHtml_(lcFirst_(action));
    if (pb) out += L(pb + '.');
  }
  if (data.details) out += L('The reporter noted: &ldquo;' + escapeHtml_(data.details) + '&rdquo;');
  if (!infoOnly) out += L(escapeHtml_(sev.ask) + '.' + (dl ? ' The deadline is <b>' + escapeHtml_(fmtDate_(dl)) + '</b>.' : ''));
  if (typeof icIsOpsTeam_ === 'function' && icIsOpsTeam_(team)) {
    out += L('Tap <b>Mark complete</b> to mark it done. No photo is needed.');
  }
  out += addressedButton_(token, team);
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
  const infoOnly = norm_(data.issueType) === 'informational';
  const sentBack = data.state === 'sentback';   // set by reminderPlan_
  if (infoOnly) {
    out += L('This is a follow-up on a note for your team'
      + (reportedOn ? ', logged on <b>' + escapeHtml_(reportedOn) + '</b>' : '')
      + '. No action was requested. It is still open.');
  } else {
    if (sentBack) {
      // The team did mark it complete, so "not yet marked as addressed" would be wrong.
      out += L('This is a follow-up on a space issue in your team\'s space'
        + (reportedOn ? ', reported on <b>' + escapeHtml_(reportedOn) + '</b>' : '')
        + '. Your team marked it complete and it was sent back, so it is still open.');
      if (data.sentBackReason) out += L('<b>Reason:</b> ' + escapeHtml_(data.sentBackReason));
    } else {
      out += L('This is a follow-up on a space issue in your team\'s space'
        + (reportedOn ? ', reported on <b>' + escapeHtml_(reportedOn) + '</b>,' : '')
        + ' that has not yet been marked as addressed. It was due on <b>' + escapeHtml_(fmtDate_(deadline)) + '</b>.');
    }
    let pb = '';
    if (issue) pb += 'The concern relates to ' + escapeHtml_(lcFirst_(issue));
    if (action) pb += (pb ? ', and we ask that you ' : 'We ask that you ') + escapeHtml_(lcFirst_(action));
    if (pb) out += L(pb + '.');
  }
  if (data.details) out += L('The reporter noted: &ldquo;' + escapeHtml_(data.details) + '&rdquo;');
  if (!infoOnly) out += L(sentBack ? 'Please make the correction, then mark it complete again.' : 'Please resolve it as soon as possible.');
  if (typeof icIsOpsTeam_ === 'function' && icIsOpsTeam_(team)) {
    out += L('Tap <b>Mark complete</b> to mark it done. No photo is needed.');
  }
  out += addressedButton_(token, team);
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
  if (p.module === 'proposals') return proposalsPage_(embed, admin, p.back);   // Honeybee-swarm improvement voting (07_proposals.gs); ?back= is the hub link for the back button
  if (p.module === 'projects') return projectsPage_(embed, admin, p.back);   // Multi-user projects (04_tasks_projects.gs)
  if (p.module === 'projects-dash') return projectsDashboardPage_(embed, p.back);   // Projects dashboard (04_tasks_projects.gs)
  // The old equipment registry was retired. Old links and printed QR labels land here
  // instead of falling through to the ?id= handler below.
  if (p.registry || p.module === 'registry-dash') return htmlPage_('Inventory has moved', 'The old equipment registry was retired. Tools and equipment are now in Tool inventory on the Ops Hub. <a href="https://cornellprojectteams.github.io/centralhub/inventory.html" target="_top" style="color:#b31b1b;font-weight:700">Open the inventory</a>');
  if (p.team) { const t = lookupTeamByToken_(p.team); return t ? teamPortal_(t, false) : htmlPage_('Invalid link', 'This team link is not recognized.'); }
  if (p.view === 'all') return allIssuesPage_(embed, admin, p.back);
  // Reminder emails: what was sent, to whom, and what the next run sends. Admin page only.
  if (p.view === 'emails') return admin ? emailsPage_(embed, p.back) : htmlPage_('Invalid link', 'This page opens from the admin page.');
  if (p.view) return teamPortal_(String(p.view), !(p.act === '1' || p.act === 'true'), p.embed === '1' || p.embed === 'true');   // ?view=<team> read-only; &act=1 markable; &embed=1 no top bar
  if (p.id) return confirmPage_(p.id);
  return pickerPage_();
}

// The webfonts are pulled in by prFontLinks_ in the document head, NOT by an @import
// here. An @import inside a <style> is only discovered after this whole sheet parses,
// and it blocks rendering while it resolves, so it serialised two round trips in front
// of first paint on every page.
function portalStyles_() {
  return ''
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
    // Admin quick-close: a distinct warm-toned override, set apart from the doer buttons.
    + '.btn-skip{color:#9a6a12;background:#fbf4e4;border:1.5px solid #eeddb4;font-weight:700}'
    + '.btn-skip svg{opacity:.85}'
    + '.btn-skip:hover{color:#7a5410;background:#f7edd2;border-color:#e2c882}'
    + '.btn-skip:active{background:#f0e4c4}'
    + '.btn-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end}'
    + '.stats{display:flex;flex-wrap:wrap;gap:10px;margin-top:14px}'
    + '.stat{flex:1 1 8rem;min-width:0;padding:16px 18px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.06);position:relative;overflow:hidden}'
    + '.stat::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,#8f1515,#b31b1b,#f0c050)}'
    + '.stat-label{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c}'
    + '.stat-val{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:28px;font-weight:800;letter-spacing:-.04em;line-height:1.1;margin-top:6px;color:#111}'
    + '.stat-val--danger{color:#b31b1b}'
    + '.filters{display:flex;flex-wrap:wrap;gap:10px;margin:18px 0 10px;padding:14px 16px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.search-wrap{flex:1;min-width:0;position:relative}'
    + '.search-wrap input{width:100%;font:inherit;font-size:14px;padding:12px 14px 12px 40px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fafaf8;outline:none;transition:border-color .15s,box-shadow .15s}'
    + '.search-wrap input:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12);background:#fff}'
    + '.search-wrap::before{content:"";position:absolute;left:14px;top:50%;transform:translateY(-50%);width:16px;height:16px;background:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'16\' height=\'16\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%239a9a96\' stroke-width=\'2.2\' stroke-linecap=\'round\'%3E%3Ccircle cx=\'11\' cy=\'11\' r=\'7\'/%3E%3Cline x1=\'21\' y1=\'21\' x2=\'16.65\' y2=\'16.65\'/%3E%3C/svg%3E") center/contain no-repeat;pointer-events:none}'
    + '.filters select{font:inherit;font-size:14px;font-weight:600;padding:12px 14px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;outline:none;cursor:pointer;min-width:0;flex:1 1 8rem;max-width:100%;color:#333}'
    + '.filters select:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12)}'
    + '.toggle{display:inline-flex;align-items:center;gap:8px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#444;padding:10px 14px;border-radius:10px;border:1.5px solid #e0e0dc;background:#fff;cursor:pointer;user-select:none;transition:border-color .15s,background .15s,box-shadow .15s}'
    + '.toggle:hover{border-color:#ccc;box-shadow:0 2px 6px rgba(20,20,30,.06)}'
    + '.toggle.is-on,.toggle:has(input:checked){border-color:#b31b1b;background:#fff8f8;color:#8f1515;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.toggle input{width:16px;height:16px;accent-color:#b31b1b;cursor:pointer;margin:0}'
    + '.section-head{display:flex;align-items:center;gap:12px;margin:0}'
    + '.section-label{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:16px;font-weight:700;letter-spacing:-.02em;color:#14110e;flex:1;min-width:0}'
    + '.section-label--late{color:#b31b1b}'
    + '.section-label--open{color:#14110e}'
    + '.section-count{flex:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:800;font-variant-numeric:tabular-nums;min-width:1.75rem;height:1.75rem;padding:0 8px;border-radius:999px;background:#efece6;color:#14110e;display:inline-flex;align-items:center;justify-content:center}'
    + '.section[data-sec-id="overdue"] .section-count{background:#fdecec;color:#b31b1b}'
    + '.section[data-sec-id="pending"] .section-count{background:#fdf2df;color:#9a5a00}'
    + '.section[data-sec-id="done"] .section-count{background:#e7f6ee;color:#157a47}'
    + '.section[data-sec-id="active"] .section-count{background:#fdf3e3;color:#c56a10}'
    // iOS-style grouped accordion. One tappable row per group: title, count badge,
    // chevron that points down (expands) not right (navigates). 52px hit target.
    + '.sec-list{margin-top:14px;background:#fff;border:1.5px solid #e2ddd6;border-radius:16px;overflow:hidden;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.section{margin:0;border-bottom:1px solid #efece6}'
    + '.section:last-child{border-bottom:0}'
    + '.section > .section-head{cursor:pointer;list-style:none;display:flex;align-items:center;width:100%;box-sizing:border-box;min-height:52px;margin:0;padding:12px 16px;background:#fff;border:0;border-radius:0;-webkit-user-select:none;user-select:none;-webkit-tap-highlight-color:transparent;touch-action:manipulation}'
    + '.section > .section-head::-webkit-details-marker{display:none}'
    + '.section > .section-head::marker{content:none}'
    + '.section > .section-head:active{background:#f3f0ea}'
    + '@media(hover:hover) and (pointer:fine){.section > .section-head:hover{background:#faf9f6}}'
    + '.section[open] > .section-head{background:#faf8f4}'
    + '.section-chev{flex:none;width:28px;height:28px;border-radius:50%;background:#efece6;position:relative}'
    + '.section-chev::before{content:"";position:absolute;top:9px;left:8px;width:8px;height:8px;border-right:2.5px solid #5c574f;border-bottom:2.5px solid #5c574f;transform:rotate(45deg);transition:transform .2s cubic-bezier(.2,.8,.3,1),top .2s}'
    + '.section[open] > .section-head .section-chev,.pr-card[open] > .pr-sum .section-chev{background:#f3e4e4}'
    + '.section[open] > .section-head .section-chev::before,.pr-card[open] > .pr-sum .section-chev::before{top:11px;border-color:#8f1515;transform:rotate(-135deg)}'
    + '.section-body{padding:8px 10px 12px;background:#f6f4ef;border:0;border-radius:0;animation:secOpen .22s cubic-bezier(.2,.8,.3,1) both}'
    + '.sec-empty{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:600;color:#8a857c;text-align:center;padding:18px 12px}'
    + '.ic-toolbar{display:flex;align-items:center;justify-content:flex-end;gap:10px;margin:12px 0 0}'
    + '.card-sum{display:flex;align-items:center;gap:10px;min-height:52px;padding:13px 14px 10px;cursor:pointer;list-style:none;width:100%;box-sizing:border-box;position:relative;-webkit-user-select:none;user-select:none;-webkit-tap-highlight-color:transparent;touch-action:manipulation}'
    + '.card-sum .card-accent{position:absolute;left:0;right:0;top:0;height:3px;display:block}'
    + '.card-sum::-webkit-details-marker{display:none}'
    + '.card-sum::marker{content:none}'
    + '.card-sum:active{background:#f3f0ea}'
    + '@media(hover:hover) and (pointer:fine){.card-sum:hover{background:#faf9f6}}'
    + '.card-sum-main{flex:1;min-width:0}'
    + '.card-sum-title{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:16px;font-weight:700;letter-spacing:-.02em;line-height:1.25;color:#14110e;overflow-wrap:anywhere}'
    + '.card-sum-action{display:block;margin-top:3px;font-size:13.5px;font-weight:600;color:#26231f;line-height:1.35;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    + '.card-sum-alabel{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#a8a29e;margin-bottom:1px}'
    + '.card-sum-sub{display:block;margin-top:2px;font-size:12.5px;font-weight:600;color:#8a857c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    + '.card-sum-stat{flex:none;max-width:42%}'
    + 'details.card:not([open]){margin-bottom:8px}'
    + 'details.card[open]>.card-sum{border-bottom:1px solid #f0eee8}'
    + 'details.card[open]>.card-sum .card-sum-action{white-space:normal}'
    + 'details.card[open]>.card-sum .section-chev{background:#f3e4e4}'
    + 'details.card[open]>.card-sum .section-chev::before{top:11px;border-color:#8f1515;transform:rotate(-135deg)}'
    + '@keyframes secOpen{from{opacity:0}to{opacity:1}}'
    + '@media(prefers-reduced-motion:reduce){.section-body{animation:none}.section-chev::before{transition:none}}'
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
    + '.card-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;letter-spacing:-.02em;line-height:1.25;margin-top:3px;color:#111;overflow-wrap:anywhere}'
    + '.card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;min-width:0}'
    + '.card-field{margin-top:13px}'
    + '.card-flabel{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#a8a29e;margin-bottom:4px}'
    + '.card-action{font-size:14.5px;color:#26231f;font-weight:600;line-height:1.5;overflow-wrap:anywhere}'
    + '.card-details{font-size:13.5px;color:#57534e;line-height:1.62;white-space:pre-line;padding-left:12px;border-left:2px solid #ece9e2;overflow-wrap:anywhere}'
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
    + '@media(max-width:600px){.card-foot{flex-direction:column;align-items:stretch}.card-foot>:first-child{margin-right:0}.btn-row{width:100%}.search-wrap{flex:1 1 100%}.filters select{width:100%;flex-basis:100%}.stat-val{font-size:24px}.card-body{padding:15px 14px}.card-foot{padding:12px 14px}.swh{font-size:26px!important}}';
}

// Fonts, fetched without blocking first paint: warm the connections up front, then load
// the stylesheet at print media so it is non-render-blocking and flip it to all once it
// arrives. Inter/Jakarta already carry display=swap, so text paints immediately in the
// fallback and swaps in place rather than sitting invisible.
function fontLinks_() {
  var href = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap';
  return '<link rel="preconnect" href="https://fonts.googleapis.com">'
    + '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
    + '<link rel="stylesheet" href="' + href + '" media="print" onload="this.media=\'all\';this.onload=null">'
    + '<noscript><link rel="stylesheet" href="' + href + '"></noscript>';
}

function swissShell_(innerHtml, pageTitle, wide, embedded, maxWidth, bg) {
  const maxW = maxWidth || (wide ? '960px' : '600px');
  const pad = embedded ? '14px 12px 28px' : '40px 28px 64px';
  const padSm = embedded ? '10px 10px 24px' : '24px 14px 40px';
  const topBar = embedded ? '' : '<div style="height:5px;background:linear-gradient(90deg,#8f1515,#b31b1b,#8f1515)"></div>';
  const enh = fontLinks_() + '<style>' + portalStyles_()
    + '.swiss-inner{max-width:' + maxW + ';width:100%;margin:0 auto;padding:' + pad + ';min-width:0;box-sizing:border-box}'
    + '@media(max-width:600px){.page-title{font-size:24px !important}.swiss-inner{padding:' + padSm + ' !important}}'
    + '</style>';
  const html = enh
    + '<div style="margin:0;background:' + (bg || '#f5f4f0') + ';min-height:100vh;max-width:100%;font-family:Helvetica,Arial,sans-serif;color:#111">'
    + topBar
    + '<div class="swiss-inner">' + innerHtml + '</div>'
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

// Reached from the action button in the notification email (?id=token).
// Operations Team mail goes to an admin: mark complete with no photo, after confirm.
// Other teams need an evidence photo (or a photo-optional Mark done) -> Pending approval.
function confirmPage_(id) {
  const info = findIssue_(id);
  if (!info) return htmlPage_('Not found', 'We could not find that item. It may have been removed.');
  if (info.addressed) return htmlPage_('Already completed', 'This was approved as complete on ' + escapeHtml_(fmtShort_(info.addressed)) + '.');

  // The Operations email goes to Noah (an admin), so it marks complete with no photo
  // after a confirm — even if a student already submitted it for review.
  if (icIsOpsTeam_(info.team)) {
    const opsBanner = info.sentBackReason
      ? '<div style="margin-top:14px;font:600 14px/1.6 Arial,sans-serif;color:#8a4b00;background:#fdf2df;border:1px solid #f4dfb0;border-radius:10px;padding:11px 14px"><b>Sent back:</b> ' + escapeHtml_(info.sentBackReason) + '</div>'
      : '';
    const opsInner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
      + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">Mark complete</div>'
      + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:12px">This marks the task complete with no photo. You will be asked to confirm.</div>'
      + opsBanner
      + '<div id="act" style="margin-top:24px"><button type="button" class="btn btn-primary" onclick="cfAsk()">Mark complete</button></div>'
      + '<div id="done" style="display:none;font-size:22px;font-weight:800;letter-spacing:-.02em;color:#1d7a46;margin-top:14px"></div>'
      + '<script>function cfAsk(){var a=document.getElementById("act");a.innerHTML="<span class=\\"tp-hint\\" style=\\"margin-right:6px\\">Mark this complete with no photo?</span><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"cfResolve()\\">Yes, mark complete</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"cfCancel()\\">Cancel</button>";}'
      + 'function cfCancel(){document.getElementById("act").innerHTML="<button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"cfAsk()\\">Mark complete</button>";}'
      + 'function cfResolve(){var a=document.getElementById("act");a.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
      + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Could not save")+"</span>";return;}try{tpConfetti();}catch(e){}a.style.display="none";var d=document.getElementById("done");d.style.display="block";d.innerHTML="\\u2713 Completed";'
      + '}).withFailureHandler(function(){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Could not save. Please retry.</span>";}).resolveIssueComplete(' + JSON.stringify(id) + ');}</script>';
    return swissShell_(tpStyles_() + opsInner + tpSharedJs_(), 'Space Status');
  }

  if (info.completedAt) return htmlPage_('Pending approval', 'This was already submitted on ' + escapeHtml_(fmtShort_(info.completedAt)) + '. It is waiting for an admin to review.');

  const photoOptional = icPhotoOptional_(info.action, info.issueType);
  const sentBackBanner = info.sentBackReason
    ? '<div style="margin-top:14px;font:600 14px/1.6 Arial,sans-serif;color:#8a4b00;background:#fdf2df;border:1px solid #f4dfb0;border-radius:10px;padding:11px 14px"><b>Sent back:</b> ' + escapeHtml_(info.sentBackReason) + '</div>'
    : '';
  const lead = photoOptional
    ? 'Tap <b>Mark done</b> and an admin gives it a quick review. A photo or a note is optional — add either if it helps.'
    : 'Add one or more photos of the completed work. A note is optional.';
  const controls = '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'cf-file\').click()">Add photos</button>'
    + '<button type="button" class="btn btn-primary" id="cf-done" onclick="cfComplete()">' + (photoOptional ? 'Mark done' : 'Complete') + '</button>'
    + '<span id="cf-stage" class="tp-hint"></span>';
  const inner = '<div style="font-size:11px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:#999">Space Status</div>'
    + '<div class="swh" style="font-size:30px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-top:16px">Mark complete</div>'
    + '<div style="font-size:16px;line-height:1.7;color:#555;margin-top:12px">' + lead + '</div>'
    + sentBackBanner
    + '<div class="ic-reply" style="margin-top:18px">'
    +   '<label class="ic-reply-lbl" for="cf-note">Note <span class="tp-opt">optional</span></label>'
    +   '<textarea id="cf-note" class="ic-reply-in" rows="3" placeholder="What you did, or anything the reviewer should know">' + escapeHtml_(info.completionNote || '') + '</textarea>'
    +   '<span class="ic-reply-help">This is sent with your completion. Leave it blank if the photo is enough.</span>'
    + '</div>'
    + '<div id="act" style="margin-top:16px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
    +   '<input type="file" accept="image/*" multiple id="cf-file" style="display:none" onchange="cfPick(this)">'
    +   controls
    + '</div>'
    + '<div id="cf-photo"></div>'
    + '<div id="done" style="display:none;font-size:22px;font-weight:800;letter-spacing:-.02em;color:#1d7a46;margin-top:14px"></div>'
    + '<script>'
    + 'var CFBUF=[];'
    + 'function cfNote(){var el=document.getElementById("cf-note");return el?(el.value||"").trim():"";}'
    + 'function cfPick(input){tpStageRead(input,CFBUF,cfRenderStage);}'
    + 'function cfRenderStage(){document.getElementById("cf-photo").innerHTML=tpStagePreview(CFBUF,"cfUnstage(","document.getElementById(\'cf-file\').click()");var h=document.getElementById("cf-stage");if(h){h.style.color="";h.textContent="";}var db=document.getElementById("cf-done");if(db)db.textContent=CFBUF.length?("' + (photoOptional ? 'Submit' : 'Complete') + ' \\u00b7 "+CFBUF.length+" photo"+(CFBUF.length===1?"":"s")):"' + (photoOptional ? 'Mark done' : 'Complete') + '";}'
    + 'function cfUnstage(idx){CFBUF.splice(idx,1);cfRenderStage();}'
    + 'function cfDoneUi(ids,closed){try{tpConfetti();}catch(e){}var a=document.getElementById("act");a.style.display="none";ids=Array.isArray(ids)?ids:(ids?[ids]:[]);document.getElementById("cf-photo").innerHTML=tpThumbs(ids);var n=document.getElementById("cf-note");if(n)n.readOnly=true;var d=document.getElementById("done");d.style.display="block";d.innerHTML=closed?"\\u2713 Completed":"\\u2713 Submitted, pending approval";}'
    + 'function cfComplete(){var a=document.getElementById("act");var h=document.getElementById("cf-stage");var note=cfNote();'
    + 'if(' + (photoOptional ? 'false' : '!CFBUF.length') + '){if(h){h.style.color="#b31b1b";h.textContent="Add at least one photo first.";}return;}'
    + 'a.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'if(!CFBUF.length){google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Could not save")+"</span>";return;}cfDoneUi([],r.status==="Completed");}).withFailureHandler(function(){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Could not save. Please retry.</span>";}).submitIssueCompletion(' + JSON.stringify(id) + ',"","",false,note);return;}'
    + 'tpUploadStaged(CFBUF,' + JSON.stringify(id) + ',function(done,total){a.innerHTML="<span class=\\"tp-hint\\">Uploading photo "+(done+1)+" of "+total+"\\u2026</span>";},'
    + 'function(ids){cfDoneUi(ids);},'
    + 'function(msg){a.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+msg+"</span>";},note);}'
    + '</script>';
  return swissShell_(tpStyles_() + inner + tpSharedJs_(), 'Space Status');
}

// Completed-item cards. Folded like the rest; shown when the "recently completed" chip is on.
function icDoneCards_(doneList, showTeam, startOpen) {
  if (!doneList || !doneList.length) return '';
  let s = '';
  doneList.forEach(function (d) {
    // Show the completion photo(s); if the task was finished without one (one-tap /
    // photo-optional / admin close), fall back to the originally reported photo so the
    // card still has an image. onerror hides any thumbnail whose Drive file will not load.
    let ids = extractFileIds_(d.completionPhoto);
    let capt = 'Completed';
    if (!ids.length) { ids = extractFileIds_(d.reportedPhoto || ''); if (ids.length) capt = 'Reported'; }
    const thumb = ids.length
      ? '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:10px">' + ids.map(function (fid) {
          const e = encodeURIComponent(fid);
          return '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" style="display:inline-block;line-height:0" title="' + capt + ' photo"><img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w240" loading="lazy" decoding="async" onerror="this.closest(\'a\').style.display=\'none\'" style="max-height:120px;border-radius:10px;border:1px solid #ececec"></a>';
        }).join('') + '</div>'
      : '';
    const title = escapeHtml_((showTeam && d.team ? d.team + ' · ' : '') + (d.issueType ? phrase_(d.issueType) : 'Issue'));
    const when = escapeHtml_(fmtShort_(d.addressed));
    const hay = ((d.team || '') + ' ' + (d.issueType || '') + ' ' + (d.details || '') + ' ' + (d.completionNote || '')).toLowerCase();
    const desc = d.details ? escapeHtml_(d.details) : '';
    s += '<details class="card" data-bucket="done" data-hay="' + escapeHtml_(hay) + '"' + (d.team ? ' data-team="' + escapeHtml_(d.team) + '"' : '') + icOpenAttr_(startOpen) + ' style="display:none">'
      + cardSum_(title, 'Completed ' + when, '<span class="tp-pill tp-pill--done">Completed</span>',
          '<span class="card-accent" style="background:#157a47"></span>', desc, 'Details')
      + '<div class="card-body">'
      +   (d.completionNote ? '<div class="card-field"><span class="card-flabel">Response</span><div class="card-details">' + escapeHtml_(d.completionNote) + '</div></div>' : '')
      +   thumb
      + '</div>'
      + '<div class="card-foot"><span class="due due--done">&#10003; Completed ' + when + '</span></div>'
      + '</details>';
  });
  return s;
}

// ---- Open issues pages: look ----
// The all-teams dashboard and the team pages share their base styles with Projects
// and Proposals (portalStyles_, tpStyles_). This sheet goes after those two, on the
// issues pages only, and restyles them: rows on paper in place of shadowed cards, one
// color per state, pill buttons. It changes looks only; wording and markup are as before.
function icSkin_() {
  const hair = 'rgba(20,17,14,.14)';
  const edge = 'rgba(20,17,14,.2)';
  // The photos come from Drive, by way of a second host. Opening both connections
  // while the page is still being read saves a round trip on the first photo.
  return '<link rel="preconnect" href="https://drive.google.com"><link rel="preconnect" href="https://lh3.googleusercontent.com">'
    + '<style>'
    + '.page-kicker{color:#b31b1b}'
    + '.page-title{font-size:38px;letter-spacing:-.045em;color:#14110e}'
    + '.page-rule{width:64px;height:5px;background:linear-gradient(90deg,#b31b1b,#e08a1e,#f0c050)}'
    // Open / pending / completed: pills. The one that is on fills with its color.
    + '.ic-summary{margin:18px 0 16px}'
    + '.ic-seg{gap:8px;padding:0;background:none;border:0;border-radius:0}'
    + '.ic-sum{font-size:13.5px;font-weight:700;color:#14110e;background:#f1e9dc;border-radius:999px;padding:9px 16px}'
    + '.ic-sum:hover:not(:disabled):not(.is-on){background:#e8dcc9;color:#14110e}'
    + '.ic-sum:focus-visible{box-shadow:0 0 0 3px rgba(179,27,27,.25)}'
    + '.ic-sum b{color:#b31b1b}'
    + '.ic-sum--pending b{color:#8a5200}'
    + '.ic-sum--done b{color:#157a47}'
    + '.ic-sum.is-on{background:#b31b1b;color:#fff;box-shadow:none}'
    + '.ic-sum.is-on b{color:#fff}'
    + '.ic-sum--pending.is-on{background:#b06a00;color:#fff;box-shadow:none}'
    + '.ic-sum--pending.is-on b{color:#fff}'
    + '.ic-sum--done.is-on{background:#157a47;color:#fff;box-shadow:none}'
    + '.ic-sum--done.is-on b{color:#fff}'
    + '.ic-unlock,.ic-refresh{color:#14110e;background:none;border:1.5px solid #14110e;border-radius:999px;padding:8px 15px}'
    + '.ic-unlock:hover,.ic-refresh:hover{color:#fff;background:#14110e;border-color:#14110e;box-shadow:none}'
    + '.ic-unlock.is-on{color:#fff;background:#14110e;border-color:#14110e;box-shadow:none}'
    + '.ic-unlock:focus-visible{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.2)}'
    // Search, team and overdue: no box around them.
    + '.filters{gap:10px;margin:6px 0 4px;padding:0;background:none;border:0;border-radius:0;box-shadow:none}'
    + '.search-wrap input{padding:11px 16px 11px 42px;border:1.5px solid ' + edge + ';border-radius:999px;background:#fff}'
    + '.search-wrap input:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.14)}'
    + '.search-wrap::before{left:16px}'
    + '.filters select{padding:11px 16px;border:1.5px solid ' + edge + ';border-radius:999px;font-weight:700;color:#14110e}'
    + '.toggle{padding:10px 16px;border:1.5px solid ' + edge + ';border-radius:999px;color:#14110e}'
    + '.toggle:hover{border-color:#14110e;box-shadow:none}'
    + '.toggle.is-on,.toggle:has(input:checked){border-color:#b31b1b;background:#b31b1b;color:#fff;box-shadow:none}'
    + '.toggle.is-on input,.toggle:has(input:checked) input{accent-color:#fff}'
    // Each issue: a row with its urgency color down the left edge. Open, it becomes a
    // white sheet so its details and buttons hold together.
    + '.card{margin-bottom:0;border:0;border-bottom:1px solid ' + hair + ';border-radius:0;background:none;box-shadow:none}'
    + '.card:hover{box-shadow:none;transform:none}'
    + 'details.card:not([open]){margin-bottom:0}'
    + 'details.card[open]{margin:12px 0;border:1.5px solid rgba(20,17,14,.12);border-radius:18px;background:#fff}'
    + '.card-sum{gap:14px;min-height:0;padding:16px 4px 16px 20px}'
    + '.card-sum .card-accent{left:0;right:auto;top:15px;bottom:15px;width:6px;height:auto;border-radius:99px}'
    + 'details.card[open]>.card-sum{padding:16px 16px 16px 30px;border-bottom:1px solid ' + hair + '}'
    + 'details.card[open]>.card-sum .card-accent{left:12px}'
    + '.card-sum:active{background:none}'
    + '@media(hover:hover) and (pointer:fine){.card-sum:hover{background:none}.card-sum:hover .card-sum-title{color:#b31b1b}}'
    + '.card-sum-title{font-size:16.5px;font-weight:800;color:#14110e;transition:color .15s}'
    + '.card-sum-action{font-size:14px;color:#14110e}'
    + '.card-sum-alabel{color:#b31b1b}'
    + '.card-sum-sub{font-size:13px;color:#3d3832}'
    + '.section-chev{width:22px;height:22px;background:none}'
    + '.section-chev::before{top:5px;left:6px;border-color:#14110e}'
    + 'details.card[open]>.card-sum .section-chev{background:none}'
    + 'details.card[open]>.card-sum .section-chev::before{top:8px;border-color:#b31b1b}'
    // State pills, filled.
    + '.chip,.tp-pill{padding:6px 12px;font-size:10.5px;letter-spacing:.08em}'
    + '.chip--overdue{color:#fff;background:#b31b1b;border:0}'
    + '.chip--due{color:#1a468f;background:#e2ecfc;border:0}'
    + '.tp-pill--pending{color:#fff;background:#b06a00;border:0}'
    + '.tp-pill--sent{color:#8f1515;background:#fde0e0;border:0}'
    + '.tp-pill--done{color:#fff;background:#157a47;border:0}'
    + '.card-body{padding:18px 22px 16px 30px}'
    + '.card-flabel,.tp-photo figcaption{color:#b31b1b}'
    + '.card-action{color:#14110e}'
    + '.card-details{padding-left:14px;border-left:3px solid #f0c050;color:#14110e}'
    + '.card-body img{border-radius:14px!important;border-color:rgba(20,17,14,.1)!important}'
    + '.card-foot{padding:14px 22px 16px 30px;border-top:1px solid ' + hair + ';background:none}'
    + '.due{font-weight:700;color:#14110e}'
    + '.due--late{color:#b31b1b}'
    + '.due--done{color:#157a47}'
    // Buttons: pills. Solid for the main step, outlined for the rest.
    + '.btn{border-radius:999px}'
    + '.btn-primary{background:#b31b1b;box-shadow:none}'
    + '.btn-primary:hover{background:#8f1515;transform:none;box-shadow:none}'
    + '.btn-confirm{background:#157a47;box-shadow:none}'
    + '.btn-confirm:hover{background:#0f6a3c}'
    + '.btn-ghost{padding:9px 16px;border:1.5px solid #14110e;border-radius:999px;background:none;color:#14110e}'
    + '.btn-ghost:hover,.btn-ghost:active{border-color:#14110e;background:#14110e;color:#fff}'
    + '.btn-warn{padding:9px 16px;border:1.5px solid #b06a00;border-radius:999px;background:none;color:#8a4b00}'
    + '.btn-warn:hover{border-color:#b06a00;background:#b06a00;color:#fff}'
    + '.btn-danger{padding:9px 16px;border:1.5px solid #b31b1b;border-radius:999px;background:none;color:#b31b1b}'
    + '.btn-danger:hover{border-color:#b31b1b;background:#b31b1b;color:#fff}'
    + '.btn-skip{border-color:#fdeccb;background:#fdeccb;color:#8a5200}'
    + '.btn-skip:hover,.btn-skip:active{border-color:#b06a00;background:#b06a00;color:#fff}'
    + '.card-foot .btn{padding:9px 16px;border-radius:999px;font-size:13px}'
    + '.act-tools{border-left:1px solid ' + hair + '}'
    + '.act-panel{padding:14px;border:0;border-radius:16px;background:#f6efe2}'
    + '.act-panel--ask{background:#f6efe2}'
    + '.act-panel--danger{border:0;background:#fde4e4}'
    + '.act-ask{font-weight:700;color:#14110e}'
    // Notes and fields.
    + '.ic-note{padding:11px 14px;border:0;border-radius:14px;background:#e9f0fd;color:#14110e}'
    + '.ic-note b{color:#1a468f}'
    + '.ic-note--warn{background:#fdeccb;color:#6f3d00}'
    + '.ic-note--warn b{color:#6f3d00}'
    + '.ic-reply-lbl,.ic-edit-lbl{color:#b31b1b}'
    + '.ic-reply-help,.tp-opt,.ic-edit-msg{color:#3d3832}'
    + '.ic-reply-in,.ic-reason,.ic-edit-in{border:1.5px solid ' + edge + ';border-radius:14px;background:#fff}'
    + '.ic-reply-in:focus,.ic-reason:focus,.ic-edit-in:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.14)}'
    + '.ic-edit{padding:16px;border:0;border-radius:16px;background:#f6efe2}'
    + '.empty{margin-top:14px;padding:20px;border:0;border-radius:16px;background:#f1e9dc;font-weight:700;color:#14110e;text-align:left}'
    + '#ic-list>.empty:only-child{background:#e3f5ea;color:#0f6a3c}'
    + '@media(max-width:600px){'
    +   '.page-title{font-size:28px!important}'
    +   '.card-sum{padding:14px 2px 14px 18px}'
    +   'details.card[open]>.card-sum{padding:14px 12px 14px 26px}'
    +   '.card-body{padding:16px 14px 14px 26px}'
    +   '.card-foot{padding:12px 14px 14px 26px}'
    + '}'
    + '@media(prefers-reduced-motion:reduce){.card-sum-title{transition:none}}'
    + '</style>';
}

// Photos sit inside folded rows, so the browser does not fetch them until a row
// opens, and a Drive thumbnail then takes a moment to arrive. This starts the fetch
// early: for the first rows once the page has settled, and for any row as soon as the
// pointer, a finger or the keyboard reaches it. By the time the row opens, the photo
// is usually there.
function icFastImagesJs_() {
  return '<script>'
    + '(function(){'
    + 'function warm(c){if(!c||c.__warm)return;c.__warm=1;c.querySelectorAll("img[loading=lazy]").forEach(function(i){i.loading="eager";});}'
    + 'function near(e){var t=e.target;if(t&&t.closest)warm(t.closest("details.card"));}'
    + 'document.addEventListener("pointerover",near,{passive:true});'
    + 'document.addEventListener("touchstart",near,{passive:true});'
    + 'document.addEventListener("focusin",near);'
    + 'function first(){var n=0;document.querySelectorAll("#ic-list details.card").forEach(function(c){if(n<6&&c.style.display!=="none"){warm(c);n++;}});}'
    + 'if(document.readyState==="complete")setTimeout(first,300);else window.addEventListener("load",function(){setTimeout(first,300);});'
    + '})();'
    + '</script>';
}

// A "Refresh" button for the open-issues pages. Re-renders the list in place via
// icRefresh() (a full reload would blank the sandboxed iframe).
function icRefreshBtn_() {
  return '<span id="ic-refreshed" class="ic-refreshed"></span>'
    + '<button type="button" id="ic-refresh" class="ic-refresh" onclick="icRefresh()" title="Load the latest">'
    +   '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"></path><path d="M21 3v6h-6"></path></svg>'
    +   '<span>Refresh</span></button>';
}

// Filter chip at the top of a board. Tapping one shows that bucket of folded cards.
function sumBtn_(bucket, countId, n, label, extraCls) {
  const num = Number(n) || 0;
  const on = bucket === 'open';
  const cls = 'ic-sum' + (extraCls ? ' ' + extraCls : '') + (on ? ' is-on' : '');
  return '<button type="button" class="' + cls + '" data-bucket="' + bucket + '" onclick="icFilt(\'' + bucket + '\')"'
    + ' aria-pressed="' + (on ? 'true' : 'false') + '">'
    + '<b id="' + countId + '">' + num + '</b> ' + label + '</button>';
}

function cardSum_(title, sub, chipHtml, accentHtml, extraHtml, extraLabel) {
  return '<summary class="card-sum">'
    + (accentHtml || '')
    + '<span class="card-sum-main">'
    +   '<span class="card-sum-title">' + title + '</span>'
    +   (extraHtml ? '<span class="card-sum-action"><span class="card-sum-alabel">' + (extraLabel || 'Required action') + '</span>' + extraHtml + '</span>' : '')
    +   (sub ? '<span class="card-sum-sub">' + sub + '</span>' : '')
    + '</span>'
    + '<span class="card-sum-stat">' + (chipHtml || '') + '</span>'
    + '<span class="section-chev" aria-hidden="true"></span></summary>';
}

function icActionSum_(action, id) {
  if (!action) return '';
  var txt = escapeHtml_(phrase_(action));
  return id ? '<span id="' + id + '">' + txt + '</span>' : txt;
}

function icOpenAttr_(startOpen) {
  return startOpen ? ' open' : '';
}

function icChipBar_(openN, pendingN, doneN) {
  return '<div class="ic-summary">'
    + '<div class="ic-seg" role="group" aria-label="Filter list">'
    +   sumBtn_('open', 'sum-open', openN, 'open')
    +   sumBtn_('pending', 'sum-pending', pendingN, 'pending', 'ic-sum--pending')
    +   sumBtn_('done', 'sum-done', doneN, 'recently completed', 'ic-sum--done')
    + '</div>'
    + '<div class="ic-tools">'
    +   icUnlockBtn_()
    +   icRefreshBtn_()
    + '</div>'
    + '</div>';
}

function icUnlockBtn_() {
  return '<button type="button" class="ic-unlock" id="ic-unlock" onclick="icUnlockAll()" aria-pressed="false">Expand all</button>';
}

function icFiltEmpty_(msg) {
  return '<div id="empty" class="empty" style="display:none">' + (msg || 'Nothing in this view.') + '</div>';
}

// Client logic for the Refresh button, shared by both open-issues pages. Reads the
// page-level IC_SCOPE ("*" for all teams, else the team name) and IC_READONLY.
function icRefreshJs_() {
  return '<script>'
    + 'function icSetNum(id,n){if(n==null)return;var e=document.getElementById(id);if(e)e.textContent=n;}'
    + 'function icRefresh(){var b=document.getElementById("ic-refresh");if(b){b.disabled=true;b.classList.add("is-spin");}'
    + 'google.script.run.withSuccessHandler(function(r){if(b){b.disabled=false;b.classList.remove("is-spin");}if(!r||!r.ok)return;'
    + 'var el=document.getElementById("ic-list");if(el)el.innerHTML=r.html;secRestore();'
    + 'var sh=document.getElementById("pr-strip-host");if(sh&&r.strip!=null)sh.innerHTML=r.strip;if(typeof prAnnounceStrip==="function")prAnnounceStrip();'
    + 'icSetNum("sum-open",r.open);icSetNum("sum-pending",r.pending);icSetNum("sum-done",r.done);'
    + 'if(typeof ADMIN_PASS!=="undefined"&&ADMIN_PASS)document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});'
    + 'if(typeof flt==="function")flt();'
    + 'var t=document.getElementById("ic-refreshed");if(t){t.textContent="Updated just now";clearTimeout(window.__icrt);window.__icrt=setTimeout(function(){if(t)t.textContent="";},4000);}'
      + '}).withFailureHandler(function(){if(b){b.disabled=false;b.classList.remove("is-spin");}var t=document.getElementById("ic-refreshed");if(t){t.style.color="#b31b1b";t.textContent="Could not refresh";}}).icRefreshData(IC_SCOPE,IC_READONLY,(typeof IC_ADMIN!=="undefined")?!!IC_ADMIN:!!(typeof CARD_ALL!=="undefined"&&CARD_ALL));}'
    + '</script>';
}

// Client-callable: fresh list HTML + counts for a Refresh. scope "*" = all teams.
// Third arg: for "*" it is the admin flag (Close/Edit/Delete vs student Complete).
// For a team portal it is startOpen (keep Expand all after refresh). Never mix them —
// the client passes IC_ADMIN on the all-issues page and CARD_ALL on team portals.
function icRefreshData(scope, readOnly, startOpen) {
  const strip = (typeof prTasksStripHtml_ === 'function') ? prTasksStripHtml_() : '';
  if (String(scope) === '*') {
    const data = listAllIssues_();
    const pend = data.open.filter(function (x) { return x.pending; }).length;
    return { ok: true, html: icAllSectionsHtml_(data, !!startOpen), open: data.open.length - pend, pending: pend, done: data.resolved, strip: strip };
  }
  const d = listTeamIssues_(String(scope));
  const pend2 = d.open.filter(function (x) { return x.pending; }).length;
  return { ok: true, html: icTeamSectionsHtml_(d, !!readOnly, !!startOpen), open: d.open.length - pend2, pending: pend2, done: d.resolved, strip: strip };
}

function teamPortal_(team, readOnly, embedded) {
  if (!team) return htmlPage_('Invalid link', 'This team link is not recognized.');
  const data = listTeamIssues_(team);

  let inner = '<div style="font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#b31b1b">Project Teams Ops Hub</div>'
    + '<div class="swh" style="font-size:36px;font-weight:800;letter-spacing:-.045em;line-height:1;margin-top:10px;color:#14110e">' + escapeHtml_(team) + '</div>'
    + '<div style="width:64px;height:5px;border-radius:99px;background:linear-gradient(90deg,#b31b1b,#e08a1e,#f0c050);margin-top:12px"></div>';

  const pendN = (data.open || []).filter(function (x) { return x.pending; }).length;
  inner += icChipBar_((data.open || []).length - pendN, pendN, data.resolved || 0);

  inner += '<div id="pr-strip-host">' + prTasksStripHtml_() + '</div>';
  inner += '<div id="ic-list">' + icTeamSectionsHtml_(data, readOnly) + '</div>';

  const scripts = tpSharedJs_() + (readOnly ? '' : icClientJs_()) + icRefreshJs_()
    + '<script>var IC_SCOPE=' + JSON.stringify(team) + ';var IC_READONLY=' + (readOnly ? 'true' : 'false') + ';</script>';
  return swissShell_(tpStyles_() + prStripStyles_() + icSkin_() + inner + scripts + prStripJs_() + icFastImagesJs_(), 'Space Status - ' + team, true, embedded, null, '#fffdf9');
}

// Folded issue cards for a team portal. Open includes overdue; chips above the list
// filter which bucket is visible. Extracted so a Refresh re-renders identically.
function icTeamSectionsHtml_(data, readOnly, startOpen) {
  const issues = data.open || [];
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
    const title = escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue');
    const sub = isPending ? 'Awaiting approval' : (isSentBack ? 'Sent back' : (it.overdue ? 'Overdue' : (it.deadline ? 'Due ' + escapeHtml_(dl) : '')));
    const bucket = isPending ? 'pending' : 'open';
    const statusTxt = isPending ? 'Submitted, awaiting approval' : (isSentBack ? 'Sent back' : ('Due ' + escapeHtml_(dl)));
    const accent = COLOR_HEX[it.color] || '#d6d3ce';
    let foot = '';
    if (!readOnly) foot = isPending ? icPendingFoot_(rid, it.token) : icOpenFoot_(rid, it.token, it.photoOptional);
    return '<details class="card" id="' + rid + '" data-bucket="' + bucket + '" data-tok="' + it.token + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-over="' + (it.overdue ? '1' : '0') + '" data-po="' + (it.photoOptional ? '1' : '0') + '"' + icOpenAttr_(startOpen) + (bucket === 'open' ? '' : ' style="display:none"') + '>'
      + cardSum_(title, sub, '<span id="' + rid + '-pill">' + chip + '</span>',
          '<span class="card-accent" id="' + rid + '-accent" style="background:' + accent + '"></span>',
          icActionSum_(it.action))
      + '<div class="card-body">'
      +   (it.details ? '<div class="card-field"><span class="card-flabel">Details</span><div class="card-details">' + escapeHtml_(it.details) + '</div></div>' : '')
      +   (it.photos && it.photos.length ? photoStrip_(it.photos) : '')
      +   icPhotoBlock_(rid, it.completionPhoto)
      +   icNoteContainer_(rid, isPending, it.sentBackReason, !readOnly, it.photoOptional)
      +   icReplyBlock_(rid, it.completionNote, isPending, !readOnly)
      + '</div>'
      + '<div class="card-foot">'
      +   '<span id="' + rid + '-status" class="' + (isPending ? 'due' : dueCls) + '">' + statusTxt + '</span>'
      +   (readOnly ? '' : '<span id="' + rid + '-act" class="btn-row">' + foot + '</span>')
      + '</div>'
      + '</details>';
  };

  const pendingList = issues.filter(function (x) { return x.pending; });
  const activeList = issues.filter(function (x) { return !x.pending; });
  const out = activeList.map(renderCard).join('') + pendingList.map(renderCard).join('') + icDoneCards_(data.done, false, startOpen);
  if (!out) return '<div class="empty">No open issues. Everything is in good shape.</div>';
  return out + icFiltEmpty_('Nothing in this view.');
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
    +   '<select name="view" required style="flex:1 1 12rem;min-width:0;width:100%;font-size:15px;padding:12px 14px;border:1.5px solid #ddd;border-radius:8px;background:#fff;color:#111">' + opts + '</select>'
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
  const cTeam = ci(CONFIG.headers.team), cStatus = findHeader_(H, statusHeaderNames_()), cIssue = ci(CONFIG.headers.issueType),
        cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details), cTs = ci(CONFIG.headers.timestamp),
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader), cPhoto = ci(CONFIG.headers.photo),
        cCompletedAt = ci(CONFIG.completedAtHeader), cCompletionPhoto = ci(CONFIG.completionPhotoHeader),
        cCompletionNote = ci(CONFIG.completionNoteHeader), cSentBack = ci(CONFIG.sentBackHeader);
  const open = [], done = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTeam < 0 || norm_(v[i][cTeam]) !== norm_(teamName)) continue;
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) {
      resolved++;
      done.push({
        issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
        details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
        completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
        completionNote: cCompletionNote >= 0 ? String(v[i][cCompletionNote] || '').trim() : '',
        reportedPhoto: cPhoto >= 0 ? String(v[i][cPhoto] || '').trim() : '',
        addressed: new Date(v[i][cAddr])
      });
      continue;
    }
    const pending = cCompletedAt >= 0 && v[i][cCompletedAt] ? true : false;
    const color = parseColor_(statusFromRow_(v[i], H));
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
      photoOptional: icPhotoOptional_(cAction >= 0 ? v[i][cAction] : '', cIssue >= 0 ? v[i][cIssue] : ''),
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      photos: cPhoto >= 0 ? extractFileIds_(v[i][cPhoto]) : [],
      color: color, deadline: deadline, overdue: overdue,
      pending: pending, completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
      completionNote: cCompletionNote >= 0 ? String(v[i][cCompletionNote] || '').trim() : '',
      sentBackReason: (!pending && cSentBack >= 0) ? String(v[i][cSentBack] || '').trim() : '',
    });
  }
  open.sort(function (a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (a.deadline && b.deadline) return a.deadline - b.deadline;
    if (a.deadline) return -1; if (b.deadline) return 1; return 0;
  });
  done.sort(function (a, b) { return (b.addressed || 0) - (a.addressed || 0); });
  return { open: open, resolved: resolved, done: done.slice(0, 25) };
}

function listAllIssues_() {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const ci = function (n) { return H.indexOf(norm_(n)); };
  const cTeam = ci(CONFIG.headers.team), cStatus = findHeader_(H, statusHeaderNames_()), cIssue = ci(CONFIG.headers.issueType),
        cAction = ci(CONFIG.headers.action), cDetails = ci(CONFIG.headers.details), cTs = ci(CONFIG.headers.timestamp),
        cTok = ci(CONFIG.issueTokenHeader), cAddr = ci(CONFIG.addressedHeader), cPhoto = ci(CONFIG.headers.photo),
        cCompletedAt = ci(CONFIG.completedAtHeader), cCompletionPhoto = ci(CONFIG.completionPhotoHeader),
        cCompletionNote = ci(CONFIG.completionNoteHeader), cSentBack = ci(CONFIG.sentBackHeader);
  const open = [], done = [];
  let resolved = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  for (let i = 1; i < v.length; i++) {
    if (cTok < 0 || !v[i][cTok]) continue;
    if (cAddr >= 0 && v[i][cAddr]) {
      resolved++;
      done.push({
        team: cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
        issueType: cIssue >= 0 ? String(v[i][cIssue]).trim() : '',
        details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
        completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
        completionNote: cCompletionNote >= 0 ? String(v[i][cCompletionNote] || '').trim() : '',
        reportedPhoto: cPhoto >= 0 ? String(v[i][cPhoto] || '').trim() : '',
        addressed: new Date(v[i][cAddr])
      });
      continue;
    }
    const pending = cCompletedAt >= 0 && v[i][cCompletedAt] ? true : false;
    const color = parseColor_(statusFromRow_(v[i], H));
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
      photoOptional: icPhotoOptional_(cAction >= 0 ? v[i][cAction] : '', cIssue >= 0 ? v[i][cIssue] : ''),
      details: cDetails >= 0 ? String(v[i][cDetails]).trim() : '',
      photos: cPhoto >= 0 ? extractFileIds_(v[i][cPhoto]) : [],
      color: color, deadline: deadline, overdue: overdue,
      pending: pending, completionPhoto: cCompletionPhoto >= 0 ? String(v[i][cCompletionPhoto] || '').trim() : '',
      completionNote: cCompletionNote >= 0 ? String(v[i][cCompletionNote] || '').trim() : '',
      sentBackReason: (!pending && cSentBack >= 0) ? String(v[i][cSentBack] || '').trim() : '',
    });
  }
  open.sort(function (a, b) {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    if (a.deadline && b.deadline) return a.deadline - b.deadline;
    if (a.deadline) return -1; if (b.deadline) return 1; return 0;
  });
  done.sort(function (a, b) { return (b.addressed || 0) - (a.addressed || 0); });
  return { open: open, resolved: resolved, done: done.slice(0, 25) };
}

// Every open issue across all teams. Staff complete from the team portal (photo
// evidence). The admin hub (?admin=1) is Mark complete / Edit / Delete on open
// items; pending keeps Approve / Send back. No photo. Embedded when embed=1.
function allIssuesPage_(embedded, admin, back) {
  const data = listAllIssues_();
  const issues = data.open;
  const teamSet = {};
  issues.forEach(function (it) { if (it.team) teamSet[it.team] = 1; });
  const teamOpts = Object.keys(teamSet).sort().map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('');

  let inner = '';
  if (!embedded) {
    inner += hubBackLink_(back) + '<div class="page-head">'
      + '<div class="page-kicker">Project Teams Ops Hub</div>'
      + '<div class="page-title">Open space issues</div>'
      + '<div class="page-rule"></div>'
      + '</div>';
  }


  const pendN = issues.filter(function (x) { return x.pending; }).length;
  inner += icChipBar_(issues.length - pendN, pendN, data.resolved || 0);

  inner += '<div class="filters">'
    + '<div class="search-wrap"><input id="q" type="search" placeholder="Search team, issue, details…" oninput="flt()"></div>'
    + '<select id="team" onchange="flt()"><option value="">All teams</option>' + teamOpts + '</select>'
    + '<label class="toggle"><input id="odue" type="checkbox" onchange="this.closest(\'.toggle\').classList.toggle(\'is-on\', this.checked); flt()"> Overdue only</label>'
    + '</div>';

  inner += '<div id="pr-strip-host">' + prTasksStripHtml_() + '</div>';
  inner += '<div id="ic-list">' + icAllSectionsHtml_(data, !!admin) + '</div>';

  inner += '<script>var IC_SCOPE="*";var IC_READONLY=false;var IC_ADMIN=' + (admin ? 'true' : 'false') + ';function flt(){icApplyFilt();}</script>';
  return swissShell_(tpStyles_() + prStripStyles_() + icSkin_() + inner + tpSharedJs_() + icClientJs_() + icRefreshJs_() + tpAdminRevealJs_(admin) + prStripJs_() + icFastImagesJs_(), 'Open issues', true, embedded, null, '#fffdf9');
}

// Folded cards for the all-teams dashboard. When admin is true (hub ?admin=1),
// open cards are Mark complete (no photo, with confirm) / Edit / Delete;
// pending cards keep Approve / Send back plus Edit / Delete. Student photo
// Complete is off.
// Chips above the list pick the bucket; search / team / overdue-only further narrow.
function icAllSectionsHtml_(data, admin) {
  const startOpen = !!admin;
  const issues = data.open || [];
  const pendingList = issues.filter(function (x) { return x.pending; });
  const active = issues.filter(function (x) { return !x.pending; });
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
    const hay = ((it.team || '') + ' ' + (it.issueType || '') + ' ' + (it.action || '') + ' ' + (it.details || '') + ' ' + (it.completionNote || '')).toLowerCase();
    const accent = COLOR_HEX[it.color] || '#d6d3ce';
    const statusTxt = isPending ? 'Submitted, awaiting approval' : (isSentBack ? 'Sent back' : ('Due ' + escapeHtml_(dl)));
    const bucket = isPending ? 'pending' : 'open';
    const title = '<span id="' + rid + '-vtype">' + escapeHtml_(it.issueType ? phrase_(it.issueType) : 'Reported issue') + '</span>';
    const sub = '<span id="' + rid + '-vteam">' + escapeHtml_(it.team || 'Unassigned') + '</span>'
      + (isPending ? ' · Awaiting approval' : (isSentBack ? ' · Sent back' : (it.overdue ? ' · Overdue' : (it.deadline ? ' · Due ' + escapeHtml_(dl) : ''))));
    const foot = admin
      ? icAdminSimpleFoot_(rid, it.token, isPending)
      : (isPending ? icPendingFoot_(rid, it.token) : icOpenFoot_(rid, it.token, it.photoOptional));
    const extraAdmin = admin
      ? ''
      : '<span class="tp-admin act-tools" hidden><button type="button" class="btn btn-ghost" onclick="icEditOpen(\'' + rid + '\')">Edit</button>'
        + '<span id="' + rid + '-delwrap"><button type="button" class="btn btn-danger" onclick="icDelOpen(\'' + rid + '\',\'' + it.token + '\')">Delete</button></span></span>';
    return '<details class="card" id="' + rid + '" data-bucket="' + bucket + '" data-tok="' + it.token + '" data-team="' + escapeHtml_(it.team) + '" data-over="' + (it.overdue ? '1' : '0') + '" data-state="' + (isPending ? 'pending' : 'open') + '" data-po="' + (it.photoOptional ? '1' : '0') + '" data-hay="' + escapeHtml_(hay) + '"' + icOpenAttr_(startOpen) + (bucket === 'open' ? '' : ' style="display:none"') + '>'
      + cardSum_(title, sub, '<span id="' + rid + '-pill">' + chip + '</span>',
          '<span class="card-accent" id="' + rid + '-accent" style="background:' + accent + '"></span>',
          icActionSum_(it.action, rid + '-vaction'))
      + '<div class="card-body">'
      +   '<div id="' + rid + '-view">'
      +     '<div class="card-field" id="' + rid + '-vdetails-wrap"' + (it.details ? '' : ' hidden') + '><span class="card-flabel">Details</span><div class="card-details" id="' + rid + '-vdetails">' + escapeHtml_(it.details) + '</div></div>'
      +     (it.photos && it.photos.length ? photoStrip_(it.photos) : '')
      +     icPhotoBlock_(rid, it.completionPhoto)
      +     icNoteContainer_(rid, isPending, it.sentBackReason, admin ? !!it.sentBackReason : true, it.photoOptional)
      +     icReplyBlock_(rid, it.completionNote, isPending, !admin)
      +   '</div>'
      +   icEditForm_(rid, it, teams, fieldOpts.issueTypes)
      + '</div>'
      + '<div class="card-foot">'
      +   '<span id="' + rid + '-status" class="' + (isPending ? 'due' : dueCls) + '">' + statusTxt + '</span>'
      +   '<span id="' + rid + '-act" class="btn-row">' + foot + '</span>'
      +   extraAdmin
      + '</div>'
      + '</details>';
  };
  const out = active.map(renderCard).join('') + pendingList.map(renderCard).join('') + icDoneCards_(data.done, true, startOpen);
  if (!out) return '<div class="empty">No open issues. Everything is in good shape.</div>';
  return out + icFiltEmpty_('No issues match those filters.');
}


// ---- Reminder emails page (?view=emails&admin=1) ----
// What the script emailed, to whom, and what the next reminder run sends. Opened
// from the Reminder emails pop-up on the admin page. Upcoming rows come from
// reminderPlan_ (the list sendReminders uses). Sent rows come from the Email log
// tab, plus each row's Last reminder stamp for reminders older than the log.

const EMAIL_LOG_SHOWN = 120;   // newest log rows on the page; the tab keeps them all

function emDay_(d) {
  const tz = Session.getScriptTimeZone();
  const thisYear = Utilities.formatDate(d, tz, 'yyyy') === Utilities.formatDate(new Date(), tz, 'yyyy');
  return Utilities.formatDate(d, tz, thisYear ? 'MMM d' : 'MMM d, yyyy');
}
function emTime_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a'); }
function emStamp_(d) { return emDay_(d) + ', ' + emTime_(d); }

// When the Email log began: the date on its first row, or null with no log yet.
function emailLogStart_() {
  const sh = emailLogSheet_(false);
  if (!sh || sh.getLastRow() < 2) return null;
  const c = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_).indexOf(norm_('Sent at'));
  return c < 0 ? null : dateOrNull_(sh.getRange(2, c + 1).getValue());
}

function emailsModel_() {
  const plan = reminderPlan_();
  const contacts = teamContacts_();
  const rows = plan.open.concat(plan.closed);
  const byToken = {};
  rows.forEach(function (p) { byToken[p.token] = p; });

  // Open issues: going out on the next run, then scheduled by date, then not reminded.
  const upcoming = plan.open.map(function (p) {
    return { p: p, mail: p.next ? reminderMail_(p, contacts) : null };
  });
  const rank = function (u) { return u.p.send ? 0 : (u.p.next ? 1 : 2); };
  upcoming.sort(function (a, b) {
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 1 && a.p.next.getTime() !== b.p.next.getTime()) return a.p.next - b.p.next;
    const x = norm_(a.p.team), y = norm_(b.p.team);
    return x < y ? -1 : (x > y ? 1 : 0);
  });

  const sent = emailLogRead_(EMAIL_LOG_SHOWN);
  // Reminders from before the log started: the Last reminder stamp is all that is left.
  const logStart = emailLogStart_();
  const logged = {};
  sent.forEach(function (e) {
    if (e.type === 'Reminder' && e.token && !/^failed/i.test(e.result)) logged[e.token] = Math.max(logged[e.token] || 0, e.at.getTime());
  });
  rows.forEach(function (p) {
    if (!p.lastReminder) return;
    const t = p.lastReminder.getTime();
    if (logStart && t >= logStart.getTime()) return;        // the log has it
    if (logged[p.token] && logged[p.token] >= t - 120000) return;
    sent.push({
      at: p.lastReminder, type: 'Reminder', team: p.team, to: '', cc: '', subject: reminderSubject_(p),
      token: p.token, why: 'Sent before the email log started. Only the date was kept.', result: 'Sent', body: '', logged: false,
    });
  });
  sent.sort(function (a, b) { return b.at - a.at; });

  sent.forEach(function (e) {
    const row = e.token ? (byToken[e.token] || null) : null;
    e.row = row;
    e.failed = /^failed/i.test(e.result);
    // The check for "reminders after completing": a reminder dated later than the
    // completion that is on the row now.
    const doneAt = row ? (row.completedAt || row.addressed) : null;
    e.after = e.type === 'Reminder' && !e.failed && !!doneAt && e.at.getTime() > doneAt.getTime();
    e.doneAt = doneAt;
  });

  const weekAgo = Date.now() - 7 * 86400000;
  return {
    upcoming: upcoming, sent: sent,
    counts: {
      next: upcoming.filter(function (u) { return u.p.send; }).length,
      later: upcoming.filter(function (u) { return !u.p.send; }).length,
      sent: sent.length,
      flag: sent.filter(function (e) { return e.after || e.failed; }).length,
    },
    week: sent.filter(function (e) { return !e.failed && e.at.getTime() >= weekAgo; }).length,
    after: sent.filter(function (e) { return e.after; }).length,
  };
}

function emStatsHtml_(m) {
  const stat = function (n, label, cls) {
    return '<div class="em-stat' + (cls ? ' ' + cls : '') + '"><b>' + n + '</b><span>' + label + '</span></div>';
  };
  return stat(m.counts.next, m.counts.next === 1 ? 'reminder goes out on the next run' : 'reminders go out on the next run', 'em-stat--red')
    + stat(m.week, (m.week === 1 ? 'email' : 'emails') + ' sent in the last 7 days', 'em-stat--blue')
    + stat(m.after, (m.after === 1 ? 'reminder' : 'reminders') + ' sent after the team finished', m.after ? 'em-stat--bad' : 'em-stat--ok');
}

function emListHtml_(m) {
  const today = dayStart_(new Date());
  let out = '', i = 0, lastGroup = '';
  const n = m.counts.next;
  if (n > 1) {
    out += '<div class="em-bulk" data-tab="next">'
      + '<span class="em-bulk-row" id="em-bulk-a"><span>' + n + ' reminders go out on the next run.</span>'
      +   '<button type="button" class="em-btn em-btn--ghost" onclick="emBulk(true)">Cancel all ' + n + '</button></span>'
      + '<span class="em-bulk-row" id="em-bulk-b" hidden><span>Cancel all ' + n + ' on this run? Each job gets its next reminder on schedule.</span>'
      +   '<span class="em-acts"><button type="button" class="em-btn" onclick="emBulkDo(this)">Yes, cancel them</button>'
      +   '<button type="button" class="em-btn em-btn--ghost" onclick="emBulk(false)">Keep them</button></span></span>'
      + '</div>';
  }
  m.upcoming.forEach(function (u) {
    const group = u.p.send ? '' : (u.p.next ? 'Scheduled' : 'Not reminded');
    if (group !== lastGroup) {
      i = 0;
      if (group) out += '<div class="em-group' + (u.p.next ? ' em-group--blue' : ' em-group--amber') + '" data-tab="later" style="display:none">' + group + '</div>';
    }
    lastGroup = group;
    out += emUpcomingRow_(u, i++, today);
  });
  i = 0;
  m.sent.forEach(function (e) { out += emSentRow_(e, i++); });
  return out;
}

// One folded row. tab lists the tabs it shows under ("sent flag"). when/sub/title/line
// are plain text; tagTxt, bodyHtml and endHtml (a control beside the tag) are already
// HTML. id is optional.
function emRow_(tab, i, hay, when, sub, title, line, tagCls, tagTxt, bodyHtml, id, endHtml) {
  const shown = (' ' + tab + ' ').indexOf(' next ') >= 0;
  // The same team gets the same color on every row, so a team is easy to follow down the list.
  let hue = 0;
  const name = String(title || '');
  for (let k = 0; k < name.length; k++) hue = (hue * 31 + name.charCodeAt(k)) % 360;
  const dot = '<span class="em-dot" style="background:hsl(' + hue + ',68%,45%)" aria-hidden="true"></span>';
  return '<details class="em-row"' + (id ? ' id="' + id + '"' : '') + ' data-tab="' + tab + '" data-hay="' + escapeHtml_(String(hay).toLowerCase()) + '" style="--i:' + Math.min(i, 14) + (shown ? '' : ';display:none') + '">'
    + '<summary class="em-sum">'
    +   '<span class="em-when' + (tagCls === 'em-tag--rem' && tab === 'next' ? ' em-when--hot' : '') + '"><b>' + escapeHtml_(when) + '</b><i>' + escapeHtml_(sub) + '</i></span>'
    +   '<span class="em-main"><span class="em-team">' + dot + escapeHtml_(title) + '</span><span class="em-subj">' + escapeHtml_(line) + '</span></span>'
    +   '<span class="em-end"><span class="em-tag ' + tagCls + '">' + tagTxt + '</span>' + (endHtml || '') + '</span>'
    +   '<span class="em-chev" aria-hidden="true"></span>'
    + '</summary>'
    + '<div class="em-body">' + bodyHtml + '</div>'
    + '</details>';
}

// Label and value lines. Values are already HTML; empty ones are left out.
function emMeta_(pairs) {
  let s = '<dl class="em-meta">';
  pairs.forEach(function (kv) { if (kv[1]) s += '<dt>' + kv[0] + '</dt><dd>' + kv[1] + '</dd>'; });
  return s + '</dl>';
}

function emAddr_(list) { return escapeHtml_(String(list || '').split(',').join(', ')); }

// The email itself, for reading only: its button and links do nothing here, and an
// inline photo (cid:) shows as a note because the image is not on this page.
function emLetter_(html) {
  const safe = String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/<img[^>]+src="cid:[^>]*>/gi, '<span class="em-photo">Photo attached</span>')
    .replace(/<a\s+href="[^"]*"/gi, '<a');
  return safe ? '<div class="em-letter"><div class="em-cap">The email</div>' + safe + '</div>' : '';
}

function emReportLine_(p) {
  const bits = [];
  if (p.issueType) bits.push(phrase_(p.issueType));
  if (p.reported) bits.push('reported ' + emDay_(p.reported));
  if (p.deadline) bits.push('due ' + emDay_(p.deadline));
  return bits.join(' · ') + (p.details ? '. ' + p.details : '');
}

function emUpcomingRow_(u, i, today) {
  const p = u.p, mail = u.mail;
  let when, sub, tagCls = 'em-tag--ink', tagTxt;
  if (p.send) { when = 'Today'; sub = 'next run'; tagCls = 'em-tag--rem'; tagTxt = 'Reminder'; }
  else if (p.next) {
    const d = dayDiff_(p.next, today);
    when = emDay_(p.next); sub = d <= 1 ? 'tomorrow' : 'in ' + d + ' days'; tagCls = 'em-tag--blue'; tagTxt = 'Scheduled';
    if (p.hold === 'skip') { tagCls = 'em-tag--amber'; tagTxt = 'Rescheduled'; }
  }
  else if (p.state === 'pending') { when = 'Paused'; sub = 'in review'; tagCls = 'em-tag--amber'; tagTxt = 'Paused'; }
  else if (p.hold === 'off') { when = 'Off'; sub = 'cancelled'; tagCls = 'em-tag--amber'; tagTxt = 'Cancelled'; }
  else { when = 'None'; sub = 'not reminded'; tagTxt = 'No reminders'; }
  const line = mail ? mail.subject : ((p.issueType ? phrase_(p.issueType) : 'Reported issue') + (p.details ? '. ' + p.details : ''));
  const hay = [p.team, p.issueType, p.details, mail ? mail.to + ' ' + mail.cc + ' ' + mail.subject : ''].join(' ');
  let body = emActs_(p) + emMeta_([
    ['To', mail ? emAddr_(mail.to) : ''],
    ['Cc', mail ? emAddr_(mail.cc) : ''],
    ['Why', escapeHtml_(p.why)],
    ['Report', escapeHtml_(emReportLine_(p))],
  ]);
  p.hints.forEach(function (h) { body += '<p class="em-hint">' + escapeHtml_(h) + '</p>'; });
  if (mail) body += emLetter_(mail.body);
  // Cancel on the row opens it and puts the choices under the cursor, so nothing is
  // cancelled by one stray click.
  const quick = p.next ? '<button type="button" class="em-x" onclick="emAsk(event,\'' + p.token + '\')">Cancel</button>' : '';
  return emRow_(p.send ? 'next' : 'later', i, hay, when, sub, p.team || 'No team', line, tagCls, tagTxt, body, 'emr-' + p.token, quick);
}

// The cancel controls for one open job, by what is in its Reminder hold cell.
function emActs_(p) {
  const btn = function (mode, label, ghost) {
    return '<button type="button" class="em-btn' + (ghost ? ' em-btn--ghost' : '') + '" onclick="emAct(this,\'' + p.token + '\',\'' + mode + '\')">' + label + '</button>';
  };
  let inner = '';
  if (p.hold === 'off') {
    inner = '<span class="em-acts-note">No reminders go out for this job.</span>' + btn('resume', 'Turn reminders back on');
  } else if (p.hold === 'skip' && !p.send) {
    inner = '<span class="em-acts-note">The ' + escapeHtml_(emDay_(p.skipped)) + ' reminder is cancelled.</span>'
      + btn('resume', 'Undo') + btn('all', 'Cancel all for this job', true);
  } else if (p.next) {
    inner = btn('next', p.send ? 'Cancel this reminder' : 'Cancel the next reminder') + btn('all', 'Cancel all for this job', true);
  }
  return inner ? '<div class="em-acts">' + inner + '</div>' : '';
}

function emSentRow_(e, i) {
  const row = e.row;
  let tagCls = 'em-tag--ink', tagTxt = escapeHtml_(e.type || 'Email');
  if (e.type === 'Reminder') tagCls = 'em-tag--rem';
  else if (e.type === 'Sent back') tagCls = 'em-tag--amber';
  else if (e.type === 'Notice') tagCls = 'em-tag--violet';
  if (e.failed) { tagCls = 'em-tag--bad'; tagTxt = 'Failed'; }
  else if (e.after) { tagCls = 'em-tag--bad'; tagTxt = 'After completion'; }
  let now = '';
  if (e.token) {
    now = !row ? 'This report is no longer in the sheet.'
      : row.state === 'done' ? 'Completed' + (row.addressed ? ' ' + emStamp_(row.addressed) : '')
      : row.state === 'pending' ? 'Submitted' + (row.completedAt ? ' ' + emStamp_(row.completedAt) : '') + ', waiting for approval'
      : row.state === 'sentback' ? 'Open. Sent back to the team.'
      : 'Open';
  }
  let body = emMeta_([
    ['To', e.to ? emAddr_(e.to) : (e.logged ? '' : 'Not recorded')],
    ['Cc', emAddr_(e.cc)],
    ['Sent', escapeHtml_(emStamp_(e.at))],
    ['Why', escapeHtml_(e.why)],
    ['Now', escapeHtml_(now)],
  ]);
  if (e.failed) body += '<p class="em-hint em-hint--bad">' + escapeHtml_('This email did not send. ' + e.result.replace(/^failed:?\s*/i, '')) + '</p>';
  if (e.after) body += '<p class="em-hint em-hint--bad">' + escapeHtml_('This reminder went out after the team finished the report (' + emStamp_(e.doneAt) + ').') + '</p>';
  // Why an open report keeps getting reminders: sent back, or filed twice.
  if (row && e.type === 'Reminder' && !e.after) row.hints.forEach(function (h) { body += '<p class="em-hint">' + escapeHtml_(h) + '</p>'; });
  body += emLetter_(e.body);
  const hay = [e.team, e.subject, e.to, e.cc, e.type].join(' ');
  return emRow_('sent' + (e.failed || e.after ? ' flag' : ''), i, hay, emDay_(e.at), emTime_(e.at), e.team || 'No team', e.subject, tagCls, tagTxt, body);
}

// Flat, and led by color: rules and type, no cards. Red for what goes out, blue for
// what is scheduled, amber for what is held, green for what is fine. Text stays ink.
function emStyles_() {
  const disp = '"Plus Jakarta Sans",Helvetica,Arial,sans-serif';
  const hair = 'rgba(20,17,14,.14)';
  return '<style>'
    + '@media(min-width:601px){.swiss-inner{padding-left:26px;padding-right:26px}}'
    + '.em{font-family:Inter,Helvetica,Arial,sans-serif;color:#14110e}'
    + '.em-kicker{font-family:' + disp + ';font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#b31b1b}'
    + '.em-title{font-family:' + disp + ';font-size:38px;font-weight:800;letter-spacing:-.045em;line-height:1;margin:10px 0 22px;color:#14110e}'
    // Stats under a five-color bar.
    + '.em-stats{position:relative;display:flex;flex-wrap:wrap;padding-top:5px;border-bottom:1px solid ' + hair + '}'
    + '.em-stats::before{content:"";position:absolute;left:0;right:0;top:0;height:5px;border-radius:99px;background:linear-gradient(90deg,#b31b1b 0%,#e08a1e 30%,#f0c050 50%,#1f9d5b 72%,#2563c9 100%)}'
    + '.em-stat{flex:1 1 9rem;min-width:0;padding:16px 18px 16px 0;margin-right:18px;border-right:1px solid ' + hair + ';animation:emIn .5s cubic-bezier(.2,.8,.3,1) both}'
    + '.em-stat:nth-child(2){animation-delay:70ms}.em-stat:nth-child(3){animation-delay:140ms}'
    + '.em-stat:last-child{border-right:0;margin-right:0;padding-right:0}'
    + '.em-stat b{display:block;font-family:' + disp + ';font-size:46px;font-weight:800;letter-spacing:-.055em;line-height:1;font-variant-numeric:tabular-nums}'
    + '.em-stat span{display:block;margin-top:7px;font-size:12.5px;font-weight:600;line-height:1.35}'
    + '.em-stat--red b{color:#b31b1b}.em-stat--blue b{color:#2563c9}.em-stat--bad b{color:#b31b1b}.em-stat--ok b{color:#157a47}'
    // Tabs with count pills.
    + '.em-bar{display:flex;align-items:flex-end;gap:10px 18px;flex-wrap:wrap;margin-top:24px;border-bottom:1px solid ' + hair + '}'
    + '.em-tabs{display:flex;gap:22px;flex:1 1 auto;min-width:0;overflow-x:auto;scrollbar-width:none}'
    + '.em-tabs::-webkit-scrollbar{display:none}'
    + '.em-tab{appearance:none;-webkit-appearance:none;border:0;border-bottom:3px solid transparent;background:none;margin:0 0 -1px;padding:0 0 10px;font-family:' + disp + ';font-size:14px;font-weight:700;letter-spacing:-.01em;color:#14110e;cursor:pointer;white-space:nowrap}'
    + '.em-tab b{display:inline-block;min-width:22px;margin-left:7px;padding:2px 7px;border-radius:999px;background:#f1e9dc;font-size:11.5px;font-weight:800;line-height:1.4;text-align:center;font-variant-numeric:tabular-nums;color:#14110e}'
    + '.em-tab:hover{border-bottom-color:rgba(179,27,27,.35)}'
    + '.em-tab.is-on{color:#b31b1b;border-bottom-color:#b31b1b}'
    + '.em-tab.is-on b{background:#b31b1b;color:#fff}'
    + '.em-tab:focus-visible,.em-refresh:focus-visible,.em-sum:focus-visible{outline:2px solid #b31b1b;outline-offset:3px}'
    + '.em-q{flex:0 1 13rem;min-width:9rem;margin-bottom:8px;padding:6px 0;border:0;border-bottom:1.5px solid #14110e;border-radius:0;background:transparent;font:500 14px Inter,Helvetica,Arial,sans-serif;color:#14110e;outline:none}'
    + '.em-q::placeholder{color:#57534e}'
    + '.em-q:focus{border-bottom-color:#b31b1b}'
    + '.em-refresh{appearance:none;-webkit-appearance:none;border:0;background:none;margin:0 0 9px;padding:0;font-family:' + disp + ';font-size:13px;font-weight:700;color:#b31b1b;text-decoration:underline;text-underline-offset:3px;cursor:pointer}'
    + '.em-refresh:hover{color:#8f1515}.em-refresh:disabled{cursor:default;text-decoration:none}'
    + '.em-msg{display:inline-block;margin:14px 0 0;padding:8px 15px;border-radius:999px;background:#e3f5ea;font-family:' + disp + ';font-size:13.5px;font-weight:700;color:#0f6a3c}'
    + '.em-msg:empty{display:none}'
    + '.em-msg--bad{background:#fde4e4;color:#b31b1b}'
    + '.em-bulk{margin-top:14px;padding:12px 16px;border-radius:14px;background:#fdeeee;font-size:13.5px;font-weight:600;line-height:1.45}'
    + '.em-bulk-row{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px 16px}'
    + '.em-bulk-row[hidden]{display:none}'
    + '.em-bulk .em-acts{margin:0}'
    + '.em-group{margin-top:28px;padding-bottom:8px;border-bottom:3px solid #14110e;font-family:' + disp + ';font-size:11px;font-weight:800;letter-spacing:.16em;text-transform:uppercase}'
    + '.em-group--blue{border-bottom-color:#2563c9;color:#1a468f}'
    + '.em-group--amber{border-bottom-color:#e08a1e;color:#8a5200}'
    // Rows.
    + '.em-row{border-bottom:1px solid ' + hair + ';animation:emIn .42s cubic-bezier(.2,.8,.3,1) both;animation-delay:calc(var(--i,0)*32ms)}'
    + '.em-sum{display:grid;grid-template-columns:6rem minmax(0,1fr) auto 18px;align-items:center;gap:14px;padding:15px 2px;cursor:pointer;list-style:none;-webkit-tap-highlight-color:transparent}'
    + '.em-sum::-webkit-details-marker{display:none}.em-sum::marker{content:none}'
    + '.em-when b{display:block;font-family:' + disp + ';font-size:15px;font-weight:800;letter-spacing:-.02em;line-height:1.2}'
    + '.em-when i{display:block;margin-top:3px;font-style:normal;font-size:10.5px;font-weight:700;letter-spacing:.09em;text-transform:uppercase}'
    + '.em-when--hot b,.em-when--hot i{color:#b31b1b}'
    + '.em-main{min-width:0}'
    + '.em-team{display:block;font-family:' + disp + ';font-size:16px;font-weight:800;letter-spacing:-.02em;line-height:1.25;overflow-wrap:anywhere;transition:color .15s}'
    + '.em-dot{display:inline-block;width:10px;height:10px;margin-right:9px;border-radius:50%;vertical-align:1px}'
    + '.em-sum:hover .em-team{color:#b31b1b}'
    + '.em-subj{display:block;margin-top:3px;font-size:13.5px;line-height:1.4;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    + '.em-row[open] .em-subj{white-space:normal;overflow-wrap:anywhere}'
    + '.em-end{display:inline-flex;align-items:center;gap:12px}'
    // Tags: filled pills, one color per meaning.
    + '.em-tag{display:inline-flex;align-items:center;gap:6px;padding:5px 11px;border-radius:999px;background:#f1e9dc;font-family:' + disp + ';font-size:10.5px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;white-space:nowrap;color:#14110e}'
    + '.em-tag::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor}'
    + '.em-tag--rem{background:#b31b1b;color:#fff}'
    + '.em-tag--blue{background:#e2ecfc;color:#1a468f}'
    + '.em-tag--violet{background:#ece4fd;color:#5b27b0}'
    + '.em-tag--amber{background:#fdeccb;color:#8a5200}'
    + '.em-tag--bad{background:#14110e;color:#fff}.em-tag--bad::before{background:#ff5a5a}'
    + '.em-chev{justify-self:center;width:9px;height:9px;margin-top:-4px;border-right:2px solid #14110e;border-bottom:2px solid #14110e;transform:rotate(45deg);transition:transform .22s cubic-bezier(.2,.8,.3,1)}'
    + '.em-row[open] .em-chev{margin-top:4px;border-color:#b31b1b;transform:rotate(-135deg)}'
    + '.em-body{padding:0 2px 24px calc(6rem + 16px);animation:emIn .3s ease both}'
    // Buttons.
    + '.em-x{appearance:none;-webkit-appearance:none;border:1.5px solid #b31b1b;border-radius:999px;background:none;padding:5px 12px;font-family:' + disp + ';font-size:12px;font-weight:700;line-height:1.2;color:#b31b1b;cursor:pointer;transition:background .15s,color .15s}'
    + '.em-x:hover{background:#b31b1b;color:#fff}'
    + '.em-acts{display:flex;flex-wrap:wrap;align-items:center;gap:8px 10px;margin:0 0 16px}'
    + '.em-acts-note{font-size:13.5px;font-weight:700;margin-right:4px;color:#8a5200}'
    + '.em-btn{appearance:none;-webkit-appearance:none;border:1.5px solid #b31b1b;border-radius:999px;background:#b31b1b;padding:8px 16px;font-family:' + disp + ';font-size:13px;font-weight:700;line-height:1.2;color:#fff;cursor:pointer;transition:background .15s,color .15s,border-color .15s}'
    + '.em-btn:hover{background:#8f1515;border-color:#8f1515}'
    + '.em-btn--ghost{border-color:#14110e;background:none;color:#14110e}'
    + '.em-btn--ghost:hover{background:#14110e;border-color:#14110e;color:#fff}'
    + '.em-btn:disabled,.em-x:disabled{opacity:.45;cursor:default}'
    + '.em-btn:focus-visible,.em-x:focus-visible{outline:2px solid #b31b1b;outline-offset:3px}'
    + '.em-meta{display:grid;grid-template-columns:auto minmax(0,1fr);gap:5px 16px;margin:0;font-size:13.5px;line-height:1.5}'
    + '.em-meta dt{padding-top:3px;font-family:' + disp + ';font-size:10.5px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#b31b1b}'
    + '.em-meta dd{margin:0;overflow-wrap:anywhere}'
    + '.em-hint{margin:14px 0 0;padding:9px 13px;border-left:4px solid #e08a1e;border-radius:0 10px 10px 0;background:#fdf1d9;font-size:13.5px;font-weight:600;line-height:1.5}'
    + '.em-hint--bad{border-left-color:#b31b1b;background:#fde4e4;color:#8f1515}'
    + '.em-letter{margin-top:18px;padding:2px 0 2px 18px;border-left:3px solid #b31b1b;overflow-wrap:anywhere}'
    + '.em-letter a{pointer-events:none;cursor:default}'
    + '.em-letter img{max-width:100%;height:auto}'
    + '.em-cap{margin-bottom:6px;font-family:' + disp + ';font-size:10.5px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#b31b1b}'
    + '.em-photo{display:inline-block;font:italic 13px Georgia,serif}'
    + '.em-empty{margin-top:16px;padding:18px 20px;border-radius:14px;background:#e3f5ea;font-family:' + disp + ';font-size:15px;font-weight:700;letter-spacing:-.01em;color:#0f6a3c}'
    + '.em-foot{margin:22px 0 0;font-size:12.5px;line-height:1.55;max-width:44rem}'
    + '@keyframes emIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}'
    + '@media(prefers-reduced-motion:reduce){.em-row,.em-stat,.em-body{animation:none}.em-chev,.em-team,.em-btn,.em-x{transition:none}}'
    + '@media(max-width:600px){'
    +   '.em-title{font-size:30px}.em-stat{flex-basis:5.5rem;padding:13px 10px 13px 0;margin-right:10px}.em-stat b{font-size:34px}'
    +   '.em-sum{grid-template-columns:4.6rem minmax(0,1fr) 16px;gap:4px 10px}'
    +   '.em-end{grid-column:2;grid-row:2;justify-self:start}.em-chev{grid-column:3;grid-row:1}'
    +   '.em-body{padding-left:2px}.em-q{flex:1 1 100%}'
    + '}'
    + '</style>';
}

function emJs_() {
  return '<script>'
    + 'var EM_TAB="next";'
    + 'var EM_EMPTY={next:"Nothing goes out on the next run.",later:"No open issue is waiting on a reminder.",sent:"No emails on record yet.",flag:"Nothing flagged. No reminder went to a finished report, and no email failed."};'
    + 'function emTab(t){EM_TAB=t;emApply();}'
    + 'function emApply(){var qe=document.getElementById("em-q");var q=qe?(qe.value||"").toLowerCase().trim():"";var n=0;'
    + 'document.querySelectorAll(".em-tab").forEach(function(b){var on=b.getAttribute("data-tab")===EM_TAB;b.classList.toggle("is-on",on);b.setAttribute("aria-selected",on?"true":"false");});'
    + 'document.querySelectorAll("#em-list [data-tab]").forEach(function(r){var ok=(" "+r.getAttribute("data-tab")+" ").indexOf(" "+EM_TAB+" ")>=0;'
    + 'if(ok&&q&&(r.getAttribute("data-hay")||"").indexOf(q)<0)ok=false;r.style.display=ok?"":"none";if(ok&&r.classList.contains("em-row"))n++;});'
    + 'var e=document.getElementById("em-empty");if(e){e.textContent=q?"Nothing matches that filter.":(EM_EMPTY[EM_TAB]||"");e.style.display=n?"none":"block";}}'
    + 'function emPaint(r){document.getElementById("em-stats").innerHTML=r.stats;document.getElementById("em-list").innerHTML=r.list;'
    + 'for(var k in r.counts){var c=document.getElementById("em-n-"+k);if(c)c.textContent=r.counts[k];}emApply();}'
    + 'function emMsg(t,bad){var m=document.getElementById("em-msg");if(!m)return;m.textContent=t||"";m.className="em-msg"+(bad?" em-msg--bad":"");clearTimeout(window.__emt);if(t&&!bad)window.__emt=setTimeout(function(){m.textContent="";},8000);}'
    + 'function emRefresh(){var b=document.getElementById("em-refresh");if(b){b.disabled=true;b.textContent="Refreshing\\u2026";}'
    + 'google.script.run.withSuccessHandler(function(r){if(b){b.disabled=false;b.textContent="Refresh";}if(!r||!r.ok)return;emPaint(r);'
    + '}).withFailureHandler(function(){if(b){b.disabled=false;b.textContent="Could not refresh. Retry";}}).emRefreshData();}'
    // Cancel on a row: open it and move to the choices. The click must not fold the row.
    + 'function emAsk(ev,tok){if(ev){ev.preventDefault();ev.stopPropagation();}var r=document.getElementById("emr-"+tok);if(!r)return;r.open=true;var b=r.querySelector(".em-acts button");if(b)b.focus();}'
    + 'function emLock(el,on){var bar=el.closest(".em-acts")||el.parentNode;bar.querySelectorAll("button").forEach(function(b){b.disabled=on;});}'
    + 'function emSaved(el){return function(r){if(!r||!r.ok){emLock(el,false);emMsg((r&&r.error)||"Could not save.",true);return;}emPaint(r);emMsg(r.msg||"Saved.");};}'
    + 'function emFailed(el){return function(){emLock(el,false);emMsg("Could not save. Try again.",true);};}'
    + 'function emAct(el,tok,mode){emLock(el,true);emMsg("Saving\\u2026");google.script.run.withSuccessHandler(emSaved(el)).withFailureHandler(emFailed(el)).emCancelReminder(tok,mode);}'
    + 'function emBulk(ask){var a=document.getElementById("em-bulk-a"),b=document.getElementById("em-bulk-b");if(a)a.hidden=!!ask;if(b)b.hidden=!ask;}'
    + 'function emBulkDo(el){emLock(el,true);emMsg("Saving\\u2026");google.script.run.withSuccessHandler(emSaved(el)).withFailureHandler(emFailed(el)).emCancelNextRun();}'
    + 'emApply();'
    + '</script>';
}

function emailsPage_(embedded, back) {
  const m = emailsModel_();
  const tab = function (key, label, n) {
    const on = key === 'next';
    return '<button type="button" class="em-tab' + (on ? ' is-on' : '') + '" role="tab" data-tab="' + key + '" aria-selected="' + (on ? 'true' : 'false') + '" onclick="emTab(\'' + key + '\')">'
      + label + '<b id="em-n-' + key + '">' + n + '</b></button>';
  };
  let inner = '<div class="em">';
  if (!embedded) {
    inner += hubBackLink_(back)
      + '<div class="em-kicker">Space Status</div>'
      + '<h1 class="em-title">Reminder emails</h1>';
  }
  inner += '<div id="em-stats" class="em-stats">' + emStatsHtml_(m) + '</div>'
    + '<div class="em-bar">'
    +   '<div class="em-tabs" role="tablist" aria-label="Which emails">'
    +     tab('next', 'Next run', m.counts.next) + tab('later', 'Later', m.counts.later)
    +     tab('sent', 'Sent', m.counts.sent) + tab('flag', 'Flagged', m.counts.flag)
    +   '</div>'
    +   '<input id="em-q" class="em-q" type="search" placeholder="Filter by team or address" aria-label="Filter by team or address" oninput="emApply()">'
    +   '<button type="button" id="em-refresh" class="em-refresh" onclick="emRefresh()">Refresh</button>'
    + '</div>'
    + '<div id="em-msg" class="em-msg" role="status" aria-live="polite"></div>'
    + '<div id="em-list">' + emListHtml_(m) + '</div>'
    + '<div id="em-empty" class="em-empty" style="display:none"></div>'
    + '<p class="em-foot">Next run and Later are worked out from the sheet each time this opens. Sent is the newest ' + EMAIL_LOG_SHOWN
    +   ' rows of the <b>Email log</b> tab in the Space Status sheet. A reminder from before the log started shows its date and no recipients.</p>'
    + '</div>';
  return swissShell_(emStyles_() + inner + emJs_(), 'Reminder emails', true, embedded, null, '#fffdf9');
}

// Fresh numbers and rows, with a line to show the admin what just happened.
function emFresh_(msg) {
  const m = emailsModel_();
  return { ok: true, msg: msg || '', stats: emStatsHtml_(m), list: emListHtml_(m), counts: m.counts };
}

// Client-callable: the Refresh button.
function emRefreshData() { return emFresh_(''); }

// ---- Cancelling reminders ----
// The Reminder hold column on Form Responses holds the admin's choice:
//   blank   reminders run as usual
//   Off     no reminders for this job
//   a date  no reminder before that date (one reminder cancelled)
// reminderPlan_ reads it; sendReminders checks it again just before each send.

function setReminderHold_(p, value) {
  const sh = ss_().getSheetByName(CONFIG.responsesSheet);
  sh.getRange(p.rowNum, ensureColumn_(sh, CONFIG.reminderHoldHeader)).setValue(value);
}

// Client-callable. mode: 'next' cancels the coming reminder only, 'all' cancels every
// reminder for the job, 'resume' undoes either.
function emCancelReminder(token, mode) {
  const p = reminderPlan_().open.filter(function (x) { return x.token === String(token); })[0];
  if (!p) return { ok: false, error: 'That job is no longer open.' };
  const who = p.team || 'this job';
  let value = '', msg = '';
  if (mode === 'all') {
    value = 'Off';
    msg = 'Reminders cancelled for ' + who + '.';
  } else if (mode === 'resume') {
    msg = 'Reminders are back on for ' + who + '.';
  } else if (mode === 'next') {
    if (!p.next) return { ok: false, error: 'No reminder is scheduled for this job.' };
    const today = dayStart_(new Date());
    value = addDays_(p.next > today ? p.next : today, p.interval);   // the following reminder keeps its place
    msg = 'Cancelled. The next reminder for ' + who + ' goes out ' + emDay_(value) + '.';
  } else {
    return { ok: false, error: 'Unknown action.' };
  }
  setReminderHold_(p, value);
  return emFresh_(msg);
}

// Client-callable: cancel every reminder on the next run, once each.
function emCancelNextRun() {
  const today = dayStart_(new Date());
  const due = reminderPlan_().open.filter(function (p) { return p.send; });
  due.forEach(function (p) { setReminderHold_(p, addDays_(today, p.interval)); });
  return emFresh_(due.length
    ? 'Cancelled ' + due.length + ' reminder' + (due.length === 1 ? '' : 's') + '. Each job gets its next one on schedule.'
    : 'Nothing was due.');
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
      + '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w480" alt="Reported photo" loading="lazy" decoding="async" style="max-width:100%;max-height:240px;border-radius:10px;border:1px solid #ececec">'
      + '</a>';
  }
  let s = '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px">';
  ids.forEach(function (id) {
    s += '<a href="https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view" target="_blank" rel="noopener" style="display:block;line-height:0">'
      + '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w240" alt="Reported photo" loading="lazy" decoding="async" style="height:96px;width:auto;border-radius:8px;border:1px solid #ececec">'
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
  const cNote = H.indexOf(norm_(CONFIG.completionNoteHeader));
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
        completionNote: cNote >= 0 ? String(v[i][cNote] || '').trim() : '',
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
// The return link at the top of a full page opened from the hub. back= is the hub page
// that opened it (the hub adds it to its links); the label follows that page, so it reads
// "Back to admin" from the admin page. Same look as .back-to-admin in the hub's styles.css.
function hubBackLink_(back) {
  const url = String(back || '').trim();
  if (!/^https?:\/\/[^\s"'<>]+$/i.test(url)) return '';
  const label = /admin\.html/i.test(url) ? 'Back to admin' : /developer\.html/i.test(url) ? 'Back to the developer page' : 'Back to the hub';
  return '<a class="hub-back" href="' + escapeHtml_(url) + '" target="_top" '
    + 'style="display:inline-flex;align-items:center;gap:6px;margin:0 0 18px;font-family:Inter,Helvetica,Arial,sans-serif;font-size:14px;font-weight:700;color:#b31b1b;text-decoration:none">'
    + '<span aria-hidden="true">&#8592;</span> ' + label + '</a>';
}

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