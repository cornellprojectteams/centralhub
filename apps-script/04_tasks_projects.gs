/**
 * Completion evidence + Projects, served by the same Space Status web app.
 *
 * TASKS ARE THE OPEN ISSUES. There is no separate task list. Completing any open
 * action item (in the all-issues dashboard, a team portal, or the "mark complete"
 * button in the notification email) now works like this:
 *   1. The doer uploads a photo as EVIDENCE of completion (required).
 *   2. Submitting fires a client-side confetti burst and moves the item to
 *      "Pending approval" (stamps `Completed at`; reminders pause).
 *   3. An admin (passcode) Approves -> stamps `Addressed at` (Completed), or
 *      Sends back -> clears the submission (back to Open / Uncompleted).
 * This reuses the one Form Responses tracking sheet, with no parallel system.
 *
 *   ?module=projects  Multi-user projects created/assigned by an admin. Any
 *                     student can pick (join) a project; the first pick flips it
 *                     Assigned -> In Progress. Completion is a manual trigger that
 *                     ENFORCES both a "before" and an "after" photo.
 *
 * Roles (no login): everyone is a doer by default. Admin-only actions (approve or
 * send back completions, create/assign projects) are gated by the shared passcode,
 * validated server-side on every mutating call against CONFIG.toolPassHash
 * (SHA-256; default "bigred"). Project completion is open to assignees.
 *
 * Reuses engine helpers in 02_notify_on_submit.gs (swissShell_, portalStyles_,
 * escapeHtml_, phrase_, fmtShort_, newToken_, ensureColumn_, norm_, ss_,
 * extractFileIds_, findIssue_). Apps Script shares one global scope, so the doGet
 * router in 02 delegates ?module=projects here, the issue pages in 02 call the
 * ic*/tp* UI helpers below, and the client pages call the server functions via
 * google.script.run.
 *
 * Any change to web-app logic needs a NEW deployment version. Same URL after.
 */

var TP = {
  projectsSheet: 'Projects',
  projectHeaders: ['Project ID', 'Title', 'Description', 'Status', 'Assignees', 'Before photo', 'After photo', 'Hours', 'Started at', 'Completed at', 'Created at', 'Sent back reason', 'Attachment', 'Link'],
  projectStatus: { assigned: 'Assigned', active: 'In Progress', pending: 'Pending', done: 'Completed' },
  uploadsFolderName: 'Ops Hub completion and project uploads',
};

// ---- One-time setup (mirrors setupRegistry). Safe to re-run. ----

function setupTasksProjects() {
  var ss = ss_();

  // Projects tab.
  tpEnsureTab_(ss, TP.projectsSheet, TP.projectHeaders,
    ['', 'Example: repaint the mezzanine railing', 'Sand, prime, and repaint the north railing safety yellow.', TP.projectStatus.assigned, '', '', '', '', '', new Date()]);

  // Completion-evidence columns on the existing Form Responses sheet (added on
  // demand anyway, but created here so they are visible in the sheet from day one).
  var resp = ss.getSheetByName(CONFIG.responsesSheet);
  if (resp) {
    ensureColumn_(resp, CONFIG.completionPhotoHeader);
    ensureColumn_(resp, CONFIG.completedAtHeader);
    ensureColumn_(resp, CONFIG.sentBackHeader);
    Logger.log('Ensured completion columns on "' + CONFIG.responsesSheet + '".');
  } else {
    Logger.log('WARNING: no "' + CONFIG.responsesSheet + '" tab found for completion columns.');
  }
  Logger.log('Projects tab ready; completing an open issue now asks for an evidence photo.');
}

function tpEnsureTab_(ss, name, headers, exampleRow) {
  var sh = ss.getSheetByName(name);
  var created = false;
  if (!sh) { sh = ss.insertSheet(name); created = true; }
  if (!String(sh.getRange(1, 1).getValue()).trim()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (exampleRow) sh.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
  }
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#8f1515').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  Logger.log((created ? 'Created' : 'Found') + ' "' + name + '" tab.');
}

// ---- Shared helpers ----

function tpSheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing "' + name + '" tab. Run setupTasksProjects() once.');
  return sh;
}

// Open a sheet and return {sh, col:{Header->1-based index}}. Reads the header row
// ONCE and only appends genuinely-missing headers, so it stays self-healing without
// a Sheets round-trip per column (that per-column check was the slow part of a load).
function tpOpen_(name, headers) {
  var sh = tpSheet_(name);
  var lastCol = sh.getLastColumn();
  var existing = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_) : [];
  var col = {}, nextCol = lastCol;
  headers.forEach(function (h) {
    var i = existing.indexOf(norm_(h));
    if (i >= 0) { col[h] = i + 1; }
    else { nextCol++; sh.getRange(1, nextCol).setValue(h); existing.push(norm_(h)); col[h] = nextCol; }
  });
  return { sh: sh, col: col };
}

function tpFindRow_(sh, idCol, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function tpSplitList_(v) {
  return String(v == null ? '' : v)
    .split(/[,\n;]+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length; });
}

function tpSha256Hex_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s == null ? '' : s), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function tpIsAdmin_(pass) {
  return !!pass && !!CONFIG.toolPassHash && tpSha256Hex_(pass) === CONFIG.toolPassHash;
}

// Client-callable: lets the unlock bar confirm a passcode before revealing controls.
function tpCheckPass(pass) { return tpIsAdmin_(pass); }

// Save a base64 data URL to Drive, return the file id. Sharing is inherited from the
// uploads folder (set once in tpUploadsFolder_), so there is no per-upload Drive
// permission round-trip here.
function tpSaveUpload_(dataUrl, filename) {
  var m = /^data:([^;]+);base64,(.*)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('That upload was not a valid image.');
  var bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, m[1], filename || ('upload_' + Date.now()));
  return tpUploadsFolder_().createFile(blob).getId();
}

// Find-or-create the uploads folder and link-share it ONCE (tracked by a script
// property), so every file dropped in it is viewable by link without a per-file call.
function tpUploadsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TP_UPLOADS_FOLDER');
  var folder = null;
  if (id) { try { folder = DriveApp.getFolderById(id); } catch (err) { /* recreate below */ } }
  if (!folder) {
    var it = DriveApp.getFoldersByName(TP.uploadsFolderName);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(TP.uploadsFolderName);
    props.setProperty('TP_UPLOADS_FOLDER', folder.getId());
    props.deleteProperty('TP_UPLOADS_SHARED');   // re-share if the folder changed
  }
  if (props.getProperty('TP_UPLOADS_SHARED') !== '1') {
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); props.setProperty('TP_UPLOADS_SHARED', '1'); }
    catch (err) { Logger.log('Folder share failed: ' + err); }
  }
  return folder;
}

function tpViewUrl_(id) { return 'https://drive.google.com/file/d/' + id + '/view'; }

function tpHoursLabel_(h) { var n = Number(h) || 0; return (n % 1 === 0 ? n : n.toFixed(1)) + (n === 1 ? ' hour' : ' hours'); }

function tpThumb_(url, maxH) {
  var ids = extractFileIds_(url);
  if (!ids.length) return '';
  var e = encodeURIComponent(ids[0]);
  return '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" style="display:inline-block;line-height:0">'
    + '<img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w600" loading="lazy" alt="Uploaded photo" style="max-width:100%;max-height:' + (maxH || 200) + 'px;border-radius:10px;border:1px solid #ececec">'
    + '</a>';
}

// One .tp-photo figure per file id in the cell, so a task with several completion
// photos shows them all. label captions each ("Completion photo", or numbered when
// there is more than one).
function tpPhotoCells_(url, maxH, label) {
  var ids = extractFileIds_(url);
  var out = '';
  for (var i = 0; i < ids.length; i++) {
    var e = encodeURIComponent(ids[i]);
    var cap = label ? (ids.length > 1 ? label + ' ' + (i + 1) : label) : '';
    out += '<div class="tp-photo"><figure>'
      + (cap ? '<figcaption>' + escapeHtml_(cap) + '</figcaption>' : '')
      + '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" style="display:inline-block;line-height:0">'
      + '<img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w600" loading="lazy" alt="Uploaded photo" style="max-width:100%;max-height:' + (maxH || 200) + 'px;border-radius:10px;border:1px solid #ececec">'
      + '</a></figure></div>';
  }
  return out;
}

// ---- Completion evidence: server mutations (operate on Form Responses) ----

// Locate an issue row by token; returns {sh, row, addressed, completedAt, col(name)} or null.
function icLocate_(token) {
  var sh = ss_().getSheetByName(CONFIG.responsesSheet);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  var H = v[0].map(norm_);
  var cTok = H.indexOf(norm_(CONFIG.issueTokenHeader));
  if (cTok < 0 || !token) return null;
  var cAddr = H.indexOf(norm_(CONFIG.addressedHeader));
  var cComp = H.indexOf(norm_(CONFIG.completedAtHeader));
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][cTok]) === String(token)) {
      return {
        sh: sh, row: i + 1,
        addressed: cAddr >= 0 ? v[i][cAddr] : '',
        completedAt: cComp >= 0 ? v[i][cComp] : '',
        col: function (name) { return ensureColumn_(sh, name); },
      };
    }
  }
  return null;
}

// Which Required-action types can be finished without a photo (they usually have no
// visible before/after): informational notices, facilities work orders, unsafe-
// practice corrections, and "Other". Everything else needs evidence.
function icPhotoOptional_(action) {
  var a = norm_(action);
  if (!a) return false;
  return a === 'other'
    || a.indexOf('no action needed') >= 0 || a.indexOf('informational') >= 0
    || a.indexOf('facilities work order') >= 0 || a.indexOf('unsafe work practice') >= 0;
}

// True for the Operations team. Their notification email goes to Noah (an admin), so
// the "Mark complete" button on that email may close the issue with no photo. Every
// other team's email is student staff, who always submit a photo.
function icIsOpsTeam_(team) {
  return norm_(team).indexOf('operations') >= 0;
}

// Admin bypass: mark an issue complete with no photo and no approval step, stamping
// Completed at + Addressed at (fully resolved). This is an admin-side action - the
// button is only shown on the admin dashboard (revealed by ?admin=1, like approve /
// send back). Student pages never surface it, so staff always submit a photo.
function resolveIssueComplete(token) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  if (loc.addressed) return { ok: false, error: 'This was already completed.' };
  var now = new Date();
  loc.sh.getRange(loc.row, loc.col(CONFIG.completedAtHeader)).setValue(now);
  loc.sh.getRange(loc.row, loc.col(CONFIG.addressedHeader)).setValue(now);
  loc.sh.getRange(loc.row, loc.col(CONFIG.sentBackHeader)).setValue('');
  return { ok: true, status: 'Completed' };
}

// Doer submits completion evidence -> Pending approval. A photo is required unless
// the action type is photo-optional.
// Doers can attach more than one completion photo. The client uploads each selected
// image in turn: the first call (append falsy) replaces the cell so a re-submission
// drops any stale photo from a prior sent-back attempt; later calls (append true)
// tack onto the comma-separated list.
function submitIssueCompletion(token, dataUrl, filename, append) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  if (loc.addressed) return { ok: false, error: 'This was already approved as complete.' };
  var photoCell = loc.sh.getRange(loc.row, loc.col(CONFIG.completionPhotoHeader));
  var prior = append ? String(photoCell.getValue() || '').trim() : '';
  if (!dataUrl && !prior && !icPhotoOptional_(loc.sh.getRange(loc.row, loc.col(CONFIG.headers.action)).getValue())) {
    return { ok: false, error: 'A photo is required to complete this task.' };
  }
  var id = '';
  if (dataUrl) {
    try { id = tpSaveUpload_(dataUrl, filename); }
    catch (err) { return { ok: false, error: String(err.message || err) }; }
  }
  if (id) {
    var url = tpViewUrl_(id);
    photoCell.setValue(prior ? prior + ', ' + url : url);
  } else if (!append) {
    photoCell.setValue('');
  }
  loc.sh.getRange(loc.row, loc.col(CONFIG.completedAtHeader)).setValue(new Date());
  loc.sh.getRange(loc.row, loc.col(CONFIG.sentBackHeader)).setValue('');   // clear any prior send-back reason
  return { ok: true, photoId: id, status: 'Pending' };
}

// Admin edits an issue's fields in place. Team drives reminder routing, so the
// dashboard offers it as a picklist. Required action is fixed at submission.
function updateIssueFields(token, team, issueType, details) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  var t = String(team || '').trim(), iss = String(issueType || '').trim(), det = String(details || '').trim();
  loc.sh.getRange(loc.row, loc.col(CONFIG.headers.team)).setValue(t);
  loc.sh.getRange(loc.row, loc.col(CONFIG.headers.issueType)).setValue(iss);
  loc.sh.getRange(loc.row, loc.col(CONFIG.headers.details)).setValue(det);
  return { ok: true, team: t, issueType: phrase_(iss), details: det };
}

// Admin deletes an issue -> the row is removed for good.
function deleteIssue(token) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  loc.sh.deleteRow(loc.row);
  return { ok: true };
}

// Admin approves -> Completed (stamps Addressed at, which is what "resolved" reads).
function approveIssueCompletion(token, pass) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  loc.sh.getRange(loc.row, loc.col(CONFIG.addressedHeader)).setValue(new Date());
  return { ok: true, status: 'Completed' };
}

// Admin sends back -> Uncompleted, with an optional reason. The doer's photo is KEPT
// (not cleared) so they can see what they submitted and what to fix.
function rejectIssueCompletion(token, reason, pass) {
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  loc.sh.getRange(loc.row, loc.col(CONFIG.completedAtHeader)).setValue('');   // no longer pending
  loc.sh.getRange(loc.row, loc.col(CONFIG.addressedHeader)).setValue('');
  loc.sh.getRange(loc.row, loc.col(CONFIG.sentBackHeader)).setValue(String(reason || '').trim());
  // NOTE: Completion photo is intentionally left in place.
  return { ok: true, status: 'Open' };
}

// ---- Projects: data ----

function tpListProjects_() {
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var sh = o.sh, col = o.col;
  var last = sh.getLastRow();
  var out = [];
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var pending = [];
  for (var i = 0; i < v.length; i++) {
    var row = v[i];
    var title = String(row[col['Title'] - 1] || '').trim();
    if (!title) continue;
    var id = String(row[col['Project ID'] - 1] || '').trim();
    if (!id) { id = newToken_(); pending.push({ r: i + 2, c: col['Project ID'], val: id }); }
    var status = String(row[col['Status'] - 1] || '').trim();
    if (!status) { status = TP.projectStatus.assigned; pending.push({ r: i + 2, c: col['Status'], val: status }); }
    out.push({
      id: id,
      title: title,
      description: String(row[col['Description'] - 1] || '').trim(),
      status: status,
      assignees: tpSplitList_(row[col['Assignees'] - 1]),
      before: String(row[col['Before photo'] - 1] || '').trim(),
      after: String(row[col['After photo'] - 1] || '').trim(),
      hours: Number(row[col['Hours'] - 1]) || 0,
      completedAt: row[col['Completed at'] - 1] || '',
      sentBackReason: String(row[col['Sent back reason'] - 1] || '').trim(),
      attachment: String(row[col['Attachment'] - 1] || '').trim(),
      link: String(row[col['Link'] - 1] || '').trim(),
    });
  }
  pending.forEach(function (w) { sh.getRange(w.r, w.c).setValue(w.val); });
  return out;
}

// ---- Projects: mutations ----

// Student picks (joins) a project. The first pick flips Assigned -> In Progress.
function tpJoinProject(projectId, name) {
  name = String(name || '').trim();
  if (!name) return { ok: false, error: 'Enter your name to join.' };
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  var list = tpSplitList_(o.sh.getRange(r, o.col['Assignees']).getValue());
  var lower = list.map(function (s) { return s.toLowerCase(); });
  if (lower.indexOf(name.toLowerCase()) < 0) list.push(name);
  o.sh.getRange(r, o.col['Assignees']).setValue(list.join(', '));

  var status = String(o.sh.getRange(r, o.col['Status']).getValue() || '').trim();
  if (!status || norm_(status) === norm_(TP.projectStatus.assigned)) {
    status = TP.projectStatus.active;
    o.sh.getRange(r, o.col['Status']).setValue(status);
    if (!String(o.sh.getRange(r, o.col['Started at']).getValue()).trim()) {
      o.sh.getRange(r, o.col['Started at']).setValue(new Date());
    }
  }
  return { ok: true, status: status, assignees: list };
}

// Admin edits a project's entered fields in place: title, scope, and the assignees.
function tpUpdateProject(projectId, title, description, assignees) {
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  var t = String(title || '').trim();
  if (!t) return { ok: false, error: 'A title is required.' };
  var desc = String(description || '').trim();
  var list = tpSplitList_(assignees);
  o.sh.getRange(r, o.col['Title']).setValue(t);
  o.sh.getRange(r, o.col['Description']).setValue(desc);
  o.sh.getRange(r, o.col['Assignees']).setValue(list.join(', '));
  return { ok: true, title: t, description: desc, assignees: list };
}

// Assignee submits a finished project for approval (after photo + hours required).
// The "before" photo was captured at creation. Moves the project to Pending; an
// admin then approves it (-> Completed) or sends it back (-> In Progress).
function tpCompleteProject(projectId, afterUrl, afterName, hours) {
  if (!afterUrl) return { ok: false, error: 'An "after" photo is required.' };
  var h = Number(hours);
  if (!(h > 0)) return { ok: false, error: 'Enter how many hours the project took.' };
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  var afterId;
  try { afterId = tpSaveUpload_(afterUrl, afterName || 'after'); }
  catch (err) { return { ok: false, error: String(err.message || err) }; }
  o.sh.getRange(r, o.col['After photo']).setValue(tpViewUrl_(afterId));
  o.sh.getRange(r, o.col['Hours']).setValue(h);
  o.sh.getRange(r, o.col['Status']).setValue(TP.projectStatus.pending);
  o.sh.getRange(r, o.col['Sent back reason']).setValue('');
  var beforeIds = extractFileIds_(String(o.sh.getRange(r, o.col['Before photo']).getValue() || ''));
  return { ok: true, status: TP.projectStatus.pending, afterId: afterId, beforeId: (beforeIds[0] || ''), hours: h };
}

// Admin approves a submitted project -> Completed.
function tpApproveProject(projectId, pass) {
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  o.sh.getRange(r, o.col['Status']).setValue(TP.projectStatus.done);
  o.sh.getRange(r, o.col['Completed at']).setValue(new Date());
  return { ok: true, status: TP.projectStatus.done };
}

// Admin sends a submitted project back -> In Progress, with an optional reason.
// The after photo and hours are kept for reference.
function tpRejectProject(projectId, reason, pass) {
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  o.sh.getRange(r, o.col['Status']).setValue(TP.projectStatus.active);
  o.sh.getRange(r, o.col['Sent back reason']).setValue(String(reason || '').trim());
  return { ok: true, status: TP.projectStatus.active };
}

// Admin creates/assigns a project, optionally with a "before" photo of the starting state.
function tpCreateProject(title, description, assignees, beforeUrl, beforeName, fileUrl, fileName, link, pass) {
  title = String(title || '').trim();
  if (!title) return { ok: false, error: 'A project title is required.' };
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var list = tpSplitList_(assignees);
  var id = newToken_();
  var r = o.sh.getLastRow() + 1;
  o.sh.getRange(r, o.col['Project ID']).setValue(id);
  o.sh.getRange(r, o.col['Title']).setValue(title);
  o.sh.getRange(r, o.col['Description']).setValue(String(description || '').trim());
  o.sh.getRange(r, o.col['Status']).setValue(TP.projectStatus.assigned);
  o.sh.getRange(r, o.col['Assignees']).setValue(list.join(', '));
  o.sh.getRange(r, o.col['Created at']).setValue(new Date());
  if (beforeUrl) {
    try { o.sh.getRange(r, o.col['Before photo']).setValue(tpViewUrl_(tpSaveUpload_(beforeUrl, beforeName || 'before'))); }
    catch (err) { Logger.log('Before photo save failed: ' + err); }
  }
  if (fileUrl) {
    try { o.sh.getRange(r, o.col['Attachment']).setValue(tpViewUrl_(tpSaveUpload_(fileUrl, fileName || 'attachment'))); }
    catch (err) { Logger.log('Attachment save failed: ' + err); }
  }
  var lk = String(link || '').trim();
  if (lk) {
    if (!/^https?:\/\//i.test(lk)) lk = 'https://' + lk.replace(/^\/+/, '');
    o.sh.getRange(r, o.col['Link']).setValue(lk);
  }
  return { ok: true, project: { id: id, title: title } };
}

// Admin deletes a project (removes the row). Uploaded photos are left in Drive.
function tpDeleteProject(projectId, pass) {
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  o.sh.deleteRow(r);
  return { ok: true };
}

// ---- Shared page furniture ----

function tpStyles_() {
  return '<style>'
    + '[hidden]{display:none!important}'   // hidden attr must beat class display rules (.btn-row/.ic-admin-fields) so admin controls stay hidden until unlock
    + '.ic-summary{display:flex;flex-wrap:wrap;align-items:center;gap:12px;margin:12px 0 2px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13.5px;font-weight:600;color:#8a857c}'
    + '.ic-sum b{font-weight:800;color:#111;font-size:16px;margin-right:5px}'
    + '.ic-sum--danger b{color:#b31b1b}'
    + '.ic-dot{width:4px;height:4px;border-radius:50%;background:#d6d3ce}'
    + '.ic-adminbar{display:flex;align-items:center;gap:10px;justify-content:flex-end;flex-wrap:wrap;margin:2px 0}'
    + '.ic-admin-toggle{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#b5b0a8;background:none;border:none;cursor:pointer;padding:4px 2px}'
    + '.ic-admin-toggle:hover{color:#8f1515}'
    + '.ic-admin-fields{display:inline-flex;gap:8px;align-items:center;flex-wrap:wrap}'
    + '.ic-admin-fields input{font:inherit;font-size:13px;padding:8px 11px;border:1.5px solid #e0e0dc;border-radius:9px;background:#fff;outline:none;min-width:150px}'
    + '.ic-admin-fields input:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.tp-lock-msg{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700}'
    + '.tp-lock-msg.ok{color:#157a47}.tp-lock-msg.bad{color:#b31b1b}'
    + '.tp-assignees{display:flex;flex-wrap:wrap;gap:6px;margin-top:12px}'
    + '.tp-chip{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#3f3a34;background:#f0efe9;border:1px solid #e5e4de;border-radius:999px;padding:4px 11px}'
    + '.tp-attach{display:inline-flex;align-items:center;gap:6px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#8f1515;text-decoration:none;border:1px solid #ece9e2;border-radius:9px;padding:6px 11px;background:#faf9f6}'
    + '.tp-attach:hover{border-color:#d6b26a;background:#fff}'
    + '.tp-refs{display:flex;gap:8px;flex-wrap:wrap}'
    + '.card-accent{height:3px}'
    + '@keyframes cardIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}'
    + '.card{animation:cardIn .3s ease}'
    + '.tp-chip--empty{color:#a8a29e;background:transparent;border-style:dashed}'
    + '.tp-pill{flex:0 0 auto;white-space:nowrap;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;padding:5px 11px;border-radius:999px}'
    + '.tp-pill--todo{color:#6b665e;background:#f0efe9;border:1px solid #e5e4de}'
    + '.tp-pill--assigned{color:#2563c9;background:#eaf1fd;border:1px solid #cfe0fb}'
    + '.tp-pill--active{color:#e08a1e;background:#fdf3e3;border:1px solid #f6e2bf}'
    + '.tp-pill--pending{color:#b06a00;background:#fdf2df;border:1px solid #f4dfb0}'
    + '.tp-pill--done{color:#157a47;background:#e7f6ee;border:1px solid #c7e9d5}'
    + '.tp-hint{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#a8a29e}'
    + '.tp-photos{display:flex;flex-wrap:wrap;gap:14px;margin-top:14px}'
    + '.tp-photo figure{margin:0}'
    + '.tp-photo figcaption{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#a8a29e;margin-bottom:5px}'
    // ---- staged-photo tray (add photos before submitting) ----
    + '.stage{display:grid;grid-template-columns:repeat(auto-fill,minmax(108px,1fr));gap:11px;margin-top:16px}'
    + '.stage-item{position:relative;aspect-ratio:1;border-radius:14px;overflow:hidden;background:#efeee9;border:1px solid #e6e3db;box-shadow:0 3px 12px rgba(20,20,30,.08);animation:stageIn .3s cubic-bezier(.2,.9,.3,1.35) both}'
    + '@keyframes stageIn{from{opacity:0;transform:scale(.85)}to{opacity:1;transform:scale(1)}}'
    + '.stage-item img{width:100%;height:100%;object-fit:cover;display:block}'
    + '.stage-item::before{content:"";position:absolute;left:0;right:0;top:0;height:52%;background:linear-gradient(180deg,rgba(20,17,14,.34),transparent);pointer-events:none}'
    + '.stage-rm{position:absolute;top:6px;right:6px;width:26px;height:26px;padding:0;border:none;border-radius:50%;background:rgba(20,17,14,.62);color:#fff;font-size:16px;line-height:1;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .16s,transform .16s}'
    + '.stage-rm:hover{background:#b31b1b;transform:scale(1.09)}'
    + '.stage-rm:active{transform:scale(.96)}'
    + '.stage-idx{position:absolute;left:9px;top:7px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.04em;color:#fff;text-shadow:0 1px 3px rgba(0,0,0,.45)}'
    + '.stage-add{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;aspect-ratio:1;border-radius:14px;border:1.6px dashed #d3cdc2;background:#faf9f6;color:#9a948a;cursor:pointer;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;transition:border-color .18s,color .18s,background .18s,transform .12s}'
    + '.stage-add:hover{border-color:#b31b1b;color:#8f1515;background:#fff}'
    + '.stage-add:active{transform:scale(.97)}'
    + '.stage-add-i{font-size:26px;font-weight:300;line-height:1}'
    // ---- refresh control on the open-issues pages ----
    + '.ic-refresh{display:inline-flex;align-items:center;gap:7px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#57534e;background:#fff;border:1.5px solid #e2ddd6;border-radius:10px;padding:8px 13px;cursor:pointer;transition:border-color .15s,color .15s,box-shadow .15s}'
    + '.ic-refresh:hover{border-color:#b5b0a8;color:#292524;box-shadow:0 2px 6px rgba(20,20,30,.06)}'
    + '.ic-refresh:disabled{opacity:.65;cursor:default}'
    + '.ic-refresh svg{display:block}'
    + '.ic-refresh.is-spin svg{animation:icspin .8s linear infinite}'
    + '@keyframes icspin{to{transform:rotate(360deg)}}'
    + '.ic-refreshed{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:600;color:#157a47;white-space:nowrap}'
    + '.tp-uploader{margin-top:14px;padding:14px 16px;background:#fbfbf9;border:1.5px dashed #ddd;border-radius:12px}'
    + '.tp-drop{display:flex;flex-wrap:wrap;gap:10px}'
    + '.tp-slot{flex:1;min-width:150px}'
    + '.tp-slot-btn{display:flex;align-items:center;justify-content:center;gap:8px;width:100%;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#57534e;padding:14px 12px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;cursor:pointer;text-align:center}'
    + '.tp-slot-btn:hover{border-color:#b31b1b;color:#8f1515}'
    + '.tp-slot.is-set .tp-slot-btn{border-color:#157a47;color:#157a47;background:#f2fbf6}'
    + '.tp-hours-field{flex:1;min-width:150px;display:flex;flex-direction:column;justify-content:center}'
    + '.tp-hours-label{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c;margin-bottom:5px}'
    + '.tp-hours-input{font:inherit;font-size:14px;padding:12px 12px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;outline:none;width:100%}'
    + '.tp-hours-input:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.tp-create-wrap{margin:6px 0 20px}'
    + '.tp-create-toggle{display:inline-flex;align-items:center;gap:8px;cursor:pointer;list-style:none;user-select:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;letter-spacing:.02em;color:#8f1515;background:#fff;border:1.5px dashed #e0ddd6;border-radius:10px;padding:9px 14px;transition:border-color .15s ease}'
    + '.tp-create-toggle:hover{border-color:#d6b26a}'
    + '.tp-req{font-size:9px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b;margin-left:5px;vertical-align:1px}'
    + '.tp-opt{font-size:9px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#a8a29e;margin-left:5px;vertical-align:1px}'
    + '.tp-attach-row{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}'
    + '.tp-attach-btn{display:inline-flex;align-items:center;gap:7px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#57534e;background:#faf9f6;border:1.5px dashed #d6d3ce;border-radius:9px;padding:8px 13px;cursor:pointer;transition:border-color .15s,color .15s,background .15s}'
    + '.tp-attach-btn:hover{border-color:#b5b0a8;color:#292524}'
    + '.tp-attach-btn.is-set{border-style:solid;border-color:#157a47;color:#157a47;background:#eef6f0}'
    + '.tp-attach-link{display:flex;align-items:center;gap:8px;border:1.5px solid #e2ddd6;border-radius:9px;padding:0 11px;background:#fff;transition:border-color .15s}'
    + '.tp-attach-link:focus-within{border-color:#b31b1b}'
    + '.tp-link-icon{color:#a8a29e;font-size:14px;flex-shrink:0}'
    + '.tp-attach-link input{flex:1;min-width:0;border:none;background:none;padding:9px 0;font:inherit;font-size:13px;outline:none;color:#111}'
    + '.tp-create-toggle::-webkit-details-marker{display:none}'
    + '.tp-create-toggle::marker{content:""}'
    + '.tp-create-caret{display:inline-block;font-weight:800;font-size:15px;line-height:1;transition:transform .15s ease}'
    + '.tp-create-wrap[open] .tp-create-caret{transform:rotate(45deg)}'
    + '.tp-create-wrap[open] .tp-create-toggle{border-style:solid;border-color:#e7e7e3;color:#111}'
    + '.tp-create{margin:10px 0 0;padding:18px 20px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.tp-create h3{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:800;letter-spacing:.04em;color:#111;margin:0 0 12px}'
    + '.tp-field{margin-bottom:12px}'
    + '.tp-field label{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c;margin-bottom:5px}'
    + '.tp-field input,.tp-field textarea{width:100%;font:inherit;font-size:14px;padding:10px 12px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fafaf8;outline:none;resize:vertical}'
    + '.tp-field input:focus,.tp-field textarea:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12);background:#fff}'
    + '.tp-inline-join{display:flex;gap:8px;flex-wrap:wrap;align-items:center}'
    + '.tp-inline-join input{font:inherit;font-size:14px;padding:9px 12px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;outline:none;min-width:160px}'
    + '.tp-inline-join input:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12)}'
    + '#tp-cfx{position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:9999}'
    + '.ic-note{margin-top:12px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:600;line-height:1.55;color:#6b665e;background:#f7f6f2;border:1px solid #ece9e2;border-radius:10px;padding:10px 13px}'
    + '.ic-note b{color:#8f1515}'
    + '.ic-note--warn{color:#8a4b00;background:#fdf2df;border-color:#f4dfb0}'
    + '.ic-note--warn b{color:#8a4b00}'
    + '.ic-note-cta{display:block;margin-top:5px;font-weight:600;opacity:.82}'
    + '.ic-edit{margin-top:8px;padding:14px 15px;background:#faf9f6;border:1px solid #e9e6df;border-radius:12px}'
    + '.ic-edit-lbl{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.11em;text-transform:uppercase;color:#a8a29e;margin-bottom:12px}'
    + '.ic-edit-in{display:block;width:100%;box-sizing:border-box;margin-top:5px;padding:9px 11px;border:1.5px solid #e2ddd6;border-radius:9px;font-family:inherit;font-size:13.5px;color:#111;background:#fff}'
    + '.ic-edit-in:focus{outline:none;border-color:#b31b1b}'
    + 'textarea.ic-edit-in{resize:vertical;line-height:1.5}'
    + '.ic-edit-btns{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:2px}'
    + '.ic-edit-msg{font-size:12px;font-weight:600;color:#8a857c}'
    + '.tp-pill--sent{color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0}'
    + '.tp-del{color:#b31b1b;border-color:#f0d0d0}'
    + '.tp-del:hover{border-color:#e0a0a0;background:#fdf6f6;color:#8f1515}'
    + '.ic-reason{font:inherit;font-size:13px;padding:9px 12px;border:1.5px solid #e0e0dc;border-radius:9px;background:#fff;outline:none;min-width:170px;flex:1}'
    + '.ic-reason:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '#tp-cheer{position:fixed;left:50%;top:20%;transform:translate(-50%,-10px) scale(.92);z-index:10000;pointer-events:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-weight:800;font-size:18px;color:#fff;background:linear-gradient(180deg,#1d9d5b 0%,#157a47 100%);padding:14px 22px;border-radius:14px;box-shadow:0 14px 34px rgba(21,122,71,.34);opacity:0;transition:opacity .3s ease,transform .35s cubic-bezier(.2,.9,.3,1.4);max-width:88vw;text-align:center}'
    + '#tp-cheer.show{opacity:1;transform:translate(-50%,0) scale(1)}'
    // Project cards: a status-colored left rail, a soft hover lift, and refined chips.
    + '#tp-proj-list .card{border-left:4px solid #d6d3ce;transition:transform .18s ease,box-shadow .18s ease}'
    + '#tp-proj-list .card:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(20,20,30,.10)}'
    + '#tp-proj-list .card[data-status="Assigned"]{border-left-color:#2563c9}'
    + '#tp-proj-list .card[data-status="In Progress"]{border-left-color:#e08a1e}'
    + '#tp-proj-list .card[data-status="Pending"]{border-left-color:#b06a00}'
    + '#tp-proj-list .card[data-status="Completed"]{border-left-color:#157a47}'
    + '#tp-proj-list .card-team{color:#a8a29e}'
    + '#tp-proj-list .card-title{font-size:19px;margin-top:5px}'
    + '#tp-proj-list .card-body{padding:20px 22px 18px}'
    + '#tp-proj-list .tp-chip{background:#f4f2ec;border-color:#e8e3d7;color:#4a453d;padding:5px 12px}'
    + '.tp-create{border-radius:16px;box-shadow:0 6px 20px rgba(20,20,30,.06)}'
    + '.tp-create h3{font-size:14px;letter-spacing:-.01em;text-transform:none}'
    + '.tp-c-photo{max-width:none;width:100%}'
    + '.tp-c-photo .tp-slot-btn{border-style:dashed;padding:18px 14px;font-size:13px}'
    + '.tp-c-cam{font-size:16px;margin-right:2px}'
    + '.section-head{margin:30px 0 14px}'
    // Mobile: stack the completion uploader, admin bar, join box and buttons; shrink photos.
    + '@media(max-width:600px){'
    +   '.ic-adminbar{justify-content:flex-start}'
    +   '.ic-admin-fields{width:100%}.ic-admin-fields input{flex:1;min-width:0}'
    +   '.tp-drop{flex-direction:column;gap:12px}.tp-slot{min-width:0}.tp-hours-field{min-width:0}'
    +   '.tp-inline-join{width:100%}.tp-inline-join input{flex:1;min-width:0}'
    +   '.tp-create{padding:16px 14px}.tp-uploader{padding:13px}'
    +   '.ic-reason{min-width:0;width:100%}'
    +   '.tp-photos{gap:10px}.tp-photo img{max-height:150px}'
    +   '.card-foot{gap:8px}.card-foot .btn-row{width:100%;flex-wrap:wrap}.card-foot .btn-row>.btn{flex:1 1 auto}'
    +   '.tp-cheer,#tp-cheer{font-size:15px;padding:12px 16px;max-width:92vw}'
    +   '.ic-summary{gap:7px 12px;font-size:12.5px}.ic-sum b{font-size:15px}'
    + '}'
    + '@media(max-width:400px){.tp-photos{flex-direction:column}.tp-photo img{max-height:none;width:100%}}'
    + '</style>';
}

// Vanilla, dependency-free confetti burst + the admin-unlock JS + a file reader.
// Shared by the issue pages (in 02) and the projects page (below).
function tpSharedJs_() {
  return '<script>'
    + 'var ADMIN_PASS="";'
    + 'function tpUnlock(){var p=(document.getElementById("tp-pass")||{}).value||"";var m=document.getElementById("tp-lock-msg");'
    + 'if(m){m.textContent="Checking\\u2026";m.className="tp-lock-msg";}'
    + 'google.script.run.withSuccessHandler(function(ok){'
    + 'if(ok){ADMIN_PASS=p;document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});'
    + 'if(m){m.textContent="\\u2713 Admin unlocked";m.className="tp-lock-msg ok";}'
    + 'var lb=document.getElementById("tp-lock-fields");if(lb)lb.style.display="none";}'
    + 'else if(m){m.textContent="Incorrect passcode";m.className="tp-lock-msg bad";}'
    + '}).withFailureHandler(function(){if(m){m.textContent="Could not verify. Try again.";m.className="tp-lock-msg bad";}}).tpCheckPass(p);}'
    + 'function tpAdminToggle(){var f=document.getElementById("tp-lock-fields");if(f)f.hidden=false;var t=document.getElementById("tp-admin-toggle");if(t)t.style.display="none";var p=document.getElementById("tp-pass");if(p)p.focus();}'
    + 'function tpApproveFx(rid){var c=document.getElementById(rid);if(!c)return;c.classList.remove("card--approved");void c.offsetWidth;c.classList.add("card--approved");setTimeout(function(){if(c)c.classList.remove("card--approved");},900);}'
    + 'function tpAdvance(rid,hex){var a=document.getElementById(rid+"-accent");if(a)a.style.background=hex;}'
    + 'function tpConfetti(msg){var c=document.getElementById("tp-cfx");if(!c){c=document.createElement("canvas");c.id="tp-cfx";document.body.appendChild(c);}'
    + 'var cheer=document.getElementById("tp-cheer");if(!cheer){cheer=document.createElement("div");cheer.id="tp-cheer";document.body.appendChild(cheer);}'
    + 'var msgs=["\\uD83C\\uDF89 Boom! Thanks for taking care of that.","\\uD83D\\uDE4C Nice work. The space thanks you!","\\u2B50 Legend. Thanks for closing that out!","\\u2728 Done and dusted. Thank you!","\\uD83D\\uDCAA You crushed it. Thanks a ton!","\\uD83D\\uDE80 One down. Thanks for handling it!","\\uD83E\\uDD73 High five! Thanks for getting it done."];'
    + 'cheer.textContent=msg||msgs[Math.floor(Math.random()*msgs.length)];cheer.className="";void cheer.offsetWidth;cheer.className="show";clearTimeout(cheer._t);cheer._t=setTimeout(function(){cheer.className="";},2600);'
    + 'var ctx=c.getContext("2d");var W=c.width=window.innerWidth,H=c.height=window.innerHeight;'
    + 'var cols=["#b31b1b","#e08a1e","#f0c050","#1d9d5b","#2563c9","#7c3aed"];var P=[];'
    + 'for(var i=0;i<150;i++){P.push({x:W/2+(Math.random()-0.5)*W*0.35,y:H*0.34+(Math.random()-0.5)*60,vx:(Math.random()-0.5)*15,vy:Math.random()*-15-4,g:0.30+Math.random()*0.22,s:6+Math.random()*7,rot:Math.random()*6.28,vr:(Math.random()-0.5)*0.45,col:cols[i%cols.length]});}'
    + 'var start=Date.now();(function frame(){var t=Date.now()-start;ctx.clearRect(0,0,W,H);'
    + 'P.forEach(function(p){p.vy+=p.g;p.x+=p.vx;p.y+=p.vy;p.rot+=p.vr;p.vx*=0.99;ctx.save();ctx.translate(p.x,p.y);ctx.rotate(p.rot);ctx.globalAlpha=Math.max(0,1-t/2600);ctx.fillStyle=p.col;ctx.fillRect(-p.s/2,-p.s/2,p.s,p.s*0.62);ctx.restore();});'
    + 'if(t<2600){requestAnimationFrame(frame);}else{ctx.clearRect(0,0,W,H);}})();}'
    // Downscale + JPEG-compress in the browser before upload. A 12MP phone photo
    // (~8MB) becomes a ~1600px JPEG (~0.3MB), so google.script.run and the Drive
    // write handle a fraction of the bytes. Falls back to the original if the image
    // cannot be decoded (e.g. an unusual format).
    + 'function tpReadOne(f,cb,silent){if(!f){cb(null);return;}if(!/^image\\//.test(f.type)){if(!silent)alert("Please choose an image file.");cb(null);return;}'
    + 'var r=new FileReader();r.onerror=function(){if(!silent)alert("Could not read that file.");cb(null);};'
    + 'r.onload=function(){var img=new Image();'
    + 'img.onload=function(){var max=1600,w=img.width,h=img.height;if(w>max||h>max){var s=Math.min(max/w,max/h);w=Math.round(w*s);h=Math.round(h*s);}'
    + 'try{var c=document.createElement("canvas");c.width=w;c.height=h;var x=c.getContext("2d");x.fillStyle="#fff";x.fillRect(0,0,w,h);x.drawImage(img,0,0,w,h);'
    + 'cb({dataUrl:c.toDataURL("image/jpeg",0.82),name:(f.name||"photo").replace(/\\.[^.]+$/,"")+".jpg"});}catch(e){cb({dataUrl:r.result,name:f.name});}};'
    + 'img.onerror=function(){cb({dataUrl:r.result,name:f.name});};img.src=r.result;};'
    + 'r.readAsDataURL(f);}'
    + 'function tpReadFile(input,cb){tpReadOne((input.files&&input.files[0])||null,function(res){if(!res)input.value="";cb(res);});}'
    // Staging model: read the picked images into a client-side buffer (arr) WITHOUT
    // uploading, so photos can be added one pick at a time and reviewed before the
    // doer commits. cb() fires once every file in this pick has been read.
    + 'function tpStageRead(input,arr,cb){var files=[];for(var k=0;input.files&&k<input.files.length;k++){if(/^image\\//.test(input.files[k].type))files.push(input.files[k]);}input.value="";if(!files.length){cb();return;}var i=0;(function next(){if(i>=files.length){cb();return;}tpReadOne(files[i],function(res){if(res)arr.push(res);i++;next();},true);})();}'
    // Upload an already-read buffer (from tpStageRead) in order. The first upload
    // replaces the cell (dropping any stale sent-back photo); the rest append.
    + 'function tpUploadStaged(arr,token,onProgress,onDone,onError){if(!arr||!arr.length){onDone([]);return;}var ids=[],i=0;function step(){if(i>=arr.length){onDone(ids);return;}if(onProgress)onProgress(i,arr.length);'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){onError((r&&r.error)||"Upload failed");return;}if(r.photoId)ids.push(r.photoId);i++;step();}).withFailureHandler(function(){onError("Upload failed. Please retry.");}).submitIssueCompletion(token,arr[i].dataUrl,arr[i].name,ids.length>0);}step();}'
    // The staged-photo tray (local data URLs, not yet uploaded): a grid of cover-fit
    // thumbnails each with a remove (x), and a trailing "Add more" tile. removeFn is the
    // call prefix, e.g. "icUnstage(\'iss0\'," -> icUnstage(\'iss0\',2); addExpr is the
    // onclick that reopens the picker. Empty tray renders nothing (the foot button adds).
    + 'function tpStagePreview(arr,removeFn,addExpr){if(!arr||!arr.length)return "";var h="<div class=\\"stage\\">";for(var i=0;i<arr.length;i++){h+="<div class=\\"stage-item\\"><img src=\\""+arr[i].dataUrl+"\\"><span class=\\"stage-idx\\">"+(i+1)+"</span><button type=\\"button\\" class=\\"stage-rm\\" title=\\"Remove\\" onclick=\\""+removeFn+i+")\\">&times;</button></div>";}if(addExpr){h+="<button type=\\"button\\" class=\\"stage-add\\" onclick=\\""+addExpr+"\\"><span class=\\"stage-add-i\\">+</span>Add more</button>";}return h+"</div>";}'
    // Render a strip of completion-photo thumbnails from an array of Drive file ids.
    + 'function tpThumbs(ids){if(!ids||!ids.length)return "";var h="<div class=\\"tp-photos\\">";for(var i=0;i<ids.length;i++){var e=encodeURIComponent(ids[i]);var cap=ids.length>1?"Completion photo "+(i+1):"Completion photo";'
    + 'h+="<div class=\\"tp-photo\\"><figure><figcaption>"+cap+"</figcaption><a href=\\"https://drive.google.com/file/d/"+e+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+e+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div>";}return h+"</div>";}'
    + 'function tpReadAnyFile(input,cb){var f=input.files&&input.files[0];if(!f){cb(null);return;}if(f.size>10485760){alert("That file is too large (max 10 MB).");input.value="";cb(null);return;}var r=new FileReader();r.onerror=function(){alert("Could not read that file.");cb(null);};r.onload=function(){cb({dataUrl:r.result,name:f.name});};r.readAsDataURL(f);}'
    + '</script>';
}

// Admin mode is signalled by the ?admin=1 flag on links from the (unlisted) admin
// page - no passcode. This reveals the .tp-admin controls and marks the session
// admin so the dynamic foot builders (icAdminFootJs / tpProjAdminFootJs / tpDelWrapJs)
// include their admin buttons. Renders nothing for regular doers.
function tpAdminRevealJs_(admin) {
  return admin
    ? '<script>ADMIN_PASS="admin";document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});</script>'
    : '';
}

// ---- Completion evidence: UI helpers used by the issue pages in 02 ----

// Empty (or pre-filled) container for the completion photo on a card.
function icPhotoBlock_(rid, url) {
  var cells = tpPhotoCells_(url, 200, 'Completion photo');
  var inner = cells ? '<div class="tp-photos">' + cells + '</div>' : '';
  return '<div id="' + rid + '-photo">' + inner + '</div>';
}

// Plain-language instruction on an open item. Wording adapts to whether the task
// needs a photo.
function icNoteText_(photoOptional) {
  return photoOptional
    ? 'Finished? Tap <b>Mark done</b> and an admin gives it a quick review. A photo is optional here, so add one only if it helps.'
    : 'Finished? Add a photo of the completed work and an admin gives it a quick review.';
}
function icNoteInner_(photoOptional) { return '<div class="ic-note">' + icNoteText_(photoOptional) + '</div>'; }

// Amber banner shown when a submission was sent back, with the admin's reason.
function icSentBackInner_(reason) {
  var r = String(reason || '').trim();
  var lead = r ? '<b>Sent back:</b> ' + escapeHtml_(r) : '<b>Sent back.</b>';
  return '<div class="ic-note ic-note--warn">' + lead + '<span class="ic-note-cta">Make the fix, then send it in again.</span></div>';
}

// The card's note slot: sent-back banner if there's a reason, else the plain
// instruction on a fresh open item, else nothing (pending / read-only).
function icNoteContainer_(rid, isPending, sentBackReason, show, photoOptional) {
  var inner = '';
  if (show) {
    if (sentBackReason) inner = icSentBackInner_(sentBackReason);
    else if (!isPending) inner = icNoteInner_(photoOptional);
  }
  return '<div id="' + rid + '-note">' + inner + '</div>';
}

// Foot action for an OPEN item. Photos are STAGED (added one pick at a time and
// previewed) and only uploaded when the doer taps Complete, so several photos can be
// attached without each one closing the task. Photo-optional tasks can Complete with
// none. Admins get an extra "Complete without a photo" bypass, revealed on unlock.
function icOpenFoot_(rid, token, photoOptional) {
  var input = '<input type="file" accept="image/*" multiple id="' + rid + '-file" style="display:none" onchange="icPick(this,\'' + rid + '\')">';
  var addBtn = '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'' + rid + '-file\').click()">Add photos</button>';
  var doneBtn = '<button type="button" class="btn btn-primary" id="' + rid + '-done" onclick="icComplete(\'' + rid + '\',\'' + token + '\',' + (photoOptional ? 'true' : 'false') + ')">' + (photoOptional ? 'Mark done' : 'Complete') + '</button>';
  var hint = '<span id="' + rid + '-stagehint" class="tp-hint"></span>';
  // Admin no-photo bypass only where it waives a real requirement: photo-required
  // tasks. Photo-optional tasks already have "Mark done", so it would be redundant.
  var admin = photoOptional ? '' : '<span class="tp-admin btn-row" hidden><button type="button" class="btn btn-ghost" onclick="icResolve(\'' + rid + '\',\'' + token + '\')">Complete without a photo</button></span>';
  return input + addBtn + doneBtn + hint + admin;
}

// Foot action for a PENDING item: admin-only Approve / Send back (revealed on unlock).
function icPendingFoot_(rid, token) {
  return '<span class="tp-admin btn-row" hidden>'
    + '<button type="button" class="btn btn-confirm" onclick="icApprove(\'' + rid + '\',\'' + token + '\')">Approve</button>'
    + '<button type="button" class="btn btn-ghost" onclick="icRejectOpen(\'' + rid + '\',\'' + token + '\')">Send back</button>'
    + '</span>';
}

// The client logic that drives completion on the issue pages. Requires tpSharedJs_
// (ADMIN_PASS, tpConfetti, tpReadFile) on the same page. Stat counters (#sum-open,
// #sum-pending, #sum-over) are updated only if present.
function icClientJs_() {
  return '<script>'
    + 'function icSetPill(rid,cls,txt){var p=document.getElementById(rid+"-pill");if(p)p.innerHTML="<span class=\\"tp-pill "+cls+"\\">"+txt+"</span>";}'
    + 'function icBump(id,d){var e=document.getElementById(id);if(e)e.textContent=Math.max(0,(parseInt(e.textContent,10)||0)+d);}'
    + 'function icOpenFootJs(rid,token){var c=document.getElementById(rid);var po=c&&c.dataset.po==="1";var input="<input type=\\"file\\" accept=\\"image/*\\" multiple id=\\""+rid+"-file\\" style=\\"display:none\\" onchange=\\"icPick(this,\'"+rid+"\')\\">";var addBtn="<button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"document.getElementById(\'"+rid+"-file\').click()\\">Add photos</button>";var doneBtn="<button type=\\"button\\" class=\\"btn btn-primary\\" id=\\""+rid+"-done\\" onclick=\\"icComplete(\'"+rid+"\',\'"+token+"\',"+(po?"true":"false")+")\\">"+(po?"Mark done":"Complete")+"</button>";var hint="<span id=\\""+rid+"-stagehint\\" class=\\"tp-hint\\"></span>";var admin=(ADMIN_PASS&&!po)?"<button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"icResolve(\'"+rid+"\',\'"+token+"\')\\">Complete without a photo</button>":"";return input+addBtn+doneBtn+hint+admin;}'
    + 'function icResolve(rid,token){var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed. Retry.")+"</span>";return;}'
    + 'icSetPill(rid,"tp-pill--done","Completed");act.innerHTML="";var s=document.getElementById(rid+"-status");if(s){s.innerHTML="\\u2713 Completed";s.className="due due--done";}tpApproveFx(rid);tpAdvance(rid,"#157a47","#e7f3ec",2);var c=document.getElementById(rid);if(c){c.style.opacity="0.72";icBump("sum-open",-1);if(c.dataset.over==="1")icBump("sum-over",-1);}'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).resolveIssueComplete(token);}'
    + 'function icNoteInner(){return "<div class=\\"ic-note\\">Finished? Add a photo of the completed work and an admin gives it a quick review.</div>";}'
    + 'function icAdminFootJs(rid,token){return ADMIN_PASS?"<span class=\\"tp-admin btn-row\\"><button type=\\"button\\" class=\\"btn btn-confirm\\" onclick=\\"icApprove(\'"+rid+"\',\'"+token+"\')\\">Approve</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"icRejectOpen(\'"+rid+"\',\'"+token+"\')\\">Send back</button></span>":"";}'
    + 'function icSentBackInner(reason){var r=(reason||"").trim();var esc=function(s){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");};var lead=r?"<b>Sent back:</b> "+esc(r):"<b>Sent back.</b>";return "<div class=\\"ic-note ic-note--warn\\">"+lead+"<span class=\\"ic-note-cta\\">Make the fix, then send it in again.</span></div>";}'
    + 'function icAfterSubmit(rid,token,photoIds){tpConfetti();icSetPill(rid,"tp-pill--pending","Pending approval");tpAdvance(rid,"#b06a00","#f7edd8",1);var ids=Array.isArray(photoIds)?photoIds:(photoIds?[photoIds]:[]);var ph=document.getElementById(rid+"-photo");if(ph)ph.innerHTML=tpThumbs(ids);'
    + 'var s=document.getElementById(rid+"-status");if(s){s.textContent="Submitted, awaiting approval";s.className="due";}'
    + 'var nt=document.getElementById(rid+"-note");if(nt)nt.innerHTML="";'
    + 'var act=document.getElementById(rid+"-act");if(act)act.innerHTML=icAdminFootJs(rid,token);var c=document.getElementById(rid);if(c){if(c.dataset.over==="1"){icBump("sum-over",-1);c.dataset.over="0";}c.dataset.state="pending";}'
    + 'icBump("sum-open",-1);icBump("sum-pending",1);}'
    + 'var ICBUF={};'
    + 'function icPick(input,rid){ICBUF[rid]=ICBUF[rid]||[];tpStageRead(input,ICBUF[rid],function(){icRenderStage(rid);});}'
    + 'function icRenderStage(rid){var buf=ICBUF[rid]||[];var ph=document.getElementById(rid+"-photo");if(ph)ph.innerHTML=tpStagePreview(buf,"icUnstage(\'"+rid+"\',","document.getElementById(\'"+rid+"-file\').click()");var h=document.getElementById(rid+"-stagehint");if(h){h.style.color="";h.textContent="";}var c=document.getElementById(rid);var po=c&&c.dataset.po==="1";var db=document.getElementById(rid+"-done");if(db)db.textContent=buf.length?((po?"Submit":"Complete")+" \\u00b7 "+buf.length+" photo"+(buf.length===1?"":"s")):(po?"Mark done":"Complete");}'
    + 'function icUnstage(rid,idx){if(ICBUF[rid]){ICBUF[rid].splice(idx,1);icRenderStage(rid);}}'
    + 'function icComplete(rid,token,po){var buf=ICBUF[rid]||[];var act=document.getElementById(rid+"-act");var h=document.getElementById(rid+"-stagehint");'
    + 'if(!po&&!buf.length){if(h){h.style.color="#b31b1b";h.textContent="Add at least one photo first.";}return;}'
    + 'act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'if(!buf.length){google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}icAfterSubmit(rid,token,[]);}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).submitIssueCompletion(token,"","");return;}'
    + 'tpUploadStaged(buf,token,function(done,total){act.innerHTML="<span class=\\"tp-hint\\">Uploading photo "+(done+1)+" of "+total+"\\u2026</span>";},'
    + 'function(ids){delete ICBUF[rid];icAfterSubmit(rid,token,ids);},'
    + 'function(msg){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+msg+"</span>";});}'
    + 'function icEditOpen(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),c=document.getElementById(rid);if(v)v.hidden=true;if(e)e.hidden=false;var f=c&&c.querySelector(".card-foot");if(f)f.style.display="none";}'
    + 'function icEditCancel(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),m=document.getElementById(rid+"-emsg"),c=document.getElementById(rid);if(e)e.hidden=true;if(v)v.hidden=false;if(m)m.textContent="";var f=c&&c.querySelector(".card-foot");if(f)f.style.display="";}'
    + 'function icEditSave(rid,token){var g=function(s){var el=document.getElementById(rid+s);return el?el.value:"";};var team=g("-eteam"),type=g("-etype"),det=g("-edetails");var m=document.getElementById(rid+"-emsg");if(m){m.style.color="";m.textContent="Saving\\u2026";}'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){if(m){m.style.color="#b31b1b";m.textContent=(r&&r.error)||"Could not save";}return;}'
    + 'var st=function(s,t){var el=document.getElementById(rid+s);if(el)el.textContent=t;};st("-vteam",r.team||"Unassigned");st("-vtype",r.issueType||"Reported issue");'
    + 'var dw=document.getElementById(rid+"-vdetails-wrap");if(dw){if(r.details){dw.hidden=false;st("-vdetails",r.details);}else dw.hidden=true;}'
    + 'var c=document.getElementById(rid);if(c){var ca=(document.getElementById(rid+"-vaction")||{}).textContent||"";c.dataset.team=r.team||"";c.dataset.hay=((r.team||"")+" "+(r.issueType||"")+" "+ca+" "+(r.details||"")).toLowerCase();}'
    + 'icEditCancel(rid);'
    + '}).withFailureHandler(function(){if(m){m.style.color="#b31b1b";m.textContent="Could not save. Retry.";}}).updateIssueFields(token,team,type,det);}'
    + 'function icDelOpen(rid,token){document.getElementById(rid+"-delwrap").innerHTML="<span class=\\"tp-hint\\" style=\\"margin-right:6px\\">Delete this task?</span><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"icDelDo(\'"+rid+"\',\'"+token+"\')\\">Yes, delete</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"icDelCancel(\'"+rid+"\',\'"+token+"\')\\">Cancel</button>";}'
    + 'function icDelCancel(rid,token){document.getElementById(rid+"-delwrap").innerHTML="<button type=\\"button\\" class=\\"btn btn-ghost tp-del\\" onclick=\\"icDelOpen(\'"+rid+"\',\'"+token+"\')\\">Delete</button>";}'
    + 'function icDelDo(rid,token){var w=document.getElementById(rid+"-delwrap");w.innerHTML="<span class=\\"tp-hint\\">Deleting\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){w.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'var c=document.getElementById(rid);if(c){if(c.dataset.state==="pending")icBump("sum-pending",-1);else{icBump("sum-open",-1);if(c.dataset.over==="1")icBump("sum-over",-1);}c.parentNode.removeChild(c);}'
    + '}).withFailureHandler(function(){w.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).deleteIssue(token);}'
    + 'function icApprove(rid,token){var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'icSetPill(rid,"tp-pill--done","Completed");act.innerHTML="";var s=document.getElementById(rid+"-status");if(s){s.innerHTML="\\u2713 Completed";s.className="due due--done";}tpApproveFx(rid);tpAdvance(rid,"#157a47","#e7f3ec",2);var c=document.getElementById(rid);if(c)c.style.opacity="0.72";icBump("sum-pending",-1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).approveIssueCompletion(token,ADMIN_PASS);}'
    + 'function icRejectOpen(rid,token){var act=document.getElementById(rid+"-act");act.innerHTML="<input id=\\""+rid+"-reason\\" class=\\"ic-reason\\" placeholder=\\"Reason (optional)\\" onkeydown=\\"if(event.key===\'Enter\')icRejectDo(\'"+rid+"\',\'"+token+"\')\\"><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"icRejectDo(\'"+rid+"\',\'"+token+"\')\\">Send back</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"icRejectCancel(\'"+rid+"\',\'"+token+"\')\\">Cancel</button>";var i=document.getElementById(rid+"-reason");if(i)i.focus();}'
    + 'function icRejectCancel(rid,token){document.getElementById(rid+"-act").innerHTML=icAdminFootJs(rid,token);}'
    + 'function icRejectDo(rid,token){var act=document.getElementById(rid+"-act");var reason=(document.getElementById(rid+"-reason")||{}).value||"";act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'icSetPill(rid,"tp-pill--sent","Sent back");var s=document.getElementById(rid+"-status");if(s){s.textContent="Sent back";s.className="due";}'
    + 'var nt=document.getElementById(rid+"-note");if(nt)nt.innerHTML=icSentBackInner(reason);'   // photo left in place
    + 'act.innerHTML=icOpenFootJs(rid,token);var c=document.getElementById(rid);if(c)c.dataset.state="open";icBump("sum-pending",-1);icBump("sum-open",1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).rejectIssueCompletion(token,reason,ADMIN_PASS);}'
    + '</script>';
}

// Admin-only Delete control (hidden until unlock).
function tpDelWrap_(rid, pid) {
  return '<span id="' + rid + '-delwrap" class="tp-admin" hidden><button type="button" class="btn btn-ghost tp-del" onclick="tpDelOpen(\'' + rid + '\',\'' + pid + '\')">Delete</button></span>';
}

// Admin-only Approve / Send back for a project awaiting approval (hidden until unlock).
function tpProjPendingFoot_(rid, pid) {
  return '<span class="tp-admin btn-row" hidden>'
    + '<button type="button" class="btn btn-confirm" onclick="tpProjApprove(\'' + rid + '\',\'' + pid + '\')">Approve</button>'
    + '<button type="button" class="btn btn-ghost" onclick="tpProjRejectOpen(\'' + rid + '\',\'' + pid + '\')">Send back</button>'
    + '</span>';
}

// The after-photo completion uploader shown when a doer taps Complete.
function tpProjUploader_(rid, pid) {
  return '<div id="' + rid + '-uploader" class="tp-uploader" style="display:none">'
    + '<div class="tp-hint" style="margin-bottom:10px">Add an <b>after</b> photo and the <b>hours</b> it took, then submit for approval.</div>'
    + '<div class="tp-drop">'
    + '<div class="tp-slot" id="' + rid + '-slot-a"><input type="file" accept="image/*" id="' + rid + '-after" style="display:none" onchange="tpSlot(this,\'' + rid + '\',\'a\')"><div class="tp-slot-btn" onclick="document.getElementById(\'' + rid + '-after\').click()"><span id="' + rid + '-alabel">Add “after” photo</span></div></div>'
    + '<div class="tp-hours-field"><label class="tp-hours-label">Hours it took</label><input type="number" min="0.5" step="0.5" id="' + rid + '-hours" class="tp-hours-input" placeholder="e.g. 6"></div>'
    + '</div>'
    + '<div style="margin-top:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">'
    + '<button type="button" class="btn btn-confirm" id="' + rid + '-finish" onclick="tpComplete(\'' + rid + '\',\'' + pid + '\')">Submit for approval</button>'
    + '<button type="button" class="btn btn-ghost" onclick="tpCompleteClose(\'' + rid + '\')">Cancel</button>'
    + '<span id="' + rid + '-cmsg" class="tp-hint"></span>'
    + '</div></div>';
}

// Admin-only edit form for one project card (hidden until Edit is tapped).
function tpEditForm_(rid, p) {
  return '<div id="' + rid + '-edit" class="ic-edit" hidden>'
    + '<label class="ic-edit-lbl">Title<input id="' + rid + '-etitle" class="ic-edit-in" value="' + escapeHtml_(p.title || '') + '"></label>'
    + '<label class="ic-edit-lbl">Scope<textarea id="' + rid + '-edesc" class="ic-edit-in" rows="3">' + escapeHtml_(p.description || '') + '</textarea></label>'
    + '<label class="ic-edit-lbl">Assignees (comma-separated)<input id="' + rid + '-eassignees" class="ic-edit-in" value="' + escapeHtml_(p.assignees.join(', ')) + '"></label>'
    + '<div class="ic-edit-btns"><button type="button" class="btn btn-primary" onclick="tpEditSave(\'' + rid + '\',\'' + p.id + '\')">Save changes</button>'
    + '<button type="button" class="btn btn-ghost" onclick="tpEditCancel(\'' + rid + '\')">Cancel</button>'
    + '<span id="' + rid + '-emsg" class="ic-edit-msg"></span></div>'
    + '</div>';
}

// Admin-only Edit button (revealed with the other .tp-admin controls).
function tpEditBtn_(rid) {
  return '<span class="tp-admin" hidden><button type="button" class="btn btn-ghost" onclick="tpEditOpen(\'' + rid + '\')">Edit</button></span>';
}

// Reference material for a project card: an uploaded file and/or an external link,
// as chips. Empty when the project has neither.
function tpRefsHtml_(p) {
  var refs = '';
  if (p.attachment) refs += '<a class="tp-attach" href="' + escapeHtml_(p.attachment) + '" target="_blank" rel="noopener">&#128206; File</a>';
  if (p.link) refs += '<a class="tp-attach" href="' + escapeHtml_(p.link) + '" target="_blank" rel="noopener">&#128279; Link</a>';
  if (!refs) return '';
  return '<div class="card-field"><span class="card-flabel">Attachments</span><div class="tp-refs">' + refs + '</div></div>';
}

// One project card. Extracted so the initial page and the post-create refresh
// (tpProjectsListHtml) render identically. Handles Assigned / In Progress (incl.
// sent-back) / Pending approval / Completed.
function tpRenderProjectCard_(p, rid) {
  var st = norm_(p.status);
  var isDone = st === norm_(TP.projectStatus.done);
  var isPending = st === norm_(TP.projectStatus.pending);
  var isActive = st === norm_(TP.projectStatus.active);
  var sentBack = isActive && !!p.sentBackReason;

  var pill = isDone ? '<span class="tp-pill tp-pill--done">Completed</span>'
    : isPending ? '<span class="tp-pill tp-pill--pending">Pending approval</span>'
    : isActive ? '<span class="tp-pill tp-pill--active">In progress</span>'
    : '<span class="tp-pill tp-pill--assigned">Assigned</span>';

  var chips = p.assignees.length
    ? p.assignees.map(function (a) { return '<span class="tp-chip">' + escapeHtml_(a) + '</span>'; }).join('')
    : '<span class="tp-chip tp-chip--empty">No one yet</span>';

  var accentHex = isDone ? '#157a47' : isPending ? '#b06a00' : isActive ? '#d97a12' : '#2563c9';

  var photoInner = '';
  if (p.before) photoInner += '<div class="tp-photo"><figure><figcaption>Before</figcaption>' + tpThumb_(p.before) + '</figure></div>';
  if (p.after && (isPending || isDone || sentBack)) photoInner += '<div class="tp-photo"><figure><figcaption>After</figcaption>' + tpThumb_(p.after) + '</figure></div>';
  var photos = '<div id="' + rid + '-photos">' + (photoInner ? '<div class="tp-photos">' + photoInner + '</div>' : '') + '</div>';

  var hoursField = ((isPending || isDone) && p.hours)
    ? '<div class="card-field"><span class="card-flabel">Hours</span><div class="card-action">' + escapeHtml_(tpHoursLabel_(p.hours)) + '</div></div>' : '';

  var body = '<div class="card-body">'
    + '<div id="' + rid + '-view">'
    + '<div class="card-head"><div><div class="card-team">Project</div>'
    + '<div class="card-title" id="' + rid + '-vtitle">' + escapeHtml_(p.title) + '</div></div>'
    + '<span id="' + rid + '-pill">' + pill + '</span></div>'
    + '<div class="card-field" id="' + rid + '-vscope-wrap"' + (p.description ? '' : ' hidden') + '><span class="card-flabel">Scope</span><div class="card-details" id="' + rid + '-vscope">' + escapeHtml_(p.description) + '</div></div>'
    + '<div class="card-field"><span class="card-flabel">Assignees</span><div class="tp-assignees" id="' + rid + '-chips">' + chips + '</div></div>'
    + hoursField
    + tpRefsHtml_(p)
    + '<div id="' + rid + '-note">' + (sentBack ? icSentBackInner_(p.sentBackReason) : '') + '</div>'
    + photos
    + '</div>'
    + tpEditForm_(rid, p)
    + '</div>';

  var foot;
  if (isDone) {
    foot = '<div class="card-foot"><span id="' + rid + '-status" class="due due--done">✓ Completed' + (p.hours ? ' &middot; ' + escapeHtml_(tpHoursLabel_(p.hours)) : '') + '</span>'
      + '<span class="btn-row">' + tpEditBtn_(rid) + tpDelWrap_(rid, p.id) + '</span></div>';
  } else if (isPending) {
    foot = '<div class="card-foot"><span id="' + rid + '-status" class="due">Submitted, awaiting approval</span>'
      + '<span id="' + rid + '-act" class="btn-row">' + tpProjPendingFoot_(rid, p.id) + '</span>'
      + '<span class="btn-row">' + tpEditBtn_(rid) + tpDelWrap_(rid, p.id) + '</span></div>';
  } else {
    foot = '<div class="card-foot"><span id="' + rid + '-status" class="due">' + (isActive ? 'Work in progress' : 'Waiting to be picked up') + '</span>'
      + '<span class="btn-row">'
      + '<span id="' + rid + '-join"><button type="button" class="btn btn-ghost" onclick="tpJoinOpen(\'' + rid + '\')">Join project</button></span>'
      + '<span id="' + rid + '-joinbox" class="tp-inline-join" style="display:none"><input id="' + rid + '-name" placeholder="Your name" onkeydown="if(event.key===\'Enter\')tpJoin(\'' + rid + '\',\'' + p.id + '\')"><button type="button" class="btn btn-primary" onclick="tpJoin(\'' + rid + '\',\'' + p.id + '\')">Join</button></span>'
      + '<button type="button" class="btn btn-primary" onclick="tpCompleteOpen(\'' + rid + '\')">Complete</button>'
      + tpEditBtn_(rid) + tpDelWrap_(rid, p.id)
      + '</span></div>'
      + tpProjUploader_(rid, p.id);
  }

  return '<div class="card" id="' + rid + '" data-status="' + escapeHtml_(p.status) + '"><div class="card-accent" id="' + rid + '-accent" style="background:' + accentHex + '"></div>' + body + foot + '</div>';
}

// The grouped Assigned / In progress / Completed sections for #tp-proj-list.
function tpProjectsSectionsHtml_(projects) {
  if (!projects) projects = tpListProjects_();
  var assigned = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.assigned) || (!p.status); });
  var active = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.active); });
  var pending = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.pending); });
  var done = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.done); });
  if (!projects.length) return '<div class="empty">No projects yet.</div>';
  var sectionHead = function (label, count, cls) {
    return '<div class="section-head"><span class="section-label ' + cls + '">' + label + '</span>'
      + '<span class="section-count">' + count + '</span><span class="section-rule"></span></div>';
  };
  var idx = 0;
  var render = function (p) { return tpRenderProjectCard_(p, 'pj' + (idx++)); };
  var out = '';
  if (pending.length) { out += sectionHead('Pending approval', pending.length, 'section-label--late'); pending.forEach(function (p) { out += render(p); }); }
  if (active.length) { out += sectionHead('In progress', active.length, 'section-label--late'); active.forEach(function (p) { out += render(p); }); }
  if (assigned.length) { out += sectionHead('Assigned', assigned.length, 'section-label--open'); assigned.forEach(function (p) { out += render(p); }); }
  if (done.length) { out += sectionHead('Completed', done.length, 'section-label--open'); done.forEach(function (p) { out += render(p); }); }
  return out;
}

// Client-callable: re-fetch the project list HTML so a create/refresh needs no reload.
function tpProjectsListHtml() { return tpProjectsSectionsHtml_(); }

// ---- Projects dashboard ----

function tpProjectStats_() {
  var projects = tpListProjects_();
  var byStatus = { assigned: 0, active: 0, pending: 0, done: 0 };
  var totalHours = 0, completed = [], byPerson = {};
  projects.forEach(function (p) {
    var st = norm_(p.status);
    if (st === norm_(TP.projectStatus.done)) { byStatus.done++; totalHours += (p.hours || 0); completed.push(p); }
    else if (st === norm_(TP.projectStatus.pending)) byStatus.pending++;
    else if (st === norm_(TP.projectStatus.active)) byStatus.active++;
    else byStatus.assigned++;
    p.assignees.forEach(function (a) { var k = String(a).trim(); if (k) byPerson[k] = (byPerson[k] || 0) + 1; });
  });
  completed.sort(function (a, b) { return (new Date(b.completedAt || 0)) - (new Date(a.completedAt || 0)); });
  var people = Object.keys(byPerson).map(function (k) { return { name: k, count: byPerson[k] }; })
    .sort(function (a, b) { return b.count - a.count; });
  return { total: projects.length, assigned: byStatus.assigned, active: byStatus.active, pending: byStatus.pending, done: byStatus.done, totalHours: totalHours, completed: completed, people: people };
}

function tpDashStyles_() {
  return '<style>'
    + '.dash{max-width:1040px;margin:0 auto}'
    + '.dash-hero{position:relative;overflow:hidden;border-radius:22px;padding:34px 34px 30px;color:#fff;background:radial-gradient(120% 140% at 0% 0%,#d62b2b 0%,#8f1515 46%,#5c0d0d 100%);box-shadow:0 24px 60px rgba(143,21,21,.34)}'
    + '.dash-hero-glow{position:absolute;right:-80px;top:-90px;width:320px;height:320px;background:radial-gradient(circle,rgba(240,192,80,.55),transparent 65%);pointer-events:none}'
    + '.dash-kicker{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.7)}'
    + '.dash-h1{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:38px;font-weight:800;letter-spacing:-.04em;line-height:1;margin:10px 0 0}'
    + '.dash-hero-row{position:relative;display:flex;align-items:flex-end;gap:34px;margin-top:26px;flex-wrap:wrap}'
    + '.dash-hero-num{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:76px;font-weight:800;letter-spacing:-.05em;line-height:.9;background:linear-gradient(180deg,#fff,#ffe6a8);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}'
    + '.dash-hero-lbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:rgba(255,255,255,.78);margin-top:6px}'
    + '.dash-hero-mini{display:flex;gap:26px;padding-bottom:8px}'
    + '.dash-hero-mini>div{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:600;color:rgba(255,255,255,.82)}'
    + '.dash-hero-mini b{display:block;font-size:26px;font-weight:800;letter-spacing:-.03em;color:#fff}'
    + '.dash-tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:16px}'
    + '.dash-tile{position:relative;overflow:hidden;background:#fff;border:1.5px solid #ececea;border-radius:16px;padding:18px 18px 16px;box-shadow:0 4px 16px rgba(20,20,30,.06)}'
    + '.dash-tile::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--c)}'
    + '.dash-tile-num{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:34px;font-weight:800;letter-spacing:-.04em;color:#14110e;line-height:1}'
    + '.dash-tile-lbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c;margin-top:8px}'
    + '.dash-grid{display:grid;grid-template-columns:300px 1fr;gap:14px;margin-top:14px}'
    + '.dash-card{background:#fff;border:1.5px solid #ececea;border-radius:18px;padding:22px;box-shadow:0 4px 16px rgba(20,20,30,.05)}'
    + '.dash-card-h{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c;margin:0 0 16px}'
    + '.dash-ringwrap{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px}'
    + '.dash-ring{width:190px;height:190px}'
    + '.dash-ring-bg{fill:none;stroke:#f0efe9;stroke-width:3.6}'
    + '.dash-ring-fg{fill:none;stroke:url(#dashgrad);stroke-width:3.6;stroke-linecap:round;transition:stroke-dashoffset 1.2s cubic-bezier(.2,.8,.2,1)}'
    + '.dash-ring-num{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:7px;font-weight:800;fill:#14110e}'
    + '.dash-ring-sub{font-size:3px;letter-spacing:.4px;fill:#9a958c}'
    + '.dash-ring-cap{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#57534e}'
    + '.dash-bars{display:flex;flex-direction:column;gap:14px}'
    + '.dash-bar-row{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center}'
    + '.dash-bar-name{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#26231f;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}'
    + '.dash-bar-val{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;color:#8f1515;white-space:nowrap}'
    + '.dash-bar-track{grid-column:1 / -1;height:12px;border-radius:99px;background:#f2f1ec;overflow:hidden}'
    + '.dash-bar-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,#8f1515,#d62b2b 60%,#f0c050);transform-origin:left;transform:scaleX(var(--w));animation:dashGrow 1.1s cubic-bezier(.2,.8,.2,1) both}'
    + '@keyframes dashGrow{from{transform:scaleX(0)}}'
    + '.dash-ava{flex:0 0 auto;width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:800;color:#fff}'
    + '.dash-chips{display:flex;flex-wrap:wrap;gap:8px}'
    + '.dash-chip{display:inline-flex;align-items:center;gap:8px;background:#faf9f6;border:1.5px solid #ececea;border-radius:99px;padding:4px 6px 4px 4px}'
    + '.dash-chip .dash-ava{width:26px;height:26px;font-size:11px}'
    + '.dash-chip-name{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#26231f}'
    + '.dash-chip-count{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;color:#8f1515;background:#fdecec;border:1px solid #f5d0d0;border-radius:99px;padding:1px 8px}'
    + '.dash-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px}'
    + '.dash-gal-item{border:1.5px solid #eee;border-radius:14px;overflow:hidden;background:#fcfcfb}'
    + '.dash-ba{display:grid;grid-template-columns:1fr 1fr;gap:2px;background:#eee}'
    + '.dash-ba figure{margin:0;position:relative;aspect-ratio:4/3;background:#eceae4}'
    + '.dash-ba img{width:100%;height:100%;object-fit:cover;display:block}'
    + '.dash-ba figcaption{position:absolute;left:6px;top:6px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:8.5px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#fff;background:rgba(20,17,14,.66);padding:2px 7px;border-radius:99px}'
    + '.dash-gal-body{padding:12px 14px 14px}'
    + '.dash-gal-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:800;letter-spacing:-.01em;color:#14110e;line-height:1.25}'
    + '.dash-gal-meta{display:flex;align-items:center;gap:8px;margin-top:7px;flex-wrap:wrap}'
    + '.dash-hchip{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;letter-spacing:.04em;color:#8f1515;background:#fdecec;border:1px solid #f5d0d0;border-radius:99px;padding:3px 9px}'
    + '.dash-gal-who{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11.5px;font-weight:600;color:#8a857c}'
    + '.dash-empty{background:#fff;border:1.5px dashed #ddd;border-radius:16px;padding:30px;text-align:center;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;color:#8a857c;margin-top:14px}'
    + '.dash-sec{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#8a857c;margin:26px 0 12px}'
    + '@media(max-width:760px){.dash-tiles{grid-template-columns:repeat(2,1fr)}.dash-grid{grid-template-columns:1fr}.dash-h1{font-size:30px}.dash-hero-num{font-size:60px}.dash-hero{padding:26px 22px 24px}.dash-hero-row{gap:22px}}'
    + '@media(max-width:440px){.dash-hero-num{font-size:50px}.dash-hero-mini{gap:16px}.dash-hero-mini b{font-size:22px}.dash-card{padding:18px 16px}.dash-ring{width:158px;height:158px}.dash-tile{padding:15px 14px}.dash-tile-num{font-size:28px}}'
    + '</style>';
}

// Projects dashboard: hero hours counter, status tiles, completion ring,
// hours-by-project bars, a compact contributor row, and a before/after gallery.
function projectsDashboardPage_(embedded) {
  var s = tpProjectStats_();
  var pct = s.total > 0 ? Math.round(s.done / s.total * 100) : 0;
  var AVA = ['#b31b1b', '#e08a1e', '#157a47', '#2563c9', '#7c3aed', '#0d9488', '#b06a00'];
  var initials = function (n) { var parts = String(n).trim().split(/\s+/); return ((parts[0] || '')[0] || '?').toUpperCase() + (parts.length > 1 ? (parts[parts.length - 1][0] || '').toUpperCase() : ''); };

  var inner = tpDashStyles_() + '<div class="dash">';

  inner += '<div class="dash-hero"><div class="dash-hero-glow"></div>'
    + '<div class="dash-kicker">Project Teams Ops Hub</div>'
    + '<h1 class="dash-h1">Projects Dashboard</h1>'
    + '<div class="dash-hero-row">'
    +   '<div><div class="dash-hero-num">' + (Math.round(s.totalHours * 10) / 10) + '</div><div class="dash-hero-lbl">hours logged across ' + s.total + ' project' + (s.total === 1 ? '' : 's') + '</div></div>'
    + '</div></div>';

  // Tiles are the status breakdown (hours live in the hero, completion % in the ring).
  var tile = function (n, lbl, c) { return '<div class="dash-tile" style="--c:' + c + '"><div class="dash-tile-num">' + n + '</div><div class="dash-tile-lbl">' + lbl + '</div></div>'; };
  inner += '<div class="dash-tiles">'
    + tile(s.assigned, 'Assigned', '#2563c9')
    + tile(s.active, 'In progress', '#e08a1e')
    + tile(s.pending, 'Pending approval', '#b06a00')
    + tile(s.done, 'Completed', '#157a47')
    + '</div>';

  if (!s.total) {
    inner += '<div class="dash-empty">No projects yet. Assign a project and complete a few to light this up.</div></div>';
    return swissShell_(inner, 'Projects Dashboard', true, embedded);
  }

  // Completion ring + hours-by-project
  var offset = 100 - pct;
  var ring = '<div class="dash-card"><div class="dash-card-h">Completion</div><div class="dash-ringwrap">'
    + '<svg class="dash-ring" viewBox="0 0 42 42" aria-hidden="true">'
    +   '<defs><linearGradient id="dashgrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8f1515"/><stop offset="0.6" stop-color="#d62b2b"/><stop offset="1" stop-color="#f0c050"/></linearGradient></defs>'
    +   '<circle class="dash-ring-bg" cx="21" cy="21" r="15.915"/>'
    +   '<circle class="dash-ring-fg" cx="21" cy="21" r="15.915" stroke-dasharray="100" stroke-dashoffset="100" data-off="' + offset + '" transform="rotate(-90 21 21)"/>'
    +   '<text class="dash-ring-num" x="21" y="20.4" text-anchor="middle">' + pct + '%</text>'
    +   '<text class="dash-ring-sub" x="21" y="26" text-anchor="middle">COMPLETE</text>'
    + '</svg>'
    + '</div></div>';

  var hoursProjects = s.completed.filter(function (p) { return p.hours > 0; }).slice().sort(function (a, b) { return b.hours - a.hours; }).slice(0, 7);
  var maxH = hoursProjects.reduce(function (m, p) { return Math.max(m, p.hours); }, 0) || 1;
  var bars = '<div class="dash-card"><div class="dash-card-h">Hours by project</div>';
  if (!hoursProjects.length) {
    bars += '<div class="dash-gal-who">No completed projects with hours yet.</div>';
  } else {
    bars += '<div class="dash-bars">';
    hoursProjects.forEach(function (p) {
      bars += '<div><div class="dash-bar-row"><span class="dash-bar-name">' + escapeHtml_(p.title) + '</span><span class="dash-bar-val">' + escapeHtml_(tpHoursLabel_(p.hours)) + '</span></div>'
        + '<div class="dash-bar-track"><div class="dash-bar-fill" style="--w:' + (p.hours / maxH).toFixed(3) + '"></div></div></div>';
    });
    bars += '</div>';
  }
  bars += '</div>';

  inner += '<div class="dash-grid">' + ring + bars + '</div>';

  // Contributors: a compact chip row (name + project count), kept small on purpose.
  if (s.people.length) {
    inner += '<div class="dash-sec">Contributors</div><div class="dash-chips">';
    s.people.forEach(function (pp, i) {
      inner += '<span class="dash-chip">'
        + '<span class="dash-ava" style="background:' + AVA[i % AVA.length] + '">' + escapeHtml_(initials(pp.name)) + '</span>'
        + '<span class="dash-chip-name">' + escapeHtml_(pp.name) + '</span>'
        + '<span class="dash-chip-count">' + pp.count + '</span>'
        + '</span>';
    });
    inner += '</div>';
  }

  // Recent completions gallery (before/after)
  var gal = s.completed.filter(function (p) { return p.before || p.after; }).slice(0, 6);
  if (gal.length) {
    inner += '<div class="dash-sec">Recent completions</div><div class="dash-gal">';
    gal.forEach(function (p) {
      var b = extractFileIds_(p.before)[0], a = extractFileIds_(p.after)[0];
      var img = function (id, cap) { return '<figure>' + (id ? '<img loading="lazy" src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w400">' : '') + '<figcaption>' + cap + '</figcaption></figure>'; };
      var who = p.assignees.length ? p.assignees.join(', ') : '';
      inner += '<div class="dash-gal-item"><div class="dash-ba">' + img(b, 'Before') + img(a, 'After') + '</div>'
        + '<div class="dash-gal-body"><div class="dash-gal-title">' + escapeHtml_(p.title) + '</div>'
        + '<div class="dash-gal-meta">' + (p.hours ? '<span class="dash-hchip">' + escapeHtml_(tpHoursLabel_(p.hours)) + '</span>' : '') + (who ? '<span class="dash-gal-who">' + escapeHtml_(who) + '</span>' : '') + '</div>'
        + '</div></div>';
    });
    inner += '</div>';
  }

  inner += '</div>';   // .dash
  inner += '<script>requestAnimationFrame(function(){var c=document.querySelector(".dash-ring-fg");if(c)c.style.strokeDashoffset=c.getAttribute("data-off");});</script>';

  return swissShell_(inner, 'Projects Dashboard', true, embedded);
}

// ---- Projects page ----

function projectsPage_(embedded, admin) {
  var projects = tpListProjects_();
  var assigned = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.assigned) || (!p.status); });
  var active = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.active); });
  var pending = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.pending); });
  var done = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.done); });

  var inner = '';
  if (!embedded) {
    inner += '<div class="page-head"><div class="page-kicker">Project Teams Ops Hub</div>'
      + '<div class="page-title">Projects</div><div class="page-rule"></div></div>';
  }

  // Admin create form (revealed only in admin mode via tpAdminRevealJs_).
  // Collapsed into a disclosure so it stays out of the way until needed.
  inner += '<details class="tp-create-wrap tp-admin" hidden>'
    + '<summary class="tp-create-toggle"><span class="tp-create-caret">&#43;</span> New project</summary>'
    + '<form class="tp-create" onsubmit="return tpCreate(event)">'
    + '<div class="tp-field"><label>Title <span class="tp-req">Required</span></label><input id="tp-c-title" placeholder="e.g. Rebuild the tool crib shelving" required></div>'
    + '<div class="tp-field"><label>Description</label><textarea id="tp-c-desc" rows="2" placeholder="Scope, location, and what done looks like"></textarea></div>'
    + '<div class="tp-field"><label>Assignees</label><input id="tp-c-assignees" placeholder="Alex Rivera, Sam Chen (comma-separated)"></div>'
    + '<div class="tp-field"><label>Reference material <span class="tp-opt">Optional</span></label>'
    +   '<div class="tp-attach-row">'
    +     '<input type="file" accept="image/*" id="tp-c-before" style="display:none" onchange="tpCreateSlot(this)">'
    +     '<button type="button" class="tp-attach-btn" id="tp-c-slot-b" onclick="document.getElementById(\'tp-c-before\').click()"><span aria-hidden="true">&#128247;</span> <span id="tp-c-blabel">Before photo</span></button>'
    +     '<input type="file" id="tp-c-file" style="display:none" onchange="tpCreateFile(this)">'
    +     '<button type="button" class="tp-attach-btn" id="tp-c-slot-f" onclick="document.getElementById(\'tp-c-file\').click()"><span aria-hidden="true">&#128206;</span> <span id="tp-c-flabel">File</span></button>'
    +   '</div>'
    +   '<div class="tp-attach-link"><span class="tp-link-icon" aria-hidden="true">&#128279;</span><input id="tp-c-link" type="url" placeholder="Paste a link: spreadsheet, doc, or Drive folder"></div>'
    + '</div>'
    + '<button type="submit" class="btn btn-primary">Create project</button>'
    + '<span id="tp-c-msg" class="tp-lock-msg" style="margin-left:10px"></span>'
    + '</form>'
    + '</details>';

  inner += '<div class="ic-summary">'
    + '<span class="ic-sum"><b id="tp-n-assigned">' + assigned.length + '</b> assigned</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum"><b id="tp-n-active">' + active.length + '</b> in progress</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum"><b id="tp-n-pending">' + pending.length + '</b> pending</span>'
    + '<span class="ic-dot"></span>'
    + '<span class="ic-sum"><b id="tp-n-done">' + done.length + '</b> completed</span>'
    + '</div>';

  inner += '<div id="tp-proj-list">' + tpProjectsSectionsHtml_(projects) + '</div>';

  inner += '<script>'
    + 'var TPUP={};var TPCB=null;var TPCF=null;'
    + 'function tpCreateSlot(input){tpReadFile(input,function(res){if(!res)return;TPCB=res;var s=document.getElementById("tp-c-slot-b");if(s)s.classList.add("is-set");var l=document.getElementById("tp-c-blabel");if(l)l.textContent="\\u2713 Before: "+res.name;});}'
    + 'function tpCreateFile(input){tpReadAnyFile(input,function(res){if(!res)return;TPCF=res;var s=document.getElementById("tp-c-slot-f");if(s)s.classList.add("is-set");var l=document.getElementById("tp-c-flabel");if(l)l.textContent="\\u2713 "+res.name;});}'
    + 'function tpBump(id,d){var e=document.getElementById(id);if(e)e.textContent=Math.max(0,(parseInt(e.textContent,10)||0)+d);}'
    + 'function tpSetPill(rid,cls,txt){document.getElementById(rid+"-pill").innerHTML="<span class=\\"tp-pill "+cls+"\\">"+txt+"</span>";}'
    + 'function tpDelOpen(rid,pid){document.getElementById(rid+"-delwrap").innerHTML="<span class=\\"tp-hint\\" style=\\"margin-right:6px\\">Delete this project?</span><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"tpDelDo(\'"+rid+"\',\'"+pid+"\')\\">Yes, delete</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"tpDelCancel(\'"+rid+"\',\'"+pid+"\')\\">Cancel</button>";}'
    + 'function tpDelCancel(rid,pid){document.getElementById(rid+"-delwrap").innerHTML="<button type=\\"button\\" class=\\"btn btn-ghost tp-del\\" onclick=\\"tpDelOpen(\'"+rid+"\',\'"+pid+"\')\\">Delete</button>";}'
    + 'function tpDelDo(rid,pid){var w=document.getElementById(rid+"-delwrap");w.innerHTML="<span class=\\"tp-hint\\">Deleting\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){w.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'var card=document.getElementById(rid);var st=(card.dataset.status||"").toLowerCase();'
    + 'if(st==="completed")tpBump("tp-n-done",-1);else if(st==="in progress")tpBump("tp-n-active",-1);else tpBump("tp-n-assigned",-1);'
    + 'card.parentNode.removeChild(card);'
    + '}).withFailureHandler(function(){w.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).tpDeleteProject(pid,ADMIN_PASS);}'
    + 'function tpJoinOpen(rid){document.getElementById(rid+"-join").style.display="none";document.getElementById(rid+"-joinbox").style.display="inline-flex";document.getElementById(rid+"-name").focus();}'
    + 'function tpJoin(rid,pid){var nm=document.getElementById(rid+"-name").value.trim();if(!nm){document.getElementById(rid+"-name").focus();return;}'
    + 'var box=document.getElementById(rid+"-joinbox");box.innerHTML="<span class=\\"tp-hint\\">Joining\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){box.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'var chips=document.getElementById(rid+"-chips");if(chips)chips.innerHTML=r.assignees.length?r.assignees.map(function(a){return "<span class=\\"tp-chip\\">"+String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;")+"</span>";}).join(""):"<span class=\\"tp-chip tp-chip--empty\\">No one yet</span>";'
    + 'var card=document.getElementById(rid);var was=card.dataset.status;card.dataset.status=r.status;'
    + 'if((was||"").toLowerCase()!=="in progress"&&r.status==="In Progress"){tpSetPill(rid,"tp-pill--active","In progress");document.getElementById(rid+"-status").textContent="Work in progress";tpAdvance(rid,"#d97a12","#fdf1e0",1);tpBump("tp-n-assigned",-1);tpBump("tp-n-active",1);}'
    + 'document.getElementById(rid+"-joinbox").style.display="none";document.getElementById(rid+"-join").style.display="inline";'
    + 'document.getElementById(rid+"-join").innerHTML="<span class=\\"tp-hint\\">\\u2713 You\'re on this project</span>";'
    + '}).withFailureHandler(function(){box.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).tpJoinProject(pid,nm);}'
    + 'function tpEditOpen(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),c=document.getElementById(rid);if(v)v.hidden=true;if(e)e.hidden=false;var f=c&&c.querySelector(".card-foot");if(f)f.style.display="none";}'
    + 'function tpEditCancel(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),m=document.getElementById(rid+"-emsg"),c=document.getElementById(rid);if(e)e.hidden=true;if(v)v.hidden=false;if(m)m.textContent="";var f=c&&c.querySelector(".card-foot");if(f)f.style.display="";}'
    + 'function tpEditSave(rid,pid){var g=function(s){var el=document.getElementById(rid+s);return el?el.value:"";};var title=g("-etitle"),desc=g("-edesc"),asg=g("-eassignees");var m=document.getElementById(rid+"-emsg");if(m){m.style.color="";m.textContent="Saving\\u2026";}'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){if(m){m.style.color="#b31b1b";m.textContent=(r&&r.error)||"Could not save";}return;}'
    + 'var esc=function(s){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;");};'
    + 'var vt=document.getElementById(rid+"-vtitle");if(vt)vt.textContent=r.title;'
    + 'var sw=document.getElementById(rid+"-vscope-wrap");if(sw){if(r.description){sw.hidden=false;var sv=document.getElementById(rid+"-vscope");if(sv)sv.textContent=r.description;}else sw.hidden=true;}'
    + 'var chips=document.getElementById(rid+"-chips");if(chips)chips.innerHTML=r.assignees.length?r.assignees.map(function(a){return "<span class=\\"tp-chip\\">"+String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;")+"</span>";}).join(""):"<span class=\\"tp-chip tp-chip--empty\\">No one yet</span>";'
    + 'tpEditCancel(rid);'
    + '}).withFailureHandler(function(){if(m){m.style.color="#b31b1b";m.textContent="Could not save. Retry.";}}).tpUpdateProject(pid,title,desc,asg);}'
    + 'function tpCompleteOpen(rid){document.getElementById(rid+"-uploader").style.display="block";}'
    + 'function tpCompleteClose(rid){document.getElementById(rid+"-uploader").style.display="none";}'
    + 'function tpSlot(input,rid,which){tpReadFile(input,function(res){if(!res)return;TPUP[rid]=TPUP[rid]||{};TPUP[rid][which]=res;'
    + 'var slot=document.getElementById(rid+"-slot-"+which);slot.classList.add("is-set");'
    + 'document.getElementById(rid+"-"+(which==="b"?"blabel":"alabel")).textContent=(which==="b"?"\\u2713 Before: ":"\\u2713 After: ")+res.name;});}'
    + 'function tpComplete(rid,pid){var buf=TPUP[rid]||{};var msg=document.getElementById(rid+"-cmsg");'
    + 'if(!buf.a){msg.style.color="#b31b1b";msg.textContent="Add an after photo first.";return;}'
    + 'var hrs=parseFloat((document.getElementById(rid+"-hours")||{}).value);if(!(hrs>0)){msg.style.color="#b31b1b";msg.textContent="Enter how many hours it took.";return;}'
    + 'msg.style.color="";msg.textContent="Uploading photo\\u2026";document.getElementById(rid+"-finish").disabled=true;'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.style.color="#b31b1b";msg.textContent=(r&&r.error)||"Failed";document.getElementById(rid+"-finish").disabled=false;return;}'
    + 'tpConfetti();tpSetPill(rid,"tp-pill--pending","Pending approval");tpAdvance(rid,"#b06a00","#f7edd8",2);'
    + 'var bh=r.beforeId?"<div class=\\"tp-photo\\"><figure><figcaption>Before</figcaption><a href=\\"https://drive.google.com/file/d/"+r.beforeId+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+r.beforeId+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div>":"";'
    + 'document.getElementById(rid+"-photos").innerHTML="<div class=\\"tp-photos\\">"+bh+"<div class=\\"tp-photo\\"><figure><figcaption>After</figcaption><a href=\\"https://drive.google.com/file/d/"+r.afterId+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+r.afterId+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div></div>";'
    + 'var card=document.getElementById(rid);var was=(card.dataset.status||"").toLowerCase();'
    + 'var up=document.getElementById(rid+"-uploader");if(up)up.parentNode.removeChild(up);'
    + 'var foot=card.querySelector(".card-foot");if(foot)foot.innerHTML="<span id=\\""+rid+"-status\\" class=\\"due\\">Submitted, awaiting approval</span><span id=\\""+rid+"-act\\" class=\\"btn-row\\">"+tpProjAdminFootJs(rid,pid)+"</span><span class=\\"btn-row\\">"+tpDelWrapJs(rid,pid)+"</span>";'
    + 'if(was==="in progress")tpBump("tp-n-active",-1);else tpBump("tp-n-assigned",-1);tpBump("tp-n-pending",1);card.dataset.status="Pending";'
    + '}).withFailureHandler(function(){msg.style.color="#b31b1b";msg.textContent="Upload failed. Please retry.";document.getElementById(rid+"-finish").disabled=false;}).tpCompleteProject(pid,buf.a.dataUrl,buf.a.name,hrs);}'
    + 'function tpProjAdminFootJs(rid,pid){return ADMIN_PASS?"<span class=\\"tp-admin btn-row\\"><button type=\\"button\\" class=\\"btn btn-confirm\\" onclick=\\"tpProjApprove(\'"+rid+"\',\'"+pid+"\')\\">Approve</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"tpProjRejectOpen(\'"+rid+"\',\'"+pid+"\')\\">Send back</button></span>":"";}'
    + 'function tpDelWrapJs(rid,pid){return "<span id=\\""+rid+"-delwrap\\" class=\\"tp-admin\\""+(ADMIN_PASS?"":" hidden")+"><button type=\\"button\\" class=\\"btn btn-ghost tp-del\\" onclick=\\"tpDelOpen(\'"+rid+"\',\'"+pid+"\')\\">Delete</button></span>";}'
    + 'function tpProjApprove(rid,pid){var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'tpApproveFx(rid);tpAdvance(rid,"#157a47","#e7f3ec",3);tpSetPill(rid,"tp-pill--done","Completed");act.innerHTML="";var s=document.getElementById(rid+"-status");if(s){s.innerHTML="\\u2713 Completed";s.className="due due--done";}var c=document.getElementById(rid);if(c){c.style.opacity="0.72";c.dataset.status="Completed";}tpBump("tp-n-pending",-1);tpBump("tp-n-done",1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).tpApproveProject(pid,ADMIN_PASS);}'
    + 'function tpProjRejectOpen(rid,pid){var act=document.getElementById(rid+"-act");act.innerHTML="<input id=\\""+rid+"-reason\\" class=\\"ic-reason\\" placeholder=\\"Reason (optional)\\" onkeydown=\\"if(event.key===\'Enter\')tpProjRejectDo(\'"+rid+"\',\'"+pid+"\')\\"><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"tpProjRejectDo(\'"+rid+"\',\'"+pid+"\')\\">Send back</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"tpProjRejectCancel(\'"+rid+"\',\'"+pid+"\')\\">Cancel</button>";var i=document.getElementById(rid+"-reason");if(i)i.focus();}'
    + 'function tpProjRejectCancel(rid,pid){document.getElementById(rid+"-act").innerHTML=tpProjAdminFootJs(rid,pid);}'
    + 'function tpProjRejectDo(rid,pid){var act=document.getElementById(rid+"-act");var reason=(document.getElementById(rid+"-reason")||{}).value||"";act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'google.script.run.withSuccessHandler(function(html){document.getElementById("tp-proj-list").innerHTML=html;document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});tpBump("tp-n-pending",-1);tpBump("tp-n-active",1);}).withFailureHandler(function(){}).tpProjectsListHtml();'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).tpRejectProject(pid,reason,ADMIN_PASS);}'
    + 'function tpCreate(ev){ev.preventDefault();var t=document.getElementById("tp-c-title").value.trim();var d=document.getElementById("tp-c-desc").value.trim();var a=document.getElementById("tp-c-assignees").value.trim();var lk=(document.getElementById("tp-c-link")||{}).value||"";var msg=document.getElementById("tp-c-msg");'
    + 'if(!t){msg.className="tp-lock-msg bad";msg.textContent="A title is required.";return false;}'
    + 'msg.className="tp-lock-msg";msg.textContent="Creating\\u2026";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Failed";return;}'
    // Refresh the list in place (no location.reload: that reloads the sandboxed iframe to a blank page).
    + 'google.script.run.withSuccessHandler(function(html){document.getElementById("tp-proj-list").innerHTML=html;document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});tpBump("tp-n-assigned",1);'
    + 'document.getElementById("tp-c-title").value="";document.getElementById("tp-c-desc").value="";document.getElementById("tp-c-assignees").value="";TPCB=null;var sb=document.getElementById("tp-c-slot-b");if(sb)sb.classList.remove("is-set");var bl=document.getElementById("tp-c-blabel");if(bl)bl.textContent="Before photo";var bf=document.getElementById("tp-c-before");if(bf)bf.value="";'
    + 'TPCF=null;var sf=document.getElementById("tp-c-slot-f");if(sf)sf.classList.remove("is-set");var fl=document.getElementById("tp-c-flabel");if(fl)fl.textContent="File";var ff=document.getElementById("tp-c-file");if(ff)ff.value="";var lkf=document.getElementById("tp-c-link");if(lkf)lkf.value="";'
    + 'msg.className="tp-lock-msg ok";msg.textContent="\\u2713 Project created";'
    + '}).withFailureHandler(function(){msg.className="tp-lock-msg ok";msg.textContent="\\u2713 Created (refresh to see it)";}).tpProjectsListHtml();'
    + '}).withFailureHandler(function(){msg.className="tp-lock-msg bad";msg.textContent="Failed. Retry.";}).tpCreateProject(t,d,a,TPCB?TPCB.dataUrl:"",TPCB?TPCB.name:"",TPCF?TPCF.dataUrl:"",TPCF?TPCF.name:"",lk,ADMIN_PASS);return false;}'
    + '</script>';

  return swissShell_(tpStyles_() + inner + tpSharedJs_() + tpAdminRevealJs_(admin), 'Projects', true, embedded);
}
