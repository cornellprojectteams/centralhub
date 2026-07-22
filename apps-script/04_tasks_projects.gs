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
  projectHeaders: ['Project ID', 'Title', 'Description', 'Status', 'Assignees', 'Before photo', 'After photo', 'Started at', 'Completed at', 'Created at'],
  projectStatus: { assigned: 'Assigned', active: 'In Progress', done: 'Completed' },
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

// Open a sheet and return {sh, col:{Header->1-based index}}. ensureColumn_ appends
// any header that is missing, so the tabs are self-healing.
function tpOpen_(name, headers) {
  var sh = tpSheet_(name);
  var col = {};
  headers.forEach(function (h) { col[h] = ensureColumn_(sh, h); });
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

// Save a base64 data URL to Drive, share by link (so thumbnails render), return the file id.
function tpSaveUpload_(dataUrl, filename) {
  var m = /^data:([^;]+);base64,(.*)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('That upload was not a valid image.');
  var bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, m[1], filename || ('upload_' + Date.now()));
  var file = tpUploadsFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (err) { Logger.log('Upload share failed: ' + err); }
  return file.getId();
}

function tpUploadsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TP_UPLOADS_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) { /* recreate below */ } }
  var it = DriveApp.getFoldersByName(TP.uploadsFolderName);
  var f = it.hasNext() ? it.next() : DriveApp.createFolder(TP.uploadsFolderName);
  props.setProperty('TP_UPLOADS_FOLDER', f.getId());
  return f;
}

function tpViewUrl_(id) { return 'https://drive.google.com/file/d/' + id + '/view'; }

function tpThumb_(url, maxH) {
  var ids = extractFileIds_(url);
  if (!ids.length) return '';
  var e = encodeURIComponent(ids[0]);
  return '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" style="display:inline-block;line-height:0">'
    + '<img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w600" loading="lazy" alt="Uploaded photo" style="max-width:100%;max-height:' + (maxH || 200) + 'px;border-radius:10px;border:1px solid #ececec">'
    + '</a>';
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

// Doer submits completion evidence. A photo is REQUIRED -> Pending approval.
function submitIssueCompletion(token, dataUrl, filename) {
  if (!dataUrl) return { ok: false, error: 'A photo is required to mark this complete.' };
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  if (loc.addressed) return { ok: false, error: 'This was already approved as complete.' };
  var id;
  try { id = tpSaveUpload_(dataUrl, filename); }
  catch (err) { return { ok: false, error: String(err.message || err) }; }
  loc.sh.getRange(loc.row, loc.col(CONFIG.completionPhotoHeader)).setValue(tpViewUrl_(id));
  loc.sh.getRange(loc.row, loc.col(CONFIG.completedAtHeader)).setValue(new Date());
  return { ok: true, photoId: id, status: 'Pending' };
}

// Admin approves -> Completed (stamps Addressed at, which is what "resolved" reads).
function approveIssueCompletion(token, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Admin passcode required.' };
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  loc.sh.getRange(loc.row, loc.col(CONFIG.addressedHeader)).setValue(new Date());
  return { ok: true, status: 'Completed' };
}

// Admin sends back -> Uncompleted: clear the submission so the doer redoes it.
function rejectIssueCompletion(token, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Admin passcode required.' };
  var loc = icLocate_(token);
  if (!loc) return { ok: false, error: 'That item could not be found.' };
  loc.sh.getRange(loc.row, loc.col(CONFIG.completedAtHeader)).setValue('');
  loc.sh.getRange(loc.row, loc.col(CONFIG.completionPhotoHeader)).setValue('');
  loc.sh.getRange(loc.row, loc.col(CONFIG.addressedHeader)).setValue('');
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

// Manual completion. ENFORCES both a before and an after photo (open to assignees).
function tpCompleteProject(projectId, beforeUrl, afterUrl, beforeName, afterName) {
  if (!beforeUrl || !afterUrl) {
    return { ok: false, error: 'Both a "before" and an "after" photo are required to complete a project.' };
  }
  var o = tpOpen_(TP.projectsSheet, TP.projectHeaders);
  var r = tpFindRow_(o.sh, o.col['Project ID'], projectId);
  if (r < 0) return { ok: false, error: 'That project could not be found.' };
  var beforeId, afterId;
  try {
    beforeId = tpSaveUpload_(beforeUrl, beforeName || 'before');
    afterId = tpSaveUpload_(afterUrl, afterName || 'after');
  } catch (err) { return { ok: false, error: String(err.message || err) }; }
  o.sh.getRange(r, o.col['Before photo']).setValue(tpViewUrl_(beforeId));
  o.sh.getRange(r, o.col['After photo']).setValue(tpViewUrl_(afterId));
  o.sh.getRange(r, o.col['Status']).setValue(TP.projectStatus.done);
  o.sh.getRange(r, o.col['Completed at']).setValue(new Date());
  return { ok: true, status: TP.projectStatus.done, beforeId: beforeId, afterId: afterId };
}

// Admin creates/assigns a project.
function tpCreateProject(title, description, assignees, pass) {
  if (!tpIsAdmin_(pass)) return { ok: false, error: 'Admin passcode required.' };
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
  return { ok: true, project: { id: id, title: title, description: String(description || '').trim(), status: TP.projectStatus.assigned, assignees: list } };
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
    + '.tp-chip--empty{color:#a8a29e;background:transparent;border-style:dashed}'
    + '.tp-pill{flex:0 0 auto;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;padding:5px 11px;border-radius:999px}'
    + '.tp-pill--todo{color:#6b665e;background:#f0efe9;border:1px solid #e5e4de}'
    + '.tp-pill--assigned{color:#2563c9;background:#eaf1fd;border:1px solid #cfe0fb}'
    + '.tp-pill--active{color:#e08a1e;background:#fdf3e3;border:1px solid #f6e2bf}'
    + '.tp-pill--pending{color:#b06a00;background:#fdf2df;border:1px solid #f4dfb0}'
    + '.tp-pill--done{color:#157a47;background:#e7f6ee;border:1px solid #c7e9d5}'
    + '.tp-hint{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#a8a29e}'
    + '.tp-photos{display:flex;flex-wrap:wrap;gap:14px;margin-top:14px}'
    + '.tp-photo figure{margin:0}'
    + '.tp-photo figcaption{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#a8a29e;margin-bottom:5px}'
    + '.tp-uploader{margin-top:14px;padding:14px 16px;background:#fbfbf9;border:1.5px dashed #ddd;border-radius:12px}'
    + '.tp-drop{display:flex;flex-wrap:wrap;gap:10px}'
    + '.tp-slot{flex:1;min-width:150px}'
    + '.tp-slot-btn{display:flex;align-items:center;justify-content:center;gap:8px;width:100%;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#57534e;padding:14px 12px;border:1.5px solid #e0e0dc;border-radius:10px;background:#fff;cursor:pointer;text-align:center}'
    + '.tp-slot-btn:hover{border-color:#b31b1b;color:#8f1515}'
    + '.tp-slot.is-set .tp-slot-btn{border-color:#157a47;color:#157a47;background:#f2fbf6}'
    + '.tp-create{margin:6px 0 20px;padding:18px 20px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
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
    + '#tp-cheer{position:fixed;left:50%;top:20%;transform:translate(-50%,-10px) scale(.92);z-index:10000;pointer-events:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-weight:800;font-size:18px;color:#fff;background:linear-gradient(180deg,#1d9d5b 0%,#157a47 100%);padding:14px 22px;border-radius:14px;box-shadow:0 14px 34px rgba(21,122,71,.34);opacity:0;transition:opacity .3s ease,transform .35s cubic-bezier(.2,.9,.3,1.4);max-width:88vw;text-align:center}'
    + '#tp-cheer.show{opacity:1;transform:translate(-50%,0) scale(1)}'
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
    + 'function tpReadFile(input,cb){var f=input.files&&input.files[0];if(!f){cb(null);return;}if(!/^image\\//.test(f.type)){alert("Please choose an image file.");input.value="";cb(null);return;}var r=new FileReader();r.onload=function(){cb({dataUrl:r.result,name:f.name});};r.onerror=function(){alert("Could not read that file.");cb(null);};r.readAsDataURL(f);}'
    + '</script>';
}

// Discreet admin entry: a small "Admin" link that expands a passcode field on click.
// Kept low-key so it does not shout at doers who never need it.
function tpAdminBar_(purpose) {
  return '<div class="ic-adminbar">'
    + '<span id="tp-lock-msg" class="tp-lock-msg"></span>'
    + '<button type="button" class="ic-admin-toggle" id="tp-admin-toggle" onclick="tpAdminToggle()">Admin</button>'
    + '<span id="tp-lock-fields" class="ic-admin-fields" hidden>'
    + '<input type="password" id="tp-pass" placeholder="Passcode to ' + escapeHtml_(purpose) + '" onkeydown="if(event.key===\'Enter\')tpUnlock()">'
    + '<button type="button" class="btn btn-ghost" onclick="tpUnlock()">Unlock</button>'
    + '</span>'
    + '</div>';
}

// ---- Completion evidence: UI helpers used by the issue pages in 02 ----

// Empty (or pre-filled) container for the completion photo on a card.
function icPhotoBlock_(rid, url) {
  var inner = url
    ? '<div class="tp-photos"><div class="tp-photo"><figure><figcaption>Completion photo</figcaption>' + tpThumb_(url) + '</figure></div></div>'
    : '';
  return '<div id="' + rid + '-photo">' + inner + '</div>';
}

// Plain-language instruction shown on an open item so people know what to do.
function icNoteText_() {
  return 'Done with this? Take a photo of the finished work, then tap <b>Mark complete</b> to upload it.';
}
function icNoteInner_() { return '<div class="ic-note">' + icNoteText_() + '</div>'; }
function icDoerNote_(rid, show) { return '<div id="' + rid + '-note">' + (show ? icNoteInner_() : '') + '</div>'; }

// Foot action for an OPEN item: hidden file input + "Mark complete" (photo required).
function icOpenFoot_(rid, token) {
  return '<input type="file" accept="image/*" id="' + rid + '-file" style="display:none" onchange="icFile(this,\'' + rid + '\',\'' + token + '\')">'
    + '<button type="button" class="btn btn-primary" onclick="document.getElementById(\'' + rid + '-file\').click()">Mark complete</button>';
}

// Foot action for a PENDING item: admin-only Approve / Send back (revealed on unlock).
function icPendingFoot_(rid, token) {
  return '<span class="tp-admin btn-row" hidden>'
    + '<button type="button" class="btn btn-confirm" onclick="icApprove(\'' + rid + '\',\'' + token + '\')">Approve</button>'
    + '<button type="button" class="btn btn-ghost" onclick="icReject(\'' + rid + '\',\'' + token + '\')">Send back</button>'
    + '</span>';
}

// The client logic that drives completion on the issue pages. Requires tpSharedJs_
// (ADMIN_PASS, tpConfetti, tpReadFile) on the same page. Stat counters (#sum-open,
// #sum-pending, #sum-over) are updated only if present.
function icClientJs_() {
  return '<script>'
    + 'function icSetPill(rid,cls,txt){var p=document.getElementById(rid+"-pill");if(p)p.innerHTML="<span class=\\"tp-pill "+cls+"\\">"+txt+"</span>";}'
    + 'function icBump(id,d){var e=document.getElementById(id);if(e)e.textContent=Math.max(0,(parseInt(e.textContent,10)||0)+d);}'
    + 'function icThumb(pid){return "<div class=\\"tp-photos\\"><div class=\\"tp-photo\\"><figure><figcaption>Completion photo</figcaption><a href=\\"https://drive.google.com/file/d/"+pid+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+pid+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div></div>";}'
    + 'function icOpenFootJs(rid,token){return "<input type=\\"file\\" accept=\\"image/*\\" id=\\""+rid+"-file\\" style=\\"display:none\\" onchange=\\"icFile(this,\'"+rid+"\',\'"+token+"\')\\"><button type=\\"button\\" class=\\"btn btn-primary\\" onclick=\\"document.getElementById(\'"+rid+"-file\').click()\\">Mark complete</button>";}'
    + 'function icNoteInner(){return "<div class=\\"ic-note\\">Done with this? Take a photo of the finished work, then tap <b>Mark complete</b> to upload it.</div>";}'
    + 'function icAdminFootJs(rid,token){return ADMIN_PASS?"<span class=\\"tp-admin btn-row\\"><button type=\\"button\\" class=\\"btn btn-confirm\\" onclick=\\"icApprove(\'"+rid+"\',\'"+token+"\')\\">Approve</button><button type=\\"button\\" class=\\"btn btn-ghost\\" onclick=\\"icReject(\'"+rid+"\',\'"+token+"\')\\">Send back</button></span>":"";}'
    + 'function icFile(input,rid,token){tpReadFile(input,function(res){if(!res)return;var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Uploading photo\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Upload failed")+"</span>";return;}'
    + 'tpConfetti();icSetPill(rid,"tp-pill--pending","Pending approval");var ph=document.getElementById(rid+"-photo");if(ph)ph.innerHTML=icThumb(r.photoId);'
    + 'var s=document.getElementById(rid+"-status");if(s){s.textContent="Submitted, awaiting approval";s.className="due";}'
    + 'var nt=document.getElementById(rid+"-note");if(nt)nt.innerHTML="";'
    + 'act.innerHTML=icAdminFootJs(rid,token);var c=document.getElementById(rid);if(c){if(c.dataset.over==="1"){icBump("sum-over",-1);c.dataset.over="0";}c.dataset.state="pending";}'
    + 'icBump("sum-open",-1);icBump("sum-pending",1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Upload failed. Please retry.</span>";}).submitIssueCompletion(token,res.dataUrl,res.name);});}'
    + 'function icApprove(rid,token){var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'icSetPill(rid,"tp-pill--done","Completed");act.innerHTML="";var s=document.getElementById(rid+"-status");if(s){s.innerHTML="\\u2713 Completed";s.className="due due--done";}var c=document.getElementById(rid);if(c)c.style.opacity="0.72";icBump("sum-pending",-1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).approveIssueCompletion(token,ADMIN_PASS);}'
    + 'function icReject(rid,token){var act=document.getElementById(rid+"-act");act.innerHTML="<span class=\\"tp-hint\\">Saving\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'icSetPill(rid,"tp-pill--todo","Open");var ph=document.getElementById(rid+"-photo");if(ph)ph.innerHTML="";var s=document.getElementById(rid+"-status");if(s){s.textContent="Sent back, please add a new photo";s.className="due";}'
    + 'var nt=document.getElementById(rid+"-note");if(nt)nt.innerHTML=icNoteInner();'
    + 'act.innerHTML=icOpenFootJs(rid,token);var c=document.getElementById(rid);if(c)c.dataset.state="open";icBump("sum-pending",-1);icBump("sum-open",1);'
    + '}).withFailureHandler(function(){act.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).rejectIssueCompletion(token,ADMIN_PASS);}'
    + '</script>';
}

// ---- Projects page ----

function projectsPage_(embedded) {
  var projects = tpListProjects_();
  var assigned = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.assigned) || (!p.status); });
  var active = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.active); });
  var done = projects.filter(function (p) { return norm_(p.status) === norm_(TP.projectStatus.done); });

  var inner = '';
  if (!embedded) {
    inner += '<div class="page-head"><div class="page-kicker">Project Teams Ops Hub</div>'
      + '<div class="page-title">Projects</div><div class="page-rule"></div></div>';
  }

  inner += tpAdminBar_('assign projects');

  // Admin create form (revealed on unlock).
  inner += '<form class="tp-create tp-admin" hidden onsubmit="return tpCreate(event)">'
    + '<h3>Assign a new project</h3>'
    + '<div class="tp-field"><label>Title</label><input id="tp-c-title" placeholder="e.g. Rebuild the tool crib shelving" required></div>'
    + '<div class="tp-field"><label>Description</label><textarea id="tp-c-desc" rows="2" placeholder="Scope, location, and what done looks like"></textarea></div>'
    + '<div class="tp-field"><label>Assignees (optional, comma-separated)</label><input id="tp-c-assignees" placeholder="Alex Rivera, Sam Chen"></div>'
    + '<button type="submit" class="btn btn-primary">Create project</button>'
    + '<span id="tp-c-msg" class="tp-lock-msg" style="margin-left:10px"></span>'
    + '</form>';

  inner += '<div class="stats">'
    + '<div class="stat"><div class="stat-label">Assigned</div><div class="stat-val" id="tp-n-assigned">' + assigned.length + '</div></div>'
    + '<div class="stat"><div class="stat-label">In progress</div><div class="stat-val" id="tp-n-active">' + active.length + '</div></div>'
    + '<div class="stat"><div class="stat-label">Completed</div><div class="stat-val" id="tp-n-done">' + done.length + '</div></div>'
    + '</div>';

  var sectionHead = function (label, count, cls) {
    return '<div class="section-head"><span class="section-label ' + cls + '">' + label + '</span>'
      + '<span class="section-count">' + count + '</span><span class="section-rule"></span></div>';
  };

  inner += '<div id="tp-proj-list">';

  var idx = 0;
  var renderProject = function (p) {
    var rid = 'pj' + (idx++);
    var st = norm_(p.status);
    var isDone = st === norm_(TP.projectStatus.done);
    var isActive = st === norm_(TP.projectStatus.active);
    var pill = isDone ? '<span class="tp-pill tp-pill--done">Completed</span>'
      : isActive ? '<span class="tp-pill tp-pill--active">In progress</span>'
      : '<span class="tp-pill tp-pill--assigned">Assigned</span>';

    var chips = p.assignees.length
      ? p.assignees.map(function (a) { return '<span class="tp-chip">' + escapeHtml_(a) + '</span>'; }).join('')
      : '<span class="tp-chip tp-chip--empty">No one yet</span>';

    var photos = '';
    if (isDone && (p.before || p.after)) {
      photos = '<div class="tp-photos" id="' + rid + '-photos">'
        + '<div class="tp-photo"><figure><figcaption>Before</figcaption>' + tpThumb_(p.before) + '</figure></div>'
        + '<div class="tp-photo"><figure><figcaption>After</figcaption>' + tpThumb_(p.after) + '</figure></div>'
        + '</div>';
    } else {
      photos = '<div id="' + rid + '-photos"></div>';
    }

    var body = '<div class="card-body">'
      + '<div class="card-head"><div><div class="card-team">Project</div>'
      + '<div class="card-title">' + escapeHtml_(p.title) + '</div></div>'
      + '<span id="' + rid + '-pill">' + pill + '</span></div>'
      + (p.description ? '<div class="card-field"><span class="card-flabel">Scope</span><div class="card-details">' + escapeHtml_(p.description) + '</div></div>' : '')
      + '<div class="card-field"><span class="card-flabel">Assignees</span><div class="tp-assignees" id="' + rid + '-chips">' + chips + '</div></div>'
      + photos
      + '</div>';

    var foot = '';
    if (!isDone) {
      foot = '<div class="card-foot"><span id="' + rid + '-status" class="due">' + (isActive ? 'Work in progress' : 'Waiting to be picked up') + '</span>'
        + '<span class="btn-row">'
        + '<span id="' + rid + '-join"><button type="button" class="btn btn-ghost" onclick="tpJoinOpen(\'' + rid + '\')">Join project</button></span>'
        + '<span id="' + rid + '-joinbox" class="tp-inline-join" style="display:none"><input id="' + rid + '-name" placeholder="Your name" onkeydown="if(event.key===\'Enter\')tpJoin(\'' + rid + '\',\'' + p.id + '\')"><button type="button" class="btn btn-primary" onclick="tpJoin(\'' + rid + '\',\'' + p.id + '\')">Join</button></span>'
        + '<button type="button" class="btn btn-primary" onclick="tpCompleteOpen(\'' + rid + '\')">Complete</button>'
        + '</span></div>'
        + '<div id="' + rid + '-uploader" class="tp-uploader" style="display:none">'
        + '<div class="tp-hint" style="margin-bottom:10px">Completing a project requires both a <b>before</b> and an <b>after</b> photo.</div>'
        + '<div class="tp-drop">'
        + '<div class="tp-slot" id="' + rid + '-slot-b"><input type="file" accept="image/*" id="' + rid + '-before" style="display:none" onchange="tpSlot(this,\'' + rid + '\',\'b\')"><div class="tp-slot-btn" onclick="document.getElementById(\'' + rid + '-before\').click()"><span id="' + rid + '-blabel">Add “before” photo</span></div></div>'
        + '<div class="tp-slot" id="' + rid + '-slot-a"><input type="file" accept="image/*" id="' + rid + '-after" style="display:none" onchange="tpSlot(this,\'' + rid + '\',\'a\')"><div class="tp-slot-btn" onclick="document.getElementById(\'' + rid + '-after\').click()"><span id="' + rid + '-alabel">Add “after” photo</span></div></div>'
        + '</div>'
        + '<div style="margin-top:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">'
        + '<button type="button" class="btn btn-confirm" id="' + rid + '-finish" onclick="tpComplete(\'' + rid + '\',\'' + p.id + '\')">Complete project</button>'
        + '<button type="button" class="btn btn-ghost" onclick="tpCompleteClose(\'' + rid + '\')">Cancel</button>'
        + '<span id="' + rid + '-cmsg" class="tp-hint"></span>'
        + '</div></div>';
    } else {
      foot = '<div class="card-foot"><span id="' + rid + '-status" class="due due--done">✓ Completed</span></div>';
    }

    return '<div class="card" id="' + rid + '" data-status="' + escapeHtml_(p.status) + '">' + body + foot + '</div>';
  };

  if (!projects.length) {
    inner += '<div class="empty">No projects yet. Unlock admin above to assign one, or add rows in the "Projects" tab.</div>';
  } else {
    if (assigned.length) { inner += sectionHead('Assigned', assigned.length, 'section-label--open'); assigned.forEach(function (p) { inner += renderProject(p); }); }
    if (active.length) { inner += sectionHead('In progress', active.length, 'section-label--late'); active.forEach(function (p) { inner += renderProject(p); }); }
    if (done.length) { inner += sectionHead('Completed', done.length, 'section-label--open'); done.forEach(function (p) { inner += renderProject(p); }); }
  }
  inner += '</div>';

  inner += '<script>'
    + 'var TPUP={};'
    + 'function tpBump(id,d){var e=document.getElementById(id);if(e)e.textContent=Math.max(0,(parseInt(e.textContent,10)||0)+d);}'
    + 'function tpSetPill(rid,cls,txt){document.getElementById(rid+"-pill").innerHTML="<span class=\\"tp-pill "+cls+"\\">"+txt+"</span>";}'
    + 'function tpJoinOpen(rid){document.getElementById(rid+"-join").style.display="none";document.getElementById(rid+"-joinbox").style.display="inline-flex";document.getElementById(rid+"-name").focus();}'
    + 'function tpJoin(rid,pid){var nm=document.getElementById(rid+"-name").value.trim();if(!nm){document.getElementById(rid+"-name").focus();return;}'
    + 'var box=document.getElementById(rid+"-joinbox");box.innerHTML="<span class=\\"tp-hint\\">Joining\\u2026</span>";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){box.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">"+((r&&r.error)||"Failed")+"</span>";return;}'
    + 'var chips=document.getElementById(rid+"-chips");chips.innerHTML=r.assignees.map(function(a){return "<span class=\\"tp-chip\\">"+a.replace(/&/g,"&amp;").replace(/</g,"&lt;")+"</span>";}).join("");'
    + 'var card=document.getElementById(rid);var was=card.dataset.status;card.dataset.status=r.status;'
    + 'if((was||"").toLowerCase()!=="in progress"&&r.status==="In Progress"){tpSetPill(rid,"tp-pill--active","In progress");document.getElementById(rid+"-status").textContent="Work in progress";tpBump("tp-n-assigned",-1);tpBump("tp-n-active",1);}'
    + 'document.getElementById(rid+"-joinbox").style.display="none";document.getElementById(rid+"-join").style.display="inline";'
    + 'document.getElementById(rid+"-join").innerHTML="<span class=\\"tp-hint\\">\\u2713 You\'re on this project</span>";'
    + '}).withFailureHandler(function(){box.innerHTML="<span class=\\"tp-hint\\" style=\\"color:#b31b1b\\">Failed. Retry.</span>";}).tpJoinProject(pid,nm);}'
    + 'function tpCompleteOpen(rid){document.getElementById(rid+"-uploader").style.display="block";}'
    + 'function tpCompleteClose(rid){document.getElementById(rid+"-uploader").style.display="none";}'
    + 'function tpSlot(input,rid,which){tpReadFile(input,function(res){if(!res)return;TPUP[rid]=TPUP[rid]||{};TPUP[rid][which]=res;'
    + 'var slot=document.getElementById(rid+"-slot-"+which);slot.classList.add("is-set");'
    + 'document.getElementById(rid+"-"+(which==="b"?"blabel":"alabel")).textContent=(which==="b"?"\\u2713 Before: ":"\\u2713 After: ")+res.name;});}'
    + 'function tpComplete(rid,pid){var buf=TPUP[rid]||{};var msg=document.getElementById(rid+"-cmsg");'
    + 'if(!buf.b||!buf.a){msg.style.color="#b31b1b";msg.textContent="Add both a before and an after photo first.";return;}'
    + 'msg.style.color="";msg.textContent="Uploading both photos\\u2026";document.getElementById(rid+"-finish").disabled=true;'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.style.color="#b31b1b";msg.textContent=(r&&r.error)||"Failed";document.getElementById(rid+"-finish").disabled=false;return;}'
    + 'tpConfetti();tpSetPill(rid,"tp-pill--done","Completed");'
    + 'document.getElementById(rid+"-photos").innerHTML="<div class=\\"tp-photos\\"><div class=\\"tp-photo\\"><figure><figcaption>Before</figcaption><a href=\\"https://drive.google.com/file/d/"+r.beforeId+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+r.beforeId+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div>"'
    + '+"<div class=\\"tp-photo\\"><figure><figcaption>After</figcaption><a href=\\"https://drive.google.com/file/d/"+r.afterId+"/view\\" target=\\"_blank\\" rel=\\"noopener\\" style=\\"display:inline-block;line-height:0\\"><img src=\\"https://drive.google.com/thumbnail?id="+r.afterId+"&sz=w600\\" style=\\"max-width:100%;max-height:200px;border-radius:10px;border:1px solid #ececec\\"></a></figure></div></div>";'
    + 'var card=document.getElementById(rid);var was=(card.dataset.status||"").toLowerCase();'
    + 'var up=document.getElementById(rid+"-uploader");if(up)up.parentNode.removeChild(up);'
    + 'var foot=card.querySelector(".card-foot");if(foot)foot.innerHTML="<span class=\\"due due--done\\">\\u2713 Completed</span>";'
    + 'if(was==="in progress")tpBump("tp-n-active",-1);else tpBump("tp-n-assigned",-1);tpBump("tp-n-done",1);card.dataset.status="Completed";'
    + '}).withFailureHandler(function(){msg.style.color="#b31b1b";msg.textContent="Upload failed. Please retry.";document.getElementById(rid+"-finish").disabled=false;}).tpCompleteProject(pid,buf.b.dataUrl,buf.a.dataUrl,buf.b.name,buf.a.name);}'
    + 'function tpCreate(ev){ev.preventDefault();var t=document.getElementById("tp-c-title").value.trim();var d=document.getElementById("tp-c-desc").value.trim();var a=document.getElementById("tp-c-assignees").value.trim();var msg=document.getElementById("tp-c-msg");'
    + 'if(!t){msg.className="tp-lock-msg bad";msg.textContent="A title is required.";return false;}'
    + 'msg.className="tp-lock-msg";msg.textContent="Creating\\u2026";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Failed";return;}'
    + 'msg.className="tp-lock-msg ok";msg.textContent="\\u2713 Project created, reloading\\u2026";setTimeout(function(){location.reload();},700);'
    + '}).withFailureHandler(function(){msg.className="tp-lock-msg bad";msg.textContent="Failed. Retry.";}).tpCreateProject(t,d,a,ADMIN_PASS);return false;}'
    + '</script>';

  return swissShell_(tpStyles_() + inner + tpSharedJs_(), 'Projects', true, embedded);
}
