/**
 * Proposals - honeybee-swarm improvement voting (?module=proposals).
 *
 * Employees "scout" improvements and post them as proposals. Others rate each on
 * Impact and Effort (1-5). A proposal promotes from PROVISIONAL to REAL once it hits
 * quorum: >= PR.quorumMinVotes COUNTED reviews AND average Impact >= PR.quorumMinImpact.
 *
 * Honeybee mechanics:
 *  - Independent assessment: every voter rates Impact and Effort separately.
 *  - Quorum sensing: a real threshold of committed support flips it to Real (not a bare
 *    majority).
 *  - Earned voice: a vote only COUNTS once its voter has rated >= PR.earnedVoiceK OTHER
 *    proposals - you commit scouting effort before your signal is heard. Waived while
 *    fewer than PR.earnedVoiceBootstrap proposals are open, so the pool can bootstrap.
 *  - Honesty: one rating per proposal per NetID, deduped server-side.
 *
 * Admin mode (?admin=1, links from the unlisted admin page - same gate as projects)
 * reveals Approve now / Decline on each provisional card via tpAdminRevealJs_.
 *
 * Proposals can carry photos of the spot to improve: the create form stages them
 * client-side (tpStageRead/tpStagePreview) and uploads to Drive one at a time via
 * prUploadPhoto -> tpSaveUpload_; cards render them as linked thumbnails.
 *
 * Kept in its own file to avoid cluttering the notifier. Apps Script shares one global
 * scope, so doGet (in 02) routes ?module=proposals straight here. Reuses engine helpers
 * from 02/04: swissShell_, tpStyles_, tpSharedJs_, tpConfetti, tpAdminRevealJs_,
 * escapeHtml_, extractFileIds_, norm_, fmtShort_, newToken_, ss_, tpEnsureTab_,
 * tpOpen_, tpFindRow_, tpSaveUpload_. A change here still needs a NEW deployment version.
 */

var PR = {
  proposalsSheet: 'Proposals',
  votesSheet: 'Proposal votes',
  proposalHeaders: ['Proposal ID', 'Title', 'Description', 'Area', 'Proposed by', 'NetID', 'Status', 'Created at', 'Promoted at', 'Photos'],
  voteHeaders: ['Vote ID', 'Proposal ID', 'Voter NetID', 'Impact', 'Effort', 'Comment', 'Voted at'],
  status: { provisional: 'Provisional', real: 'Real', declined: 'Declined' },
  quorumMinVotes: 3,       // counted reviews needed to promote
  quorumMinImpact: 3.5,    // ...and the average Impact must reach this
  earnedVoiceK: 2,         // rate this many OTHER proposals before your votes count
  earnedVoiceBootstrap: 3, // ...waived while fewer than this many proposals are open
};

// One-time setup. Safe to re-run.
function setupProposals() {
  var ss = ss_();
  tpEnsureTab_(ss, PR.proposalsSheet, PR.proposalHeaders,
    ['', 'Example: add a shadow board to the hand-tool wall', 'So every tool has a home and a missing one is obvious at a glance.', 'Baja bay', 'Example Student', 'abc123', PR.status.provisional, new Date(), '', '']);
  tpEnsureTab_(ss, PR.votesSheet, PR.voteHeaders, null);
  Logger.log('Proposals ready: ' + PR.proposalsSheet + ' + ' + PR.votesSheet + '.');
}

// ---- data ----

function prProposals_() {
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders);
  var sh = o.sh, col = o.col, last = sh.getLastRow(), out = [], pending = [];
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < v.length; i++) {
    var row = v[i], title = String(row[col['Title'] - 1] || '').trim();
    if (!title) continue;
    var id = String(row[col['Proposal ID'] - 1] || '').trim();
    if (!id) { id = newToken_(); pending.push({ r: i + 2, c: col['Proposal ID'], val: id }); }
    var status = String(row[col['Status'] - 1] || '').trim() || PR.status.provisional;
    out.push({
      row: i + 2, id: id, title: title,
      description: String(row[col['Description'] - 1] || '').trim(),
      area: String(row[col['Area'] - 1] || '').trim(),
      proposedBy: String(row[col['Proposed by'] - 1] || '').trim(),
      netid: String(row[col['NetID'] - 1] || '').trim(),
      status: status,
      createdAt: row[col['Created at'] - 1] || '',
      promotedAt: row[col['Promoted at'] - 1] || '',
      photos: String(row[col['Photos'] - 1] || '').trim()
    });
  }
  pending.forEach(function (w) { sh.getRange(w.r, w.c).setValue(w.val); });
  return out;
}

function prVotes_() {
  var o = tpOpen_(PR.votesSheet, PR.voteHeaders);
  var sh = o.sh, col = o.col, last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < v.length; i++) {
    var pid = String(v[i][col['Proposal ID'] - 1] || '').trim(), netid = norm_(v[i][col['Voter NetID'] - 1]);
    if (!pid || !netid) continue;
    out.push({ row: i + 2, proposalId: pid, netid: netid,
      impact: Number(v[i][col['Impact'] - 1]) || 0, effort: Number(v[i][col['Effort'] - 1]) || 0 });
  }
  return out;
}

// netid -> count of DISTINCT proposals that voter has rated (the earned-voice signal).
function prReviewMap_(votes) {
  var seen = {}, counts = {};
  votes.forEach(function (vt) { (seen[vt.netid] = seen[vt.netid] || {})[vt.proposalId] = 1; });
  Object.keys(seen).forEach(function (n) { counts[n] = Object.keys(seen[n]).length; });
  return counts;
}

// Proposals + per-proposal vote stats. "counted" votes only come from qualified voters
// (earned-voice), unless we are still bootstrapping the pool. viewerNetid, if given,
// tags each proposal with the viewer's own prior vote.
function prListData_(viewerNetid) {
  var proposals = prProposals_(), votes = prVotes_(), reviewCounts = prReviewMap_(votes);
  var openCount = proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.provisional); }).length;
  var bootstrap = openCount < PR.earnedVoiceBootstrap;
  var vn = norm_(viewerNetid || '');
  var byProp = {};
  votes.forEach(function (vt) { (byProp[vt.proposalId] = byProp[vt.proposalId] || []).push(vt); });
  proposals.forEach(function (p) {
    var vs = byProp[p.id] || [];
    var counted = vs.filter(function (vt) {
      if (bootstrap) return true;
      return ((reviewCounts[vt.netid] || 0) - 1) >= PR.earnedVoiceK;  // rated K OTHERS besides this one
    });
    p.votes = vs.length;
    p.countedVotes = counted.length;
    p.avgImpact = counted.length ? counted.reduce(function (s, x) { return s + x.impact; }, 0) / counted.length : 0;
    p.avgEffort = counted.length ? counted.reduce(function (s, x) { return s + x.effort; }, 0) / counted.length : 0;
    if (vn) { var mine = vs.filter(function (x) { return x.netid === vn; })[0]; p.myVote = mine ? { impact: mine.impact, effort: mine.effort } : null; }
  });
  return { proposals: proposals, viewerReviewCount: vn ? (reviewCounts[vn] || 0) : 0, bootstrap: bootstrap };
}

// Promote a provisional proposal if it has reached quorum. Returns true if it just did.
function prPromoteIfQuorum_(id) {
  var p = prListData_('').proposals.filter(function (x) { return x.id === id; })[0];
  if (!p || norm_(p.status) !== norm_(PR.status.provisional)) return false;
  if (p.countedVotes >= PR.quorumMinVotes && p.avgImpact >= PR.quorumMinImpact) {
    var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), r = tpFindRow_(o.sh, o.col['Proposal ID'], id);
    if (r > 0) { o.sh.getRange(r, o.col['Status']).setValue(PR.status.real); o.sh.getRange(r, o.col['Promoted at']).setValue(new Date()); }
    return true;
  }
  return false;
}

// ---- mutations (client-callable) ----

// Admin decision on a provisional proposal: 'approve' promotes it without waiting for
// quorum, 'decline' closes it (the row stays in the sheet for the record; the page
// only lists Provisional + Real). The buttons only render revealed in admin mode.
function prAdminSetStatus(proposalId, decision) {
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), r = tpFindRow_(o.sh, o.col['Proposal ID'], proposalId);
  if (r < 0) return { ok: false, error: 'That proposal could not be found.' };
  if (decision === 'approve') {
    o.sh.getRange(r, o.col['Status']).setValue(PR.status.real);
    o.sh.getRange(r, o.col['Promoted at']).setValue(new Date());
  } else if (decision === 'decline') {
    o.sh.getRange(r, o.col['Status']).setValue(PR.status.declined);
  } else {
    return { ok: false, error: 'Unknown decision.' };
  }
  return { ok: true, status: decision === 'approve' ? PR.status.real : PR.status.declined };
}

function prSubmitProposal(title, description, area, name, netid, photoIds) {
  title = String(title || '').trim();
  if (!title) return { ok: false, error: 'A title is required.' };
  netid = String(netid || '').trim();
  if (!netid) return { ok: false, error: 'Enter your NetID first (top of the page).' };
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), id = newToken_(), r = o.sh.getLastRow() + 1;
  o.sh.getRange(r, o.col['Proposal ID']).setValue(id);
  o.sh.getRange(r, o.col['Title']).setValue(title);
  o.sh.getRange(r, o.col['Description']).setValue(String(description || '').trim());
  o.sh.getRange(r, o.col['Area']).setValue(String(area || '').trim());
  o.sh.getRange(r, o.col['Proposed by']).setValue(String(name || '').trim());
  o.sh.getRange(r, o.col['NetID']).setValue(netid);
  o.sh.getRange(r, o.col['Status']).setValue(PR.status.provisional);
  o.sh.getRange(r, o.col['Created at']).setValue(new Date());
  o.sh.getRange(r, o.col['Photos']).setValue(String(photoIds || '').trim());
  return { ok: true, id: id };
}

// One staged photo -> Drive (the create form uploads its tray one file at a time,
// same as project completion photos). Returns the file id for the Photos cell.
function prUploadPhoto(dataUrl, filename) {
  try {
    return { ok: true, photoId: tpSaveUpload_(dataUrl, filename) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function prVote(proposalId, netid, impact, effort, comment) {
  netid = String(netid || '').trim();
  if (!netid) return { ok: false, error: 'Enter your NetID to vote.' };
  var imp = Number(impact), eff = Number(effort);
  if (!(imp >= 1 && imp <= 5) || !(eff >= 1 && eff <= 5)) return { ok: false, error: 'Rate impact and effort from 1 to 5.' };
  var p = prProposals_().filter(function (x) { return x.id === proposalId; })[0];
  if (!p) return { ok: false, error: 'That proposal could not be found.' };
  if (norm_(p.status) !== norm_(PR.status.provisional)) return { ok: false, error: 'Voting is closed - this one is already decided.' };
  if (norm_(p.netid) === norm_(netid)) return { ok: false, error: 'You cannot rate your own proposal.' };
  var o = tpOpen_(PR.votesSheet, PR.voteHeaders), last = o.sh.getLastRow(), found = -1;
  if (last >= 2) {
    var vv = o.sh.getRange(2, 1, last - 1, o.sh.getLastColumn()).getValues();
    for (var i = 0; i < vv.length; i++) {
      if (String(vv[i][o.col['Proposal ID'] - 1]).trim() === proposalId && norm_(vv[i][o.col['Voter NetID'] - 1]) === norm_(netid)) { found = i + 2; break; }
    }
  }
  var r = found > 0 ? found : o.sh.getLastRow() + 1;
  if (found < 0) o.sh.getRange(r, o.col['Vote ID']).setValue(newToken_());
  o.sh.getRange(r, o.col['Proposal ID']).setValue(proposalId);
  o.sh.getRange(r, o.col['Voter NetID']).setValue(netid);
  o.sh.getRange(r, o.col['Impact']).setValue(imp);
  o.sh.getRange(r, o.col['Effort']).setValue(eff);
  o.sh.getRange(r, o.col['Comment']).setValue(String(comment || '').trim());
  o.sh.getRange(r, o.col['Voted at']).setValue(new Date());
  // This review can retroactively qualify the voter's EARLIER votes (earned voice), so
  // re-check quorum on every proposal they have rated, not just this one.
  var promoted = prPromoteIfQuorum_(proposalId);
  prVotes_().forEach(function (x) {
    if (x.netid === norm_(netid) && x.proposalId !== proposalId) prPromoteIfQuorum_(x.proposalId);
  });
  return { ok: true, promoted: promoted, updated: found > 0 };
}

// ---- page ----

function prSection_(label, count, cls) {
  return '<div class="section-head"><span class="section-label ' + (cls || 'section-label--open') + '">' + label + '</span>'
    + '<span class="section-count">' + count + '</span><span class="section-rule"></span></div>';
}

// One proposal card. votable => provisional (shows the rating control); else it's an
// approved (Real) card. `data` carries bootstrap + viewer review count for the gating note.
function prCardHtml_(p, rid, votable, data) {
  var pill = votable
    ? '<span class="tp-pill tp-pill--active">Up for review</span>'
    : '<span class="tp-pill tp-pill--done">Approved</span>';
  var impTxt = p.countedVotes ? p.avgImpact.toFixed(1) : '–';
  var effTxt = p.countedVotes ? p.avgEffort.toFixed(1) : '–';
  var need = PR.quorumMinVotes;
  var progress = votable
    ? '<div class="pr-stats"><span class="pr-stat"><b>' + p.countedVotes + '</b>/' + need + ' reviews</span>'
      + '<span class="pr-stat">Impact <b>' + impTxt + '</b></span><span class="pr-stat">Effort <b>' + effTxt + '</b></span>'
      + (p.avgImpact >= PR.quorumMinImpact ? '<span class="pr-stat pr-stat--good">meets the bar</span>' : '<span class="pr-stat pr-stat--muted">needs Impact ≥ ' + PR.quorumMinImpact + '</span>')
      + '</div>'
    : '<div class="pr-stats"><span class="pr-stat pr-stat--good">Promoted ' + (p.promotedAt ? escapeHtml_(fmtShort_(p.promotedAt)) : '') + '</span>'
      + '<span class="pr-stat">Impact <b>' + impTxt + '</b></span><span class="pr-stat">Effort <b>' + effTxt + '</b></span></div>';

  var scale = function (k, val) {
    var out = '<div class="pr-scale" data-k="' + k + '" data-val="' + (val || '') + '">';
    for (var n = 1; n <= 5; n++) out += '<button type="button" class="pr-dot' + (val === n ? ' on' : '') + '" onclick="prPick(this)">' + n + '</button>';
    return out + '</div>';
  };
  var my = p.myVote;
  var rate = votable
    ? '<div class="pr-rate" id="' + rid + '-rate">'
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Impact <em>(1 low → 5 high)</em></span>' + scale('impact', my ? my.impact : 0) + '</div>'
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Effort <em>(1 easy → 5 hard)</em></span>' + scale('effort', my ? my.effort : 0) + '</div>'
      + '<div class="pr-rate-foot"><button type="button" class="btn btn-primary" onclick="prSubmitVote(\'' + rid + '\',\'' + p.id + '\')">' + (my ? 'Update rating' : 'Submit rating') + '</button>'
      + '<span id="' + rid + '-vmsg" class="tp-hint"></span></div></div>'
    : '';
  // Admin-only decisions (revealed by tpAdminRevealJs_ / re-revealed after prRefresh).
  var adminBar = votable
    ? '<div class="tp-admin pr-adminbar" hidden><button type="button" class="btn btn-confirm" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'approve\')">Approve now</button>'
      + '<button type="button" class="btn btn-ghost" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'decline\')">Decline</button></div>'
    : '';

  // Photos the scout attached (Drive ids). onerror hides a thumb whose file will not load.
  var photoIds = extractFileIds_(p.photos);
  var photos = photoIds.length
    ? '<div class="pr-photos">' + photoIds.map(function (fid) {
        var e = encodeURIComponent(fid);
        return '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" title="Proposal photo"><img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w400" loading="lazy" onerror="this.closest(\'a\').style.display=\'none\'"></a>';
      }).join('') + '</div>'
    : '';

  return '<div class="card pr-card" id="' + rid + '" data-id="' + p.id + '" data-owner="' + escapeHtml_(p.netid) + '">'
    + '<div class="card-body">'
    +   '<div class="card-head"><div><div class="card-team">Proposal' + (p.area ? ' · ' + escapeHtml_(p.area) : '') + '</div>'
    +     '<div class="card-title">' + escapeHtml_(p.title) + '</div></div><span>' + pill + '</span></div>'
    +   (p.description ? '<div class="card-field"><div class="card-details">' + escapeHtml_(p.description) + '</div></div>' : '')
    +   (p.proposedBy ? '<div class="pr-by">Proposed by ' + escapeHtml_(p.proposedBy) + '</div>' : '')
    +   photos
    +   progress
    +   rate
    +   adminBar
    + '</div></div>';
}

function prListSectionsHtml_(data) {
  var provisional = data.proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.provisional); });
  var real = data.proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.real); });
  provisional.sort(function (a, b) { return (b.avgImpact - a.avgImpact) || (b.countedVotes - a.countedVotes) || (new Date(b.createdAt || 0) - new Date(a.createdAt || 0)); });
  real.sort(function (a, b) { return new Date(b.promotedAt || 0) - new Date(a.promotedAt || 0); });
  var idx = 0, out = '';
  out += prSection_('Up for review', provisional.length, 'section-label--late');
  if (!provisional.length) out += '<div class="empty">Nothing up for review yet. Scout an improvement and post it above.</div>';
  else provisional.forEach(function (p) { out += prCardHtml_(p, 'pr' + (idx++), true, data); });
  if (real.length) { out += prSection_('Approved', real.length, 'section-label--open'); real.forEach(function (p) { out += prCardHtml_(p, 'pr' + (idx++), false, data); }); }
  return out;
}

// Client-callable: fresh list HTML for the viewer (re-render after a vote/post).
function prListHtml(netid) {
  var data = prListData_(netid || '');
  return { ok: true, html: prListSectionsHtml_(data), reviewCount: data.viewerReviewCount, bootstrap: data.bootstrap, k: PR.earnedVoiceK };
}

function proposalsPage_(embedded, admin) {
  var inner = '';
  if (!embedded) inner += '<div class="page-head"><div class="page-kicker">Project Teams Ops Hub</div><div class="page-title">Proposals</div><div class="page-rule"></div></div>';
  inner += '<div class="pr-intro">Scout improvements to our spaces and post them. Everyone rates each on impact and effort - the ones that reach a quorum of support become real projects.</div>';
  inner += '<div class="pr-idbar"><label class="pr-idlbl" for="pr-netid">Your NetID</label>'
    + '<input id="pr-netid" class="pr-idinput" placeholder="e.g. abc123" autocomplete="off" oninput="prSaveNetid()"><span id="pr-idnote" class="pr-idnote"></span></div>';
  inner += '<details class="tp-create-wrap"><summary class="tp-create-toggle"><span class="tp-create-caret">&#43;</span> Propose an improvement</summary>'
    + '<div class="tp-create">'
    + '<div class="tp-field"><label>Title <span class="tp-req">Required</span></label><input id="pr-c-title" placeholder="e.g. Add a shadow board to the hand-tool wall"></div>'
    + '<div class="tp-field"><label>What &amp; why</label><textarea id="pr-c-desc" rows="3" placeholder="What to change, and why it helps"></textarea></div>'
    + '<div class="tp-field"><label>Area / space</label><input id="pr-c-area" placeholder="e.g. Baja bay, composites lab"></div>'
    + '<div class="tp-field"><label>Your name</label><input id="pr-c-name" placeholder="Your name"></div>'
    + '<div class="tp-field"><label>Photos <span class="tp-hint" style="text-transform:none;letter-spacing:0">Optional - show the spot you want to improve</span></label>'
    +   '<input type="file" accept="image/*" multiple id="pr-c-file" style="display:none" onchange="prPickPhotos(this)">'
    +   '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'pr-c-file\').click()">Add photos</button>'
    +   '<div id="pr-c-stage"></div></div>'
    + '<button type="button" class="btn btn-primary" onclick="prCreate()">Post for review</button>'
    + '<span id="pr-c-msg" class="tp-lock-msg" style="margin-left:10px"></span>'
    + '</div></details>';
  inner += '<div id="pr-list">' + prListSectionsHtml_(prListData_('')) + '</div>';
  return swissShell_(tpStyles_() + prStyles_() + inner + tpSharedJs_() + prClientJs_() + tpAdminRevealJs_(admin), 'Proposals', true, embedded);
}

function prStyles_() {
  return '<style>'
    + '.pr-intro{font-size:14.5px;line-height:1.6;color:#57534e;margin:14px 0 4px;max-width:640px}'
    + '.pr-idbar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:18px 0 6px;padding:12px 16px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.pr-idlbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c}'
    + '.pr-idinput{font:inherit;font-size:14px;padding:10px 13px;border:1.5px solid #e0e0dc;border-radius:10px;outline:none;background:#fafaf8;min-width:160px}'
    + '.pr-idinput:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12);background:#fff}'
    + '.pr-idnote{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:600;color:#8a857c}'
    + '.pr-by{font-size:12.5px;font-weight:600;color:#8a857c;margin-top:10px}'
    + '.pr-photos{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}'
    + '.pr-photos a{display:inline-block;line-height:0}'
    + '.pr-photos img{max-height:120px;border-radius:10px;border:1px solid #ececec}'
    + '.pr-stats{display:flex;flex-wrap:wrap;gap:8px 14px;margin-top:14px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;color:#57534e}'
    + '.pr-stat b{color:#14110e;font-weight:800}'
    + '.pr-stat--good{color:#157a47;font-weight:700}.pr-stat--muted{color:#a8a29e}'
    + '.pr-rate{margin-top:16px;padding:14px 16px;background:#faf9f6;border:1.5px solid #ece9e2;border-radius:12px}'
    + '.pr-rate-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:12px}'
    + '.pr-rate-lbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;letter-spacing:.03em;color:#57534e;min-width:150px}'
    + '.pr-rate-lbl em{font-style:normal;font-weight:600;color:#a8a29e;letter-spacing:0}'
    + '.pr-scale{display:inline-flex;gap:6px}'
    + '.pr-dot{width:38px;height:38px;border-radius:10px;border:1.5px solid #e0e0dc;background:#fff;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;font-weight:800;color:#8a857c;cursor:pointer;transition:border-color .12s,background .12s,color .12s,transform .1s}'
    + '.pr-dot:hover{border-color:#b5b0a8}.pr-dot:active{transform:scale(.94)}'
    + '.pr-dot.on{background:#b31b1b;border-color:#b31b1b;color:#fff}'
    + '.pr-rate-foot{display:flex;align-items:center;gap:12px;flex-wrap:wrap}'
    + '.pr-adminbar{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;padding-top:14px;border-top:1.5px dashed #ece9e2}'
    + '@media(max-width:600px){.pr-rate-lbl{min-width:0;width:100%}.pr-dot{flex:1;min-width:0}}'
    + '</style>';
}

function prClientJs_() {
  return '<script>'
    + 'var PR_NETID="";'
    + 'function prNetid(){var el=document.getElementById("pr-netid");return el?el.value.trim():"";}'
    + 'function prSaveNetid(){PR_NETID=prNetid();try{localStorage.setItem("prNetid",PR_NETID);}catch(e){}}'
    + 'function prPick(btn){var sc=btn.parentNode;[].slice.call(sc.querySelectorAll(".pr-dot")).forEach(function(d){d.classList.remove("on");});btn.classList.add("on");sc.setAttribute("data-val",btn.textContent.trim());}'
    + 'function prScaleVal(rid,k){var sc=document.querySelector("#"+rid+"-rate .pr-scale[data-k=\\""+k+"\\"]");return sc?parseInt(sc.getAttribute("data-val"),10)||0:0;}'
    // Staged photos for the create form (client-side buffer, uploaded on Post).
    + 'var PRSTAGE=[];'
    + 'function prStageRender(){var el=document.getElementById("pr-c-stage");if(el)el.innerHTML=tpStagePreview(PRSTAGE,"prUnstage(","document.getElementById(\'pr-c-file\').click()");}'
    + 'function prPickPhotos(input){tpStageRead(input,PRSTAGE,prStageRender);}'
    + 'function prUnstage(i){PRSTAGE.splice(i,1);prStageRender();}'
    + 'function prCreate(){var t=(document.getElementById("pr-c-title")||{}).value||"";var d=(document.getElementById("pr-c-desc")||{}).value||"";var a=(document.getElementById("pr-c-area")||{}).value||"";var nm=(document.getElementById("pr-c-name")||{}).value||"";var msg=document.getElementById("pr-c-msg");var nid=prNetid();'
    + 'if(!t.trim()){msg.className="tp-lock-msg bad";msg.textContent="A title is required.";return;}if(!nid){msg.className="tp-lock-msg bad";msg.textContent="Enter your NetID at the top first.";return;}'
    + 'msg.className="tp-lock-msg";'
    // Upload the tray one photo at a time, then post the proposal with the Drive ids.
    + 'var ids=[],q=PRSTAGE.slice();'
    + '(function step(){if(q.length){msg.textContent="Uploading photo "+(ids.length+1)+" of "+PRSTAGE.length+"\\u2026";var ph=q.shift();'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Photo upload failed. Retry.";return;}ids.push(r.photoId);step();}).withFailureHandler(function(){msg.className="tp-lock-msg bad";msg.textContent="Photo upload failed. Retry.";}).prUploadPhoto(ph.dataUrl,ph.name);return;}'
    + 'msg.textContent="Posting\\u2026";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Failed";return;}'
    + 'document.getElementById("pr-c-title").value="";document.getElementById("pr-c-desc").value="";document.getElementById("pr-c-area").value="";PRSTAGE=[];prStageRender();'
    + 'msg.className="tp-lock-msg ok";msg.textContent="\\u2713 Posted for review";prRefresh();}).withFailureHandler(function(){msg.className="tp-lock-msg bad";msg.textContent="Failed. Retry.";}).prSubmitProposal(t,d,a,nm,nid,ids.join(", "));})();}'
    + 'function prSubmitVote(rid,pid){var msg=document.getElementById(rid+"-vmsg");var nid=prNetid();if(!nid){msg.style.color="#b31b1b";msg.textContent="Enter your NetID at the top to vote.";return;}'
    + 'var imp=prScaleVal(rid,"impact"),eff=prScaleVal(rid,"effort");if(!imp||!eff){msg.style.color="#b31b1b";msg.textContent="Rate both impact and effort.";return;}'
    + 'msg.style.color="";msg.textContent="Saving\\u2026";'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){msg.style.color="#b31b1b";msg.textContent=(r&&r.error)||"Failed";return;}'
    + 'if(r.promoted)tpConfetti("\\uD83D\\uDC1D Quorum reached - it is a real project now!");prRefresh();}).withFailureHandler(function(){msg.style.color="#b31b1b";msg.textContent="Failed. Retry.";}).prVote(pid,nid,imp,eff,"");}'
    + 'function prAdmin(rid,pid,dec){var msg=document.getElementById(rid+"-vmsg");if(msg){msg.style.color="";msg.textContent=(dec==="approve"?"Approving":"Declining")+"\\u2026";}'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){if(msg){msg.style.color="#b31b1b";msg.textContent=(r&&r.error)||"Failed";}return;}'
    + 'if(dec==="approve")tpConfetti("\\uD83D\\uDC1D Approved - it is a real project now!");prRefresh();}).withFailureHandler(function(){if(msg){msg.style.color="#b31b1b";msg.textContent="Failed. Retry.";}}).prAdminSetStatus(pid,dec);}'
    + 'function prRefresh(){google.script.run.withSuccessHandler(function(r){if(!r||!r.ok)return;var el=document.getElementById("pr-list");if(el)el.innerHTML=r.html;'
    + 'if(ADMIN_PASS)document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});'
    + 'var note=document.getElementById("pr-idnote");if(note){if(!prNetid())note.textContent="";else if(r.bootstrap)note.textContent="Your votes count now (early days).";else{var left=r.k+1-r.reviewCount;note.textContent=left>0?("Review "+left+" more proposal"+(left>1?"s":"")+" to have your votes counted."):"\\u2713 Your votes are counted.";}}'
    + '}).withFailureHandler(function(){}).prListHtml(prNetid());}'
    + 'window.addEventListener("DOMContentLoaded",function(){try{var s=localStorage.getItem("prNetid");if(s){var el=document.getElementById("pr-netid");if(el)el.value=s;PR_NETID=s;}}catch(e){}if(prNetid())prRefresh();});'
    + '</script>';
}
