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
 *    majority). Hitting quorum (or an admin Approve) writes a row on the Projects
 *    sheet so staff can join it from Your work → Projects. The proposals board then
 *    shows a compact “in Projects” row instead of keeping it up for review.
 *  - Earned voice: a vote only COUNTS once its voter has rated >= PR.earnedVoiceK OTHER
 *    proposals - you commit scouting effort before your signal is heard. Waived while
 *    fewer than PR.earnedVoiceBootstrap proposals are open, so the pool can bootstrap.
 *  - Honesty: one rating per proposal per NetID, deduped server-side.
 *
 * Admin mode (?admin=1, links from the unlisted admin page - same gate as projects)
 * reveals Edit / Approve now / Decline / Delete on each card via tpAdminRevealJs_.
 * Edit changes the scout-entered fields in place. Delete is for proposals the swarm
 * has clearly rejected; it removes the row and its votes. Decline keeps the row for
 * the record.
 *
 * Proposals can carry photos of the spot to improve. The create form stages them
 * client-side (tpStageRead) and sends them to Drive AS SOON AS THEY ARE PICKED, all in
 * parallel, via prUploadPhoto -> tpSaveUpload_. By the time the scout finishes typing
 * the files are usually already up, so posting costs one round trip rather than one per
 * photo plus the post. A photo removed after it uploaded is trashed by prTrashUpload so
 * abandoned picks do not pile up. Cards render the photos as linked thumbnails.
 *
 * SPEED. Sheet round-trips dominate the response time here, so:
 *  - prProposals_/prVotes_ memoize per execution (PR_MEMO); any write clears it.
 *  - The rendered list is netid-agnostic and cached in CacheService (PR.cacheKey),
 *    dropped on every write. A warm load skips the sheets entirely.
 *  - The viewer's own state (their ratings + earned-voice progress) arrives from
 *    prViewerState, which reads ONLY the votes sheet and patches the rendered cards
 *    client-side. The page never re-renders its whole list just to personalise it.
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
  cacheKey: 'pr_list_v5',  // rendered list HTML (netid-agnostic)
  stripKey: 'pr_strip_v5', // one-line banner on the tasks pages
  cacheSeconds: 120,
  freshDays: 14,           // a promotion counts as news on the tasks page for this long
};

// One-time setup. Safe to re-run.
function setupProposals() {
  var ss = ss_();
  tpEnsureTab_(ss, PR.proposalsSheet, PR.proposalHeaders,
    ['', 'Example: add a shadow board to the hand-tool wall', 'So every tool has a home and a missing one is obvious at a glance.', 'Baja bay', 'Example Student', 'abc123', PR.status.provisional, new Date(), '', '']);
  tpEnsureTab_(ss, PR.votesSheet, PR.voteHeaders, null);
  Logger.log('Proposals ready: ' + PR.proposalsSheet + ' + ' + PR.votesSheet + '.');
}

// ---- caching ----

// Per-execution memo of the two sheets. Several helpers want the same rows in one
// request (a vote reads proposals, writes, then recounts); without this each one paid
// for its own round-trip.
var PR_MEMO = {};

function prCacheSvc_() { try { return CacheService.getScriptCache(); } catch (e) { return null; } }

// Called after every write. Drops the memo and the rendered-list caches so the next
// reader sees the new numbers.
function prInvalidate_() {
  PR_MEMO = {};
  var c = prCacheSvc_();
  if (!c) return;
  var keys = [PR.cacheKey, PR.stripKey];
  if (typeof TP !== 'undefined' && TP.cacheKey) keys.push(TP.cacheKey);
  try { c.removeAll(keys); } catch (e) { /* cache is best-effort */ }
}

function prWebHref_(module) {
  var base = (typeof CONFIG !== 'undefined' && CONFIG.webAppUrl) ? String(CONFIG.webAppUrl) : '';
  var join = base.indexOf('?') >= 0 ? '&' : '?';
  return base + join + 'module=' + encodeURIComponent(module);
}

// Any Real proposal that does not yet have a Projects row gets one. Cheap when the
// two already match; used as a backfill for proposals promoted before spawn existed.
function prSpawnMissing_() {
  var reals = prProposals_().filter(function (p) { return norm_(p.status) === norm_(PR.status.real); });
  if (!reals.length || typeof tpSpawnFromProposal_ !== 'function') return 0;
  var n = 0;
  reals.forEach(function (p) {
    try { if (tpSpawnFromProposal_(p, true)) n++; } catch (e) { Logger.log('spawn ' + p.id + ': ' + e); }
  });
  if (n && typeof tpInvalidate_ === 'function') tpInvalidate_();
  return n;
}

// ---- data ----

function prProposals_() {
  if (PR_MEMO.proposals) return PR_MEMO.proposals;
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders);
  var sh = o.sh, col = o.col, last = sh.getLastRow(), out = [], pending = [];
  if (last < 2) { PR_MEMO.proposals = out; return out; }
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
  PR_MEMO.proposals = out;
  return out;
}

function prVotes_() {
  if (PR_MEMO.votes) return PR_MEMO.votes;
  var o = tpOpen_(PR.votesSheet, PR.voteHeaders);
  var sh = o.sh, col = o.col, last = sh.getLastRow(), out = [];
  if (last < 2) { PR_MEMO.votes = out; return out; }
  var v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < v.length; i++) {
    var pid = String(v[i][col['Proposal ID'] - 1] || '').trim(), netid = norm_(v[i][col['Voter NetID'] - 1]);
    if (!pid || !netid) continue;
    out.push({ row: i + 2, proposalId: pid, netid: netid,
      impact: Number(v[i][col['Impact'] - 1]) || 0, effort: Number(v[i][col['Effort'] - 1]) || 0 });
  }
  PR_MEMO.votes = out;
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

// The viewer's earned-voice standing: how many DISTINCT proposals they must rate before
// any of their votes count, and how far along they are. The page shows this as a small
// progress meter, so the remaining step is always a visible, near goal.
function prVoiceState_(data) {
  var need = PR.earnedVoiceK + 1;   // K others PLUS the one being rated
  var have = Math.min(data.viewerReviewCount, need);
  return { need: need, have: have, bootstrap: data.bootstrap, counted: data.bootstrap || data.viewerReviewCount >= need };
}

// Promote every provisional proposal in `data` that has reached quorum, and mark it
// promoted in the snapshot. Checked across the whole pool, not just the one last rated:
// a new review retroactively qualifies that voter's EARLIER votes (earned voice), which
// can carry another proposal over the line. One snapshot, so this costs a single read.
function prPromoteAll_(data) {
  var promoted = [], o = null;
  data.proposals.forEach(function (p) {
    if (norm_(p.status) !== norm_(PR.status.provisional)) return;
    if (p.countedVotes < PR.quorumMinVotes || p.avgImpact < PR.quorumMinImpact) return;
    o = o || tpOpen_(PR.proposalsSheet, PR.proposalHeaders);
    var r = tpFindRow_(o.sh, o.col['Proposal ID'], p.id);
    if (r < 0) return;
    o.sh.getRange(r, o.col['Status']).setValue(PR.status.real);
    o.sh.getRange(r, o.col['Promoted at']).setValue(new Date());
    p.status = PR.status.real;
    p.promotedAt = new Date();
    promoted.push(p.id);
    try { if (typeof tpSpawnFromProposal_ === 'function') tpSpawnFromProposal_(p, true); } catch (err) { Logger.log('spawn failed: ' + err); }
  });
  if (promoted.length && typeof tpInvalidate_ === 'function') tpInvalidate_();
  return promoted;
}

// ---- mutations (client-callable) ----

// Admin decision on a provisional proposal: 'approve' promotes it without waiting for
// quorum (and spawns it as an Open project), 'decline' closes it (the row stays in the
// sheet for the record; the board only lists what is still up for review, plus a compact
// "in Projects" row for recently promoted ones).
function prAdminSetStatus(proposalId, decision) {
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), r = tpFindRow_(o.sh, o.col['Proposal ID'], proposalId);
  if (r < 0) return { ok: false, error: 'That proposal could not be found.' };
  if (decision === 'approve') {
    o.sh.getRange(r, o.col['Status']).setValue(PR.status.real);
    o.sh.getRange(r, o.col['Promoted at']).setValue(new Date());
    prInvalidate_();
    var p = prProposals_().filter(function (x) { return x.id === proposalId; })[0];
    try { if (p && typeof tpSpawnFromProposal_ === 'function') tpSpawnFromProposal_(p); } catch (err) { Logger.log('spawn failed: ' + err); }
  } else if (decision === 'decline') {
    o.sh.getRange(r, o.col['Status']).setValue(PR.status.declined);
  } else {
    return { ok: false, error: 'Unknown decision.' };
  }
  prInvalidate_();
  return { ok: true, status: decision === 'approve' ? PR.status.real : PR.status.declined };
}

// Admin: remove a proposal the swarm rejected, along with its ratings. Unlike Decline
// this leaves no row behind, so it is the one for junk and clear no-votes. Same gate as
// the projects module: only reachable from the unlisted admin page.
function prDeleteProposal(proposalId) {
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), r = tpFindRow_(o.sh, o.col['Proposal ID'], proposalId);
  if (r < 0) return { ok: false, error: 'That proposal could not be found.' };
  o.sh.deleteRow(r);
  // Drop its ratings too, bottom-up so earlier row numbers stay valid as we delete.
  var vo = tpOpen_(PR.votesSheet, PR.voteHeaders), last = vo.sh.getLastRow();
  if (last >= 2) {
    var vv = vo.sh.getRange(2, 1, last - 1, vo.sh.getLastColumn()).getValues(), kill = [];
    for (var i = 0; i < vv.length; i++) {
      if (String(vv[i][vo.col['Proposal ID'] - 1]).trim() === proposalId) kill.push(i + 2);
    }
    for (var j = kill.length - 1; j >= 0; j--) vo.sh.deleteRow(kill[j]);
  }
  prInvalidate_();
  return { ok: true };
}

// Admin edits the scout-entered fields in place: title, what & why, area, and name.
// Status, NetID, photos and ratings stay put.
function prUpdateProposal(proposalId, title, description, area, name) {
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders);
  var r = tpFindRow_(o.sh, o.col['Proposal ID'], proposalId);
  if (r < 0) return { ok: false, error: 'That proposal could not be found.' };
  var t = String(title || '').trim();
  if (!t) return { ok: false, error: 'A title is required.' };
  var desc = String(description || '').trim();
  var ar = String(area || '').trim();
  var nm = String(name || '').trim();
  o.sh.getRange(r, o.col['Title'], 1, 4).setValues([[t, desc, ar, nm]]);
  prInvalidate_();
  return { ok: true, title: t, description: desc, area: ar, proposedBy: nm };
}

function prSubmitProposal(title, description, area, name, netid, photoIds) {
  title = String(title || '').trim();
  if (!title) return { ok: false, error: 'A title is required.' };
  netid = String(netid || '').trim();
  if (!netid) return { ok: false, error: 'Enter your NetID first (top of the page).' };
  var o = tpOpen_(PR.proposalsSheet, PR.proposalHeaders), id = newToken_(), r = o.sh.getLastRow() + 1;
  // One setValues for the whole row instead of ten setValue calls.
  var row = [];
  for (var i = 0; i < PR.proposalHeaders.length; i++) row.push('');
  row[o.col['Proposal ID'] - 1] = id;
  row[o.col['Title'] - 1] = title;
  row[o.col['Description'] - 1] = String(description || '').trim();
  row[o.col['Area'] - 1] = String(area || '').trim();
  row[o.col['Proposed by'] - 1] = String(name || '').trim();
  row[o.col['NetID'] - 1] = netid;
  row[o.col['Status'] - 1] = PR.status.provisional;
  row[o.col['Created at'] - 1] = new Date();
  row[o.col['Photos'] - 1] = String(photoIds || '').trim();
  o.sh.getRange(r, 1, 1, row.length).setValues([row]);
  prInvalidate_();
  return { ok: true, id: id };
}

// One staged photo -> Drive. The create form fires these as soon as photos are picked,
// in parallel, so the files are already up by the time the proposal is posted.
// Returns the file id for the Photos cell.
function prUploadPhoto(dataUrl, filename) {
  try {
    return { ok: true, photoId: tpSaveUpload_(dataUrl, filename) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

// Because photos upload before the proposal is posted, one taken back out of the form
// would otherwise sit in the uploads folder forever. Only ever reached for a file this
// same form just created and then dropped.
function prTrashUpload(fileId) {
  try {
    DriveApp.getFileById(String(fileId || '')).setTrashed(true);
    return { ok: true };
  } catch (err) {
    return { ok: false };   // already gone, or not ours to remove
  }
}

// Cheap personalisation call. Reads ONLY the votes sheet: the page is rendered without
// any viewer state, then this fills in the viewer's own ratings and earned-voice
// progress. Deliberately does not rebuild the list.
function prViewerState(netid) {
  var vn = norm_(netid || '');
  if (!vn) return { ok: true, reviewCount: 0, myVotes: {} };
  var votes = prVotes_(), mine = {}, seen = {};
  votes.forEach(function (vt) {
    if (vt.netid !== vn) return;
    seen[vt.proposalId] = 1;
    mine[vt.proposalId] = { impact: vt.impact, effort: vt.effort };
  });
  return { ok: true, reviewCount: Object.keys(seen).length, myVotes: mine };
}

function prVote(proposalId, netid, impact, effort, comment) {
  netid = String(netid || '').trim();
  if (!netid) return { ok: false, error: 'Enter your NetID to vote.' };
  var imp = Number(impact), eff = Number(effort);
  if (!(imp >= 1 && imp <= 5) || !(eff >= 1 && eff <= 5)) return { ok: false, error: 'Rate impact and effort from 1 to 5.' };
  var p = prProposals_().filter(function (x) { return x.id === proposalId; })[0];
  if (!p) return { ok: false, error: 'That proposal could not be found.' };
  if (norm_(p.status) !== norm_(PR.status.provisional)) return { ok: false, error: 'Voting is closed, this one is already decided.' };
  if (norm_(p.netid) === norm_(netid)) return { ok: false, error: 'You cannot rate your own proposal.' };

  var o = tpOpen_(PR.votesSheet, PR.voteHeaders), last = o.sh.getLastRow(), found = -1;
  if (last >= 2) {
    var vv = o.sh.getRange(2, 1, last - 1, o.sh.getLastColumn()).getValues();
    for (var i = 0; i < vv.length; i++) {
      if (String(vv[i][o.col['Proposal ID'] - 1]).trim() === proposalId && norm_(vv[i][o.col['Voter NetID'] - 1]) === norm_(netid)) { found = i + 2; break; }
    }
  }
  var r = found > 0 ? found : o.sh.getLastRow() + 1;
  var row = [];
  for (var k = 0; k < PR.voteHeaders.length; k++) row.push('');
  row[o.col['Vote ID'] - 1] = found > 0 ? '' : newToken_();
  row[o.col['Proposal ID'] - 1] = proposalId;
  row[o.col['Voter NetID'] - 1] = netid;
  row[o.col['Impact'] - 1] = imp;
  row[o.col['Effort'] - 1] = eff;
  row[o.col['Comment'] - 1] = String(comment || '').trim();
  row[o.col['Voted at'] - 1] = new Date();
  if (found > 0) {
    // Keep the existing Vote ID on an update; overwrite the rest in one call.
    var startCol = o.col['Proposal ID'];
    o.sh.getRange(r, startCol, 1, row.length - startCol + 1).setValues([row.slice(startCol - 1)]);
  } else {
    o.sh.getRange(r, 1, 1, row.length).setValues([row]);
  }

  // Recount from a single fresh snapshot, promote anything that now qualifies, and hand
  // the client the numbers it needs to animate without a second round-trip.
  prInvalidate_();
  var data = prListData_(netid);
  var promotedIds = prPromoteAll_(data);
  prInvalidate_();
  var voice = prVoiceState_(data);
  var mine = data.proposals.filter(function (x) { return x.id === proposalId; })[0] || p;
  return {
    ok: true,
    updated: found > 0,
    promoted: promotedIds.indexOf(proposalId) >= 0,
    alsoPromoted: promotedIds.length - (promotedIds.indexOf(proposalId) >= 0 ? 1 : 0),
    counted: mine.countedVotes || 0,
    need: PR.quorumMinVotes,
    avgImpact: mine.avgImpact || 0,
    avgEffort: mine.avgEffort || 0,
    bar: PR.quorumMinImpact,
    voice: voice
  };
}

// ---- page ----

// A collapsible group: header + its cards. Native <details>, so it folds away with no
// JS, and secRestore remembers the choice per section.
function prSection_(label, count, cls, body, isOpen) {
  return '<details class="section"' + (isOpen ? ' open' : '') + '><summary class="section-head">'
    + '<span class="section-label ' + (cls || '') + '">' + label + '</span>'
    + '<span class="section-count">' + count + '</span>'
    + '<span class="section-chev" aria-hidden="true"></span></summary>'
    + '<div class="section-body">' + (body || '') + '</div></details>';
}

// The line under the quorum meter. Its whole job is to name the SMALLEST remaining step,
// because a goal that looks one move away pulls far harder than a distant one. The
// client mirrors these rules in prSay() after a vote lands.
function prSayText_(counted, need, avg, hasVotes) {
  if (!hasVotes) return 'be the first to weigh in';
  if (counted >= need && avg < PR.quorumMinImpact) return 'enough reviews, impact below the bar';
  var left = need - counted;
  if (left === 1) return 'one more review and it is in';
  if (left > 1) return left + ' more reviews to decide';
  return 'quorum reached';
}

// One line, three non-overlapping facts: how far along (segments), what would move it
// (the say line), and how it is scoring (the two averages). Deliberately does NOT also
// spell out "N of M reviews" in words, because the segments already say exactly that.
function prMeterHtml_(p, rid) {
  var need = PR.quorumMinVotes, counted = p.countedVotes || 0;
  var segs = '';
  for (var i = 0; i < need; i++) {
    segs += '<i class="pr-seg' + (i < counted ? ' on' : '') + '" style="animation-delay:' + (i * 70) + 'ms"></i>';
  }
  var good = counted && p.avgImpact >= PR.quorumMinImpact;
  return '<div class="pr-quorum">'
    + '<span class="pr-segs" id="' + rid + '-segs" title="' + counted + ' of ' + need + ' counted reviews">' + segs + '</span>'
    + '<span class="pr-qsay' + (need - counted === 1 ? ' pr-qsay--close' : (good && counted >= need ? ' pr-qsay--good' : '')) + '" id="' + rid + '-say">'
    +   prSayText_(counted, need, p.avgImpact, p.votes) + '</span>'
    + '<span class="pr-qvals" title="Average impact / effort">'
    +   '<b class="pr-qv' + (good ? ' good' : '') + '" id="' + rid + '-vi">' + (counted ? p.avgImpact.toFixed(1) : '&ndash;') + '</b>'
    +   '<span class="pr-qslash" aria-hidden="true">/</span>'
    +   '<b class="pr-qv" id="' + rid + '-ve">' + (counted ? p.avgEffort.toFixed(1) : '&ndash;') + '</b></span>'
    + '</div>';
}

// One proposal card. votable => provisional (shows the rating control); else it's an
// approved (Real) card.
function prCardHtml_(p, rid, votable, openByDefault) {
  var scale = function (k, val) {
    var out = '<span class="pr-scale" data-k="' + k + '" data-val="' + (val || '') + '">';
    for (var n = 1; n <= 5; n++) {
      out += '<button type="button" class="pr-dot' + (val && n <= val ? ' fill' : '') + (val === n ? ' on' : '') + '"'
        + ' data-n="' + n + '" aria-label="' + n + ' of 5" onclick="prPick(this)">' + n + '</button>';
    }
    return out + '</span>';
  };
  var my = p.myVote;
  var rate = votable
    ? '<div class="pr-rate" id="' + rid + '-rate">'
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Impact</span>'
      +   '<span class="pr-scale-wrap">' + scale('impact', my ? my.impact : 0) + '<span class="pr-word" id="' + rid + '-w-impact">' + (my ? prWordFor_('impact', my.impact) : '') + '</span></span></div>'
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Effort</span>'
      +   '<span class="pr-scale-wrap">' + scale('effort', my ? my.effort : 0) + '<span class="pr-word" id="' + rid + '-w-effort">' + (my ? prWordFor_('effort', my.effort) : '') + '</span></span></div>'
      + '<div class="pr-rate-foot"><button type="button" class="btn btn-primary pr-submit" id="' + rid + '-go" onclick="prSubmitVote(\'' + rid + '\',\'' + p.id + '\')">' + (my ? 'Update rating' : 'Submit rating') + '</button>'
      + '<span id="' + rid + '-vmsg" class="pr-vmsg"></span></div></div>'
    : '';

  // Admin-only decisions (revealed by tpAdminRevealJs_ / re-revealed after prRefresh).
  var adminBar = '<div class="tp-admin pr-adminbar" hidden id="' + rid + '-admin">'
    + '<button type="button" class="btn btn-ghost" onclick="prEditOpen(\'' + rid + '\')">Edit</button>'
    + (votable ? '<button type="button" class="btn btn-confirm" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'approve\')">Approve now</button>'
      + '<button type="button" class="btn btn-ghost pr-btn-decline" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'decline\')">Decline</button>' : '')
    + '<button type="button" class="btn btn-ghost pr-del" onclick="prDelAsk(\'' + rid + '\',\'' + p.id + '\')">Delete</button></div>';

  var edit = '<div id="' + rid + '-edit" class="ic-edit" hidden>'
    + '<label class="ic-edit-lbl">Title<input id="' + rid + '-etitle" class="ic-edit-in" value="' + escapeHtml_(p.title || '') + '"></label>'
    + '<label class="ic-edit-lbl">What &amp; why<textarea id="' + rid + '-edesc" class="ic-edit-in" rows="3">' + escapeHtml_(p.description || '') + '</textarea></label>'
    + '<label class="ic-edit-lbl">Area / space<input id="' + rid + '-earea" class="ic-edit-in" value="' + escapeHtml_(p.area || '') + '"></label>'
    + '<label class="ic-edit-lbl">Scouted by<input id="' + rid + '-ename" class="ic-edit-in" value="' + escapeHtml_(p.proposedBy || '') + '"></label>'
    + '<div class="ic-edit-btns"><button type="button" class="btn btn-primary" onclick="prEditSave(\'' + rid + '\',\'' + p.id + '\')">Save changes</button>'
    + '<button type="button" class="btn btn-ghost" onclick="prEditCancel(\'' + rid + '\')">Cancel</button>'
    + '<span id="' + rid + '-emsg" class="ic-edit-msg"></span></div></div>';

  // Photos the scout attached (Drive ids). onerror hides a thumb whose file will not load.
  var photoIds = extractFileIds_(p.photos);
  var photos = photoIds.length
    ? '<div class="pr-photos">' + photoIds.map(function (fid) {
        var e = encodeURIComponent(fid);
        return '<a href="https://drive.google.com/file/d/' + e + '/view" target="_blank" rel="noopener" title="Proposal photo"><img src="https://drive.google.com/thumbnail?id=' + e + '&sz=w400" loading="lazy" alt="" onerror="this.closest(\'a\').style.display=\'none\'"></a>';
      }).join('') + '</div>'
    : '';

  var meta = votable
    ? prMeterHtml_(p, rid)
    : '<div class="pr-promoted"><span class="pr-promoted-mark" aria-hidden="true">&#10003;</span>Promoted ' + (p.promotedAt ? escapeHtml_(fmtShort_(p.promotedAt)) : '') + ' with '
      + (p.countedVotes ? p.avgImpact.toFixed(1) : '&ndash;') + ' impact from ' + (p.countedVotes || 0) + ' reviews</div>';

  // Collapsed, a card is one scannable row: title, where it is, how far along, and how
  // it is scoring. Everything that takes vertical space (description, photos, the rating
  // control) lives in the body, so a board of dozens stays navigable.
  var summary = '<summary class="pr-sum">'
    + '<span class="pr-sum-main">'
    +   '<span class="pr-sum-title" id="' + rid + '-title">' + escapeHtml_(p.title) + '</span>'
    +   '<span class="pr-sum-sub">'
    +     '<span class="pr-area" id="' + rid + '-area"' + (p.area ? '' : ' hidden') + '>' + escapeHtml_(p.area || '') + '</span>'
    +     (photoIds.length ? '<span class="pr-sum-ph" title="' + photoIds.length + ' photo' + (photoIds.length === 1 ? '' : 's') + '">&#9634; ' + photoIds.length + '</span>' : '')
    +     '<span class="pr-mine" id="' + rid + '-mine"' + (my ? '' : ' hidden') + '>you rated this</span>'
    +   '</span>'
    + '</span>'
    + '<span class="pr-sum-stat">' + prSumStatHtml_(p, rid, votable) + '</span>'
    + '<span class="section-chev" aria-hidden="true"></span>'
    + '</summary>';

  return '<details class="card pr-card" id="' + rid + '" data-id="' + p.id + '" data-owner="' + escapeHtml_(p.netid) + '" data-status="' + (votable ? 'review' : 'real') + '"'
    + (openByDefault ? ' open' : '') + ' ontoggle="prCardToggled(this)">'
    + summary
    + '<div class="card-body">'
    +   '<div id="' + rid + '-view">'
    +   '<p class="pr-desc" id="' + rid + '-desc"' + (p.description ? '' : ' hidden') + '>' + escapeHtml_(p.description) + '</p>'
    +   '<div class="pr-meta" id="' + rid + '-meta"' + (p.proposedBy ? '' : ' hidden') + '>'
    +     '<span class="pr-by" id="' + rid + '-by"' + (p.proposedBy ? '' : ' hidden') + '>' + escapeHtml_(p.proposedBy || '') + '</span>'
    +   '</div>'
    +   photos
    +   meta
    +   rate
    +   '</div>'
    +   edit
    +   adminBar
    + '</div></details>';
}

// The compact right-hand stat on a collapsed card: the same quorum segments used in the
// open meter, plus the review count and average impact. Mirrors prMeterHtml_'s numbers so
// a folded board still shows at a glance which ideas have traction.
function prSumStatHtml_(p, rid, votable) {
  var counted = p.countedVotes || 0;
  if (!votable) {
    return '<span class="pr-sum-real"><span aria-hidden="true">&#10003;</span> In Projects</span>';
  }
  var need = PR.quorumMinVotes, segs = '';
  for (var i = 0; i < need; i++) segs += '<i class="pr-seg' + (i < counted ? ' on' : '') + '"></i>';
  var good = counted && p.avgImpact >= PR.quorumMinImpact;
  return '<span class="pr-sum-segs" id="' + rid + '-ssegs" title="' + counted + ' of ' + need + ' counted reviews">' + segs + '</span>'
    + '<span class="pr-sum-n" id="' + rid + '-sn">' + counted + '/' + need + '</span>'
    + '<span class="pr-sum-imp' + (good ? ' good' : '') + '" id="' + rid + '-si" title="Average impact">'
    +   (counted ? p.avgImpact.toFixed(1) : '&ndash;') + '</span>';
}

function prListSectionsHtml_(data) {
  var provisional = data.proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.provisional); });
  var cutoff = Date.now() - PR.freshDays * 86400000;
  var real = data.proposals.filter(function (p) {
    return norm_(p.status) === norm_(PR.status.real) && p.promotedAt && new Date(p.promotedAt).getTime() >= cutoff;
  });
  provisional.sort(function (a, b) { return (b.avgImpact - a.avgImpact) || (b.countedVotes - a.countedVotes) || (new Date(b.createdAt || 0) - new Date(a.createdAt || 0)); });
  real.sort(function (a, b) { return new Date(b.promotedAt || 0) - new Date(a.promotedAt || 0); });
  // Every idea starts as a summary row. Open one to rate it; the rest stay a list.
  var idx = 0, out = '';
  var body = provisional.length
    ? provisional.map(function (p) { return prCardHtml_(p, 'pr' + (idx++), true, false); }).join('')
    : '<div class="empty pr-empty"><span class="pr-empty-bee" aria-hidden="true">&#128029;</span>Nothing up for review yet. Scout an improvement and post it above.</div>';
  if (provisional.length > 1) {
    body = '<div class="pr-expandbar">'
      + '<button type="button" class="pr-expand" id="pr-expand" onclick="prExpandAll()">Expand all</button>'
      + '<span class="pr-expandhint">' + provisional.length + ' ideas, folded so you can scan them</span></div>' + body;
  }
  // The review group is open so you can scan summary rows. Each card stays folded
  // until tapped. "Now in Projects" stays shut.
  out += prSection_('Up for review', provisional.length, 'section-label--review', body, true);
  if (real.length) {
    out += prSection_('Now in Projects', real.length, 'section-label--real',
      real.map(function (p) { return prMovedHtml_(p); }).join(''), false);
  }
  return secList_(out);
}

function prMovedHtml_(p) {
  var href = prWebHref_('projects');
  return '<a class="pr-moved" href="' + escapeHtml_(href) + '" target="_blank" rel="noopener" onclick="return prStripGo(event,\'projects\')">'
    + '<span class="pr-moved-mark" aria-hidden="true">&#10003;</span>'
    + '<span class="pr-moved-title">' + escapeHtml_(p.title) + '</span>'
    + '<span class="pr-moved-go">In Projects &rarr;</span></a>';
}

// The list is identical for every viewer (personal state is patched in client-side), so
// it can be cached whole. A warm hit answers without touching the spreadsheet.
function prListSectionsCached_() {
  var c = prCacheSvc_(), hit = null;
  if (c) { try { hit = c.get(PR.cacheKey); } catch (e) { hit = null; } }
  if (hit) return hit;
  var html = prListSectionsHtml_(prListData_(''));
  if (c && html.length < 90000) { try { c.put(PR.cacheKey, html, PR.cacheSeconds); } catch (e) { /* best effort */ } }
  return html;
}

// Client-callable: fresh list HTML (used after posting, deleting or an admin decision).
function prListHtml() {
  return { ok: true, html: prListSectionsCached_() };
}

// ---- the strip shown on the students' tasks pages ----

// A slim announcement for the open-tasks views: what just became a real project, and
// what still needs rating. Cached, and reads only the proposals sheet, so adding it to
// the hot tasks path costs almost nothing.
function prTasksStripHtml_() {
  var c = prCacheSvc_(), hit = null;
  if (c) { try { hit = c.get(PR.stripKey); } catch (e) { hit = null; } }
  if (hit !== null && hit !== undefined) return hit;

  var html = '';
  try {
    var proposals = prProposals_();
    var cutoff = new Date().getTime() - PR.freshDays * 86400000;
    var freshProj = [];
    try {
      if (typeof tpListProjects_ === 'function') {
        freshProj = tpListProjects_().filter(function (p) {
          var st = norm_(p.status);
          return (st === norm_(TP.projectStatus.assigned) || !p.status) && tpIsFresh_(p);
        }).sort(function (a, b) { return tpTime_(b.createdAt) - tpTime_(a.createdAt); });
      }
    } catch (e2) { freshProj = []; }
    var freshReal = proposals.filter(function (p) {
      return norm_(p.status) === norm_(PR.status.real) && p.promotedAt && new Date(p.promotedAt).getTime() >= cutoff;
    }).sort(function (a, b) { return new Date(b.promotedAt || 0) - new Date(a.promotedAt || 0); });
    var open = proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.provisional); });

    var news = freshProj.length ? freshProj : freshReal;
    var parts = [];
    if (news.length) {
      var line = '<b>' + escapeHtml_(news[0].title) + '</b>'
        + (news.length > 1 ? ' +' + (news.length - 1) : '')
        + ' is waiting in Projects';
      parts.push(prStripLink_('projects', line, 'pr-strip--project', news.length));
    }
    if (open.length) {
      var ideas = '<b>' + open.length + ' idea' + (open.length === 1 ? '' : 's') + '</b> awaiting ratings';
      parts.push(prStripLink_('proposals', ideas, 'pr-strip--ideas', 0));
    }
    html = parts.length ? '<div class="pr-strips" data-pr-news="' + news.length + '">' + parts.join('') + '</div>' : '';
  } catch (err) {
    html = '';   // the tasks page must never fail because proposals are not set up yet
  }
  if (c) { try { c.put(PR.stripKey, html, PR.cacheSeconds); } catch (e) { /* best effort */ } }
  return html;
}

function prStripLink_(module, line, cls, news) {
  var href = prWebHref_(module);
  return '<a class="pr-strip ' + cls + '" href="' + escapeHtml_(href) + '" target="_blank" rel="noopener" onclick="return prStripGo(event,\'' + module + '\')"'
    + (news ? ' data-pr-news="' + news + '"' : '') + '>'
    + '<span class="pr-strip-bee" aria-hidden="true">&#128029;</span>'
    + '<span class="pr-strip-title">' + line + '</span>'
    + '<span class="pr-strip-go" aria-hidden="true">&rarr;</span></a>';
}

// Styles for the strip only. The tasks pages do not load prStyles_.
function prStripStyles_() {
  return '<style>'
    + '.pr-strips{display:flex;flex-direction:column;gap:8px;margin:18px 0 2px}'
    + '.pr-strip{display:flex;align-items:center;gap:12px;padding:12px 16px;border-radius:16px;text-decoration:none;'
    +   'background:linear-gradient(135deg,#fffdf5 0%,#fff8e6 100%);border:1.5px solid #f0e2b8;box-shadow:0 2px 10px rgba(120,90,20,.07);transition:transform .16s,box-shadow .16s}'
    + '.pr-strip:hover{transform:translateY(-2px);box-shadow:0 10px 24px rgba(120,90,20,.14)}'
    + '.pr-strip--ideas{background:#faf9f6;border-color:#e7e4dc;box-shadow:none}'
    + '.pr-strip--ideas:hover{box-shadow:0 8px 18px rgba(20,20,30,.08)}'
    + '.pr-strip-bee{font-size:22px;line-height:1;flex:none}'
    + '.pr-strip-title{flex:1;min-width:0;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#3f3a31;line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
    + '.pr-strip-title b{font-weight:800;color:#14110e}'
    + '.pr-strip-go{font-size:19px;color:#a8791a;flex:none}'
    + '@media(max-width:560px){.pr-strip{padding:11px 12px;gap:10px}.pr-strip-title{font-size:14px}}'
    + '@media(prefers-reduced-motion:reduce){.pr-strip{transition:none}.pr-strip:hover{transform:none}}'
    + '</style>';
}

function prStripJs_() {
  return '<script>'
    + 'function prHubMsg(action,extra){try{var m={source:"ops-hub",action:action};if(extra)for(var k in extra)m[k]=extra[k];'
    + 'if(window.parent&&window.parent!==window)window.parent.postMessage(m,"*");'
    + 'if(window.top&&window.top!==window)window.top.postMessage(m,"*");}catch(e){}}'
    + 'function prStripGo(ev,dest){prHubMsg(dest==="projects"?"open-projects":"open-proposals");'
    + 'if(/[?&]embed=1(?:&|$)/.test(location.search||"")){if(ev)ev.preventDefault();return false;}return true;}'
    + '(function(){var n=0,el=document.querySelector("[data-pr-news]");if(el)n=parseInt(el.getAttribute("data-pr-news"),10)||0;'
    + 'if(n>0)prHubMsg("projects-news",{count:n});})();'
    + '</script>';
}

// ---- the page ----

// Only ever hand an http(s) URL to an href we build from a query parameter.
function prSafeUrl_(u) {
  u = String(u || '').trim();
  return /^https?:\/\//i.test(u) ? u : '';
}

function proposalsPage_(embedded, admin, back) {
  var backHref = prSafeUrl_(back);
  var inner = '<div class="pr-page">';
  if (!embedded) {
    inner += '<div class="pr-top">'
      + (backHref ? '<a class="pr-back" href="' + escapeHtml_(backHref) + '" target="_top" rel="noopener"><span class="pr-back-arrow" aria-hidden="true">&#8592;</span> Back to the hub</a>' : '')
      + '</div>'
      + '<div class="page-head"><div class="page-kicker">Project Teams Ops Hub</div><div class="page-title">Proposals</div><div class="page-rule"></div></div>';
  }
  inner += '<div class="pr-hero">'
    + '<p class="pr-intro">Scout an improvement to our spaces and post it. Everyone rates each one on impact and effort, and the ideas that win enough support move to Projects for the team to pick up.</p>';

  // Compact honey chip in the hero; opens in place so the rules sit next to the board
  // rather than as another stacked card.
  inner += '<details class="pr-guide"><summary class="pr-guide-toggle">'
    + '<span class="pr-guide-icon" aria-hidden="true">&#128029;</span>How this works'
    + '<span class="pr-guide-fly" aria-hidden="true">&#128029;</span>'
    + '<span class="pr-guide-caret" aria-hidden="true">&#9662;</span></summary>'
    + '<div class="pr-guide-body">'
    + '<p class="pr-guide-lede">This runs like a honeybee swarm picking a new home. Bees send scouts out, each one reports back on what it found, and the hive commits only once enough independent scouts agree. Same idea here, with our spaces.</p>'
    + '<ol class="pr-guide-steps">'
    +   '<li><b>Scout.</b> When you spot something that would make a space work better, post it. A title is all that is required; a photo of the spot helps everyone else judge it.</li>'
    +   '<li><b>Rate, independently.</b> Score every proposal twice: <b>Impact</b> (how much better it makes the place) and <b>Effort</b> (how much work it is).</li>'
    +   '<li><b>Earn your voice.</b> Your ratings start counting once you have reviewed <b>' + (PR.earnedVoiceK + 1) + ' proposals</b>. Until then they are saved but not counted.</li>'
    +   '<li><b>Quorum.</b> A proposal becomes a real project when it has <b>' + PR.quorumMinVotes + ' counted reviews</b> and an average Impact of <b>' + PR.quorumMinImpact + ' or better</b>. It then moves to <b>Projects</b>, where staff can join it.</li>'
    + '</ol>'
    + '<div class="pr-guide-notes">'
    +   '<p><b>You cannot rate your own proposal.</b></p>'
    + '</div></div></details></div>';

  // One toolbar: who you are (so ratings can count) and the act of posting. Kept on
  // one row so the board is not buried under stacked boxes.
  inner += '<div class="pr-toolbar">'
    + '<div class="pr-idbar">'
    +   '<div class="pr-idrow"><label class="pr-idlbl" for="pr-netid">Rating as</label>'
    +     '<input id="pr-netid" class="pr-idinput" placeholder="your NetID" autocomplete="username" spellcheck="false" aria-describedby="pr-idnote" oninput="prSaveNetid()"></div>'
    +   '<div class="pr-voice" id="pr-voice">'
    +     '<div class="pr-voice-segs" id="pr-voice-segs" hidden></div>'
    +     '<span class="pr-voice-note" id="pr-idnote"></span></div>'
    + '</div>'
    + '<details class="tp-create-wrap pr-propose"><summary class="tp-create-toggle"><span class="tp-create-caret">&#43;</span> Propose an improvement</summary>'
    + '<div class="tp-create">'
    + '<div class="tp-field"><label>Title <span class="tp-req">Required</span></label><input id="pr-c-title" placeholder="e.g. Add a shadow board to the hand-tool wall"></div>'
    + '<div class="tp-field"><label>What &amp; why</label><textarea id="pr-c-desc" rows="3" placeholder="What to change, and why it helps"></textarea></div>'
    + '<div class="pr-create-split">'
    +   '<div class="tp-field"><label>Area / space</label><input id="pr-c-area" placeholder="e.g. Baja bay, composites lab"></div>'
    +   '<div class="tp-field"><label>Your name</label><input id="pr-c-name" placeholder="Your name"></div></div>'
    + '<div class="tp-field"><label>Photos <span class="pr-optional">Optional, show the spot you want to improve</span></label>'
    +   '<input type="file" accept="image/*" multiple id="pr-c-file" style="display:none" onchange="prPickPhotos(this)">'
    +   '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'pr-c-file\').click()">Add photos</button>'
    +   '<div id="pr-c-stage"></div></div>'
    + '<div class="pr-create-foot"><button type="button" class="btn btn-primary pr-btn-post" id="pr-c-go" onclick="prCreate()">Post for review</button>'
    +   '<span id="pr-c-msg" class="tp-lock-msg"></span></div>'
    + '</div></details></div>';

  inner += '<div class="pr-listwrap"><div class="pr-sync" id="pr-sync" hidden></div>'
    + '<div id="pr-list">' + prListSectionsCached_() + '</div></div></div>';

  var boot = '<script>var PR_NEED=' + PR.quorumMinVotes + ',PR_BAR=' + PR.quorumMinImpact
    + ',PR_VOICE_NEED=' + (PR.earnedVoiceK + 1) + ';</script>';

  return swissShell_('<div class="pr-topbar" id="pr-top" aria-hidden="true"></div>'
    + tpStyles_() + prStyles_() + inner + boot + tpSharedJs_() + prStripJs_() + prClientJs_() + tpAdminRevealJs_(admin),
    'Proposals', true, embedded);
}

// ---- words ----

// Naming each number turns an abstract 1-5 into a judgement you can agree or disagree
// with, and gives instant feedback the moment a dot is pressed. Mirrored in prWord() JS.
function prWordFor_(kind, n) {
  var impact = ['', 'barely noticed', 'a small lift', 'a solid win', 'a big deal', 'transforms the space'];
  var effort = ['', 'minutes', 'an afternoon', 'a weekend', 'a real build', 'a major undertaking'];
  return (kind === 'impact' ? impact : effort)[n] || '';
}

// The impact/effort pair, read back as one verdict. Completing the pair is the moment
// the rating becomes a judgement, so it is worth marking.
function prVerdictText_(imp, eff) {
  if (!imp || !eff) return '';
  var hiI = imp >= 4, loE = eff <= 2;
  if (hiI && loE) return '<b>Quick win.</b> Big payoff, little work.';
  if (hiI && eff >= 4) return '<b>Big swing.</b> Worth it, but it is a project.';
  if (imp <= 2 && loE) return '<b>Easy tidy-up.</b> Small, but cheap to do.';
  if (imp <= 2 && eff >= 4) return '<b>Hard pass.</b> Lots of work for little gain.';
  return '<b>Middle of the road.</b> Reasonable either way.';
}

function prStyles_() {
  return '<style>'
    // ---- page frame ----
    + '.pr-page{max-width:100%;min-width:0}'
    + '.pr-page .page-head{margin-bottom:2px}'
    + '.pr-page .page-title{letter-spacing:-.04em}'
    + '.pr-page .page-rule{width:56px;background:linear-gradient(90deg,#c08a1e,#f0c050,#b31b1b)}'
    + '.pr-topbar{position:fixed;top:0;left:0;height:3px;width:0;z-index:60;background:linear-gradient(90deg,#b31b1b,#e08a1e);opacity:0;transition:width .3s ease,opacity .25s}'
    + '.pr-topbar.on{width:88%;opacity:1;transition:width 9s cubic-bezier(.05,.8,.25,1),opacity .2s}'
    + '.pr-topbar.done{width:100%;opacity:0;transition:width .2s,opacity .45s .15s}'
    + '.pr-top{margin-bottom:14px}'
    + '.pr-back{display:inline-flex;align-items:center;gap:8px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:800;letter-spacing:-.01em;color:#8f1515;text-decoration:none;padding:9px 14px 9px 12px;border-radius:10px;background:#fff;border:1.5px solid #ead0d0;box-shadow:0 1px 2px rgba(20,17,14,.05);transition:color .15s,background .15s,border-color .15s,box-shadow .15s,transform .15s}'
    + '.pr-back:hover{color:#fff;background:#b31b1b;border-color:#b31b1b;box-shadow:0 6px 16px rgba(179,27,27,.22);transform:translateY(-1px)}'
    + '.pr-back-arrow{font-size:15px;line-height:1}'

    // ---- hero: intro + compact honey chip ----
    + '.pr-hero{margin:14px 0 2px}'
    + '.pr-intro{font-size:15.5px;line-height:1.65;color:#57534e;margin:0 0 12px;overflow-wrap:anywhere}'
    + '.pr-guide{position:relative;display:inline-flex;max-width:100%;vertical-align:middle;margin:0;'
    +   'border:1.5px solid #ead9a8;border-radius:999px;overflow:visible;'
    +   'background:linear-gradient(180deg,#fffef6 0%,#fff4d6 100%);'
    +   'box-shadow:0 1px 2px rgba(140,100,20,.08),inset 0 1px 0 rgba(255,255,255,.8);'
    +   'transition:border-color .2s,box-shadow .2s,border-radius .2s}'
    + '.pr-guide:hover{border-color:#e0c47a;box-shadow:0 6px 18px rgba(168,121,26,.14),inset 0 1px 0 rgba(255,255,255,.9)}'
    + '.pr-guide[open]{display:block;border-radius:18px;overflow:hidden;border-color:#e0c47a;'
    +   'background:radial-gradient(ellipse 80% 120% at 0% 0%,rgba(240,192,80,.2),transparent 55%),linear-gradient(165deg,#fffef8 0%,#fff8e6 100%);'
    +   'box-shadow:0 10px 28px rgba(168,121,26,.14),inset 0 1px 0 rgba(255,255,255,.9)}'
    + '.pr-guide-toggle{display:inline-flex;align-items:center;gap:8px;padding:8px 14px 8px 12px;cursor:pointer;list-style:none;position:relative;z-index:1;'
    +   'font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:800;letter-spacing:.02em;color:#8a5f0f}'
    + '.pr-guide[open] .pr-guide-toggle{display:flex;padding:14px 18px}'
    + '.pr-guide-toggle::-webkit-details-marker{display:none}'
    + '.pr-guide-icon{font-size:16px;line-height:1;display:inline-block;transform-origin:50% 80%;'
    +   'animation:prBeeBuzz 2.15s ease-in-out infinite;filter:drop-shadow(0 2px 2px rgba(140,100,20,.22))}'
    + '.pr-guide-toggle:hover .pr-guide-icon{animation:prBeeHop .55s cubic-bezier(.2,.9,.3,1.5)}'
    + '.pr-guide[open] .pr-guide-icon{animation:prBeeHop .5s cubic-bezier(.2,.9,.3,1.4) both,prBeeBuzz 2.15s ease-in-out .5s infinite}'
    + '.pr-guide-fly{position:absolute;top:10px;right:44px;font-size:12px;opacity:0;pointer-events:none;filter:drop-shadow(0 1px 1px rgba(140,100,20,.2))}'
    + '.pr-guide[open] .pr-guide-fly{animation:prBeeDrift 5.4s ease-in-out .12s infinite}'
    + '.pr-guide-caret{margin-left:4px;font-size:10px;transition:transform .22s cubic-bezier(.2,.8,.3,1);color:#c08a1e}'
    + '.pr-guide[open] .pr-guide-caret{margin-left:auto;transform:rotate(180deg)}'
    + '.pr-guide-body{padding:4px 18px 18px;border-top:1.5px solid #f0e4c4;position:relative;z-index:1;animation:prCardIn .32s ease both}'
    + '.pr-guide-lede{font-size:14px;line-height:1.65;color:#57534e;margin:12px 0 0}'
    + '.pr-guide-steps{margin:14px 0 0;padding:0;list-style:none;display:grid;grid-template-columns:1fr 1fr;gap:10px;counter-reset:prstep}'
    + '.pr-guide-steps li{counter-increment:prstep;font-size:13.5px;line-height:1.55;color:#57534e;margin:0;padding:13px 14px 13px 42px;position:relative;'
    +   'background:rgba(255,255,255,.72);border:1px solid rgba(224,196,122,.4);border-radius:14px}'
    + '.pr-guide-steps li::before{content:counter(prstep);position:absolute;left:12px;top:13px;width:22px;height:22px;border-radius:8px;'
    +   'display:flex;align-items:center;justify-content:center;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;color:#8a5f0f;'
    +   'background:linear-gradient(180deg,#fff6d6,#f0d78a);box-shadow:inset 0 1px 0 rgba(255,255,255,.8)}'
    + '.pr-guide-steps b,.pr-guide-notes b{color:#14110e;font-weight:800}'
    + '.pr-guide-notes{margin-top:14px;padding-top:12px;border-top:1.5px dashed #f0e4cc;display:flex;flex-direction:column;gap:6px}'
    + '.pr-guide-notes p{margin:0;font-size:12.5px;line-height:1.55;color:#8a857c}'
    + '@keyframes prBeeBuzz{0%,100%{transform:translate(0,0) rotate(-7deg)}25%{transform:translate(1px,-3px) rotate(9deg)}'
    +   '50%{transform:translate(-1px,-1px) rotate(-5deg)}75%{transform:translate(2px,-4px) rotate(11deg)}}'
    + '@keyframes prBeeHop{0%{transform:translate(0,0) rotate(-6deg) scale(1)}40%{transform:translate(3px,-6px) rotate(16deg) scale(1.18)}'
    +   '100%{transform:translate(0,0) rotate(-6deg) scale(1)}}'
    + '@keyframes prBeeDrift{0%{opacity:0;right:6%;top:11px;transform:rotate(-18deg) scale(.75)}'
    +   '14%{opacity:.95}52%{right:46%;top:5px;transform:rotate(16deg) scale(1)}'
    +   '86%{opacity:.85}100%{opacity:0;right:78%;top:15px;transform:rotate(-8deg) scale(.8)}}'

    // ---- toolbar: identity + propose on one row ----
    + '.pr-toolbar{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:16px 0 6px;max-width:100%;min-width:0}'
    + '.pr-idbar{display:flex;align-items:center;gap:10px 18px;flex-wrap:wrap;flex:1 1 12rem;min-width:0;margin:0;padding:10px 14px;'
    +   'background:#fff;border:1px solid #ece6d8;border-radius:14px;box-shadow:0 1px 2px rgba(20,17,14,.04)}'
    + '.pr-idrow{display:flex;align-items:center;gap:10px;flex:0 1 auto;min-width:0}'
    + '.pr-idlbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#8a857c;white-space:nowrap}'
    + '.pr-idinput{font:inherit;font-size:14px;padding:9px 12px;border:1.5px solid #e8e2d4;border-radius:10px;outline:none;background:#fafaf8;min-width:0;width:14ch;transition:border-color .15s,box-shadow .15s,background .15s}'
    + '.pr-idinput:focus{border-color:#c08a1e;box-shadow:0 0 0 4px rgba(192,138,30,.14);background:#fff}'
    + '.pr-voice{display:flex;align-items:center;gap:10px;flex:1 1 200px;min-width:0}'
    + '.pr-voice-segs{display:inline-flex;gap:4px;flex:none}'
    + '.pr-voice-seg{width:22px;height:6px;border-radius:99px;background:#eceae3;transition:background .45s cubic-bezier(.2,.9,.3,1.3),transform .45s cubic-bezier(.2,.9,.3,1.3)}'
    + '.pr-voice-seg.on{background:linear-gradient(90deg,#e0a11e,#e08a1e)}'
    + '.pr-voice-seg.pop{transform:scaleY(1.7)}'
    + '.pr-voice.done .pr-voice-seg.on{background:linear-gradient(90deg,#1d9d5b,#157a47)}'
    + '.pr-voice-note{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:600;color:#8a857c;line-height:1.35;min-width:0}'
    + '.pr-voice.done .pr-voice-note{color:#157a47}'
    + '.pr-propose{flex:0 0 auto;margin:0}'
    + '.pr-propose[open]{flex:1 1 100%}'
    + '.pr-page .pr-propose .tp-create-toggle{color:#8a5f0f;background:linear-gradient(180deg,#fffef8,#fff4d6);border:1.5px solid #ead9a8;border-radius:12px;padding:10px 16px;font-weight:800;box-shadow:0 1px 2px rgba(140,100,20,.08),inset 0 1px 0 rgba(255,255,255,.8)}'
    + '.pr-page .pr-propose .tp-create-toggle:hover{border-color:#e0c47a;color:#6d4b0c}'
    + '.pr-page .pr-propose[open] .tp-create-toggle{background:#fff;border-color:#ece6d8;color:#14110e;box-shadow:none}'
    + '.pr-page .pr-propose .tp-create{margin-top:10px;border-color:#ece6d8;border-radius:16px;box-shadow:0 8px 24px rgba(20,17,14,.06)}'
    + '.pr-create-split{display:grid;grid-template-columns:1fr 1fr;gap:0 12px}'
    + '.pr-optional{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:600;color:#a8a29e;letter-spacing:0;text-transform:none;margin-left:6px}'
    + '.pr-create-foot{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:4px}'
    + '.pr-stage-state{position:absolute;left:6px;bottom:6px;display:inline-flex;align-items:center;justify-content:center;min-width:21px;height:21px;padding:0 7px;border-radius:99px;'
    +   'background:rgba(20,17,14,.72);color:#fff;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;line-height:1}'
    + '.stage-item.pr-stage-uploading img{opacity:.6}'
    + '.stage-item.pr-stage-done .pr-stage-state{background:#157a47}'
    + '.stage-item.pr-stage-error .pr-stage-state{background:#b31b1b;cursor:pointer}'
    + '.stage-item.pr-stage-error img{opacity:.55}'
    + '.pr-stage-state .pr-spin{width:11px;height:11px;border-width:2px}'

    // ---- list + cards ----
    + '.pr-listwrap{position:relative}'
    + '.pr-sync{position:absolute;top:-2px;left:0;right:0;height:2px;border-radius:99px;overflow:hidden;background:#eceae3;z-index:3}'
    + '.pr-sync::after{content:"";position:absolute;inset:0;width:38%;border-radius:99px;background:linear-gradient(90deg,transparent,#c08a1e,transparent);animation:prSweep 1.15s linear infinite}'
    + '@keyframes prSweep{0%{transform:translateX(-110%)}100%{transform:translateX(320%)}}'
    + '#pr-list.is-syncing{opacity:.55;transition:opacity .2s}'
    + '.pr-page .section-label--review{color:#c08a1e}'
    + '.pr-page .section-label--real{color:#157a47}'
    + '.pr-page .section-rule{background:linear-gradient(90deg,#efe6d2,transparent)}'
    + '.pr-empty{display:flex;align-items:center;justify-content:center;gap:10px;flex-wrap:wrap;padding:28px 22px;'
    +   'background:linear-gradient(180deg,#fffef8,#fff8e8);border:1.5px dashed #e8d49a;color:#8a5f0f}'
    + '.pr-empty-bee{font-size:20px;line-height:1;animation:prBeeBuzz 2.4s ease-in-out infinite}'
    + '.pr-page .pr-card{position:relative;overflow:hidden;margin-bottom:14px;border-radius:18px;border:1px solid #ece6d8;border-left:4px solid #e0a11e;'
    +   'box-shadow:0 1px 2px rgba(20,17,14,.04),0 10px 28px -16px rgba(20,17,14,.18);animation:prCardIn .42s cubic-bezier(.2,.85,.3,1.05) both}'
    + '.pr-page .pr-card[data-status="real"]{border-left-color:#157a47}'
    + '.pr-page .pr-card:hover{transform:translateY(-2px);box-shadow:0 14px 36px -14px rgba(20,17,14,.2)}'
    + '.pr-page .pr-card .card-body{padding:2px 20px 16px}'
    + '@keyframes prCardIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}'
    // ---- collapsed card: the summary row ----
    // The whole row is the hit target. Title left, progress right, so a folded column of
    // dozens can be scanned down either edge.
    + '.pr-sum{display:flex;align-items:center;gap:12px;padding:12px 14px;min-height:52px;cursor:pointer;list-style:none;-webkit-user-select:none;user-select:none;-webkit-tap-highlight-color:transparent;touch-action:manipulation;transition:background .15s}'
    + '.pr-sum::-webkit-details-marker{display:none}'
    + '.pr-sum::marker{content:""}'
    + '.pr-sum:hover{background:rgba(201,134,12,.05)}'
    + '.pr-card[open] > .pr-sum{padding-bottom:6px}'
    + '.pr-sum-main{flex:1;min-width:0}'
    + '.pr-sum-title{display:block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:clamp(15px,2.1vw,17px);font-weight:800;letter-spacing:-.025em;line-height:1.3;color:#14110e;overflow-wrap:anywhere}'
    + '.pr-sum-sub{display:flex;align-items:center;flex-wrap:wrap;gap:8px;margin-top:3px;font-size:12px;font-weight:600;color:#8a857c;line-height:1.4}'
    + '.pr-sum-sub:empty{display:none}'
    + '.pr-sum-ph{color:#a8a29e}'
    + '.pr-mine{color:#157a47;font-weight:700}'
    + '.pr-sum-stat{display:flex;align-items:center;gap:9px;flex:none}'
    + '.pr-sum-segs{display:inline-flex;gap:3px}'
    + '.pr-sum-segs .pr-seg{width:14px;height:6px;animation:none}'
    + '.pr-sum-n{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;color:#8a857c;font-variant-numeric:tabular-nums}'
    + '.pr-sum-imp{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:800;color:#a8a29e;min-width:24px;text-align:right;font-variant-numeric:tabular-nums}'
    + '.pr-sum-imp.good{color:#157a47}'
    + '.pr-sum-real{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;color:#157a47;white-space:nowrap}'
    + '.pr-sum:active{background:#f3f0ea}'
    // Folded cards sit tighter together than open ones so a long list stays compact.
    + '.pr-page .pr-card:not([open]){margin-bottom:8px}'
    + '.pr-expandbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:0 0 12px}'
    + '.pr-expand{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#57534e;background:#fff;'
    +   'border:1.5px solid #e2ddd6;border-radius:9px;padding:7px 13px;cursor:pointer;transition:border-color .15s,color .15s}'
    + '.pr-expand:hover{border-color:#c9c2b8;color:#14110e}'
    + '.pr-expandhint{font-size:12px;color:#a8a29e}'
    + '@media(max-width:560px){.pr-sum{padding:13px 14px;gap:9px}.pr-sum-n{display:none}}'
    + '.pr-desc{font-size:14.5px;line-height:1.55;color:#57534e;margin:0;overflow-wrap:anywhere}'
    + '.pr-meta{display:flex;align-items:center;flex-wrap:wrap;gap:0;margin-top:8px;font-size:12.5px;font-weight:600;color:#8a857c}'
    + '.pr-photos{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,92px),1fr));gap:8px;margin-top:12px}'
    + '.pr-photos a{display:block;aspect-ratio:1;border-radius:12px;overflow:hidden;line-height:0;box-shadow:0 2px 8px rgba(20,17,14,.08);transition:transform .16s}'
    + '.pr-photos a:hover{transform:scale(1.04)}'
    + '.pr-photos img{width:100%;height:100%;object-fit:cover;display:block;border:0}'

    // ---- quorum as an instrument, not a nested box ----
    + '.pr-quorum{display:flex;align-items:center;gap:10px 14px;flex-wrap:wrap;margin-top:16px;padding:12px 0 0;border-top:1.5px solid #f0ebe0;background:none;border-radius:0}'
    + '.pr-qsay{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#a8a29e;transition:color .3s;flex:1 1 auto;min-width:0}'
    + '.pr-qsay--close{color:#c9761a}'
    + '.pr-qsay--good{color:#157a47}'
    + '.pr-qvals{display:inline-flex;align-items:baseline;gap:5px;flex:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-variant-numeric:tabular-nums}'
    + '.pr-qv{font-size:13.5px;font-weight:800;color:#14110e}'
    + '.pr-qv.good{color:#157a47}'
    + '.pr-qslash{color:#d6d3ce;font-weight:700;font-size:12px}'
    + '.pr-segs{display:inline-flex;flex-wrap:wrap;gap:5px;flex:none;max-width:100%}'
    + '.pr-seg{width:28px;height:7px;border-radius:99px;background:#ece6d8;animation:prSegIn .4s cubic-bezier(.2,.9,.3,1.2) both}'
    + '.pr-seg.on{background:linear-gradient(90deg,#e0a11e,#c08a1e)}'
    + '.pr-seg.pending{background:repeating-linear-gradient(115deg,#e6ded2,#e6ded2 5px,#f3ede4 5px,#f3ede4 10px);animation:prPend 1s linear infinite}'
    + '.pr-seg.pop{animation:prSegPop .5s cubic-bezier(.2,.9,.3,1.5)}'
    + '@keyframes prSegIn{from{transform:scaleX(.25);opacity:0}to{transform:none;opacity:1}}'
    + '@keyframes prSegPop{0%{transform:scaleY(1)}45%{transform:scaleY(2.1)}100%{transform:scaleY(1)}}'
    + '@keyframes prPend{to{background-position:20px 0}}'
    + '.pr-promoted{display:flex;align-items:center;gap:9px;margin-top:16px;padding:11px 0 0;border-top:1.5px solid #e4f0e8;background:none;border-radius:0;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:600;color:#157a47}'
    + '.pr-promoted-mark{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:#157a47;color:#fff;font-size:11px;font-weight:800;flex:none}'
    + '.pr-moved{display:flex;align-items:center;gap:12px;margin:0 0 8px;padding:12px 16px;border-radius:14px;text-decoration:none;background:#f4faf6;border:1.5px solid #d7eadf;transition:transform .16s,box-shadow .16s}'
    + '.pr-moved:hover{transform:translateY(-1px);box-shadow:0 8px 18px rgba(21,122,71,.12)}'
    + '.pr-moved-mark{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:#157a47;color:#fff;font-size:12px;font-weight:800;flex:none}'
    + '.pr-moved-title{flex:1;min-width:0;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;font-weight:800;color:#14110e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
    + '.pr-moved-go{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:800;color:#157a47;flex:none;white-space:nowrap}'

    // ---- rating: two scales, two colours, sitting on the card ----
    + '.pr-rate{margin-top:2px;padding:14px 0 0;background:none;border:none;border-radius:0;min-width:0}'
    + '.pr-rate-row{display:flex;align-items:center;flex-wrap:wrap;gap:8px 12px;margin-bottom:10px;min-width:0}'
    + '.pr-rate-lbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:#8a857c;flex:0 0 64px}'
    + '.pr-scale-wrap{display:flex;align-items:center;flex-wrap:wrap;gap:8px 10px;flex:1 1 12rem;min-width:0}'
    + '.pr-scale{display:flex;gap:6px;flex:1 1 auto;min-width:0;max-width:100%}'
    + '.pr-dot{flex:1 1 0;min-width:0;max-width:40px;width:auto;height:36px;border-radius:10px;border:1.5px solid #e8e2d4;background:#fff;padding:0;cursor:pointer;'
    +   'font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:800;color:#a8a29e;line-height:1;'
    +   'transition:border-color .14s,background .14s,color .14s,transform .16s cubic-bezier(.2,.9,.3,1.6),box-shadow .16s;-webkit-tap-highlight-color:transparent}'
    + '.pr-dot:hover{border-color:#c4b89a;color:#57534e;background:#fffdf8}'
    + '.pr-dot.fill{background:#f3d2d2;border-color:#e6b4b4;color:#8f1515}'
    + '.pr-dot.on{background:linear-gradient(180deg,#d62b2b,#b31b1b);border-color:#b31b1b;color:#fff;box-shadow:0 3px 8px rgba(179,27,27,.28)}'
    + '.pr-dot.preview{border-color:#d99a9a;background:#fdf3f3;color:#8f1515}'
    + '.pr-scale[data-k="effort"] .pr-dot.fill{background:#f3e4b8;border-color:#e0c47a;color:#8a5f0f}'
    + '.pr-scale[data-k="effort"] .pr-dot.on{background:linear-gradient(180deg,#e0a11e,#c08a1e);border-color:#c08a1e;color:#fff;box-shadow:0 3px 8px rgba(192,138,30,.3)}'
    + '.pr-scale[data-k="effort"] .pr-dot.preview{border-color:#e0c47a;background:#fff8e6;color:#8a5f0f}'
    + '.pr-dot.bump{animation:prDotBump .42s cubic-bezier(.2,.9,.3,1.5)}'
    + '@keyframes prDotBump{0%{transform:scale(1)}40%{transform:scale(1.24)}100%{transform:scale(1)}}'
    + '.pr-dot:active{transform:scale(.93)}'
    + '.pr-word{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#8f1515;min-height:1em;min-width:0;flex:1 1 8rem;overflow-wrap:anywhere;opacity:0;transform:translateY(3px);transition:opacity .25s,transform .25s}'
    + '.pr-rate-row:has([data-k="effort"]) .pr-word{color:#8a5f0f}'
    + '.pr-word.show{opacity:1;transform:none}'
    + '.pr-rate-foot{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:4px}'
    + '.pr-submit{min-width:148px;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:background .2s,transform .12s}'
    + '.pr-submit.is-busy{opacity:.92;cursor:progress}'
    + '.pr-submit.ok{background:#157a47!important}'
    + '.pr-submit.shake{animation:prShake .4s}'
    + '@keyframes prShake{10%,90%{transform:translateX(-2px)}30%,70%{transform:translateX(4px)}50%{transform:translateX(-4px)}}'
    + '.pr-spin{display:inline-block;width:13px;height:13px;border-radius:50%;border:2px solid rgba(255,255,255,.35);border-top-color:#fff;animation:prSpin .7s linear infinite;flex:none}'
    + '@keyframes prSpin{to{transform:rotate(360deg)}}'
    + '.pr-vmsg{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#8a857c;flex:1 1 auto;min-width:0}'
    + '.pr-vmsg.bad{color:#b31b1b}.pr-vmsg.good{color:#157a47}'

    // ---- promotion moment ----
    + '.pr-card.is-promoting{animation:prLift .9s cubic-bezier(.2,.85,.3,1) both}'
    + '.pr-card.is-promoting::after{content:"";position:absolute;inset:0;pointer-events:none;background:linear-gradient(115deg,transparent 30%,rgba(240,192,80,.55) 50%,transparent 70%);transform:translateX(-100%);animation:prSweepGold 1.1s ease-out both}'
    + '@keyframes prSweepGold{to{transform:translateX(100%)}}'
    + '@keyframes prLift{0%{transform:none;box-shadow:0 2px 8px rgba(20,20,30,.05)}45%{transform:translateY(-7px) scale(1.012);box-shadow:0 20px 44px rgba(224,138,30,.3)}100%{transform:none;box-shadow:0 2px 8px rgba(20,20,30,.05)}}'
    + '.pr-card.is-leaving{animation:prLeave .45s ease-in both}'
    + '@keyframes prLeave{to{opacity:0;transform:translateX(-14px)}}'

    // ---- button colours ----
    // One colour per KIND of action, so weight matches consequence at a glance.
    //   rate      Cornell red   the main loop
    //   post      deep teal     creating, not rating
    //   approve   green         affirmative admin outcome
    //   decline   amber tint    reversible caution
    //   delete    red outline   destructive, not the primary
    + '.pr-btn-post{background:#0f5c73;box-shadow:0 1px 2px rgba(15,92,115,.25)}'
    + '.pr-btn-post:hover{background:#0a4557}'
    + '.pr-btn-post.ok{background:#157a47!important}'
    + '.pr-btn-decline{color:#8a5f0f;background:#fdf6e7;border-color:#e8d5a8}'
    + '.pr-btn-decline:hover{color:#6d4b0c;background:#fbf0d8;border-color:#d9c084}'

    // ---- admin ----
    + '.pr-adminbar{display:flex;gap:9px;flex-wrap:wrap;margin-top:14px;padding-top:13px;border-top:1.5px dashed #ece9e2}'
    + '.pr-del{color:#b31b1b;background:#fff;border-color:#f0d5d5}'
    + '.pr-del:hover{background:#fdf3f3;border-color:#e0a5a5;color:#8f1515}'
    + '.pr-del--go{color:#fff;background:#8f1515;border-color:#8f1515}'
    + '.pr-del--go:hover{color:#fff;background:#761111;border-color:#761111}'
    + '.pr-delask{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#57534e;align-self:center;margin-right:2px}'

    // ---- responsive ----
    + '@media(max-width:720px){.pr-guide-steps{grid-template-columns:1fr}.pr-create-split{grid-template-columns:1fr}}'
    + '@media(max-width:600px){'
    +   '.pr-idbar{padding:11px 12px;gap:10px}.pr-idrow{flex:1 1 100%}.pr-idinput{flex:1;width:auto}'
    +   '.pr-voice{flex:1 1 100%}'
    +   '.pr-propose{flex:1 1 100%}.pr-page .pr-propose .tp-create-toggle{width:100%;justify-content:center}'
    +   '.pr-rate-row{align-items:flex-start}'
    +   '.pr-rate-lbl{flex:1 1 100%}'
    +   '.pr-scale-wrap{flex:1 1 100%;gap:6px}'
    +   '.pr-scale{width:100%;justify-content:space-between}'
    +   '.pr-dot{flex:1 1 0;max-width:none;height:40px;font-size:14px}'
    +   '.pr-word{flex:1 1 100%}'
    +   '.pr-quorum{gap:8px 10px}.pr-qsay{flex:1 1 100%}.pr-qvals{margin-left:0}'
    +   '.pr-submit{flex:1 1 auto;min-width:0}.pr-vmsg{flex:1 1 100%;text-align:left}'
    +   '.pr-adminbar .btn{flex:1 1 auto}'
    +   '.pr-guide-body{padding:4px 14px 15px}'
    +   '.pr-guide-fly{display:none}'
    +   '.pr-page .pr-card .card-body{padding:15px 14px 14px}'
    +   '.pr-optional{display:block;margin:4px 0 0}'
    + '}'
    + '@media(max-width:380px){.pr-dot{height:36px;font-size:13px;border-radius:9px}.pr-seg{width:20px}}'

    // ---- motion preferences ----
    + '@media(prefers-reduced-motion:reduce){'
    +   '.pr-card,.pr-guide-body,.pr-empty-bee{animation:none}'
    +   '.pr-guide-icon,.pr-guide-fly,.pr-guide-toggle:hover .pr-guide-icon,'
    +   '.pr-guide[open] .pr-guide-icon,.pr-guide[open] .pr-guide-fly{animation:none}'
    +   '.pr-seg,.pr-voice-seg,.pr-dot,.pr-word{transition:none;animation:none}'
    +   '.pr-card.is-promoting,.pr-card.is-promoting::after,.pr-card.is-leaving{animation:none}'
    +   '.pr-sync::after{animation:none;width:100%}'
    +   '.pr-spin{animation:none;border-top-color:rgba(255,255,255,.9)}'
    +   '.pr-topbar.on{transition:opacity .2s}'
    +   '.pr-page .pr-card:hover{transform:none}'
    + '}'
    + '</style>';
}

function prClientJs_() {
  return '<script>'
    // ---------- identity ----------
    + 'function prNetid(){var el=document.getElementById("pr-netid");return el?el.value.trim():"";}'
    + 'var PR_SAVE_T=null;'
    + 'function prSaveNetid(){try{localStorage.setItem("prNetid",prNetid());}catch(e){}'
    + 'clearTimeout(PR_SAVE_T);PR_SAVE_T=setTimeout(prLoadViewer,600);}'

    // ---------- global loading affordances ----------
    // A blind spinner reads as "stuck"; naming the actual step reads as "working", so
    // every wait here says what the server is doing.
    + 'var PR_BUSY=0;'
    + 'function prTop(on){var b=document.getElementById("pr-top");if(!b)return;'
    + 'PR_BUSY=Math.max(0,PR_BUSY+(on?1:-1));'
    + 'if(PR_BUSY>0){b.className="pr-topbar on";}else{b.className="pr-topbar done";setTimeout(function(){if(PR_BUSY<=0)b.className="pr-topbar";},600);}}'
    // Swap a button into a working state that steps through the real phases of the call.
    + 'function prBusy(btn,stages){if(!btn)return function(){};var old=btn.innerHTML,i=0;'
    + 'btn.disabled=true;btn.classList.add("is-busy");prTop(true);'
    + 'function paint(){btn.innerHTML=\'<span class="pr-spin"></span>\'+stages[Math.min(i,stages.length-1)];}paint();'
    + 'var t=setInterval(function(){i++;if(i<stages.length)paint();else clearInterval(t);},1100);'
    + 'return function(done,cls){clearInterval(t);prTop(false);btn.disabled=false;btn.classList.remove("is-busy");'
    + 'if(!done){btn.innerHTML=old;return;}btn.innerHTML=done;if(cls)btn.classList.add(cls);'
    + 'setTimeout(function(){btn.innerHTML=old;if(cls)btn.classList.remove(cls);},1700);};}'

    // ---------- the rating scale ----------
    // Dots fill cumulatively (1..n) so the control reads as a quantity, not a radio set.
    + 'function prPaintScale(sc,val,cls){[].slice.call(sc.querySelectorAll(".pr-dot")).forEach(function(d){'
    + 'var n=parseInt(d.getAttribute("data-n"),10);d.classList.remove("preview");'
    + 'if(cls==="preview"){d.classList.toggle("preview",n<=val);return;}'
    + 'd.classList.toggle("fill",!!val&&n<=val);d.classList.toggle("on",n===val);});}'
    + 'function prWord(kind,n){var impact=["","barely noticed","a small lift","a solid win","a big deal","transforms the space"];'
    + 'var effort=["","minutes","an afternoon","a weekend","a real build","a major undertaking"];'
    + 'return (kind==="impact"?impact:effort)[n]||"";}'
    + 'function prVerdict(imp,eff){if(!imp||!eff)return "";var hiI=imp>=4,loE=eff<=2;'
    + 'if(hiI&&loE)return "<b>Quick win.</b> Big payoff, little work.";'
    + 'if(hiI&&eff>=4)return "<b>Big swing.</b> Worth it, but it is a project.";'
    + 'if(imp<=2&&loE)return "<b>Easy tidy-up.</b> Small, but cheap to do.";'
    + 'if(imp<=2&&eff>=4)return "<b>Hard pass.</b> Lots of work for little gain.";'
    + 'return "<b>Middle of the road.</b> Reasonable either way.";}'
    + 'function prPick(btn){var sc=btn.parentNode,n=parseInt(btn.getAttribute("data-n"),10);'
    + 'sc.setAttribute("data-val",n);prPaintScale(sc,n);'
    + 'btn.classList.remove("bump");void btn.offsetWidth;btn.classList.add("bump");'
    + 'var card=btn.closest(".pr-card"),rid=card?card.id:"",kind=sc.getAttribute("data-k");'
    + 'var w=document.getElementById(rid+"-w-"+kind);if(w){w.textContent=prWord(kind,n);w.classList.add("show");}'
    + 'var imp=prScaleVal(rid,"impact"),eff=prScaleVal(rid,"effort"),v=document.getElementById(rid+"-verdict");'
    + 'if(v){if(imp&&eff){v.innerHTML=prVerdict(imp,eff);v.hidden=false;}else{v.hidden=true;}}'
    + 'var msg=document.getElementById(rid+"-vmsg");if(msg&&msg.className==="pr-vmsg bad"){msg.textContent="";msg.className="pr-vmsg";}}'
    + 'function prScaleVal(rid,k){var sc=document.querySelector("#"+rid+"-rate .pr-scale[data-k=\\""+k+"\\"]");return sc?parseInt(sc.getAttribute("data-val"),10)||0:0;}'
    // Hover preview: fill up to the dot under the cursor, restore on leave.
    + 'document.addEventListener("mouseover",function(e){var d=e.target.closest?e.target.closest(".pr-dot"):null;if(!d)return;'
    + 'var sc=d.parentNode;prPaintScale(sc,parseInt(d.getAttribute("data-n"),10),"preview");});'
    + 'document.addEventListener("mouseout",function(e){var d=e.target.closest?e.target.closest(".pr-dot"):null;if(!d)return;'
    + 'var sc=d.parentNode;prPaintScale(sc,parseInt(sc.getAttribute("data-val"),10)||0);});'

    // ---------- earned voice meter ----------
    // The nearest goal on the page: a small number of reviews that switches your ratings
    // on. Rendered as segments so the remaining step is one visible gap.
    + 'var PR_REVIEWS=0,PR_VOICE_ON=false;'
    + 'function prVoiceRender(count,bump){PR_REVIEWS=count;'
    + 'var wrap=document.getElementById("pr-voice"),segs=document.getElementById("pr-voice-segs"),note=document.getElementById("pr-idnote");'
    + 'if(!wrap||!segs||!note)return;wrap.hidden=false;'
    + 'if(!prNetid()){segs.hidden=true;wrap.classList.remove("done");'
    + 'note.textContent="";'
    + 'wrap.title="Enter your own NetID. Do not add someone else.";return;}'
    + 'segs.hidden=false;'
    + 'var need=PR_VOICE_NEED,have=Math.min(count,need),done=have>=need;'
    + 'if(segs.children.length!==need){segs.innerHTML="";for(var i=0;i<need;i++)segs.appendChild(document.createElement("i"));}'
    + '[].slice.call(segs.children).forEach(function(el,i){el.className="pr-voice-seg"+(i<have?" on":"");'
    + 'if(bump&&i===have-1){el.classList.add("pop");setTimeout(function(){el.classList.remove("pop");},450);}});'
    + 'wrap.classList.toggle("done",done);'
    + 'note.textContent=done?"Ratings on":((need-have)+" to go");'
    + 'wrap.title=done?"Your ratings are counted.":("Review "+(need-have)+" more proposal"+(need-have===1?"":"s")+" and your ratings start counting.");'
    + 'PR_VOICE_ON=done;}'

    // ---------- viewer state (cheap personalisation) ----------
    // Patches the already-rendered cards instead of rebuilding the list.
    + 'function prLoadViewer(){var nid=prNetid();if(!nid){prVoiceRender(0);return;}'
    + 'prTop(true);google.script.run.withSuccessHandler(function(r){prTop(false);if(!r||!r.ok)return;'
    + 'prVoiceRender(r.reviewCount||0);'
    + '[].slice.call(document.querySelectorAll(".pr-card")).forEach(function(card){'
    + 'var mine=r.myVotes[card.getAttribute("data-id")];if(!mine)return;var rid=card.id;'
    + '[["impact",mine.impact],["effort",mine.effort]].forEach(function(pair){'
    + 'var sc=card.querySelector(\'.pr-scale[data-k="\'+pair[0]+\'"]\');if(!sc||!pair[1])return;'
    + 'sc.setAttribute("data-val",pair[1]);prPaintScale(sc,pair[1]);'
    + 'var w=document.getElementById(rid+"-w-"+pair[0]);if(w){w.textContent=prWord(pair[0],pair[1]);w.classList.add("show");}});'
    + 'var v=document.getElementById(rid+"-verdict");if(v&&mine.impact&&mine.effort){v.innerHTML=prVerdict(mine.impact,mine.effort);v.hidden=false;}'
    + 'var go=document.getElementById(rid+"-go");if(go)go.textContent="Update rating";'
    // Flag it on the folded summary too, so a scan down the list shows what is left to do.
    + 'var mk=document.getElementById(rid+"-mine");if(mk)mk.hidden=false;'
    + '});}).withFailureHandler(function(){prTop(false);}).prViewerState(nid);}'

    // ---------- folded cards ----------
    // A card the reader opened stays open across a re-render. Kept in memory rather than
    // localStorage: it is about not losing your place mid-session, not a lasting setting.
    + 'var PR_OPEN={};'
    + 'function prCardToggled(el){var id=el.getAttribute("data-id");if(!id)return;'
    + 'if(el.open)PR_OPEN[id]=1;else delete PR_OPEN[id];}'
    + 'function prCardsRestore(){[].slice.call(document.querySelectorAll(".pr-card")).forEach(function(c){'
    + 'if(PR_OPEN[c.getAttribute("data-id")])c.open=true;});}'
    + 'function prExpandAll(){var btn=document.getElementById("pr-expand");'
    + 'var cards=[].slice.call(document.querySelectorAll(".pr-card"));'
    + 'var anyClosed=cards.some(function(c){return !c.open;});'
    + 'cards.forEach(function(c){c.open=anyClosed;});'
    + 'if(btn)btn.textContent=anyClosed?"Collapse all":"Expand all";}'

    // ---------- quorum meter updates ----------
    + 'function prSay(counted,need,avg,hasVotes){if(!hasVotes)return "be the first to weigh in";'
    + 'if(counted>=need&&avg<PR_BAR)return "enough reviews, impact below the bar";'
    + 'var left=need-counted;if(left===1)return "one more review and it is in";'
    + 'if(left>1)return left+" more reviews to decide";return "quorum reached";}'
    + 'function prMeter(rid,counted,avgI,avgE){'
    + 'var segs=document.getElementById(rid+"-segs");'
    + 'if(segs){segs.title=counted+" of "+PR_NEED+" counted reviews";'
    + '[].slice.call(segs.children).forEach(function(el,i){el.classList.remove("pending");'
    + 'var on=i<counted;if(on&&!el.classList.contains("on")){el.classList.add("on","pop");setTimeout(function(){el.classList.remove("pop");},520);}'
    + 'else el.classList.toggle("on",on);});}'
    + 'var vi=document.getElementById(rid+"-vi"),ve=document.getElementById(rid+"-ve");'
    + 'if(vi){vi.textContent=avgI?avgI.toFixed(1):"\\u2013";vi.classList.toggle("good",!!counted&&avgI>=PR_BAR);}'
    + 'if(ve)ve.textContent=avgE?avgE.toFixed(1):"\\u2013";'
    + 'var say=document.getElementById(rid+"-say");'
    + 'if(say){say.textContent=prSay(counted,PR_NEED,avgI,true);'
    + 'say.className="pr-qsay"+(PR_NEED-counted===1?" pr-qsay--close":"")+(counted>=PR_NEED&&avgI>=PR_BAR?" pr-qsay--good":"");}'
    // Keep the collapsed summary in step with the open meter, so folding the card back up
    // does not show stale numbers.
    + 'var ss=document.getElementById(rid+"-ssegs");'
    + 'if(ss){ss.title=counted+" of "+PR_NEED+" counted reviews";'
    + '[].slice.call(ss.children).forEach(function(el,i){el.classList.toggle("on",i<counted);});}'
    + 'var sn=document.getElementById(rid+"-sn");if(sn)sn.textContent=counted+"/"+PR_NEED;'
    + 'var si=document.getElementById(rid+"-si");'
    + 'if(si){si.textContent=avgI?avgI.toFixed(1):"\\u2013";si.classList.toggle("good",!!counted&&avgI>=PR_BAR);}}'

    // ---------- submitting a rating ----------
    // The wait is real (a write, a recount and a quorum check), so it is narrated rather
    // than hidden, and the outcome lands as one beat: meter, then verdict.
    + 'function prSubmitVote(rid,pid){var msg=document.getElementById(rid+"-vmsg"),btn=document.getElementById(rid+"-go");'
    + 'var nid=prNetid();'
    + 'if(!nid){msg.className="pr-vmsg bad";msg.textContent="Enter your NetID at the top to rate.";'
    + 'var f=document.getElementById("pr-netid");if(f){f.focus();f.scrollIntoView({block:"center",behavior:"smooth"});}return;}'
    + 'var imp=prScaleVal(rid,"impact"),eff=prScaleVal(rid,"effort");'
    + 'if(!imp||!eff){msg.className="pr-vmsg bad";msg.textContent="Rate both impact and effort.";'
    + 'if(btn){btn.classList.remove("shake");void btn.offsetWidth;btn.classList.add("shake");setTimeout(function(){btn.classList.remove("shake");},450);}return;}'
    + 'msg.className="pr-vmsg";msg.textContent="";'
    // Show the pending slot immediately so the click has a consequence before the server answers.
    + 'var segs=document.getElementById(rid+"-segs");'
    + 'if(segs){var pend=[].slice.call(segs.children).filter(function(el){return !el.classList.contains("on");})[0];if(pend)pend.classList.add("pending");}'
    + 'var done=prBusy(btn,["Recording your rating","Counting the swarm","Checking quorum"]);'
    + 'google.script.run.withSuccessHandler(function(r){'
    + 'if(!r||!r.ok){done();if(segs)[].slice.call(segs.children).forEach(function(el){el.classList.remove("pending");});'
    + 'msg.className="pr-vmsg bad";msg.textContent=(r&&r.error)||"Could not save that.";return;}'
    + 'prMeter(rid,r.counted,r.avgImpact,r.avgEffort);'
    + 'var unlocked=r.voice&&r.voice.counted&&!PR_VOICE_ON;'
    + 'if(r.voice)prVoiceRender(Math.min(r.voice.have,PR_VOICE_NEED),true);'
    + 'if(r.promoted){done("\\u2713 It is in","ok");prPromoteFx(rid);}'
    + 'else if(unlocked){done("\\u2713 Counted","ok");'
    + 'tpConfetti("\\uD83D\\uDC1D Your ratings count from here on.");msg.className="pr-vmsg good";msg.textContent="Your ratings now count.";prRefresh(true);}'
    + 'else if(r.voice&&!r.voice.counted){var togo=r.voice.need-r.voice.have;done("\\u2713 Saved","ok");'
    + 'msg.className="pr-vmsg";msg.textContent="Saved. Review "+togo+" more proposal"+(togo===1?"":"s")+" and your ratings start counting.";}'
    + 'else{done("\\u2713 Counted","ok");msg.className="pr-vmsg good";'
    + 'msg.textContent=r.updated?"Rating updated.":(PR_NEED-r.counted===1?"Counted. One more review and this one is in.":"Rating counted.");'
    + 'if(r.alsoPromoted>0){tpConfetti("\\uD83D\\uDC1D That review carried another proposal over the line.");prRefresh(true);}}'
    + '}).withFailureHandler(function(){done();if(segs)[].slice.call(segs.children).forEach(function(el){el.classList.remove("pending");});'
    + 'msg.className="pr-vmsg bad";msg.textContent="Network hiccup. Try again.";}).prVote(pid,nid,imp,eff,"");}'

    // The promotion beat: hold on the full meter, THEN flip the badge and celebrate. The
    // pause is the point; a reveal with no anticipation barely registers.
    + 'function prPromoteFx(rid){var card=document.getElementById(rid);if(!card){prRefresh(true);return;}'
    + 'card.classList.add("is-promoting");'
    + 'setTimeout(function(){'
    + 'var say=document.getElementById(rid+"-say");if(say){say.textContent="quorum reached";say.className="pr-qsay pr-qsay--good";}'
    + 'tpConfetti("\\uD83D\\uDC1D Quorum reached. It is waiting in Projects.");},620);'
    + 'setTimeout(function(){prRefresh(true);},2400);}'

    // ---------- posting a proposal ----------
    // Photos go to Drive the moment they are picked, all at once, while the scout is
    // still typing the title. Posting then costs ONE round trip instead of one per photo
    // followed by the post, which is what made creating with photos feel slow.
    + 'var PRSTAGE=[];'
    + 'function prStageRender(){var el=document.getElementById("pr-c-stage");if(!el)return;'
    + 'if(!PRSTAGE.length){el.innerHTML="";return;}'
    + 'var h=\'<div class="stage">\';'
    + 'for(var i=0;i<PRSTAGE.length;i++){var p=PRSTAGE[i],st=p.state||"pending";'
    + 'var badge=st==="done"?"\\u2713":(st==="error"?"Retry":\'<span class="pr-spin"></span>\');'
    + 'h+=\'<div class="stage-item pr-stage-\'+st+\'"><img src="\'+p.dataUrl+\'" alt="">\''
    + '+\'<span class="stage-idx">\'+(i+1)+\'</span>\''
    + '+\'<button type="button" class="stage-rm" title="Remove" onclick="prUnstage(\'+i+\')">&times;</button>\''
    + '+\'<span class="pr-stage-state"\'+(st==="error"?\' role="button" title="Upload again" onclick="prRetryPhoto(\'+i+\')"\':"")+\'>\'+badge+\'</span></div>\';}'
    + 'h+=\'<button type="button" class="stage-add" onclick="document.getElementById(\\\'pr-c-file\\\').click()"><span class="stage-add-i">+</span>Add more</button></div>\';'
    + 'el.innerHTML=h;}'
    + 'function prPickPhotos(input){tpStageRead(input,PRSTAGE,function(){prStageRender();prUploadPending();});}'
    // Fire every not-yet-started photo at once. Uploads only touch Drive, never the
    // sheet, so there is nothing for them to contend over.
    + 'function prUploadPending(){PRSTAGE.forEach(function(p){if(p.state)return;p.state="uploading";prSendPhoto(p);});prStageRender();}'
    + 'function prSendPhoto(p){google.script.run.withSuccessHandler(function(r){'
    // Removed while it was in flight: drop the file rather than leave it orphaned.
    + 'if(PRSTAGE.indexOf(p)<0){if(r&&r.ok&&r.photoId)prTrashPhoto(r.photoId);return;}'
    + 'if(!r||!r.ok){p.state="error";}else{p.state="done";p.photoId=r.photoId;}prStageRender();'
    + '}).withFailureHandler(function(){if(PRSTAGE.indexOf(p)>=0){p.state="error";prStageRender();}}).prUploadPhoto(p.dataUrl,p.name);}'
    + 'function prRetryPhoto(i){var p=PRSTAGE[i];if(!p||p.state!=="error")return;p.state="uploading";prStageRender();prSendPhoto(p);}'
    + 'function prUnstage(i){var p=PRSTAGE[i];if(!p)return;PRSTAGE.splice(i,1);'
    + 'if(p.state==="done"&&p.photoId)prTrashPhoto(p.photoId);prStageRender();}'
    + 'function prTrashPhoto(id){google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(){}).prTrashUpload(id);}'
    + 'function prPhotosBusy(){return PRSTAGE.filter(function(p){return p.state==="uploading";}).length;}'
    + 'function prPhotosFailed(){return PRSTAGE.filter(function(p){return p.state==="error";}).length;}'
    + 'function prCreate(){var g=function(id){return (document.getElementById(id)||{}).value||"";};'
    + 'var t=g("pr-c-title"),d=g("pr-c-desc"),a=g("pr-c-area"),nm=g("pr-c-name");'
    + 'var msg=document.getElementById("pr-c-msg"),btn=document.getElementById("pr-c-go"),nid=prNetid();'
    + 'if(!t.trim()){msg.className="tp-lock-msg bad";msg.textContent="A title is required.";return;}'
    + 'if(!nid){msg.className="tp-lock-msg bad";msg.textContent="Enter your NetID at the top first.";'
    + 'var f=document.getElementById("pr-netid");if(f)f.focus();return;}'
    + 'if(prPhotosFailed()){msg.className="tp-lock-msg bad";msg.textContent="A photo did not upload. Tap Retry on it, or remove it.";return;}'
    + 'msg.className="tp-lock-msg";msg.textContent="";'
    // Usually the photos are already up by now, so this is a single call.
    + 'var waiting=prPhotosBusy();'
    + 'var done=prBusy(btn,waiting?["Finishing your photos","Posting your proposal"]:["Posting your proposal","Adding it to the board"]);'
    + '(function ready(){if(prPhotosBusy()){setTimeout(ready,150);return;}'
    + 'if(prPhotosFailed()){done();msg.className="tp-lock-msg bad";msg.textContent="A photo did not upload. Tap Retry on it, or remove it.";return;}'
    + 'var ids=PRSTAGE.map(function(p){return p.photoId;}).filter(Boolean);'
    + 'google.script.run.withSuccessHandler(function(r){'
    + 'if(!r||!r.ok){done();msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Could not post that.";return;}'
    + 'done("\\u2713 Posted","ok");'
    + '["pr-c-title","pr-c-desc","pr-c-area"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});'
    + 'PRSTAGE=[];prStageRender();'
    + 'msg.className="tp-lock-msg ok";msg.textContent="Posted. The team can rate it now.";'
    + 'var w=document.querySelector(".tp-create-wrap");if(w)w.open=false;'
    + 'tpConfetti("\\uD83D\\uDC1D Posted. Let the swarm decide.");prRefresh();'
    + '}).withFailureHandler(function(){done();msg.className="tp-lock-msg bad";msg.textContent="Network hiccup. Try again.";})'
    + '.prSubmitProposal(t,d,a,nm,nid,ids.join(", "));})();}'

    // ---------- admin ----------
    + 'function prEditOpen(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),a=document.getElementById(rid+"-admin");'
    + 'if(v)v.hidden=true;if(e)e.hidden=false;if(a)a.hidden=true;var t=document.getElementById(rid+"-etitle");if(t)t.focus();}'
    + 'function prEditCancel(rid){var v=document.getElementById(rid+"-view"),e=document.getElementById(rid+"-edit"),a=document.getElementById(rid+"-admin"),m=document.getElementById(rid+"-emsg");'
    + 'if(e)e.hidden=true;if(v)v.hidden=false;if(a)a.hidden=false;if(m){m.textContent="";m.style.color="";}}'
    + 'function prEditSave(rid,pid){var g=function(s){var el=document.getElementById(rid+s);return el?el.value:"";};'
    + 'var title=g("-etitle"),desc=g("-edesc"),area=g("-earea"),name=g("-ename"),m=document.getElementById(rid+"-emsg");'
    + 'if(!title.trim()){if(m){m.style.color="#b31b1b";m.textContent="A title is required.";}return;}'
    + 'if(m){m.style.color="";m.textContent="Saving\\u2026";}'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.ok){if(m){m.style.color="#b31b1b";m.textContent=(r&&r.error)||"Could not save.";}return;}'
    + 'var t=document.getElementById(rid+"-title");if(t)t.textContent=r.title;'
    + 'var ar=document.getElementById(rid+"-area");if(ar){if(r.area){ar.hidden=false;ar.textContent=r.area;}else{ar.hidden=true;ar.textContent="";}}'
    + 'var d=document.getElementById(rid+"-desc");if(d){if(r.description){d.hidden=false;d.textContent=r.description;}else{d.hidden=true;d.textContent="";}}'
    + 'var b=document.getElementById(rid+"-by");if(b){if(r.proposedBy){b.hidden=false;b.textContent=r.proposedBy;}else{b.hidden=true;b.textContent="";}}'
    + 'var meta=document.getElementById(rid+"-meta");if(meta)meta.hidden=!(r.area||r.proposedBy);'
    + 'prEditCancel(rid);'
    + '}).withFailureHandler(function(){if(m){m.style.color="#b31b1b";m.textContent="Network hiccup. Try again.";}}).prUpdateProposal(pid,title,desc,area,name);}'
    + 'function prAdmin(rid,pid,dec){var msg=document.getElementById(rid+"-vmsg"),card=document.getElementById(rid);'
    + 'var btn=card?card.querySelector(dec==="approve"?".btn-confirm":".pr-btn-decline"):null;'
    + 'var done=prBusy(btn,[dec==="approve"?"Approving":"Declining","Updating the board"]);'
    + 'google.script.run.withSuccessHandler(function(r){done();'
    + 'if(!r||!r.ok){if(msg){msg.className="pr-vmsg bad";msg.textContent=(r&&r.error)||"Failed.";}return;}'
    + 'if(dec==="approve"){prPromoteFx(rid);}'
    + 'else{if(card)card.classList.add("is-leaving");setTimeout(function(){prRefresh(true);},420);}'
    + '}).withFailureHandler(function(){done();if(msg){msg.className="pr-vmsg bad";msg.textContent="Network hiccup. Try again.";}}).prAdminSetStatus(pid,dec);}'
    // Delete is destructive and unlike Decline leaves no record, so it always confirms.
    + 'function prDelAsk(rid,pid){var bar=document.getElementById(rid+"-admin");if(!bar)return;'
    + 'if(bar.getAttribute("data-asking")==="1")return;bar.setAttribute("data-asking","1");'
    + 'bar.setAttribute("data-prev",bar.innerHTML);'
    + 'bar.innerHTML=\'<span class="pr-delask">Delete this proposal and its ratings?</span>\''
    + '+\'<button type="button" class="btn pr-del pr-del--go" onclick="prDelGo(\\\'\'+rid+\'\\\',\\\'\'+pid+\'\\\')">Yes, delete</button>\''
    + '+\'<button type="button" class="btn btn-ghost" onclick="prDelCancel(\\\'\'+rid+\'\\\')">Cancel</button>\';}'
    + 'function prDelCancel(rid){var bar=document.getElementById(rid+"-admin");if(!bar)return;'
    + 'bar.innerHTML=bar.getAttribute("data-prev")||"";bar.setAttribute("data-asking","0");}'
    + 'function prDelGo(rid,pid){var bar=document.getElementById(rid+"-admin"),card=document.getElementById(rid);'
    + 'var btn=bar?bar.querySelector(".pr-del"):null;var done=prBusy(btn,["Deleting","Clearing its ratings"]);'
    + 'google.script.run.withSuccessHandler(function(r){done();'
    + 'if(!r||!r.ok){var m=document.getElementById(rid+"-vmsg");if(m){m.className="pr-vmsg bad";m.textContent=(r&&r.error)||"Could not delete that.";}return;}'
    + 'if(card)card.classList.add("is-leaving");setTimeout(function(){prRefresh(true);},420);'
    + '}).withFailureHandler(function(){done();}).prDeleteProposal(pid);}'

    // ---------- list refresh ----------
    // Never blanks the list: the current cards stay put and dim while the new HTML is
    // fetched, so nothing jumps under the cursor.
    + 'function prRefresh(quiet){var list=document.getElementById("pr-list"),sync=document.getElementById("pr-sync");'
    + 'if(!list)return;if(sync)sync.hidden=false;list.classList.add("is-syncing");if(!quiet)prTop(true);'
    + 'google.script.run.withSuccessHandler(function(r){'
    + 'if(sync)sync.hidden=true;list.classList.remove("is-syncing");if(!quiet)prTop(false);'
    + 'if(!r||!r.ok)return;list.innerHTML=r.html;secRestore();prCardsRestore();'
    + '[].slice.call(list.querySelectorAll(".pr-card")).forEach(function(c,i){c.style.animationDelay=(i*45)+"ms";});'
    + 'if(typeof ADMIN_PASS!=="undefined"&&ADMIN_PASS)document.querySelectorAll(".tp-admin").forEach(function(e){e.hidden=false;});'
    + 'prLoadViewer();'
    + '}).withFailureHandler(function(){if(sync)sync.hidden=true;list.classList.remove("is-syncing");if(!quiet)prTop(false);}).prListHtml();}'

    // ---------- boot ----------
    + 'window.addEventListener("DOMContentLoaded",function(){'
    + 'try{var s=localStorage.getItem("prNetid");if(s){var el=document.getElementById("pr-netid");if(el)el.value=s;}}catch(e){}'
    + 'prLoadViewer();'
    + 'var t=document.getElementById("pr-c-title");if(t)t.addEventListener("keydown",function(e){if(e.key==="Enter"){e.preventDefault();prCreate();}});'
    + '});'
    + '</script>';
}
