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
 * reveals Approve now / Decline / Delete on each provisional card via tpAdminRevealJs_.
 * Delete is for proposals the swarm has clearly rejected; it removes the row and its
 * votes. Decline keeps the row for the record.
 *
 * Proposals can carry photos of the spot to improve: the create form stages them
 * client-side (tpStageRead/tpStagePreview) and uploads to Drive one at a time via
 * prUploadPhoto -> tpSaveUpload_; cards render them as linked thumbnails.
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
  cacheKey: 'pr_list_v1',  // rendered list HTML (netid-agnostic)
  stripKey: 'pr_strip_v1', // the "new project proposed" strip shown on the tasks pages
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
  if (c) { try { c.removeAll([PR.cacheKey, PR.stripKey]); } catch (e) { /* cache is best-effort */ } }
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
  });
  return promoted;
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

// One staged photo -> Drive (the create form uploads its tray one file at a time,
// same as project completion photos). Returns the file id for the Photos cell.
function prUploadPhoto(dataUrl, filename) {
  try {
    return { ok: true, photoId: tpSaveUpload_(dataUrl, filename) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
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

function prSection_(label, count, cls) {
  return '<div class="section-head"><span class="section-label ' + (cls || 'section-label--open') + '">' + label + '</span>'
    + '<span class="section-count">' + count + '</span><span class="section-rule"></span></div>';
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

function prMeterHtml_(p, rid) {
  var need = PR.quorumMinVotes, counted = p.countedVotes || 0;
  var segs = '';
  for (var i = 0; i < need; i++) {
    segs += '<i class="pr-seg' + (i < counted ? ' on' : '') + '" style="animation-delay:' + (i * 70) + 'ms"></i>';
  }
  var impPct = Math.max(0, Math.min(100, (p.avgImpact / 5) * 100));
  var effPct = Math.max(0, Math.min(100, (p.avgEffort / 5) * 100));
  var barPct = (PR.quorumMinImpact / 5) * 100;
  var good = p.countedVotes && p.avgImpact >= PR.quorumMinImpact;
  return '<div class="pr-quorum" id="' + rid + '-q">'
    + '<div class="pr-qtop"><span class="pr-qcount"><b id="' + rid + '-n">' + counted + '</b> of ' + need + ' reviews</span>'
    +   '<span class="pr-qsay' + (need - counted === 1 ? ' pr-qsay--close' : '') + '" id="' + rid + '-say">' + prSayText_(counted, need, p.avgImpact, p.votes) + '</span></div>'
    + '<div class="pr-segs" id="' + rid + '-segs">' + segs + '</div>'
    + '<div class="pr-gauges">'
    +   '<div class="pr-gauge"><span class="pr-glbl">Impact</span>'
    +     '<span class="pr-gbar"><i class="pr-gfill' + (good ? ' good' : '') + '" id="' + rid + '-gi" style="width:' + impPct + '%"></i>'
    +     '<i class="pr-gbar-mark" style="left:' + barPct + '%" title="the bar: ' + PR.quorumMinImpact + '"></i></span>'
    +     '<b class="pr-gval" id="' + rid + '-vi">' + (p.countedVotes ? p.avgImpact.toFixed(1) : '&ndash;') + '</b></div>'
    +   '<div class="pr-gauge"><span class="pr-glbl">Effort</span>'
    +     '<span class="pr-gbar"><i class="pr-gfill pr-gfill--eff" id="' + rid + '-ge" style="width:' + effPct + '%"></i></span>'
    +     '<b class="pr-gval" id="' + rid + '-ve">' + (p.countedVotes ? p.avgEffort.toFixed(1) : '&ndash;') + '</b></div>'
    + '</div></div>';
}

// One proposal card. votable => provisional (shows the rating control); else it's an
// approved (Real) card.
function prCardHtml_(p, rid, votable) {
  var pill = votable
    ? '<span class="tp-pill tp-pill--active">Up for review</span>'
    : '<span class="tp-pill tp-pill--done">Approved</span>';

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
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Impact <em>how much better does this make the place</em></span>'
      +   '<span class="pr-scale-wrap">' + scale('impact', my ? my.impact : 0) + '<span class="pr-word" id="' + rid + '-w-impact">' + (my ? prWordFor_('impact', my.impact) : '') + '</span></span></div>'
      + '<div class="pr-rate-row"><span class="pr-rate-lbl">Effort <em>how much work is it to pull off</em></span>'
      +   '<span class="pr-scale-wrap">' + scale('effort', my ? my.effort : 0) + '<span class="pr-word" id="' + rid + '-w-effort">' + (my ? prWordFor_('effort', my.effort) : '') + '</span></span></div>'
      + '<div class="pr-verdict" id="' + rid + '-verdict"' + (my ? '' : ' hidden') + '>' + (my ? prVerdictText_(my.impact, my.effort) : '') + '</div>'
      + '<div class="pr-rate-foot"><button type="button" class="btn btn-primary pr-submit" id="' + rid + '-go" onclick="prSubmitVote(\'' + rid + '\',\'' + p.id + '\')">' + (my ? 'Update rating' : 'Submit rating') + '</button>'
      + '<span id="' + rid + '-vmsg" class="pr-vmsg"></span></div></div>'
    : '';

  // Admin-only decisions (revealed by tpAdminRevealJs_ / re-revealed after prRefresh).
  var adminBar = '<div class="tp-admin pr-adminbar" hidden id="' + rid + '-admin">'
    + (votable ? '<button type="button" class="btn btn-confirm" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'approve\')">Approve now</button>'
      + '<button type="button" class="btn btn-ghost" onclick="prAdmin(\'' + rid + '\',\'' + p.id + '\',\'decline\')">Decline</button>' : '')
    + '<button type="button" class="btn btn-ghost pr-del" onclick="prDelAsk(\'' + rid + '\',\'' + p.id + '\')">Delete</button></div>';

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

  return '<div class="card pr-card" id="' + rid + '" data-id="' + p.id + '" data-owner="' + escapeHtml_(p.netid) + '">'
    + '<div class="card-body">'
    +   '<div class="pr-head"><div class="pr-head-main"><div class="pr-kicker">Proposal' + (p.area ? '<span class="pr-area"> &middot; ' + escapeHtml_(p.area) + '</span>' : '') + '</div>'
    +     '<h3 class="pr-title">' + escapeHtml_(p.title) + '</h3></div><span class="pr-pill-wrap" id="' + rid + '-pill">' + pill + '</span></div>'
    +   (p.description ? '<p class="pr-desc">' + escapeHtml_(p.description) + '</p>' : '')
    +   (p.proposedBy ? '<div class="pr-by">Scouted by ' + escapeHtml_(p.proposedBy) + '</div>' : '')
    +   photos
    +   meta
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
  else provisional.forEach(function (p) { out += prCardHtml_(p, 'pr' + (idx++), true); });
  if (real.length) { out += prSection_('Approved', real.length, 'section-label--open'); real.forEach(function (p) { out += prCardHtml_(p, 'pr' + (idx++), false); }); }
  return out;
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
    var fresh = proposals.filter(function (p) {
      return norm_(p.status) === norm_(PR.status.real) && p.promotedAt && new Date(p.promotedAt).getTime() >= cutoff;
    }).sort(function (a, b) { return new Date(b.promotedAt || 0) - new Date(a.promotedAt || 0); });
    var open = proposals.filter(function (p) { return norm_(p.status) === norm_(PR.status.provisional); });
    var href = CONFIG.webAppUrl + (CONFIG.webAppUrl.indexOf('?') >= 0 ? '&' : '?') + 'module=proposals';

    if (fresh.length || open.length) {
      var lead = fresh.length
        ? '<b>' + escapeHtml_(fresh[0].title) + '</b>' + (fresh.length > 1 ? ' and ' + (fresh.length - 1) + ' more' : '')
        : '<b>' + open.length + ' idea' + (open.length === 1 ? '' : 's') + '</b> waiting on the team';
      var kicker = fresh.length ? 'New project proposed' : 'Proposals need your rating';
      var sub = fresh.length
        ? 'Voted up by the team, it is a real project now.'
        : 'Rate impact and effort. Enough support turns one into a real project.';
      html = '<a class="pr-strip" href="' + escapeHtml_(href) + '" target="_blank" rel="noopener">'
        + '<span class="pr-strip-bee" aria-hidden="true">&#128029;</span>'
        + '<span class="pr-strip-main"><span class="pr-strip-kicker">' + kicker + '</span>'
        +   '<span class="pr-strip-title">' + lead + '</span>'
        +   '<span class="pr-strip-sub">' + sub + '</span></span>'
        + '<span class="pr-strip-go" aria-hidden="true">&rarr;</span></a>';
    }
  } catch (err) {
    html = '';   // the tasks page must never fail because proposals are not set up yet
  }
  if (c) { try { c.put(PR.stripKey, html, PR.cacheSeconds); } catch (e) { /* best effort */ } }
  return html;
}

// Styles for the strip only. The tasks pages do not load prStyles_.
function prStripStyles_() {
  return '<style>'
    + '.pr-strip{display:flex;align-items:center;gap:14px;margin:18px 0 2px;padding:14px 18px;border-radius:16px;text-decoration:none;'
    +   'background:linear-gradient(135deg,#fffdf5 0%,#fff8e6 100%);border:1.5px solid #f0e2b8;box-shadow:0 2px 10px rgba(120,90,20,.07);transition:transform .16s,box-shadow .16s}'
    + '.pr-strip:hover{transform:translateY(-2px);box-shadow:0 10px 24px rgba(120,90,20,.14)}'
    + '.pr-strip-bee{font-size:22px;line-height:1;flex:none}'
    + '.pr-strip-main{display:flex;flex-direction:column;gap:2px;min-width:0;flex:1}'
    + '.pr-strip-kicker{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#a8791a}'
    + '.pr-strip-title{font-size:15px;font-weight:600;color:#3f3a31;line-height:1.35}'
    + '.pr-strip-title b{font-weight:800;color:#14110e}'
    + '.pr-strip-sub{font-size:12.5px;color:#8a857c;line-height:1.4}'
    + '.pr-strip-go{font-size:19px;color:#a8791a;flex:none}'
    + '@media(max-width:560px){.pr-strip{padding:12px 14px;gap:11px}.pr-strip-go{display:none}.pr-strip-title{font-size:14px}}'
    + '@media(prefers-reduced-motion:reduce){.pr-strip{transition:none}.pr-strip:hover{transform:none}}'
    + '</style>';
}

// ---- the page ----

// Only ever hand an http(s) URL to an href we build from a query parameter.
function prSafeUrl_(u) {
  u = String(u || '').trim();
  return /^https?:\/\//i.test(u) ? u : '';
}

function proposalsPage_(embedded, admin, back) {
  var backHref = prSafeUrl_(back);
  var inner = '';
  if (!embedded) {
    inner += '<div class="pr-top">'
      + (backHref ? '<a class="pr-back" href="' + escapeHtml_(backHref) + '"><span aria-hidden="true">&#8249;</span> Back to the hub</a>' : '')
      + '</div>'
      + '<div class="page-head"><div class="page-kicker">Project Teams Ops Hub</div><div class="page-title">Proposals</div><div class="page-rule"></div></div>';
  }
  inner += '<p class="pr-intro">Scout an improvement to our spaces and post it. Everyone rates each one on impact and effort, and the ideas that win enough support become real projects.</p>';

  // Identity + the earned-voice meter. Kept together because the meter is the answer to
  // "why does my NetID matter": it is the thing that turns your ratings on.
  inner += '<div class="pr-idbar">'
    + '<div class="pr-idrow"><label class="pr-idlbl" for="pr-netid">Your NetID</label>'
    +   '<input id="pr-netid" class="pr-idinput" placeholder="e.g. abc123" autocomplete="off" spellcheck="false" oninput="prSaveNetid()"></div>'
    + '<div class="pr-voice" id="pr-voice" hidden>'
    +   '<div class="pr-voice-segs" id="pr-voice-segs"></div>'
    +   '<span class="pr-voice-note" id="pr-idnote"></span></div>'
    + '</div>';

  inner += '<details class="tp-create-wrap"><summary class="tp-create-toggle"><span class="tp-create-caret">&#43;</span> Propose an improvement</summary>'
    + '<div class="tp-create">'
    + '<div class="tp-field"><label>Title <span class="tp-req">Required</span></label><input id="pr-c-title" placeholder="e.g. Add a shadow board to the hand-tool wall"></div>'
    + '<div class="tp-field"><label>What &amp; why</label><textarea id="pr-c-desc" rows="3" placeholder="What to change, and why it helps"></textarea></div>'
    + '<div class="tp-field"><label>Area / space</label><input id="pr-c-area" placeholder="e.g. Baja bay, composites lab"></div>'
    + '<div class="tp-field"><label>Your name</label><input id="pr-c-name" placeholder="Your name"></div>'
    + '<div class="tp-field"><label>Photos <span class="pr-optional">Optional, show the spot you want to improve</span></label>'
    +   '<input type="file" accept="image/*" multiple id="pr-c-file" style="display:none" onchange="prPickPhotos(this)">'
    +   '<button type="button" class="btn btn-ghost" onclick="document.getElementById(\'pr-c-file\').click()">Add photos</button>'
    +   '<div id="pr-c-stage"></div></div>'
    + '<div class="pr-create-foot"><button type="button" class="btn btn-primary" id="pr-c-go" onclick="prCreate()">Post for review</button>'
    +   '<span id="pr-c-msg" class="tp-lock-msg"></span></div>'
    + '<div class="pr-upbar" id="pr-upbar" hidden><i id="pr-upbar-fill"></i></div>'
    + '</div></details>';

  inner += '<div class="pr-listwrap"><div class="pr-sync" id="pr-sync" hidden></div>'
    + '<div id="pr-list">' + prListSectionsCached_() + '</div></div>';

  var boot = '<script>var PR_NEED=' + PR.quorumMinVotes + ',PR_BAR=' + PR.quorumMinImpact
    + ',PR_VOICE_NEED=' + (PR.earnedVoiceK + 1) + ';</script>';

  return swissShell_('<div class="pr-topbar" id="pr-top" aria-hidden="true"></div>'
    + tpStyles_() + prStyles_() + inner + boot + tpSharedJs_() + prClientJs_() + tpAdminRevealJs_(admin),
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
    + '.pr-topbar{position:fixed;top:0;left:0;height:3px;width:0;z-index:60;background:linear-gradient(90deg,#b31b1b,#e08a1e);opacity:0;transition:width .3s ease,opacity .25s}'
    + '.pr-topbar.on{width:88%;opacity:1;transition:width 9s cubic-bezier(.05,.8,.25,1),opacity .2s}'
    + '.pr-topbar.done{width:100%;opacity:0;transition:width .2s,opacity .45s .15s}'
    + '.pr-top{margin-bottom:6px}'
    + '.pr-back{display:inline-flex;align-items:center;gap:7px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:700;color:#8a857c;text-decoration:none;padding:7px 12px 7px 9px;border-radius:9px;border:1.5px solid transparent;transition:color .15s,background .15s,border-color .15s}'
    + '.pr-back:hover{color:#b31b1b;background:#fff;border-color:#eceae3}'
    + '.pr-back span{font-size:17px;line-height:1}'
    + '.pr-intro{font-size:15px;line-height:1.65;color:#57534e;margin:14px 0 4px;max-width:60ch}'

    // ---- identity + earned voice ----
    + '.pr-idbar{display:flex;align-items:center;gap:10px 22px;flex-wrap:wrap;margin:18px 0 6px;padding:13px 16px;background:#fff;border:1.5px solid #e7e7e3;border-radius:14px;box-shadow:0 2px 8px rgba(20,20,30,.05)}'
    + '.pr-idrow{display:flex;align-items:center;gap:11px;flex:0 1 auto;min-width:0}'
    + '.pr-idlbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a857c;white-space:nowrap}'
    + '.pr-idinput{font:inherit;font-size:14px;padding:10px 13px;border:1.5px solid #e0e0dc;border-radius:10px;outline:none;background:#fafaf8;min-width:0;width:11ch;transition:border-color .15s,box-shadow .15s,background .15s}'
    + '.pr-idinput:focus{border-color:#b31b1b;box-shadow:0 0 0 4px rgba(179,27,27,.12);background:#fff}'
    + '.pr-voice{display:flex;align-items:center;gap:11px;flex:1 1 240px;min-width:0}'
    + '.pr-voice-segs{display:inline-flex;gap:5px;flex:none}'
    + '.pr-voice-seg{width:26px;height:7px;border-radius:99px;background:#eceae3;transition:background .45s cubic-bezier(.2,.9,.3,1.3),transform .45s cubic-bezier(.2,.9,.3,1.3)}'
    + '.pr-voice-seg.on{background:linear-gradient(90deg,#e0a11e,#e08a1e)}'
    + '.pr-voice-seg.pop{transform:scaleY(1.7)}'
    + '.pr-voice.done .pr-voice-seg.on{background:linear-gradient(90deg,#1d9d5b,#157a47)}'
    + '.pr-voice-note{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:600;color:#8a857c;line-height:1.35;min-width:0}'
    + '.pr-voice.done .pr-voice-note{color:#157a47}'

    // ---- create form ----
    + '.pr-optional{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:600;color:#a8a29e;letter-spacing:0;text-transform:none;margin-left:6px}'
    + '.pr-create-foot{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:4px}'
    + '.pr-upbar{height:5px;border-radius:99px;background:#eceae3;overflow:hidden;margin-top:12px}'
    + '.pr-upbar i{display:block;height:100%;width:0;border-radius:99px;background:linear-gradient(90deg,#b31b1b,#e08a1e);transition:width .35s cubic-bezier(.3,.9,.4,1)}'

    // ---- list ----
    + '.pr-listwrap{position:relative}'
    + '.pr-sync{position:absolute;top:-2px;left:0;right:0;height:2px;border-radius:99px;overflow:hidden;background:#eceae3;z-index:3}'
    + '.pr-sync::after{content:"";position:absolute;inset:0;width:38%;border-radius:99px;background:linear-gradient(90deg,transparent,#b31b1b,transparent);animation:prSweep 1.15s linear infinite}'
    + '@keyframes prSweep{0%{transform:translateX(-110%)}100%{transform:translateX(320%)}}'
    + '#pr-list.is-syncing{opacity:.55;transition:opacity .2s}'
    + '.pr-card{animation:prCardIn .42s cubic-bezier(.2,.85,.3,1.05) both}'
    + '@keyframes prCardIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}'

    // ---- card head ----
    + '.pr-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap}'
    + '.pr-head-main{min-width:0;flex:1 1 260px}'
    + '.pr-kicker{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10.5px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#b31b1b}'
    + '.pr-area{color:#a8a29e}'
    + '.pr-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:clamp(17px,2.4vw,20px);font-weight:800;letter-spacing:-.02em;line-height:1.25;color:#14110e;margin:5px 0 0}'
    + '.pr-pill-wrap{flex:none}'
    + '.pr-desc{font-size:14.5px;line-height:1.6;color:#57534e;margin:11px 0 0;max-width:62ch}'
    + '.pr-by{font-size:12.5px;font-weight:600;color:#8a857c;margin-top:9px}'
    + '.pr-photos{display:flex;flex-wrap:wrap;gap:8px;margin-top:13px}'
    + '.pr-photos a{display:inline-block;line-height:0;border-radius:10px;overflow:hidden;transition:transform .16s}'
    + '.pr-photos a:hover{transform:scale(1.03)}'
    + '.pr-photos img{max-height:120px;max-width:100%;border-radius:10px;border:1px solid #ececec;display:block}'

    // ---- quorum meter ----
    + '.pr-quorum{margin-top:16px;padding:13px 15px;border-radius:12px;background:#fbfaf7;border:1.5px solid #efece5}'
    + '.pr-qtop{display:flex;align-items:baseline;justify-content:space-between;gap:10px;flex-wrap:wrap}'
    + '.pr-qcount{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#57534e}'
    + '.pr-qcount b{font-size:16px;font-weight:800;color:#14110e}'
    + '.pr-qsay{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#a8a29e;transition:color .3s}'
    + '.pr-qsay--close{color:#e08a1e}'
    + '.pr-qsay--good{color:#157a47}'
    + '.pr-segs{display:flex;gap:6px;margin-top:9px}'
    + '.pr-seg{flex:1;height:9px;border-radius:99px;background:#eceae3;animation:prSegIn .4s cubic-bezier(.2,.9,.3,1.2) both}'
    + '.pr-seg.on{background:linear-gradient(90deg,#c62b2b,#b31b1b)}'
    + '.pr-seg.pending{background:repeating-linear-gradient(115deg,#e6ded2,#e6ded2 5px,#f3ede4 5px,#f3ede4 10px);animation:prPend 1s linear infinite}'
    + '.pr-seg.pop{animation:prSegPop .5s cubic-bezier(.2,.9,.3,1.5)}'
    + '@keyframes prSegIn{from{transform:scaleX(.25);opacity:0}to{transform:none;opacity:1}}'
    + '@keyframes prSegPop{0%{transform:scaleY(1)}45%{transform:scaleY(2.1)}100%{transform:scaleY(1)}}'
    + '@keyframes prPend{to{background-position:20px 0}}'
    + '.pr-gauges{display:flex;gap:9px 20px;margin-top:12px;flex-wrap:wrap}'
    + '.pr-gauge{display:flex;align-items:center;gap:9px;flex:1 1 190px;min-width:0}'
    + '.pr-glbl{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#a8a29e;flex:none;width:44px}'
    + '.pr-gbar{position:relative;flex:1;height:7px;border-radius:99px;background:#eceae3;min-width:60px}'
    + '.pr-gfill{position:absolute;left:0;top:0;bottom:0;border-radius:99px;background:#c9c3b8;transition:width .7s cubic-bezier(.2,.85,.3,1),background .4s}'
    + '.pr-gfill.good{background:linear-gradient(90deg,#1d9d5b,#157a47)}'
    + '.pr-gfill--eff{background:#c9c3b8}'
    + '.pr-gbar-mark{position:absolute;top:-3px;bottom:-3px;width:2px;border-radius:2px;background:#14110e;opacity:.28}'
    + '.pr-gval{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:800;color:#14110e;flex:none;width:28px;text-align:right;font-variant-numeric:tabular-nums}'
    + '.pr-promoted{display:flex;align-items:center;gap:9px;margin-top:15px;padding:11px 14px;border-radius:11px;background:#f0f9f3;border:1.5px solid #cfe9da;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:600;color:#157a47}'
    + '.pr-promoted-mark{display:inline-flex;align-items:center;justify-content:center;width:19px;height:19px;border-radius:50%;background:#157a47;color:#fff;font-size:11px;font-weight:800;flex:none}'

    // ---- rating control ----
    + '.pr-rate{margin-top:14px;padding:15px 16px;background:#faf9f6;border:1.5px solid #ece9e2;border-radius:12px}'
    + '.pr-rate-row{display:flex;align-items:center;gap:10px 16px;flex-wrap:wrap;margin-bottom:13px}'
    + '.pr-rate-lbl{display:flex;flex-direction:column;gap:2px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:800;letter-spacing:.02em;color:#3f3a31;flex:1 1 190px;min-width:0}'
    + '.pr-rate-lbl em{font-style:normal;font-size:11.5px;font-weight:600;color:#a8a29e;letter-spacing:0}'
    + '.pr-scale-wrap{display:flex;align-items:center;gap:11px;flex:1 1 260px;min-width:0;flex-wrap:wrap}'
    + '.pr-scale{display:inline-flex;gap:6px;flex:0 1 auto}'
    + '.pr-dot{width:42px;height:42px;border-radius:11px;border:1.5px solid #e0e0dc;background:#fff;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:15px;font-weight:800;color:#a8a29e;cursor:pointer;'
    +   'transition:border-color .14s,background .14s,color .14s,transform .16s cubic-bezier(.2,.9,.3,1.6),box-shadow .16s;-webkit-tap-highlight-color:transparent}'
    + '.pr-dot:hover{border-color:#b5b0a8}'
    + '.pr-dot.fill{background:#f3d2d2;border-color:#e6b4b4;color:#8f1515}'
    + '.pr-dot.on{background:#b31b1b;border-color:#b31b1b;color:#fff;box-shadow:0 4px 12px rgba(179,27,27,.3)}'
    + '.pr-dot.preview{border-color:#d99a9a;background:#fdf3f3;color:#8f1515}'
    + '.pr-dot.bump{animation:prDotBump .42s cubic-bezier(.2,.9,.3,1.5)}'
    + '@keyframes prDotBump{0%{transform:scale(1)}40%{transform:scale(1.24)}100%{transform:scale(1)}}'
    + '.pr-dot:active{transform:scale(.93)}'
    + '.pr-word{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#8f1515;min-height:1em;opacity:0;transform:translateY(3px);transition:opacity .25s,transform .25s}'
    + '.pr-word.show{opacity:1;transform:none}'
    + '.pr-verdict{font-size:13px;line-height:1.5;color:#57534e;padding:10px 13px;border-radius:10px;background:#fff;border:1.5px dashed #e6e1d8;margin-bottom:13px;animation:prCardIn .35s ease both}'
    + '.pr-verdict b{color:#14110e;font-weight:800}'
    + '.pr-verdict[hidden]{display:none}'
    + '.pr-rate-foot{display:flex;align-items:center;gap:12px;flex-wrap:wrap}'
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
    + '.pr-card{position:relative;overflow:hidden}'
    + '.pr-card.is-promoting{animation:prLift .9s cubic-bezier(.2,.85,.3,1) both}'
    + '.pr-card.is-promoting::after{content:"";position:absolute;inset:0;pointer-events:none;background:linear-gradient(115deg,transparent 30%,rgba(240,192,80,.55) 50%,transparent 70%);transform:translateX(-100%);animation:prSweepGold 1.1s ease-out both}'
    + '@keyframes prSweepGold{to{transform:translateX(100%)}}'
    + '@keyframes prLift{0%{transform:none;box-shadow:0 2px 8px rgba(20,20,30,.05)}45%{transform:translateY(-7px) scale(1.012);box-shadow:0 20px 44px rgba(224,138,30,.3)}100%{transform:none;box-shadow:0 2px 8px rgba(20,20,30,.05)}}'
    + '.pr-card.is-leaving{animation:prLeave .45s ease-in both}'
    + '@keyframes prLeave{to{opacity:0;transform:translateX(-14px)}}'

    // ---- admin ----
    + '.pr-adminbar{display:flex;gap:9px;flex-wrap:wrap;margin-top:14px;padding-top:13px;border-top:1.5px dashed #ece9e2}'
    + '.pr-del{color:#b31b1b;border-color:#f0d5d5}'
    + '.pr-del:hover{background:#fdf3f3;border-color:#e0a5a5;color:#8f1515}'
    + '.pr-delask{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#57534e;align-self:center;margin-right:2px}'

    // ---- responsive ----
    + '@media(max-width:720px){.pr-quorum{padding:12px 13px}.pr-rate{padding:13px}}'
    + '@media(max-width:600px){'
    +   '.pr-idbar{padding:12px 13px;gap:12px}.pr-idrow{flex:1 1 100%}.pr-idinput{flex:1;width:auto}'
    +   '.pr-voice{flex:1 1 100%}'
    +   '.pr-head{flex-wrap:nowrap}.pr-pill-wrap{align-self:flex-start}'
    +   '.pr-rate-lbl{flex:1 1 100%}'
    +   '.pr-scale-wrap{flex:1 1 100%;gap:8px}'
    +   '.pr-scale{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;width:100%}'
    +   '.pr-dot{width:100%;height:46px;font-size:16px}'
    +   '.pr-gauge{flex:1 1 100%}'
    +   '.pr-submit{flex:1 1 100%}.pr-vmsg{flex:1 1 100%;text-align:left}'
    +   '.pr-adminbar .btn{flex:1 1 auto}'
    +   '.pr-photos img{max-height:96px}'
    + '}'
    + '@media(max-width:380px){.pr-dot{height:42px;font-size:14px;border-radius:9px}.pr-glbl{width:38px}}'

    // ---- motion preferences ----
    + '@media(prefers-reduced-motion:reduce){'
    +   '.pr-card,.pr-verdict{animation:none}'
    +   '.pr-seg,.pr-voice-seg,.pr-gfill,.pr-dot,.pr-word{transition:none;animation:none}'
    +   '.pr-card.is-promoting,.pr-card.is-promoting::after,.pr-card.is-leaving{animation:none}'
    +   '.pr-sync::after{animation:none;width:100%}'
    +   '.pr-spin{animation:none;border-top-color:rgba(255,255,255,.9)}'
    +   '.pr-topbar.on{transition:opacity .2s}'
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
    + 'if(!wrap||!segs||!note)return;'
    + 'if(!prNetid()){wrap.hidden=true;return;}wrap.hidden=false;'
    + 'var need=PR_VOICE_NEED,have=Math.min(count,need),done=have>=need;'
    + 'if(segs.children.length!==need){segs.innerHTML="";for(var i=0;i<need;i++)segs.appendChild(document.createElement("i"));}'
    + '[].slice.call(segs.children).forEach(function(el,i){el.className="pr-voice-seg"+(i<have?" on":"");'
    + 'if(bump&&i===have-1){el.classList.add("pop");setTimeout(function(){el.classList.remove("pop");},450);}});'
    + 'wrap.classList.toggle("done",done);'
    + 'note.textContent=done?"Your ratings count.":(have+" of "+need+" reviews. "+(need-have===1?"One more turns your ratings on.":(need-have)+" more turn your ratings on."));'
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
    + '});}).withFailureHandler(function(){prTop(false);}).prViewerState(nid);}'

    // ---------- quorum meter updates ----------
    + 'function prSay(counted,need,avg,hasVotes){if(!hasVotes)return "be the first to weigh in";'
    + 'if(counted>=need&&avg<PR_BAR)return "enough reviews, impact below the bar";'
    + 'var left=need-counted;if(left===1)return "one more review and it is in";'
    + 'if(left>1)return left+" more reviews to decide";return "quorum reached";}'
    + 'function prRoll(el,to){if(!el)return;var from=parseFloat(el.textContent)||0;if(from===to){el.textContent=to;return;}'
    + 'var t0=Date.now();(function step(){var k=Math.min(1,(Date.now()-t0)/450);'
    + 'el.textContent=Math.round(from+(to-from)*(1-Math.pow(1-k,3)));if(k<1)requestAnimationFrame(step);else el.textContent=to;})();}'
    + 'function prMeter(rid,counted,avgI,avgE){'
    + 'prRoll(document.getElementById(rid+"-n"),counted);'
    + 'var segs=document.getElementById(rid+"-segs");'
    + 'if(segs){[].slice.call(segs.children).forEach(function(el,i){el.classList.remove("pending");'
    + 'var on=i<counted;if(on&&!el.classList.contains("on")){el.classList.add("on","pop");setTimeout(function(){el.classList.remove("pop");},520);}'
    + 'else el.classList.toggle("on",on);});}'
    + 'var gi=document.getElementById(rid+"-gi"),ge=document.getElementById(rid+"-ge");'
    + 'if(gi){gi.style.width=Math.max(0,Math.min(100,avgI/5*100))+"%";gi.classList.toggle("good",avgI>=PR_BAR);}'
    + 'if(ge)ge.style.width=Math.max(0,Math.min(100,avgE/5*100))+"%";'
    + 'var vi=document.getElementById(rid+"-vi"),ve=document.getElementById(rid+"-ve");'
    + 'if(vi)vi.textContent=avgI?avgI.toFixed(1):"\\u2013";if(ve)ve.textContent=avgE?avgE.toFixed(1):"\\u2013";'
    + 'var say=document.getElementById(rid+"-say");'
    + 'if(say){say.textContent=prSay(counted,PR_NEED,avgI,true);'
    + 'say.className="pr-qsay"+(PR_NEED-counted===1?" pr-qsay--close":"")+(counted>=PR_NEED&&avgI>=PR_BAR?" pr-qsay--good":"");}}'

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
    + 'tpConfetti("\\uD83D\\uDC1D Your ratings count from here on.");msg.className="pr-vmsg good";msg.textContent="Your earlier ratings just started counting too.";prRefresh(true);}'
    // Not qualified yet: the rating is saved but deliberately not counted, so say so
    // rather than claiming a result the meter does not show.
    + 'else if(r.voice&&!r.voice.counted){var togo=r.voice.need-r.voice.have;done("\\u2713 Saved","ok");'
    + 'msg.className="pr-vmsg";msg.textContent="Saved. Review "+togo+" more proposal"+(togo===1?"":"s")+" and your ratings start counting, including this one.";}'
    + 'else{done("\\u2713 Counted","ok");msg.className="pr-vmsg good";'
    + 'msg.textContent=r.updated?"Rating updated.":(PR_NEED-r.counted===1?"Counted. One more review and this one is in.":"Rating counted.");'
    + 'if(r.alsoPromoted>0){tpConfetti("\\uD83D\\uDC1D That review carried another proposal over the line.");prRefresh(true);}}'
    + '}).withFailureHandler(function(){done();if(segs)[].slice.call(segs.children).forEach(function(el){el.classList.remove("pending");});'
    + 'msg.className="pr-vmsg bad";msg.textContent="Network hiccup. Try again.";}).prVote(pid,nid,imp,eff,"");}'

    // The promotion beat: hold on the full meter, THEN flip the badge and celebrate. The
    // pause is the point; a reveal with no anticipation barely registers.
    + 'function prPromoteFx(rid){var card=document.getElementById(rid);if(!card){prRefresh(true);return;}'
    + 'card.classList.add("is-promoting");'
    + 'setTimeout(function(){var pill=document.getElementById(rid+"-pill");'
    + 'if(pill)pill.innerHTML=\'<span class="tp-pill tp-pill--done">Approved</span>\';'
    + 'var say=document.getElementById(rid+"-say");if(say){say.textContent="quorum reached";say.className="pr-qsay pr-qsay--good";}'
    + 'tpConfetti("\\uD83D\\uDC1D Quorum reached. It is a real project now.");},620);'
    + 'setTimeout(function(){prRefresh(true);},2400);}'

    // ---------- posting a proposal ----------
    + 'var PRSTAGE=[];'
    + 'function prStageRender(){var el=document.getElementById("pr-c-stage");if(el)el.innerHTML=tpStagePreview(PRSTAGE,"prUnstage(","document.getElementById(\'pr-c-file\').click()");}'
    + 'function prPickPhotos(input){tpStageRead(input,PRSTAGE,prStageRender);}'
    + 'function prUnstage(i){PRSTAGE.splice(i,1);prStageRender();}'
    + 'function prUp(pct){var b=document.getElementById("pr-upbar"),f=document.getElementById("pr-upbar-fill");'
    + 'if(!b||!f)return;b.hidden=pct<0;f.style.width=Math.max(0,pct)+"%";}'
    + 'function prCreate(){var g=function(id){return (document.getElementById(id)||{}).value||"";};'
    + 'var t=g("pr-c-title"),d=g("pr-c-desc"),a=g("pr-c-area"),nm=g("pr-c-name");'
    + 'var msg=document.getElementById("pr-c-msg"),btn=document.getElementById("pr-c-go"),nid=prNetid();'
    + 'if(!t.trim()){msg.className="tp-lock-msg bad";msg.textContent="A title is required.";return;}'
    + 'if(!nid){msg.className="tp-lock-msg bad";msg.textContent="Enter your NetID at the top first.";'
    + 'var f=document.getElementById("pr-netid");if(f)f.focus();return;}'
    + 'msg.className="tp-lock-msg";msg.textContent="";'
    + 'var total=PRSTAGE.length,ids=[],q=PRSTAGE.slice();'
    + 'var stages=total?["Uploading your photos","Saving the photos","Posting your proposal"]:["Posting your proposal","Adding it to the board"];'
    + 'var done=prBusy(btn,stages);if(total)prUp(0);'
    + '(function step(){if(q.length){var ph=q.shift();'
    + 'google.script.run.withSuccessHandler(function(r){'
    + 'if(!r||!r.ok){done();prUp(-1);msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Photo upload failed. Try again.";return;}'
    + 'ids.push(r.photoId);prUp(Math.round(ids.length/total*100));step();'
    + '}).withFailureHandler(function(){done();prUp(-1);msg.className="tp-lock-msg bad";msg.textContent="Photo upload failed. Try again.";}).prUploadPhoto(ph.dataUrl,ph.name);return;}'
    + 'google.script.run.withSuccessHandler(function(r){prUp(-1);'
    + 'if(!r||!r.ok){done();msg.className="tp-lock-msg bad";msg.textContent=(r&&r.error)||"Could not post that.";return;}'
    + 'done("\\u2713 Posted","ok");'
    + '["pr-c-title","pr-c-desc","pr-c-area"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});'
    + 'PRSTAGE=[];prStageRender();'
    + 'msg.className="tp-lock-msg ok";msg.textContent="Posted. The team can rate it now.";'
    + 'var w=document.querySelector(".tp-create-wrap");if(w)w.open=false;'
    + 'tpConfetti("\\uD83D\\uDC1D Posted. Let the swarm decide.");prRefresh();'
    + '}).withFailureHandler(function(){done();prUp(-1);msg.className="tp-lock-msg bad";msg.textContent="Network hiccup. Try again.";})'
    + '.prSubmitProposal(t,d,a,nm,nid,ids.join(", "));})();}'

    // ---------- admin ----------
    + 'function prAdmin(rid,pid,dec){var msg=document.getElementById(rid+"-vmsg"),card=document.getElementById(rid);'
    + 'var btn=card?card.querySelector(".pr-adminbar .btn"):null;'
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
    + '+\'<button type="button" class="btn btn-ghost pr-del" onclick="prDelGo(\\\'\'+rid+\'\\\',\\\'\'+pid+\'\\\')">Yes, delete</button>\''
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
    + 'if(!r||!r.ok)return;list.innerHTML=r.html;'
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
