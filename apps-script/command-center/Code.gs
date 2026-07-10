/**
 * Ops Command Center — stats API  (SEPARATE Apps Script project)
 * =============================================================================
 * Read-only aggregator. Reads the domain spreadsheets and returns ONE JSON
 * payload the Command Center page renders. This is its OWN project with its OWN
 * web-app deployment — it does NOT share code or a deployment with the Space
 * Status notifier (02_notify_on_submit.gs). Nothing here can affect that URL.
 *
 * DEPLOY (first time):
 *   1. script.google.com → New project → paste this file → name it
 *      "Ops Command Center API".
 *   2. Deploy → New deployment → Web app.
 *        Execute as: Me   ·   Who has access: Anyone
 *   3. Authorize when prompted (it needs to read the spreadsheets below).
 *   4. Copy the /exec URL → paste into command-center.html (STATS_ENDPOINT).
 *   REDEPLOY later: Deploy → Manage deployments → ✏️ → New version (URL stays).
 *     ⚠️ After changing logic, ALSO bump CACHE_VERSION below (or Run ▸
 *     clearStatsCache). Otherwise the cache serves the old code's output for
 *     up to CACHE_SECONDS and the page shows stale, self-contradicting data.
 *
 * The page calls it JSONP-style:  <script src=".../exec?callback=fn">
 * so it works cross-origin from the static site (no CORS dance).
 * =============================================================================
 */

// Source spreadsheets (IDs from SHEETS_INVENTORY.md).
const SOURCES = {
  spaceStatus: '1mZrlnA-GiVKB4_Um21aMH9jsSkf6TUJhauDWivnf7-I', // Space Status Tracking
  ell:         '1yazhxA8nwYWDRSdQEN6Cc3UKKnPn6PWF_b3zAe0KoQk', // ELL Shift Summary (Responses)
  fabman:      '1lDAA4z1YexkjaABSPuXJyEXfSxN2QAsKWVk88TQCRLI', // Project Teams Fabman Admin
  master:      '1k97Ey5aNcm58zaviaXxKvJtTgL1yeSI_fvUkXZFEyM4', // Master Sheet (access + Workday)
  drivers:     '1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M', // Project Teams Student Drivers
  inventory:   '1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU', // Tool inventory, purchasing, locks
};

// Space Status SLA windows (days) by severity — mirrors the notifier.
const SLA_DAYS = { red: 0, orange: 3, yellow: 10, ivory: 21, purple: null };

const CACHE_SECONDS = 180; // SpreadsheetApp reads are slow; serve cached JSON between refreshes.

// Bump this whenever the payload shape or any domain's logic changes.
// CacheService is keyed per SCRIPT, not per deployment version — so without a
// version in the key, a redeploy keeps serving the OLD code's output until the
// TTL expires. That once showed "supplies need restocking" over "all supplies
// stocked". Bumping the key makes a redeploy take effect immediately.
const CACHE_VERSION = 'v6';   // v6: de-duplicated subs; graduated applicants excluded
const CACHE_KEY = 'stats_' + CACHE_VERSION;

// --- Web entry point -------------------------------------------------------

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const json = getStatsJson_(p.fresh === '1');
  if (p.callback) {
    // JSONP: the page loads this via <script>, sidestepping CORS.
    return ContentService
      .createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getStatsJson_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    const hit = cache.get(CACHE_KEY);
    if (hit) return hit;
  }
  const json = JSON.stringify(buildStats_());
  cache.put(CACHE_KEY, json, CACHE_SECONDS);
  return json;
}

/**
 * Belt and braces: run this from the editor (Run ▸ clearStatsCache) right after a
 * redeploy if you didn't bump CACHE_VERSION. Zero-arg, safe to click Run on.
 */
function clearStatsCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  Logger.log('Cleared "' + CACHE_KEY + '". Next request rebuilds from the sheets.');
}

// --- Assemble the payload --------------------------------------------------

function buildStats_() {
  // Each domain returns { stat, attention[] }. Order is fixed (identity, never by rank).
  // safe_() isolates failures: one unreadable sheet degrades to a single "unavailable"
  // tile instead of 500-ing the whole dashboard.
  const domains = [
    safe_(space_,      'space',      'Space issues',   '#c0392b'),
    safe_(ell_,        'ell',        'Learning Lab',   '#c07a12'),
    safe_(fabman_,     'fabman',     'Fabman',         '#2563c9'),
    safe_(fleet_,      'fleet',      'Fleet · Tacoma', '#7c3aed'),
    safe_(inventory_,  'inventory',  'Inventory',      '#0d9488'),
    // Compliance is intentionally NOT shown. The Master Sheet has no expiry or
    // renewal date in any tab, so the tile could only ever state a passive record
    // count — nothing you can act on. compliance_() is kept below, ready to
    // re-enable the moment that sheet gains a renewal-date column.
  ];

  // Flatten attention items, urgent (crit) before soon (warn), stable within.
  // NB: don't use `rank[sev] || 9` — crit's rank is 0 (falsy) and would fall through to 9.
  const sevRank = function (sev) { return sev === 'crit' ? 0 : (sev === 'warn' ? 1 : 9); };
  const attention = [];
  domains.forEach(function (d) { (d.attention || []).forEach(function (a) { attention.push(a); }); });
  attention.sort(function (a, b) { return sevRank(a.sev) - sevRank(b.sev); });

  const urgent = attention.filter(function (a) { return a.sev === 'crit'; }).length;
  const soon = attention.filter(function (a) { return a.sev === 'warn'; }).length;

  const space = domains[0];
  return {
    generatedAt: new Date().toISOString(),
    hero: { attention: attention.length, urgent: urgent, soon: soon },
    attention: attention,
    domains: domains.map(function (d) { return d.stat; }),
    charts: {
      teams: space.charts ? space.charts.teams : [],   // top teams by open space issues
      trend: space.charts ? space.charts.trend : [],   // new space issues / day, last 14
    },
  };
}

// --- SPACE (live) ----------------------------------------------------------

function space_() {
  const sh = SpreadsheetApp.openById(SOURCES.spaceStatus).getSheetByName('Form Responses');
  const v = sh.getDataRange().getValues();
  const H = v[0].map(norm_);
  const col = function (name) { return H.indexOf(norm_(name)); };
  const cTeam = col('Responsible Team'), cStatus = col('Current Status'), cIssue = col('Issue Type'),
        cTs = col('Timestamp'), cTok = col('Issue token'), cAddr = col('Addressed at');

  const today = startOfDay_(new Date());
  const open = [];
  let resolved = 0;
  const perTeam = {};
  const trendBuckets = {};                 // yyyy-mm-dd -> count, last 14 days
  for (let d = 0; d < 14; d++) trendBuckets[dayKey_(addDays_(today, -d))] = 0;

  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    if (cTok < 0 || !row[cTok]) continue;                 // not a notified/tracked issue
    if (cAddr >= 0 && row[cAddr]) { resolved++; continue; } // closed
    const color = parseColor_(row[cStatus]);
    let deadline = null, overdue = false;
    if (color && color !== 'purple') {
      const rd = new Date(row[cTs]);
      if (!isNaN(rd.getTime())) {
        deadline = addDays_(startOfDay_(rd), SLA_DAYS[color]);
        overdue = today > deadline;
      }
    }
    const team = cTeam >= 0 ? String(row[cTeam]).trim() : '';
    const it = {
      team: team, color: color, overdue: overdue, deadline: deadline,
      issueType: cIssue >= 0 ? String(row[cIssue]).trim() : '',
      ts: new Date(row[cTs]),
    };
    open.push(it);
    if (team) perTeam[team] = (perTeam[team] || 0) + 1;
    const k = dayKey_(startOfDay_(it.ts));
    if (k in trendBuckets) trendBuckets[k]++;
  }

  const overdueList = open.filter(function (x) { return x.overdue; })
    .sort(function (a, b) { return (a.deadline || 0) - (b.deadline || 0); });
  const redOpen = open.filter(function (x) { return x.color === 'red' && !x.overdue; }).length;

  // top 5 teams by open count
  const teams = Object.keys(perTeam)
    .map(function (t) { return { name: t, value: perTeam[t] }; })
    .sort(function (a, b) { return b.value - a.value; })
    .slice(0, 5);

  // 14-day trend oldest→newest
  const trend = [];
  for (let d = 13; d >= 0; d--) trend.push(trendBuckets[dayKey_(addDays_(today, -d))]);

  const state = overdueList.length ? 'action' : (open.length ? 'watch' : 'good');
  const stat = {
    key: 'space', name: 'Space issues', dot: '#c0392b', live: true,
    value: open.length, unit: 'open', state: state,
    // Only facts that change what you'd do — never "0 red".
    // Context, not alarms — "1 overdue" already has its own row in the feed.
    sub: resolved ? resolved + ' resolved this term' : 'None resolved yet',
  };
  if (open.length + resolved > 0) {
    stat.meter = { value: open.length, of: open.length + resolved, label: 'unresolved' };
  }

  var describe = function (x) {
    var days = x.deadline ? Math.round((today - x.deadline) / 86400000) : 0;
    return joinBits_([x.team || 'Unassigned', x.issueType || 'Reported issue',
      days > 0 ? (days + (days === 1 ? ' day overdue' : ' days overdue')) : '']);
  };

  const attention = [];
  if (overdueList.length) {
    attention.push({
      sev: 'crit', domain: 'Space', dot: '#c0392b',
      title: overdueList.length + (overdueList.length === 1 ? ' space issue is overdue' : ' space issues are overdue'),
      sub: 'Past the resolution deadline for its severity',
      pill: 'Overdue', pillType: 'crit',
      items: overdueList.slice(0, 8).map(describe),
    });
  }
  if (redOpen) {
    const reds = open.filter(function (x) { return x.color === 'red' && !x.overdue; });
    attention.push({
      sev: 'warn', domain: 'Space', dot: '#c0392b',
      title: redOpen + ' red-severity ' + (redOpen === 1 ? 'issue needs' : 'issues need') + ' same-day action',
      sub: 'Red severity requires same-day resolution', pill: 'Urgent', pillType: 'warn',
      items: reds.slice(0, 8).map(describe),
    });
  }

  return { stat: stat, attention: attention, charts: { teams: teams, trend: trend } };
}

// --- STUB DOMAINS ----------------------------------------------------------
// Illustrative values so the page renders complete. Replace each body with a
// real read (pattern: SpreadsheetApp.openById(SOURCES.x) → count → return).
// live:false lets the page badge these as "sample" until wired.

// --- ELL (live) ------------------------------------------------------------
// Latest shift's activity level + any safety/supply flags from that shift.

function ell_() {
  const sh = pickSheet_(SpreadsheetApp.openById(SOURCES.ell), ['Form Responses 1']);
  if (!sh) throw new Error('no "Form Responses 1" tab');
  const v = sh.getDataRange().getValues();
  if (v.length < 2) throw new Error('no shift rows');
  const H = v[0].map(normh_);

  const cTs     = findColContains_(H, ['timestamp']);
  const cShift  = findColContains_(H, ['shift']);
  const cAct    = findColContains_(H, ['overall activity level']);
  const cSafety = findColContains_(H, ['safety and maintenance']);
  const cSupply = findColContains_(H, ['supply status']);

  // Rows aren't guaranteed sorted — take the newest timestamp, not the last row.
  let latest = null, latestTs = -1, logged = 0;
  const shiftCounts = {};
  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    const t = cTs >= 0 ? new Date(row[cTs]).getTime() : NaN;
    if (isNaN(t)) continue;
    logged++;
    if (cShift >= 0) {
      const s = String(row[cShift]).trim();
      if (s) shiftCounts[s] = (shiftCounts[s] || 0) + 1;
    }
    if (t > latestTs) { latestTs = t; latest = row; }
  }
  if (!latest) throw new Error('no timestamped rows');

  const activity = cAct >= 0 ? firstWord_(latest[cAct]) : '—';
  // "No issues" / "All supplies stocked" are the nothing-to-report options. Match on
  // the distinctive word, not the whole phrase — the sheet's wording drifts.
  const flags    = cSafety >= 0 ? splitExcluding_(latest[cSafety], ['no issue']) : [];
  const supply   = cSupply >= 0 ? splitExcluding_(latest[cSupply], ['stocked']) : [];
  const peak     = topKey_(shiftCounts);

  const stat = {
    key: 'ell', name: 'Learning Lab', dot: '#c07a12', live: true,
    value: activity, unit: 'last shift',
    state: flags.length ? 'watch' : 'good',
    sub: joinBits_([peak ? 'peak ' + peak : '', logged ? logged + ' shifts logged' : '']),
  };

  const attention = [];
  const shiftLabel = (cShift >= 0 && latest[cShift]) ? String(latest[cShift]).trim() + ' shift' : 'last shift';
  if (flags.length) {
    attention.push({
      sev: 'warn', domain: 'ELL', dot: '#c07a12',
      title: flags.length === 1 ? (flags[0] + ' flagged in the ELL')
                                : (flags.length + ' safety issues flagged in the ELL'),
      sub: 'Reported on the ' + shiftLabel,
      pill: 'Safety', pillType: 'warn',
      items: flags,
    });
  }
  if (supply.length) {
    attention.push({
      sev: 'warn', domain: 'ELL', dot: '#c07a12',
      title: 'ELL supplies need restocking',
      sub: 'Flagged on the ' + shiftLabel,
      pill: 'Supplies', pillType: 'warn',
      items: supply,
    });
  }
  return { stat: stat, attention: attention };
}

// --- FABMAN (live) ---------------------------------------------------------
// Reads the sheet that fetchFabmanMembers() already syncs — deliberately NOT the
// Fabman REST API. Reasons: (1) the API needs a secret key, and we refuse to copy
// a credential into a second project; (2) "does this member have training?" is a
// per-member call (~325 HTTP round-trips) that would blow the 6-minute execution
// limit on every page load. The synced sheet holds the same fields.
//
// Schema-tolerant: handles the SU25 "Fetch" tab (Training 1/Training 2 booleans)
// and the current "Dashboard" tab (Apron column). Run describeSources() to confirm.

function fabman_() {
  const ss = SpreadsheetApp.openById(SOURCES.fabman);
  const sh = pickSheet_(ss, ['Dashboard', 'Fetch']);
  if (!sh) throw new Error('no Dashboard/Fetch tab');

  const v = sh.getDataRange().getValues();
  if (v.length < 2) throw new Error('members tab is empty');
  const H = v[0].map(normh_);

  const cId    = findCol_(H, ['id']);              // exact: 'netid' also contains "id"
  const cState = findCol_(H, ['state', 'status']);
  const cApron = findCol_(H, ['apron']);
  const cGrad  = findCol_(H, ['graduation year']);
  const cT1    = findCol_(H, ['training 1', 'training1']);   // legacy "Fetch" schema
  const cT2    = findCol_(H, ['training 2', 'training2']);
  const knowsTraining = (cT1 >= 0 || cT2 >= 0 || cApron >= 0);

  const cTeam = findCol_(H, ['team']);
  const cutoff = gradCutoff_();
  let members = 0, locked = 0, awaiting = 0, gradActive = 0;
  const awaitingByTeam = {}, gradByTeam = {};
  const teamOf = function (row) { return (cTeam >= 0 ? String(row[cTeam]).trim() : '') || 'Unassigned'; };

  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    const present = cId >= 0 ? String(row[cId]).trim() !== '' : row.join('').trim() !== '';
    if (!present) continue;
    members++;
    if (cState >= 0 && norm_(row[cState]) === 'locked') { locked++; continue; }  // locked ≠ awaiting
    if (knowsTraining && !hasTraining_(row, cT1, cT2, cApron)) {
      awaiting++;
      awaitingByTeam[teamOf(row)] = (awaitingByTeam[teamOf(row)] || 0) + 1;
    }
    if (cGrad >= 0) {
      const gy = parseInt(String(row[cGrad]).trim(), 10);
      if (!isNaN(gy) && gy <= cutoff) {           // graduated, still unlocked
        gradActive++;
        gradByTeam[teamOf(row)] = (gradByTeam[teamOf(row)] || 0) + 1;
      }
    }
  }

  const newThisWeek = signupsSince_(ss, addDays_(startOfDay_(new Date()), -7));

  const stat = {
    key: 'fabman', name: 'Fabman', dot: '#2563c9', live: true,
    value: fmtNum_(members), unit: 'members',
    state: (awaiting > 0 || gradActive > 0) ? 'watch' : 'good',
    sub: joinBits_([
      newThisWeek ? newThisWeek + ' new this week' : '',
      locked ? locked + ' locked' : '',
    ]) || 'All members active',
  };
  // No meter here: it would restate the "graduated still active" alarm below.

  const attention = [];
  if (awaiting > 0) {
    attention.push({
      sev: 'warn', domain: 'Fabman', dot: '#2563c9',
      title: awaiting + (awaiting === 1 ? ' Fabman member is' : ' Fabman members are') + ' awaiting training',
      sub: 'They cannot use shop machines until an apron tier is assigned',
      pill: 'Pending', pillType: 'warn',
      items: groupTop_(awaitingByTeam, 8),
    });
  }
  if (gradActive > 0) {
    attention.push({
      sev: 'warn', domain: 'Fabman', dot: '#2563c9',
      title: gradActive + ' graduated ' + (gradActive === 1 ? 'member' : 'members') + ' still have active access',
      sub: 'Class of ' + cutoff + ' and earlier — lock them to revoke machine access',
      pill: 'Access', pillType: 'warn',
      items: groupTop_(gradByTeam, 8),
    });
  }
  return { stat: stat, attention: attention };
}

// Students who finished in May are "graduated" from June onward.
function gradCutoff_() {
  const now = new Date();
  return now.getMonth() >= 5 ? now.getFullYear() : now.getFullYear() - 1;
}

// A member "has training" if either boolean training column is set, or (newer
// schema) an apron tier is recorded. Unknown schema -> assume yes, so we never
// invent a backlog out of missing columns.
function hasTraining_(row, cT1, cT2, cApron) {
  if (cT1 >= 0 || cT2 >= 0) return truthy_(row[cT1]) || truthy_(row[cT2]);
  if (cApron >= 0) return String(row[cApron]).trim() !== '';
  return true;
}

function truthy_(v) {
  if (v === true) return true;
  const s = norm_(v);
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'x';
}

// Count Sign-up Form submissions on/after `since`.
function signupsSince_(ss, since) {
  const sh = pickSheet_(ss, ['Sign-up Form', 'Sign up Form', 'Signup Form', 'Form Responses 1']);
  if (!sh) return 0;
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return 0;
  const c = findCol_(v[0].map(norm_), ['timestamp']);
  if (c < 0) return 0;
  let n = 0;
  for (let i = 1; i < v.length; i++) {
    const d = new Date(v[i][c]);
    if (!isNaN(d.getTime()) && d >= since) n++;
  }
  return n;
}

// --- COMPLIANCE (live, but informational) ----------------------------------
// NOTE: every Master Sheet tab is a historical sign-up log (DATE | NAME | netID |
// TEAM). There is NO expiry or renewal date anywhere, so "expiring in 30 days" /
// "lapsed" CANNOT be computed — we do not fake it. To make this tile actionable,
// the sheet needs either a renewal-date column or a stated validity window (e.g.
// "usage agreements are valid 12 months from DATE"); then this becomes real.

function compliance_() {
  const ss = SpreadsheetApp.openById(SOURCES.master);
  const byName = sheetsByTrimmedName_(ss);
  const accessTabs = ['ELL', 'HVL', 'RHODES PENTHOUSE', 'Aquatic Center', 'Risk Waivers',
                      'Private Vehicle Risk Waiver', 'Canoe waiver', 'Tang Room 403', 'Hollister room B55'];

  let total = 0;
  const counts = {};
  accessTabs.forEach(function (name) {
    const sh = byName[norm_(name)];
    const rows = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
    counts[name] = rows;
    total += rows;
  });
  const wd = byName[norm_('Workday Trainings')];
  const workday = wd ? Math.max(0, wd.getLastRow() - 1) : 0;

  return {
    stat: {
      key: 'compliance', name: 'Compliance', dot: '#1f9d5b', live: true,
      value: fmtNum_(total), unit: 'records on file', state: 'good',
      sub: counts['ELL'] + ' ELL · ' + counts['HVL'] + ' HVL · ' + workday + ' Workday complete',
    },
    attention: [],   // nothing actionable until an expiry rule exists
  };
}

// --- FLEET (live, best-effort) ---------------------------------------------
// The Log tab's row 1 is a banner ("Project Teams Vehicle Reservations Toyota
// Tacoma…"), not headers — so we scan the first few rows for the real header row.
// If we can't find it we throw, and safe_() renders an honest "unavailable" tile
// rather than inventing a booking. Run peekRows('drivers','Log',5) to confirm.

function fleet_() {
  const sh = pickSheet_(SpreadsheetApp.openById(SOURCES.drivers), ['Log']);
  if (!sh) throw new Error('no Log tab');
  const v = sh.getDataRange().getValues();

  // Real header lives on row 3: Team | Name | NetID | Phone Number |
  // Date (doubleclick) | Checkout Time | Return Time | Destination | Parking Location
  const hr = findHeaderRow_(v, ['team', 'name', 'date'], 6);
  if (hr < 0) throw new Error('header row not found in Log — run peekFleetLog()');
  const H = v[hr].map(normh_);

  const cDate  = findColContains_(H, ['date']);
  const cTeam  = findColContains_(H, ['team']);
  const cStart = findColContains_(H, ['checkout time', 'checkout', 'start']);
  const cEnd   = findColContains_(H, ['return time', 'return', 'end']);
  const cDest  = findColContains_(H, ['destination']);
  if (cDate < 0) throw new Error('no date column in Log');
  if (cStart < 0 || cEnd < 0) throw new Error('no checkout/return time columns — cannot detect conflicts');

  const today = startOfDay_(new Date());
  const weekEnd = addDays_(today, 7);
  const todays = [], upcoming = [], incomplete = [];

  for (let i = hr + 1; i < v.length; i++) {
    const row = v[i];
    const d = new Date(row[cDate]);
    if (isNaN(d.getTime())) continue;
    const d0 = startOfDay_(d);
    if (d0.getTime() === today.getTime()) todays.push(row);
    if (d0 < today) continue;                                  // past bookings aren't actionable
    if (d0 > weekEnd) continue;                                // a trip in 3 months isn't today's problem
    upcoming.push(row);

    // A reservation that can't be honoured: no times, or no stated destination.
    const missing = [];
    if (toMinutes_(row[cStart]) == null) missing.push('checkout time');
    if (toMinutes_(row[cEnd]) == null) missing.push('return time');
    if (cDest >= 0 && !String(row[cDest]).trim()) missing.push('destination');
    if (missing.length) incomplete.push({ row: row, missing: missing, date: d0 });
  }
  incomplete.sort(function (a, b) { return a.date - b.date; });   // soonest first

  // Overlap detection on today's bookings: [checkout, return) intervals that intersect.
  // This endpoint is PUBLIC — label bookings by team only. Never fall back to the
  // driver's name; that would publish student PII to anyone holding the URL.
  const label = function (row) {
    const t = cTeam >= 0 ? String(row[cTeam]).trim() : '';
    return t || 'Unlabelled booking';
  };
  const clashes = [];
  for (let i = 0; i < todays.length; i++) {
    for (let j = i + 1; j < todays.length; j++) {
      const a0 = toMinutes_(todays[i][cStart]), a1 = toMinutes_(todays[i][cEnd]);
      const b0 = toMinutes_(todays[j][cStart]), b1 = toMinutes_(todays[j][cEnd]);
      if (a0 == null || a1 == null || b0 == null || b1 == null) continue;
      if (a0 < b1 && b0 < a1) clashes.push(label(todays[i]) + ' & ' + label(todays[j]));
    }
  }

  const appr = safeCall_(fleetApproval_, null);   // approvals must never break bookings

  const stat = {
    key: 'fleet', name: 'Fleet · Tacoma', dot: '#7c3aed', live: true,
    value: todays.length, unit: 'booked today',
    state: clashes.length ? 'action'
         : ((incomplete.length || (appr && (appr.ready || appr.blocked || appr.missingInfo))) ? 'watch' : 'good'),
    sub: (upcoming.length ? upcoming.length + ' booked this week' : 'Nothing booked this week'),
  };

  const attention = [];
  if (clashes.length) {
    attention.push({
      sev: 'crit', domain: 'Fleet', dot: '#7c3aed',
      title: 'Tacoma is double-booked today',
      sub: 'Two teams hold overlapping checkout times — one must move',
      pill: 'Conflict', pillType: 'crit',
      items: clashes.slice(0, 8),
    });
  }
  if (incomplete.length) {
    attention.push({
      sev: 'warn', domain: 'Fleet', dot: '#7c3aed',
      title: incomplete.length + ' reservation' + (incomplete.length === 1 ? '' : 's') + ' this week ' + (incomplete.length === 1 ? 'is' : 'are') + ' missing details',
      sub: 'Fleet will not release the truck without times and a destination',
      pill: 'Incomplete', pillType: 'warn',
      // Team + date + what's missing. Never the driver's name — public endpoint.
      items: incomplete.slice(0, 8).map(function (x) {
        return joinBits_([
          (cTeam >= 0 ? String(x.row[cTeam]).trim() : '') || 'Unlabelled booking',
          fmtDay_(x.date),
          'missing ' + x.missing.join(' + '),
        ]);
      }),
    });
  }

  // A `sub` says what happens NEXT. It never restates the title, and never
  // mentions spreadsheet mechanics ("column P") at a human.
  if (appr && appr.ready > 0) {
    attention.push({
      sev: 'warn', domain: 'Fleet', dot: '#7c3aed',
      title: appr.ready + ' driver application' + (appr.ready === 1 ? ' is' : 's are') + ' ready to file',
      sub: 'Email the rows to Fleet Services, then mark them requested',
      pill: 'Ready', pillType: 'warn',
      items: groupTop_(appr.readyByTeam, 8),
    });
  }
  if (appr && appr.blocked > 0) {
    attention.push({
      sev: 'warn', domain: 'Fleet', dot: '#7c3aed',
      title: appr.blocked + ' driver application' + (appr.blocked === 1 ? '' : 's') + ' cannot be processed',
      sub: 'Fleet Services will reject these until the missing item is initialled',
      pill: 'Blocked', pillType: 'warn',
      items: groupTop_(appr.byRequirement, 6),   // which requirement, not who
    });
  }
  if (appr && appr.missingInfo > 0) {
    attention.push({
      sev: 'warn', domain: 'Fleet', dot: '#7c3aed',
      title: appr.missingInfo + ' driver application' + (appr.missingInfo === 1 ? '' : 's') + ' missing required info',
      sub: 'Ask the applicant to complete their details before it can be filed',
      pill: 'Missing info', pillType: 'warn',
      items: groupTop_(appr.infoByTeam, 8),
    });
  }
  return { stat: stat, attention: attention };
}

// --- FLEET: driver applications (Approval tab) ------------------------------
// Header is on row 2; row 1 is instructions. The sheet states its own rules:
//   ">> Complete all fields (application not processed with missing info)"
//   ">> Initial each upon completion (… without safety, fleet, and release)"
// and the Process tab says "update column P when filed" — column P is `requested`.
// Columns 17+ are a summary/pivot block, not driver rows; NetID gates them out.
//
// PUBLIC ENDPOINT: never emit Email, Name or NetID. Group by team / requirement.

function fleetApproval_() {
  const sh = pickSheet_(SpreadsheetApp.openById(SOURCES.drivers), ['Approval']);
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  const hr = findHeaderRow_(v, ['netid', 'name', 'team'], 5);
  if (hr < 0) return null;
  const H = v[hr].map(normh_);

  const cNetId   = findColContains_(H, ['netid']);
  const cName    = findColContains_(H, ['name']);
  const cGrad    = findColContains_(H, ['grad year']);
  const cTeam    = findColContains_(H, ['team']);
  const cSafety  = findColContains_(H, ['driver safety']);
  const cApp     = findColContains_(H, ['fleet application']);
  const cRelease = findColContains_(H, ['sign release']);
  const cReq     = findColContains_(H, ['requested']);
  if (cNetId < 0 || cSafety < 0 || cApp < 0 || cRelease < 0) return null;

  const REQUIRED = [
    { col: cSafety,  label: 'driver safety (RMI2100)' },
    { col: cApp,     label: 'fleet application' },
    { col: cRelease, label: 'sign release' },
  ];
  const filled = function (row, c) { return c >= 0 && String(row[c]).trim() !== ''; };

  let ready = 0, blocked = 0, missingInfo = 0;
  const byRequirement = {}, readyByTeam = {}, infoByTeam = {};
  const cutoff = gradCutoff_();

  for (let i = hr + 1; i < v.length; i++) {
    const row = v[i];
    const netid = String(row[cNetId] || '').trim();
    if (!netid || norm_(netid) === 'count') continue;      // skip blanks + the pivot block
    const team = (cTeam >= 0 ? String(row[cTeam]).trim() : '') || 'Unassigned';

    // This tab is a years-long log, not a queue. An application for someone who
    // has already graduated can never be actioned — counting it drowns the feed.
    const gy = parseInt(String(row[cGrad] || '').trim(), 10);
    if (!isNaN(gy) && gy <= cutoff) continue;

    if (!filled(row, cName) || !filled(row, cGrad) || !filled(row, cTeam)) {
      missingInfo++;
      infoByTeam[team] = (infoByTeam[team] || 0) + 1;
      continue;                                            // can't be processed regardless
    }
    const gaps = REQUIRED.filter(function (r) { return !filled(row, r.col); });
    if (gaps.length) {
      blocked++;
      gaps.forEach(function (g) { byRequirement[g.label] = (byRequirement[g.label] || 0) + 1; });
    } else if (!filled(row, cReq)) {
      ready++;                                             // column P blank = never filed
      readyByTeam[team] = (readyByTeam[team] || 0) + 1;
    }
  }
  return { ready: ready, blocked: blocked, missingInfo: missingInfo,
           byRequirement: byRequirement, readyByTeam: readyByTeam, infoByTeam: infoByTeam };
}

// --- INVENTORY (live) ------------------------------------------------------
// Open order  = a "Mill room Tooling" item with a blank "Ordered" cell.
// Lock out    = a "Locks" row with a "date out" but no "returned" date.

function inventory_() {
  const ss = SpreadsheetApp.openById(SOURCES.inventory);

  let openOrders = 0, millItems = 0;
  const mill = pickSheet_(ss, ['Mill room Tooling']);
  if (mill) {
    const v = mill.getDataRange().getValues();
    const H = v[0].map(normh_);
    const cEq = findColContains_(H, ['equipment']);
    const cOrd = findColContains_(H, ['ordered']);
    for (let i = 1; i < v.length; i++) {
      const eq = cEq >= 0 ? String(v[i][cEq]).trim() : '';
      if (!eq) continue;
      millItems++;
      if (cOrd >= 0 && String(v[i][cOrd]).trim() === '') openOrders++;
    }
  }

  let locksOut = 0;
  const lockItems = [];
  const locks = pickSheet_(ss, ['Locks']);
  if (locks) {
    const v = locks.getDataRange().getValues();
    const H = v[0].map(normh_);
    const cOut = findColContains_(H, ['date out']);
    const cRet = findColContains_(H, ['returned']);
    const cTeam = findColContains_(H, ['team']);
    const cEquip = findColContains_(H, ['equipment']);
    const cLock = findColContains_(H, ['lock #', 'lock']);
    for (let i = 1; i < v.length; i++) {
      const out = cOut >= 0 ? String(v[i][cOut]).trim() : '';
      const ret = cRet >= 0 ? String(v[i][cRet]).trim() : '';
      if (!out || ret) continue;
      locksOut++;
      // NOTE: this endpoint is public — never emit "Key assigned to" or NetID.
      lockItems.push(joinBits_([
        cTeam >= 0 ? String(v[i][cTeam]).trim() : '',
        cEquip >= 0 ? String(v[i][cEquip]).trim() : '',
        cLock >= 0 && String(v[i][cLock]).trim() ? 'Lock #' + String(v[i][cLock]).trim() : '',
      ]));
    }
  }

  const stat = {
    key: 'inventory', name: 'Inventory', dot: '#0d9488', live: true,
    value: openOrders, unit: 'open orders',
    state: locksOut > 0 ? 'watch' : 'good',
    sub: millItems ? millItems + ' mill-room items tracked' : 'No items tracked',
  };
  if (millItems > 0) {
    stat.meter = { value: openOrders, of: millItems, label: 'not yet ordered' };
  }

  const attention = [];
  if (locksOut > 0) {
    attention.push({
      sev: 'warn', domain: 'Inventory', dot: '#0d9488',
      title: locksOut + ' lock' + (locksOut === 1 ? '' : 's') + ' not returned',
      sub: 'Chase the key holder recorded in the Locks tab',
      pill: 'Outstanding', pillType: 'warn',
      items: lockItems.slice(0, 8),
    });
  }
  return { stat: stat, attention: attention };
}

// --- helpers ---------------------------------------------------------------

// Run one domain; on failure log it and return an "unavailable" tile so a single
// unreadable sheet can't take down the whole dashboard.
function safe_(fn, key, name, dot) {
  try {
    return fn();
  } catch (err) {
    Logger.log('domain "' + key + '" failed: ' + err);
    return {
      stat: { key: key, name: name, dot: dot, live: false, value: '—',
              unit: 'unavailable', state: 'good', sub: 'Could not read source sheet' },
      attention: [],
    };
  }
}

// Run a sub-reader that must never take its parent domain down with it.
function safeCall_(fn, fallback) {
  try { return fn(); } catch (err) { Logger.log('sub-reader failed: ' + err); return fallback; }
}

// First existing tab from a list of candidate names.
function pickSheet_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const s = ss.getSheetByName(names[i]);
    if (s) return s;
  }
  return null;
}

// Index of the first EXACT matching header (H is already normalized), else -1.
function findCol_(H, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const j = H.indexOf(norm_(candidates[i]));
    if (j >= 0) return j;
  }
  return -1;
}

// Index of the first header CONTAINING any needle. Use when headers carry noise
// like "Safety and Maintenance Issues  (check all that apply)".
function findColContains_(H, needles) {
  for (let n = 0; n < needles.length; n++) {
    const needle = normh_(needles[n]);
    for (let i = 0; i < H.length; i++) {
      if (H[i].indexOf(needle) >= 0) return i;
    }
  }
  return -1;
}

// Some tabs open with a banner/instructions instead of headers. Find the first
// row (within maxScan) whose cells contain every required token.
function findHeaderRow_(values, tokens, maxScan) {
  const limit = Math.min(values.length, maxScan || 5);
  for (let r = 0; r < limit; r++) {
    const cells = values[r].map(normh_);
    const ok = tokens.every(function (t) {
      return cells.some(function (c) { return c.indexOf(normh_(t)) >= 0; });
    });
    if (ok) return r;
  }
  return -1;
}

// Tab names sometimes carry stray whitespace ("Tang Room 403 ").
function sheetsByTrimmedName_(ss) {
  const map = {};
  ss.getSheets().forEach(function (sh) { map[norm_(sh.getName())] = sh; });
  return map;
}

// Header normalizer: lowercase, trim, and collapse internal runs of whitespace.
function normh_(s) { return norm_(s).replace(/\s+/g, ' '); }

// Minutes since midnight from a sheet time cell (Date object or "2:00 PM").
function toMinutes_(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) return cell.getHours() * 60 + cell.getMinutes();
  const s = norm_(cell);
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[:.](\d{2})\s*(am|pm)?/);
  if (!m) return null;
  let h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  if (m[3] === 'pm' && h < 12) h += 12;
  if (m[3] === 'am' && h === 12) h = 0;
  return h * 60 + min;
}

function firstWord_(s) {
  const t = String(s == null ? '' : s).trim();
  return t ? t.split(/[\s(]/)[0] : '—';
}

// Split a multi-select cell, dropping any "nothing to report" options.
function splitExcluding_(cell, excludes) {
  const ex = (Array.isArray(excludes) ? excludes : [excludes]).map(norm_);
  return String(cell == null ? '' : cell).split(',')
    .map(function (x) { return x.trim(); })
    .filter(function (x) {
      if (!x) return false;
      const n = norm_(x);
      return !ex.some(function (e) { return n.indexOf(e) >= 0; });
    });
}

function topKey_(obj) {
  let best = null, n = -1;
  Object.keys(obj).forEach(function (k) { if (obj[k] > n) { n = obj[k]; best = k; } });
  return best;
}

function fmtNum_(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

// Join only the parts that carry information — a "0 red" is noise, not a fact.
function joinBits_(parts) {
  return parts.filter(function (p) { return p; }).join(' · ');
}

// {Baja:7, Formula:5, …} -> ["Baja — 7", "Formula — 5", "+3 more teams"]
function groupTop_(counts, limit) {
  const keys = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; });
  const out = keys.slice(0, limit).map(function (k) { return k + ' — ' + counts[k]; });
  const rest = keys.length - limit;
  if (rest > 0) out.push('+' + rest + ' more team' + (rest === 1 ? '' : 's'));
  return out;
}

/**
 * Print the first `n` rows of one tab — for tabs whose row 1 is a banner rather
 * than headers (e.g. drivers/Log).
 *
 * The editor's Run button calls functions with NO arguments, so these default to
 * the Fleet log. For another tab, call it from a zero-arg wrapper like
 * peekFleetLog() below rather than clicking Run on this one.
 */
function peekRows(sourceKey, tabName, n) {
  sourceKey = sourceKey || 'drivers';
  tabName = tabName || 'Log';
  n = n || 5;
  const id = SOURCES[sourceKey];
  if (!id) { Logger.log('unknown source "' + sourceKey + '" — one of: ' + Object.keys(SOURCES).join(', ')); return; }
  const sh = SpreadsheetApp.openById(id).getSheetByName(tabName);
  if (!sh) { Logger.log('no tab "' + tabName + '"'); return; }
  const rows = Math.min(n, sh.getLastRow());
  const v = sh.getRange(1, 1, rows, sh.getLastColumn()).getValues();
  v.forEach(function (r, i) { Logger.log('row ' + (i + 1) + ' :: ' + r.join(' | ')); });
}

// Zero-arg wrappers, safe to click Run on.
function peekFleetLog()      { peekRows('drivers', 'Log', 5); }
function peekFleetApproval() { peekRows('drivers', 'Approval', 4); }

/**
 * SCHEMA DISCOVERY — run manually from the editor (Run ▸ describeSources), then
 * copy the Execution log. Prints every tab name, row count, and header row for
 * each source spreadsheet: exactly what's needed to wire the remaining domains.
 *
 * Deliberately NOT exposed over the web app — that endpoint is public ("Anyone"),
 * and this would leak your sheet structure to the internet.
 */
function describeSources() {
  Object.keys(SOURCES).forEach(function (key) {
    try {
      const ss = SpreadsheetApp.openById(SOURCES[key]);
      Logger.log('=== ' + key + ' — ' + ss.getName() + ' ===');
      ss.getSheets().forEach(function (sh) {
        const lastCol = sh.getLastColumn();
        const hdr = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
        Logger.log('  [' + sh.getName() + '] rows=' + sh.getLastRow() + ' :: ' + hdr.join(' | '));
      });
    } catch (err) {
      Logger.log('=== ' + key + ' — UNREADABLE: ' + err);
    }
  });
}

function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
function startOfDay_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function dayKey_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmtDay_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d'); }

// "Red (same-day)" / "🟥 Red" / "orange" -> canonical severity key.
function parseColor_(s) {
  const t = norm_(s);
  if (!t) return '';
  if (t.indexOf('red') >= 0) return 'red';
  if (t.indexOf('orange') >= 0) return 'orange';
  if (t.indexOf('yellow') >= 0) return 'yellow';
  if (t.indexOf('ivory') >= 0) return 'ivory';
  if (t.indexOf('purple') >= 0) return 'purple';
  return '';
}
