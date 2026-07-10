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

const CACHE_SECONDS = 300; // 5 min — SpreadsheetApp reads are slow; serve cached JSON between refreshes.

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
    const hit = cache.get('stats');
    if (hit) return hit;
  }
  const json = JSON.stringify(buildStats_());
  cache.put('stats', json, CACHE_SECONDS);
  return json;
}

// --- Assemble the payload --------------------------------------------------

function buildStats_() {
  // Each domain returns { stat, attention[] }. Order is fixed (identity, never by rank).
  // safe_() isolates failures: one unreadable sheet degrades to a single "unavailable"
  // tile instead of 500-ing the whole dashboard.
  const domains = [
    safe_(space_,      'space',      'Space issues',   '#c0392b'),  // LIVE
    safe_(ell_,        'ell',        'Learning Lab',   '#c07a12'),  // stub — TODO wire to ELL Shift Summary
    safe_(fabman_,     'fabman',     'Fabman',         '#2563c9'),  // LIVE
    safe_(compliance_, 'compliance', 'Compliance',     '#1f9d5b'),  // stub — TODO wire to Master Sheet
    safe_(fleet_,      'fleet',      'Fleet · Tacoma', '#7c3aed'),  // stub — TODO wire to Student Drivers
    safe_(inventory_,  'inventory',  'Inventory',      '#0d9488'),  // stub — TODO wire to Tool inventory
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
    sub: overdueList.length + ' overdue · ' + redOpen + ' red · ' + resolved + ' resolved this term',
  };

  const attention = [];
  if (overdueList.length) {
    const oldest = overdueList[0];
    attention.push({
      sev: 'crit', domain: 'Space', dot: '#c0392b',
      title: overdueList.length + (overdueList.length === 1 ? ' space issue is overdue' : ' space issues are overdue'),
      sub: oldest.team ? ('Oldest: ' + oldest.team + (oldest.issueType ? ' ' + oldest.issueType.toLowerCase() : '')) : 'Past SLA deadline',
      pill: 'Overdue', pillType: 'crit',
    });
  }
  if (redOpen) {
    attention.push({
      sev: 'warn', domain: 'Space', dot: '#c0392b',
      title: redOpen + ' red-severity ' + (redOpen === 1 ? 'issue needs' : 'issues need') + ' same-day action',
      sub: 'Reported today and not yet resolved', pill: 'Urgent', pillType: 'warn',
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
  const flags    = cSafety >= 0 ? splitExcluding_(latest[cSafety], 'no issue') : [];
  const supply   = cSupply >= 0 ? splitExcluding_(latest[cSupply], 'all stocked') : [];
  const peak     = topKey_(shiftCounts);

  const bits = [flags.length + (flags.length === 1 ? ' safety flag' : ' safety flags')];
  if (peak) bits.push('peak ' + peak);
  bits.push(logged + ' logged');

  const stat = {
    key: 'ell', name: 'Learning Lab', dot: '#c07a12', live: true,
    value: activity, unit: 'last shift',
    state: flags.length ? 'watch' : 'good',
    sub: bits.join(' · '),
  };

  const attention = [];
  const shiftLabel = (cShift >= 0 && latest[cShift]) ? String(latest[cShift]).trim() + ' shift' : 'last shift';
  if (flags.length) {
    attention.push({
      sev: 'warn', domain: 'ELL', dot: '#c07a12',
      title: flags.length === 1 ? (flags[0] + ' flagged in the ELL')
                                : (flags.length + ' safety issues flagged in the ELL'),
      sub: flags.join(' · ') + ' · ' + shiftLabel,
      pill: 'Safety', pillType: 'warn',
    });
  }
  if (supply.length) {
    attention.push({
      sev: 'warn', domain: 'ELL', dot: '#c07a12',
      title: 'ELL supplies need restocking',
      sub: supply.join(' · ') + ' · ' + shiftLabel,
      pill: 'Supplies', pillType: 'warn',
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

  const cutoff = gradCutoff_();
  let members = 0, locked = 0, awaiting = 0, gradActive = 0;
  for (let i = 1; i < v.length; i++) {
    const row = v[i];
    const present = cId >= 0 ? String(row[cId]).trim() !== '' : row.join('').trim() !== '';
    if (!present) continue;
    members++;
    if (cState >= 0 && norm_(row[cState]) === 'locked') { locked++; continue; }  // locked ≠ awaiting
    if (knowsTraining && !hasTraining_(row, cT1, cT2, cApron)) awaiting++;
    if (cGrad >= 0) {
      const gy = parseInt(String(row[cGrad]).trim(), 10);
      if (!isNaN(gy) && gy <= cutoff) gradActive++;   // graduated, still unlocked
    }
  }

  const newThisWeek = signupsSince_(ss, addDays_(startOfDay_(new Date()), -7));

  const bits = [];
  if (knowsTraining) bits.push(awaiting + ' awaiting training');
  if (gradActive) bits.push(gradActive + ' graduated still active');
  bits.push(newThisWeek + ' new this week');

  const stat = {
    key: 'fabman', name: 'Fabman', dot: '#2563c9', live: true,
    value: fmtNum_(members), unit: 'members',
    state: (awaiting > 0 || gradActive > 0) ? 'watch' : 'good',
    sub: bits.join(' · '),
  };

  const attention = [];
  if (awaiting > 0) {
    attention.push({
      sev: 'warn', domain: 'Fabman', dot: '#2563c9',
      title: awaiting + (awaiting === 1 ? ' Fabman member is' : ' Fabman members are') + ' awaiting training',
      sub: 'Active members with no apron tier recorded yet',
      pill: 'Pending', pillType: 'warn',
    });
  }
  if (gradActive > 0) {
    attention.push({
      sev: 'warn', domain: 'Fabman', dot: '#2563c9',
      title: gradActive + ' graduated ' + (gradActive === 1 ? 'member' : 'members') + ' still have active access',
      sub: 'Graduated ' + cutoff + ' or earlier · machine-shop access should be locked',
      pill: 'Access', pillType: 'warn',
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

function compliance_() {
  // TODO: read SOURCES.master → waiver/Workday expiries by date.
  return {
    stat: { key: 'compliance', name: 'Compliance', dot: '#1f9d5b', live: false,
      value: 8, unit: 'expiring', state: 'action',
      sub: '2 lapsed · 79 Workday complete' },
    attention: [
      { sev: 'crit', domain: 'Compliance', dot: '#1f9d5b',
        title: '2 access waivers have lapsed', sub: 'High Voltage Lab · re-sign required before entry',
        pill: 'Lapsed', pillType: 'crit' },
      { sev: 'warn', domain: 'Compliance', dot: '#1f9d5b',
        title: '8 compliance items expire within 30 days', sub: 'Workday EHS renewals across 5 teams',
        pill: 'Soon', pillType: 'warn' },
    ],
  };
}

function fleet_() {
  // TODO: read SOURCES.drivers → Log tab → today's reservations + overlap detection.
  return {
    stat: { key: 'fleet', name: 'Fleet · Tacoma', dot: '#7c3aed', live: false,
      value: 2, unit: 'booked today', state: 'action',
      sub: '1 conflict · 4 pending approval' },
    attention: [{ sev: 'crit', domain: 'Fleet', dot: '#7c3aed',
      title: 'Tacoma is double-booked today', sub: 'Rocketry & Baja both hold 2:00–5:00 PM',
      pill: 'Conflict', pillType: 'crit' }],
  };
}

function inventory_() {
  // TODO: read SOURCES.inventory → open orders + unreturned locks.
  return {
    stat: { key: 'inventory', name: 'Inventory', dot: '#0d9488', live: false,
      value: 5, unit: 'open orders', state: 'good',
      sub: '2 locks out · 44 mill-room items' },
    attention: [],
  };
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

// First existing tab from a list of candidate names.
function pickSheet_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const s = ss.getSheetByName(names[i]);
    if (s) return s;
  }
  return null;
}

// Index of the first matching header (H is already normalized), else -1.
function findCol_(H, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const j = H.indexOf(norm_(candidates[i]));
    if (j >= 0) return j;
  }
  return -1;
}

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
