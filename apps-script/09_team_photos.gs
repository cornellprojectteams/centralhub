// =============================================================================
// Team photos: the Update button's backend
//
// The photos page (team-photos.html) is static, and the scraper needs a real browser, so
// the scrape runs as a GitHub Action (.github/workflows/team-photos.yml). This web app is
// the go-between: the Update button asks it to start the Action, then asks it how the run
// is going. It holds the GitHub token so the page never sees it.
//
// Calls (JSONP, like the inventory: <script src=".../exec?action=status&callback=fn">):
//   action=refresh&admin=1   start a refresh. Answers with a requestId, or with the run
//                            already going (busy: true), or an error code.
//   action=status&id=<requestId>   where that run is. With no id, the latest run.
//   action=status&run=<runId>      the same, by GitHub run id (for a run already going).
//
// Setup (once): new Apps Script project > paste this file > Project Settings > Script
// properties > add GITHUB_TOKEN (a fine-grained token for cornellprojectteams/centralhub
// with Actions: Read and write). Deploy > New deployment > Web app, Execute as: Me,
// Who has access: Anyone. Put the /exec URL in config.js as TEAM_PHOTOS.refreshUrl.
// =============================================================================

var TP_REPO = 'cornellprojectteams/centralhub';
var TP_WORKFLOW = 'team-photos.yml';
var TP_REF = 'main';
var TP_COOLDOWN_MIN = 5;   // a second refresh can't start this soon after the last one

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    out = tpApi_(p);
  } catch (err) {
    out = { ok: false, code: 'error', error: String((err && err.message) || err) };
  }
  var json = JSON.stringify(out);
  var cb = String(p.callback || '');
  if (/^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function tpApi_(p) {
  if (p.action === 'status' && p.run) return { ok: true, run: tpRun_(tpGithub_('get', '/actions/runs/' + encodeURIComponent(p.run))) };
  if (p.action === 'status') return tpStatus_(String(p.id || ''));
  if (p.action === 'refresh') {
    if (p.admin !== '1') return { ok: false, code: 'admin', error: 'Updates start from the admin page.' };
    return tpRefresh_();
  }
  return { ok: false, code: 'action', error: 'Unknown action.' };
}

function tpRefresh_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, code: 'busy_lock', error: 'Another update is starting. Try again in a moment.' };
  try {
    var running = tpRuns_(5).filter(function (r) { return r.status !== 'completed'; })[0];
    if (running) return { ok: true, busy: true, run: tpRun_(running) };

    var props = PropertiesService.getScriptProperties();
    var last = Number(props.getProperty('TP_LAST_START') || 0);
    var wait = Math.ceil((last + TP_COOLDOWN_MIN * 60000 - Date.now()) / 60000);
    if (wait > 0) {
      return { ok: false, code: 'cooldown', minutes: wait,
               error: 'An update just finished. Try again in ' + wait + ' minute' + (wait === 1 ? '' : 's') + '.' };
    }

    var id = Utilities.getUuid().slice(0, 8);
    tpGithub_('post', '/actions/workflows/' + TP_WORKFLOW + '/dispatches', { ref: TP_REF, inputs: { request_id: id } });
    props.setProperty('TP_LAST_START', String(Date.now()));
    return { ok: true, requestId: id };
  } finally {
    lock.releaseLock();
  }
}

function tpStatus_(id) {
  var runs = tpRuns_(10);
  var run = id ? runs.filter(function (r) { return String(r.display_title || '').indexOf(id) !== -1; })[0] : runs[0];
  if (!run) return { ok: true, run: null };   // GitHub takes a few seconds to list a new run
  return { ok: true, run: tpRun_(run) };
}

function tpRuns_(n) {
  var res = tpGithub_('get', '/actions/workflows/' + TP_WORKFLOW + '/runs?per_page=' + n);
  return (res && res.workflow_runs) || [];
}

function tpRun_(r) {
  return {
    id: r.id, status: r.status, conclusion: r.conclusion, url: r.html_url,
    title: r.display_title, startedAt: r.run_started_at || r.created_at, updatedAt: r.updated_at,
  };
}

function tpGithub_(method, path, body) {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    var e0 = new Error('The update service has no GitHub token yet. An admin needs to add GITHUB_TOKEN to its script properties.');
    e0.code = 'no_token';
    throw e0;
  }
  var opts = {
    method: method, muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + TP_REPO + path, opts);
  var code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    var text = res.getContentText();
    return text ? JSON.parse(text) : null;
  }
  var msg = '';
  try { msg = JSON.parse(res.getContentText()).message || ''; } catch (err) { msg = ''; }
  if (code === 401) throw new Error('The GitHub token has expired or is wrong. An admin needs to replace GITHUB_TOKEN.');
  if (code === 403) throw new Error(/rate limit/i.test(msg) ? 'GitHub is limiting requests right now. Try again in a few minutes.'
    : 'The GitHub token is not allowed to run Actions on ' + TP_REPO + '. It needs Actions: Read and write.');
  if (code === 404) throw new Error('GitHub could not find the refresh workflow. Check that .github/workflows/' + TP_WORKFLOW + ' is on the ' + TP_REF + ' branch.');
  throw new Error('GitHub answered ' + code + (msg ? ': ' + msg : '') + '.');
}
