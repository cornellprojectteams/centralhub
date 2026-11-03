// Team member photos page (team-photos.html).
// Reads data/team-photos.json and renders the people grouped by team, with filters for
// team, role, a text search, and whether to include people who graduated or left.
// Filter choices are kept in the URL (?team=Baja&role=lead) so a filtered view can be shared.
//
// Update button (admin only, ?admin=1): asks the update service (apps-script/09_team_photos.gs,
// URL in config.js as TEAM_PHOTOS.refreshUrl) to start the GitHub Action that rescrapes the
// team sites, follows the run, then waits for GitHub Pages to publish the new data and
// reloads it. A run in progress is remembered in the browser, so a reload keeps following it.
(function () {
  'use strict';

  var ROLE_LABEL = { lead: 'Lead', member: 'Member', advisor: 'Advisor' };
  var params = new URLSearchParams(location.search);
  var ADMIN = params.get('admin') === '1';
  var REFRESH_URL = (window.TEAM_PHOTOS && window.TEAM_PHOTOS.refreshUrl) || '';
  var ACTIONS_URL = 'https://github.com/cornellprojectteams/centralhub/actions/workflows/team-photos.yml';
  var GUIDE_URL = 'https://github.com/cornellprojectteams/centralhub/blob/main/tools/team-photos/README.md';
  var JOB_KEY = 'teamPhotosJob';

  var state = { team: '', role: '', q: '', gone: false };
  var db = null;

  var $ = function (id) { return document.getElementById(id); };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function store(k, v) { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* private mode */ } }
  function recall(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } }

  // ------------------------------------------------------------------ people

  // Class of Y graduates in May of Y, so from June of Y on they count as graduated.
  function lastGradYear() {
    var now = new Date();
    return now.getMonth() >= 5 ? now.getFullYear() : now.getFullYear() - 1;
  }

  function statusOf(p) {
    if (p.status === 'graduated' || (p.classYear && p.classYear <= lastGradYear())) return 'graduated';
    if (p.status === 'left') return 'left';
    return 'current';
  }

  function initials(name) {
    var w = name.replace(/^(prof\.?|professor|dr\.?)\s+/i, '').split(/\s+/);
    return ((w[0] || '')[0] || '') + (w.length > 1 ? w[w.length - 1][0] : '');
  }

  function matches(p) {
    if (state.team && p.team !== state.team) return false;
    if (state.role && p.roleType !== state.role) return false;
    if (!state.gone && statusOf(p) !== 'current') return false;
    if (state.q) {
      var hay = (p.name + ' ' + p.role + ' ' + (p.subteam || '') + ' ' + p.team).toLowerCase();
      if (hay.indexOf(state.q) === -1) return false;
    }
    return true;
  }

  function personHtml(p, i) {
    var st = statusOf(p);
    var badge = st !== 'current'
      ? '<span class="tp-badge tp-badge--gone">' + (st === 'graduated' ? 'Graduated' : 'Former') + '</span>'
      : p.roleType === 'lead' ? '<span class="tp-badge">Lead</span>'
      : p.roleType === 'advisor' ? '<span class="tp-badge tp-badge--advisor">Advisor</span>' : '';
    var ini = '<span class="tp-initials" aria-hidden="true">' + esc(initials(p.name)) + '</span>';
    var photo = p.photo
      ? '<a class="tp-photo" href="' + esc(p.photo) + '" target="_blank" rel="noopener" aria-label="Open photo of ' + esc(p.name) + '">' + ini +
        '<img src="' + esc(p.photo) + '" alt="' + esc(p.name) + '" loading="lazy" referrerpolicy="no-referrer">' + badge + '</a>'
      : '<span class="tp-photo" title="No photo on the team site">' + ini + badge + '</span>';
    var sub = [p.subteam && p.subteam !== p.role ? p.subteam : '', p.classYear ? "Class of " + p.classYear : '']
      .filter(Boolean).join(' · ');
    return '<li class="tp-person" style="--i:' + Math.min(i, 24) + '">' + photo +
      '<div class="tp-name">' + esc(p.name) + '</div>' +
      '<div class="tp-role-line">' + esc(p.role) + (sub ? '<span class="tp-sub">' + esc(sub) + '</span>' : '') + '</div></li>';
  }

  function skeleton() {
    var tile = '<li><div class="tp-shimmer"></div><p class="tp-shimmer"></p><p class="tp-shimmer"></p></li>';
    var block = '<div class="tp-skel" aria-hidden="true"><div class="tp-skel-head tp-shimmer"></div><ul class="tp-grid">' +
      new Array(13).join(tile) + '</ul></div>';
    $('tp-list').innerHTML = block + block;
  }

  function renderUpdated() {
    var d = new Date(db.updated + 'T12:00:00');
    $('tp-updated').textContent = 'Updated ' + d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function render() {
    var shown = db.people.filter(matches);
    var byTeam = {};
    shown.forEach(function (p) { (byTeam[p.team] = byTeam[p.team] || []).push(p); });
    var names = Object.keys(db.teams);
    var html = names.map(function (t, idx) {
      var list = byTeam[t] || [];
      var info = db.teams[t] || {};
      if (!list.length && !(state.team === t && info.note)) return '';
      var leads = list.filter(function (p) { return p.roleType === 'lead'; }).length;
      var num = String(idx + 1).padStart(2, '0');
      return '<section class="tp-team" id="team-' + esc(t.replace(/\W+/g, '-').toLowerCase()) + '">' +
        '<div class="tp-team-head"><span class="tp-team-num">' + num + '</span>' +
        '<div><h2 class="tp-team-name">' + esc(t) + '</h2><span class="tp-team-meta">' + list.length + ' ' + (list.length === 1 ? 'person' : 'people') +
        ' · ' + leads + ' lead' + (leads === 1 ? '' : 's') + '</span></div>' +
        (info.teamPage || info.site ? '<a class="tp-team-site" href="' + esc(info.teamPage || info.site) + '" target="_blank" rel="noopener">Team site &#8599;</a>' : '<span></span>') +
        '</div>' +
        (info.note ? '<p class="tp-team-note">' + esc(info.note) + '.</p>' : '') +
        (list.length ? '<ul class="tp-grid">' + list.map(personHtml).join('') + '</ul>'
          : (info.note ? '' : '<p class="tp-team-note">No one on this team matches these filters.</p>')) +
        '</section>';
    }).join('');
    $('tp-list').innerHTML = html ||
      '<div class="tp-empty"><b>No one matches.</b>Try another spelling, a different role, or turn on graduated &amp; former.<br><button type="button" id="tp-reset">Clear filters</button></div>';
    $('tp-list').setAttribute('aria-busy', 'false');
    $('tp-count').textContent = shown.length.toLocaleString() + (shown.length === 1 ? ' person' : ' people');
    var reset = $('tp-reset');
    if (reset) reset.addEventListener('click', clearFilters);
    $('tp-list').querySelectorAll('.tp-photo img').forEach(function (img) {
      var done = function () { img.classList.add('is-loaded'); };
      if (img.complete && img.naturalWidth) done();
      else img.addEventListener('load', done, { once: true });
      img.addEventListener('error', function () { img.remove(); }, { once: true });   // initials show through
    });
  }

  function syncUrl() {
    var q = new URLSearchParams();
    if (ADMIN) q.set('admin', '1');
    if (state.team) q.set('team', state.team);
    if (state.role) q.set('role', state.role);
    if (state.gone) q.set('all', '1');
    history.replaceState(null, '', location.pathname + (q.toString() ? '?' + q : '') + location.hash);
  }

  function setRole(role) {
    state.role = role;
    var active = null;
    $('tp-roles').querySelectorAll('.tp-role').forEach(function (b) {
      var on = b.dataset.role === role;
      b.classList.toggle('is-active', on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
      if (on) active = b;
    });
    var thumb = $('tp-roles-thumb');
    if (active) {
      thumb.style.width = active.offsetWidth + 'px';
      thumb.style.transform = 'translateX(' + active.offsetLeft + 'px)';
    }
  }

  function clearFilters() {
    state = { team: '', role: '', q: '', gone: false };
    $('tp-team').value = '';
    $('tp-search').value = '';
    $('tp-gone').checked = false;
    setRole('');
    syncUrl();
    render();
  }

  function csv() {
    var cols = ['team', 'name', 'role', 'roleType', 'subteam', 'classYear', 'status', 'photo', 'sourcePage'];
    var rows = db.people.filter(matches).map(function (p) {
      var r = Object.assign({}, p, { status: statusOf(p) });
      return cols.map(function (c) { return '"' + String(r[c] == null ? '' : r[c]).replace(/"/g, '""') + '"'; }).join(',');
    });
    var blob = new Blob([cols.join(',') + '\n' + rows.join('\n')], { type: 'text/csv' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'team-photos' + (state.team ? '-' + state.team.replace(/\W+/g, '-') : '') + '.csv';
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  function fillTeams() {
    var counts = {};
    db.people.forEach(function (p) { if (statusOf(p) === 'current') counts[p.team] = (counts[p.team] || 0) + 1; });
    $('tp-team').innerHTML = '<option value="">All teams</option>' + Object.keys(db.teams).map(function (t) {
      return '<option value="' + esc(t) + '">' + esc(t) + ' (' + (counts[t] || 0) + ')</option>';
    }).join('');
    $('tp-team').value = db.teams[state.team] ? state.team : '';
  }

  function loadData(bust) {
    return fetch('data/team-photos.json' + (bust ? '?t=' + Date.now() : ''), { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      });
  }

  function show(data) {
    db = data;
    fillTeams();
    renderUpdated();
    render();
  }

  // ------------------------------------------------------------------ update button

  var job = null;        // { requestId, runId, startedAt, phase }
  var pollTimer = null;
  var misses = 0;

  var STEPS = ['Starting', 'Waiting for GitHub', 'Visiting team sites', 'Publishing', 'Done'];

  function setButton(busy, label) {
    var b = $('tp-update');
    b.disabled = busy;
    b.classList.toggle('is-busy', busy);
    $('tp-update-label').textContent = label || 'Update from team sites';
  }

  // kind: busy | ok | info | error. step: index into STEPS while busy.
  function setStatus(o) {
    var box = $('tp-status');
    if (!o) { box.hidden = true; box.innerHTML = ''; return; }
    var icon = o.kind === 'busy' ? '<span class="tp-spinner" aria-hidden="true"></span>'
      : o.kind === 'ok' ? '<svg class="tp-mark" viewBox="0 0 24 24" fill="none" stroke="#1f9d5b" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"></circle><path d="M7.5 12.5l3 3 6-6.5"></path></svg>'
      : o.kind === 'info' ? '<svg class="tp-mark" viewBox="0 0 24 24" fill="none" stroke="#c88a12" stroke-width="2.6" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="10"></circle><path d="M12 8v5M12 16.5v.01"></path></svg>'
      : '<svg class="tp-mark" viewBox="0 0 24 24" fill="none" stroke="#b31b1b" stroke-width="2.6" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="10"></circle><path d="M15 9l-6 6M9 9l6 6"></path></svg>';
    var steps = o.kind === 'busy' && o.step != null
      ? '<ol class="tp-steps">' + STEPS.slice(0, 4).map(function (s, i) {
          return '<li class="tp-step' + (i < o.step ? ' is-done' : i === o.step ? ' is-now' : '') + '">' + s + '</li>';
        }).join('') + '</ol><div class="tp-bar" aria-hidden="true"></div>'
      : '';
    var actions = (o.actions || []).map(function (a) {
      return a.href
        ? '<a href="' + esc(a.href) + '" target="_blank" rel="noopener">' + esc(a.label) + '</a>'
        : '<button type="button" data-act="' + esc(a.act) + '">' + esc(a.label) + '</button>';
    }).join('');
    box.className = 'tp-status tp-status--' + o.kind;
    box.innerHTML = '<div class="tp-status-title">' + icon + '<span>' + esc(o.title) + '</span></div>' +
      (o.detail ? '<div class="tp-status-detail">' + o.detail + '</div>' : '') + steps +
      (actions ? '<div class="tp-status-actions">' + actions + '</div>' : '');
    box.hidden = false;
  }

  function fail(title, detail, extra) {
    stopPolling();
    setButton(false);
    setStatus({ kind: 'error', title: title, detail: detail, actions: extra || [{ act: 'retry', label: 'Try again' }] });
  }

  // JSONP call to the update service (no CORS setup needed from the static site).
  var seq = 0;
  function api(query, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var cb = '__tp' + Date.now().toString(36) + (++seq);
      var q = new URLSearchParams(query);
      q.set('callback', cb);
      var s = document.createElement('script');
      var done = false;
      var finish = function (fn, v) {
        if (done) return;
        done = true;
        clearTimeout(timer);
        window[cb] = function () {};   // a late reply after a timeout is ignored
        if (s.parentNode) s.parentNode.removeChild(s);
        fn(v);
      };
      var timer = setTimeout(function () {
        var e = new Error('timeout'); e.kind = 'timeout'; finish(reject, e);
      }, timeoutMs || 25000);
      window[cb] = function (data) {
        if (data && data.ok) finish(resolve, data);
        else { var e = new Error((data && data.error) || 'The update service sent back an error.'); e.kind = 'server'; e.data = data; finish(reject, e); }
      };
      s.onerror = function () { var e = new Error('network'); e.kind = 'network'; finish(reject, e); };
      s.src = REFRESH_URL + (REFRESH_URL.indexOf('?') >= 0 ? '&' : '?') + q.toString();
      document.head.appendChild(s);
    });
  }

  function elapsed() {
    var s = Math.max(0, Math.round((Date.now() - job.startedAt) / 1000));
    return Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0');
  }

  function startUpdate() {
    if (!REFRESH_URL) {
      setStatus({
        kind: 'info', title: 'The update service isn’t connected yet.',
        detail: 'An admin needs to deploy the update web app once and put its URL in <code>config.js</code>. The guide has the steps.',
        actions: [{ href: GUIDE_URL + '#connect-the-update-button', label: 'Open the guide' }],
      });
      return;
    }
    if (!navigator.onLine) { fail('You’re offline.', 'Connect to the internet and try again.'); return; }
    setButton(true, 'Starting…');
    setStatus({ kind: 'busy', title: 'Starting the update…', step: 0 });
    api({ action: 'refresh', admin: '1' }, 25000).then(function (res) {
      job = { requestId: res.requestId || '', runId: res.run ? res.run.id : null, startedAt: Date.now(), phase: 'run' };
      if (res.busy && res.run && res.run.startedAt) job.startedAt = Date.parse(res.run.startedAt) || Date.now();
      store(JOB_KEY, job);
      if (res.busy) setStatus({ kind: 'busy', title: 'An update is already running. Following it…', step: 1 });
      follow();
    }).catch(function (e) {
      if (e.kind === 'network' || e.kind === 'timeout') fail('Couldn’t reach the update service.', 'Check your connection. If this keeps happening, the web app is probably asking visitors to sign in: in Apps Script, Deploy &gt; Manage deployments &gt; edit, and set <b>Who has access</b> to <b>Anyone</b>. Otherwise the URL in <code>config.js</code> may be wrong or the deployment was removed.',
        [{ act: 'retry', label: 'Try again' }, { href: GUIDE_URL + '#connect-the-update-button', label: 'Setup guide' }]);
      else if (e.data && e.data.code === 'cooldown') {
        setButton(false);
        setStatus({ kind: 'info', title: 'Already up to date.', detail: esc(e.message) });
      } else fail('The update couldn’t start.', esc(e.message), [{ act: 'retry', label: 'Try again' }, { href: GUIDE_URL + '#connect-the-update-button', label: 'Setup guide' }]);
    });
  }

  function stopPolling() { clearTimeout(pollTimer); pollTimer = null; }

  function follow() {
    stopPolling();
    misses = 0;
    setButton(true, 'Updating…');
    poll();
  }

  function poll() {
    var minutes = (Date.now() - job.startedAt) / 60000;
    if (minutes > 40) {
      store(JOB_KEY, null);
      fail('This is taking much longer than usual.', 'Updates normally finish in under 10 minutes. Check the run on GitHub.',
        [{ href: ACTIONS_URL, label: 'Open GitHub Actions' }, { act: 'retry', label: 'Start again' }]);
      return;
    }
    if (job.phase === 'publish') { pollPublish(); return; }
    var q = job.runId ? { action: 'status', run: job.runId } : { action: 'status', id: job.requestId };
    api(q, 20000).then(function (res) {
      misses = 0;
      var run = res.run;
      if (!run) {
        if (Date.now() - job.startedAt > 120000) {
          store(JOB_KEY, null);
          fail('GitHub didn’t start the update.', 'The request was sent, but no run showed up after two minutes. Actions may be turned off for the repository.',
            [{ act: 'retry', label: 'Try again' }, { href: ACTIONS_URL, label: 'Open GitHub Actions' }]);
          return;
        }
        setStatus({ kind: 'busy', title: 'Waiting for GitHub to pick it up…', step: 1 });
        return next(5000);
      }
      job.runId = run.id;
      store(JOB_KEY, job);
      if (run.status !== 'completed') {
        var running = run.status === 'in_progress';
        setStatus({
          kind: 'busy', step: running ? 2 : 1,
          title: running ? 'Visiting the team sites… ' + elapsed() : 'Waiting for a GitHub runner…',
          detail: running ? 'Opening every team page and reading the photos. This usually takes 4 to 8 minutes. You can leave this page; it picks up again when you come back.' : '',
          actions: [{ href: run.url, label: 'Watch on GitHub' }],
        });
        return next(running ? 6000 : 4000);
      }
      if (run.conclusion !== 'success') {
        store(JOB_KEY, null);
        var why = run.conclusion === 'cancelled' ? 'The run was cancelled.' : 'Something went wrong while scraping or saving.';
        fail('The update failed on GitHub.', why + ' Nothing on this page was changed.',
          [{ href: run.url, label: 'View the run log' }, { act: 'retry', label: 'Try again' }]);
        return;
      }
      job.phase = 'publish';
      job.publishAt = Date.now();
      store(JOB_KEY, job);
      pollPublish();
    }).catch(function (e) {
      misses++;
      if (misses >= 4) {
        stopPolling();
        setButton(false);
        setStatus({
          kind: 'error', title: 'Lost contact with the update service.',
          detail: (e.kind === 'server' ? esc(e.message) + ' ' : '') + 'The update may still be running on GitHub.',
          actions: [{ act: 'resume', label: 'Check again' }, { href: ACTIONS_URL, label: 'Open GitHub Actions' }],
        });
        return;
      }
      next(5000 * misses);
    });
  }

  // The run committed new data; GitHub Pages takes a minute or two to publish it.
  function pollPublish() {
    setStatus({ kind: 'busy', step: 3, title: 'Publishing the new photos…', detail: 'GitHub Pages is putting the new data live. This takes a minute or two.' });
    loadData(true).then(function (data) {
      if (data.lastRun && String(data.lastRun.runId) === String(job.runId)) {
        store(JOB_KEY, null);
        stopPolling();
        setButton(false);
        show(data);
        var r = data.lastRun;
        var bits = [r.added + ' new', r.changed + ' changed', r.left + ' no longer listed'];
        var warn = [];
        if (r.failedTeams && r.failedTeams.length) warn.push('Couldn’t load ' + r.failedTeams.join(', ') + '; their people were kept as they were.');
        if (r.shrunkTeams && r.shrunkTeams.length) warn.push(r.shrunkTeams.join(', ') + ' listed far fewer people than before, so nobody there was marked former. The team page may have moved.');
        setStatus({
          kind: warn.length ? 'info' : 'ok',
          title: 'Updated. ' + r.found.toLocaleString() + ' people on the team sites.',
          detail: bits.join(' · ') + (warn.length ? '<br>' + esc(warn.join(' ')) : ''),
        });
        return;
      }
      throw new Error('not yet');
    }).catch(function () {
      if (Date.now() - job.publishAt > 6 * 60000) {
        store(JOB_KEY, null);
        stopPolling();
        setButton(false);
        setStatus({
          kind: 'info', title: 'The update finished, but the site hasn’t published it yet.',
          detail: 'GitHub Pages is slower than usual. Reload the page in a few minutes to see the new photos.',
          actions: [{ act: 'reload', label: 'Reload now' }],
        });
        return;
      }
      next(10000);
    });
  }

  function next(ms) { pollTimer = setTimeout(poll, ms); }

  function setupUpdate() {
    if (!ADMIN) return;
    $('tp-update').hidden = false;
    $('tp-update').addEventListener('click', startUpdate);
    $('tp-status').addEventListener('click', function (e) {
      var b = e.target.closest('[data-act]');
      if (!b) return;
      if (b.dataset.act === 'retry') { setStatus(null); startUpdate(); }
      else if (b.dataset.act === 'resume') follow();
      else if (b.dataset.act === 'reload') location.reload();
    });
    var saved = recall(JOB_KEY);
    if (saved && REFRESH_URL && Date.now() - saved.startedAt < 40 * 60000) {
      job = saved;
      setStatus({ kind: 'busy', title: 'Checking on the update you started…', step: saved.phase === 'publish' ? 3 : 1 });
      follow();
    } else if (saved) {
      store(JOB_KEY, null);
    }
  }

  // ------------------------------------------------------------------ start

  function init(data) {
    state.team = params.get('team') || '';
    state.gone = params.get('all') === '1';
    $('tp-gone').checked = state.gone;
    setRole(ROLE_LABEL[params.get('role')] ? params.get('role') : '');
    show(data);

    $('tp-team').addEventListener('change', function () { state.team = this.value; syncUrl(); render(); });
    $('tp-roles').addEventListener('click', function (e) {
      var b = e.target.closest('.tp-role');
      if (b) { setRole(b.dataset.role); syncUrl(); render(); }
    });
    var t;
    $('tp-search').addEventListener('input', function () {
      var v = this.value.trim().toLowerCase();
      clearTimeout(t);
      t = setTimeout(function () { state.q = v; render(); }, 120);
    });
    $('tp-gone').addEventListener('change', function () { state.gone = this.checked; syncUrl(); render(); });
    $('tp-csv').addEventListener('click', csv);
    window.addEventListener('resize', function () { setRole(state.role); });
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { setRole(state.role); });
  }

  skeleton();
  setupUpdate();
  loadData(false).then(init).catch(function (e) {
    $('tp-list').setAttribute('aria-busy', 'false');
    var file = location.protocol === 'file:';
    $('tp-list').innerHTML = '<div class="tp-empty"><b>Couldn’t load the photos.</b>' +
      (file ? 'This page has to be opened from the hub site or a local server, not as a file.'
        : 'The data file didn’t load (' + esc(e.message) + '). Check your connection and try again.') +
      '<br><button type="button" onclick="location.reload()">Try again</button></div>';
  });
})();
