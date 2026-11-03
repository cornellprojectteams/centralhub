// Team member photos page (team-photos.html).
// Reads data/team-photos.json and renders the people grouped by team, with filters for
// team, role, a text search, and whether to include people who graduated or left.
// Filter choices are kept in the URL (?team=Baja&role=lead) so a filtered view can be shared.
(function () {
  'use strict';

  var ROLE_LABEL = { lead: 'Lead', member: 'Member', advisor: 'Advisor' };
  var state = { team: '', role: '', q: '', gone: false };
  var db = null;

  var $ = function (id) { return document.getElementById(id); };

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

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
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

  function personHtml(p) {
    var st = statusOf(p);
    var photo = p.photo
      ? '<a class="tp-photo" href="' + esc(p.photo) + '" target="_blank" rel="noopener"><img src="' + esc(p.photo) +
        '" alt="' + esc(p.name) + '" loading="lazy" referrerpolicy="no-referrer" data-initials="' + esc(initials(p.name)) + '"></a>'
      : '<span class="tp-photo"><span class="tp-initials" title="No photo on the team site">' + esc(initials(p.name)) + '</span></span>';
    var sub = [p.subteam && p.subteam !== p.role ? p.subteam : '', p.classYear ? "'" + String(p.classYear).slice(2) : '']
      .filter(Boolean).join(' · ');
    var tag = st === 'current'
      ? '<span class="tp-tag' + (p.roleType === 'member' ? ' tp-tag--member' : '') + '">' + ROLE_LABEL[p.roleType] + '</span>'
      : '<span class="tp-tag tp-tag--gone">' + (st === 'graduated' ? 'Graduated' : 'Former') + '</span>';
    return '<li class="tp-person">' + photo +
      '<div class="tp-name">' + esc(p.name) + '</div>' +
      '<div class="tp-role-line">' + esc(p.role) + (sub ? '<br>' + esc(sub) : '') + '</div>' + tag + '</li>';
  }

  function render() {
    var shown = db.people.filter(matches);
    var byTeam = {};
    shown.forEach(function (p) { (byTeam[p.team] = byTeam[p.team] || []).push(p); });
    var teams = Object.keys(db.teams).filter(function (t) { return !state.team || t === state.team; });
    var html = teams.map(function (t) {
      var list = byTeam[t] || [];
      var info = db.teams[t] || {};
      if (!list.length && !(state.team && info.note) && !(state.team === t)) return '';
      var leads = list.filter(function (p) { return p.roleType === 'lead'; }).length;
      return '<section class="tp-team">' +
        '<div class="tp-team-head"><h2 class="tp-team-name">' + esc(t) + '</h2>' +
        '<span class="tp-team-meta">' + list.length + ' shown · ' + leads + ' lead' + (leads === 1 ? '' : 's') + '</span>' +
        (info.teamPage || info.site ? '<a class="tp-team-site" href="' + esc(info.teamPage || info.site) + '" target="_blank" rel="noopener">Team site</a>' : '') +
        '</div>' +
        (info.note ? '<p class="tp-team-note">' + esc(info.note) + '.</p>' : '') +
        (list.length ? '<ul class="tp-grid">' + list.map(personHtml).join('') + '</ul>'
          : (info.note ? '' : '<p class="tp-team-note">No one matches these filters.</p>')) +
        '</section>';
    }).join('');
    $('tp-list').innerHTML = html || '<p class="tp-empty">No one matches these filters.</p>';
    $('tp-count').textContent = shown.length + (shown.length === 1 ? ' person' : ' people');
    $('tp-list').querySelectorAll('img[data-initials]').forEach(function (img) {
      img.addEventListener('error', function () {
        var box = img.parentNode;
        box.innerHTML = '<span class="tp-initials" title="Photo no longer loads from the team site">' + esc(img.dataset.initials) + '</span>';
      }, { once: true });
    });
  }

  function syncUrl() {
    var q = new URLSearchParams();
    if (state.team) q.set('team', state.team);
    if (state.role) q.set('role', state.role);
    if (state.gone) q.set('all', '1');
    history.replaceState(null, '', location.pathname + (q.toString() ? '?' + q : ''));
  }

  function setRole(role) {
    state.role = role;
    $('tp-roles').querySelectorAll('.tp-role').forEach(function (b) {
      var on = b.dataset.role === role;
      b.classList.toggle('is-active', on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
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

  function init(data) {
    db = data;
    var params = new URLSearchParams(location.search);
    var counts = {};
    db.people.forEach(function (p) { if (statusOf(p) === 'current') counts[p.team] = (counts[p.team] || 0) + 1; });
    $('tp-team').innerHTML = '<option value="">All teams</option>' + Object.keys(db.teams).map(function (t) {
      return '<option value="' + esc(t) + '">' + esc(t) + ' (' + (counts[t] || 0) + ')</option>';
    }).join('');
    state.team = db.teams[params.get('team')] ? params.get('team') : '';
    state.gone = params.get('all') === '1';
    $('tp-team').value = state.team;
    $('tp-gone').checked = state.gone;
    setRole(ROLE_LABEL[params.get('role')] ? params.get('role') : '');
    $('tp-updated').textContent = new Date(db.updated + 'T12:00:00').toLocaleDateString(undefined, { month: 'long', day: 'numeric', year: 'numeric' });

    $('tp-team').addEventListener('change', function () { state.team = this.value; syncUrl(); render(); });
    $('tp-roles').addEventListener('click', function (e) {
      var b = e.target.closest('.tp-role');
      if (b) { setRole(b.dataset.role); syncUrl(); render(); }
    });
    $('tp-search').addEventListener('input', function () { state.q = this.value.trim().toLowerCase(); render(); });
    $('tp-gone').addEventListener('change', function () { state.gone = this.checked; syncUrl(); render(); });
    $('tp-csv').addEventListener('click', csv);
    render();
  }

  fetch('data/team-photos.json', { cache: 'no-cache' })
    .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then(init)
    .catch(function () {
      $('tp-list').innerHTML = '<p class="tp-empty">Could not load data/team-photos.json. Open this page from the hub site or a local server, not as a file.</p>';
    });
})();
