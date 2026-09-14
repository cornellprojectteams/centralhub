/**
 * Project Teams Ops Hub client script (shared by index.html and admin.html)
 *
 * Live search, the Command Center badge, the staff people directory, the
 * weekly desk calendar, the open-issues dashboard panel, and the admin
 * sidebar scroll-spy. Every feature is guarded by element presence, so the
 * same file runs on the staff page and the admin page unchanged.
 *
 * Markup conventions:
 *   - .hub-zone          The page's single panel (staff on index, admin on admin.html)
 *   - .category          Searchable block (task-block or admin-block)
 *   - .action            Tool link or button; may carry data-keywords
 */

(function () {
  'use strict';

  // Deployed Space Status web app /exec URL. Set in config.js (single source of
  // truth, shared with team.html); the literal below is only a fallback.
  const SPACE_STATUS_URL = window.SPACE_STATUS_URL || 'https://script.google.com/macros/s/AKfycbwNbGjVcBrcsMZiOl2nXzpqZHz04nvKLm9D_aC0VJDz7Xxxf_4kLKlNSOHubPXj1X74/exec';

  // Ops Command Center stats API. Used only to put a live attention count on the
  // Command button. The page still works fine if it never answers.
  const STATS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzSfUQgCfkOpXAmEExVh1bHIbIQ7LAipTzP_uW2x2XKKTmZuGvmxpPI6gS1cw7oLVOz/exec';

  /**
   * Badge the Command Center button with how many things need attention.
   * JSONP, because a plain fetch() to script.google.com is blocked by CORS.
   * Fails silently: a missing badge is better than a broken masthead.
   */
  function loadCommandBadge() {
    const badge = document.getElementById('cmd-badge');
    if (!badge || !STATS_ENDPOINT) return;

    const cb = '__hub' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    let settled = false;
    const cleanup = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    };
    const timer = setTimeout(cleanup, 12000);

    window[cb] = data => {
      const n = data && data.hero ? Number(data.hero.attention) : 0;
      cleanup();
      if (!Number.isFinite(n) || n <= 0) return;   // nothing wrong -> no badge
      badge.textContent = n > 99 ? '99+' : String(n);
      badge.hidden = false;
      const link = badge.closest('.cmd-link');
      if (link) link.setAttribute('aria-label', `Command Center, ${n} items need attention`);
    };
    script.onerror = cleanup;
    script.src = STATS_ENDPOINT + '?callback=' + cb;
    document.body.appendChild(script);
  }

  loadCommandBadge();
  loadPeopleLoop();
  loadScheduleLoop();
  bindSchedChrome();

  /**
   * Staff-page people directory. Reads name / NetID / email / resource-for /
   * Roles from the Contact Info tab (JSONP). Phone numbers and CU IDs stay in
   * the sheet. Gibran, Kate, Lauren, and Noah are Admin; their duties come
   * from “Can be a Resource for…”. Role codes (`Ops`, `Apps- dev`, …) are a
   * stream-plus-focus shorthand. People without an email still appear; their
   * card is not a mailto.
   * Fails silently: a missing directory is better than a broken staff page.
   */
  let peopleGroupFilter = 'all';

  function loadPeopleLoop() {
    const section = document.getElementById('people');
    const directory = document.getElementById('people-directory');
    const filters = document.getElementById('people-filters');
    const sheet = window.PEOPLE_SHEET;
    if (!section || !directory || !sheet || !sheet.id) return;

    const cb = '__people' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    let settled = false;
    const cleanup = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    };
    const fail = () => {
      cleanup();
      section.hidden = true;
    };
    const timer = setTimeout(fail, 12000);

    window[cb] = payload => {
      cleanup();
      const people = parsePeoplePayload(payload);
      if (!people.length) { section.hidden = true; return; }

      const keywords = ['people', 'roles', 'staff', 'team', 'contact', 'email', 'admin']
        .concat(people.flatMap(p => [p.name, p.role, p.roleLabel, p.specialty, p.resource, p.email, p.netid]))
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
      section.dataset.keywords = keywords;
      renderPeopleDirectory(people, directory, filters);
      const countEl = document.getElementById('people-count');
      if (countEl) {
        countEl.textContent = String(people.length);
        countEl.hidden = false;
      }
      directory.classList.remove('is-loading');
    };
    script.onerror = fail;

    const params = new URLSearchParams({
      gid: String(sheet.gid || '0'),
      tq: 'select A, C, E, M, O',
      tqx: 'out:json;responseHandler:' + cb,
    });
    script.src = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(sheet.id) + '/gviz/tq?' + params.toString();
    document.body.appendChild(script);
  }

  // Sheet shorthand → hub labels. Prefix is the workstream; the rest is the
  // specialty. Unknown future codes title-case instead of showing the raw cell.
  const PEOPLE_STREAM = {
    ops: { id: 'ops', label: 'Operations' },
    prj: { id: 'projects', label: 'Projects' },
    apps: { id: 'apps', label: 'Apps' },
    digi: { id: 'digital', label: 'Digital' },
    lead: { id: 'lead', label: 'Lead' },
    coord: { id: 'coord', label: 'Coordinator' },
    admin: { id: 'admin', label: 'Admin' },
    staff: { id: 'staff', label: 'Staff' },
  };
  const PEOPLE_FOCUS = {
    dev: 'Development',
    coms: 'Communications',
    canvas: 'Canvas',
    chem: 'Chemistry',
    composites: 'Composites',
    fab: 'Fabrication',
  };
  const PEOPLE_STREAM_ORDER = ['admin', 'ops', 'projects', 'apps', 'digital', 'lead', 'coord', 'staff'];
  const PEOPLE_ADMIN = { gae27: 1, kmr87: 1, ls948: 1, nhh5: 1 };

  function parsePeoplePayload(payload) {
    const rows = payload && payload.table && payload.table.rows;
    if (!Array.isArray(rows)) return [];
    const seen = {};
    const out = [];
    rows.forEach(row => {
      const name = peopleCell(row, 0);
      if (!name || /^email$/i.test(name) || /^roles?$/i.test(name)) return;
      const key = name.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      const netid = peopleCell(row, 1);
      const email = peopleMailto(peopleCell(row, 2));
      const resource = peopleCell(row, 3);
      const rawRole = peopleCell(row, 4);
      const person = { name: name, netid: netid, email: email, role: rawRole, resource: resource };
      const resolved = peopleResolveRole(person);
      out.push({
        name: name,
        netid: netid,
        email: email,
        resource: resource,
        role: rawRole,
        roleId: resolved.id,
        roleLabel: resolved.label,
        specialty: resolved.specialty,
      });
    });
    out.sort((a, b) => {
      const ao = peopleOrder_(a.roleId);
      const bo = peopleOrder_(b.roleId);
      if (ao !== bo) return ao - bo;
      return a.name.localeCompare(b.name);
    });
    return out;
  }

  function peopleOrder_(id) {
    const i = PEOPLE_STREAM_ORDER.indexOf(id || 'staff');
    return i < 0 ? PEOPLE_STREAM_ORDER.length : i;
  }

  function peopleCell(row, i) {
    const c = row && row.c && row.c[i];
    if (!c || c.v == null) return '';
    return String(c.v).trim();
  }

  function peopleIsAdmin(p) {
    const net = String(p.netid || '').toLowerCase();
    if (PEOPLE_ADMIN[net]) return true;
    const nm = String(p.name || '').toLowerCase();
    return nm === 'gibran' || nm === 'kate' || nm === 'lauren' || nm === 'noah' || nm.indexOf('noah hamm') === 0;
  }

  function peopleTitleCase(s) {
    return String(s || '').split(/\s+/).filter(Boolean).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
  }

  function peopleResolveRole(person) {
    const raw = String(person.role || '').replace(/\s+/g, ' ').trim();
    if (peopleIsAdmin(person) || !raw) {
      return { id: 'admin', label: 'Admin', specialty: String(person.resource || '').replace(/\s+/g, ' ').trim() };
    }
    const s = raw.toLowerCase();
    const parts = s.split(/[\s-]+/).filter(Boolean);
    const stream = PEOPLE_STREAM[parts[0]] || { id: parts[0], label: peopleTitleCase(parts[0]) };
    const specialty = PEOPLE_FOCUS[parts[1]] || (parts.length > 1 ? peopleTitleCase(parts.slice(1).join(' ')) : '');
    return {
      id: stream.id,
      label: specialty ? stream.label + ' \u00b7 ' + specialty : stream.label,
      specialty: specialty,
    };
  }

  function peopleGroupLabel(id) {
    if (PEOPLE_STREAM[id]) return PEOPLE_STREAM[id].label;
    return peopleTitleCase(id || 'Staff');
  }

  function peopleInitials(name) {
    const parts = String(name || '').split(/\s+/).filter(Boolean);
    if (!parts.length) return '';
    if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
  }

  function peopleMailto(email) {
    const e = String(email || '').trim();
    return /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(e) ? e : '';
  }

  function renderPeopleDirectory(people, directory, filters) {
    const groups = [];
    const byId = {};
    people.forEach(p => {
      const id = p.roleId || 'staff';
      if (!byId[id]) {
        const g = { id: id, label: peopleGroupLabel(id), people: [] };
        byId[id] = g;
        groups.push(g);
      }
      byId[id].people.push(p);
    });
    groups.sort((a, b) => peopleOrder_(a.id) - peopleOrder_(b.id));

    if (filters) {
      filters.hidden = false;
      filters.setAttribute('role', 'group');
      filters.setAttribute('aria-label', 'Filter by role');
      filters.innerHTML = '<button type="button" class="people-chip is-active" data-group="all" aria-pressed="true">All</button>'
        + groups.map(g => '<button type="button" class="people-chip" data-group="' + g.id + '" aria-pressed="false">'
          + escapePeople(g.label) + ' <span>' + g.people.length + '</span></button>').join('');
      filters.querySelectorAll('.people-chip').forEach(btn => {
        btn.addEventListener('click', () => {
          peopleGroupFilter = btn.dataset.group || 'all';
          filters.querySelectorAll('.people-chip').forEach(b => {
            const on = b === btn;
            b.classList.toggle('is-active', on);
            b.setAttribute('aria-pressed', on ? 'true' : 'false');
          });
          applyPeopleFilter();
        });
      });
    }

    directory.innerHTML = groups.map(g => {
      return '<section class="people-group" data-group="' + g.id + '">'
        + '<h4 class="people-group-label">' + escapePeople(g.label)
        + ' <span>' + g.people.length + '</span></h4>'
        + '<div class="people-list">' + g.people.map(peopleCardHtml).join('') + '</div>'
        + '</section>';
    }).join('');
  }

  function applyPeopleFilter() {
    const directory = document.getElementById('people-directory');
    if (!directory) return;
    directory.querySelectorAll('.people-group').forEach(g => {
      const show = peopleGroupFilter === 'all' || g.dataset.group === peopleGroupFilter;
      g.hidden = !show;
    });
  }

  function peopleCardHtml(p) {
    const tone = p.roleId ? ' people-card--' + p.roleId : '';
    const detail = p.specialty
      ? '<span class="people-role">' + escapePeople(p.specialty) + '</span>'
      : '';
    const mail = p.email
      ? '<span class="people-email">' + escapePeople(p.email) + '</span>'
      : '';
    const keywords = [p.name, p.role, p.roleLabel, p.specialty, p.resource, p.email, p.netid, peopleGroupLabel(p.roleId)]
      .filter(Boolean).join(' ').toLowerCase();
    const inner = '<span class="people-avatar" aria-hidden="true">' + escapePeople(peopleInitials(p.name)) + '</span>'
      + '<span class="people-body">'
      + '<span class="people-name">' + escapePeople(p.name) + '</span>'
      + detail
      + mail
      + '</span>';
    if (p.email) {
      return '<a class="people-card' + tone + '" href="mailto:' + p.email + '" data-keywords="' + escapePeople(keywords) + '">' + inner + '</a>';
    }
    return '<div class="people-card people-card--static' + tone + '" data-keywords="' + escapePeople(keywords) + '">' + inner + '</div>';
  }

  function escapePeople(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /**
   * Staff-page week calendar. Reads the FA25 schedule tab (day, shift, names)
   * via JSONP and paints a Mon–Sun desk grid. The first sheet row is the
   * header, which is also Monday’s 4–6pm shift. Fails silently.
   */
  const SCHED_DAYS = [
    { id: 'monday', short: 'Mon', full: 'Monday' },
    { id: 'tuesday', short: 'Tue', full: 'Tuesday' },
    { id: 'wednesday', short: 'Wed', full: 'Wednesday' },
    { id: 'thursday', short: 'Thu', full: 'Thursday' },
    { id: 'friday', short: 'Fri', full: 'Friday' },
    { id: 'saturday', short: 'Sat', full: 'Saturday' },
    { id: 'sunday', short: 'Sun', full: 'Sunday' },
  ];
  const SCHED_PALETTE = [
    { bg: '#fde8e8', fg: '#6e1010', bar: '#b31b1b' },
    { bg: '#f8edd4', fg: '#6a4a10', bar: '#c4891a' },
    { bg: '#e7f0e4', fg: '#1d4a32', bar: '#2f6f4e' },
    { bg: '#e7eef7', fg: '#1e3a5f', bar: '#3d5a80' },
    { bg: '#f3e8f6', fg: '#4a2c5a', bar: '#7c4d8a' },
    { bg: '#fdeee4', fg: '#6b3010', bar: '#c45c26' },
    { bg: '#e6f3f1', fg: '#1a4540', bar: '#2a7a70' },
    { bg: '#efe8dc', fg: '#5a4630', bar: '#8a6a3a' },
    { bg: '#f6e4ec', fg: '#6a2040', bar: '#b44a6a' },
  ];
  let schedShifts = [];
  let schedFilter = 'all';
  let schedMode = 'day';
  let schedFocusDay = '';
  let schedTickTimer = 0;
  let schedScrollObs = null;

  function loadScheduleLoop() {
    const section = document.getElementById('schedule');
    const board = document.getElementById('sched-board');
    const sheet = window.SCHEDULE_SHEET || window.PEOPLE_SHEET;
    if (!section || !board || !sheet || !sheet.id) return;

    const cb = '__sched' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    let settled = false;
    const cleanup = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    };
    const fail = () => {
      cleanup();
      section.hidden = true;
    };
    const timer = setTimeout(fail, 12000);

    window[cb] = payload => {
      cleanup();
      schedShifts = parseSchedulePayload(payload);
      if (!schedShifts.length) { section.hidden = true; return; }
      const names = Array.from(new Set(schedShifts.flatMap(s => s.people))).sort((a, b) => a.localeCompare(b));
      section.dataset.keywords = [
        section.dataset.keywords || '',
        'schedule', 'calendar', 'shift', 'hours', 'desk',
        names.join(' '),
        SCHED_DAYS.map(d => d.full).join(' '),
      ].join(' ').toLowerCase();
      renderSchedule(schedShifts, names);
      board.classList.remove('is-loading');
      bindSchedChrome();
      if (schedTickTimer) clearInterval(schedTickTimer);
      schedTick();
      schedTickTimer = setInterval(schedTick, 30000);
    };
    script.onerror = fail;

    const params = new URLSearchParams({
      gid: String(sheet.gid || '0'),
      tq: 'select A, B, C',
      tqx: 'out:json;responseHandler:' + cb,
    });
    script.src = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(sheet.id) + '/gviz/tq?' + params.toString();
    document.body.appendChild(script);
  }

  function parseSchedulePayload(payload) {
    const table = payload && payload.table;
    const cols = (table && table.cols) || [];
    const rows = (table && table.rows) || [];
    const out = [];
    let day = schedDayId(cols[0] && cols[0].label);
    const headerTime = parseShiftTime(schedTimeLabel(cols[1] && cols[1].label));
    const headerWho = splitShiftNames(schedWhoLabel(cols[2] && cols[2].label));
    if (day && headerTime && headerWho.length) {
      out.push({ day: day, time: headerTime, people: headerWho });
    }
    rows.forEach(row => {
      const d = peopleCell(row, 0);
      const t = peopleCell(row, 1);
      const w = peopleCell(row, 2);
      if (d) day = schedDayId(d) || day;
      const time = parseShiftTime(t);
      const who = splitShiftNames(w);
      if (!day || !time || !who.length) return;
      out.push({ day: day, time: time, people: who });
    });
    return mergeScheduleShifts(out);
  }

  function mergeScheduleShifts(shifts) {
    const out = [];
    SCHED_DAYS.forEach(day => {
      shifts.filter(s => s.day === day.id)
        .sort((a, b) => a.time.start - b.time.start)
        .forEach(s => {
          const last = out.length ? out[out.length - 1] : null;
          const same = last && last.day === s.day
            && last.people.join('|').toLowerCase() === s.people.join('|').toLowerCase()
            && last.time.end === s.time.start;
          if (same) {
            last.time.end = s.time.end;
            last.time.label = schedHourNum(last.time.start) + '–' + schedFmtHour(last.time.end);
          } else {
            out.push({ day: s.day, time: { start: s.time.start, end: s.time.end, label: s.time.label, raw: s.time.raw }, people: s.people.slice() });
          }
        });
    });
    return out;
  }

  function schedDayId(s) {
    const n = String(s || '').trim().toLowerCase();
    const hit = SCHED_DAYS.find(d => n === d.id || n === d.short.toLowerCase() || n.indexOf(d.id) === 0);
    return hit ? hit.id : '';
  }

  function schedTimeLabel(s) {
    return String(s || '').replace(/^\s*shift\s+/i, '').trim();
  }

  function schedWhoLabel(s) {
    return String(s || '').replace(/^\s*employee\s+/i, '').trim();
  }

  function splitShiftNames(s) {
    return String(s || '')
      .split(/\s*\+\s*|\s+and\s+|\s*&\s*|\s*,\s*/i)
      .map(n => n.trim())
      .filter(Boolean);
  }

  function parseShiftTime(raw) {
    const label = schedTimeLabel(raw);
    const m = label.match(/(\d{1,2})\s*[-–]\s*(\d{1,2})\s*(am|pm)/i);
    if (!m) return null;
    const a = parseInt(m[1], 10);
    const b = parseInt(m[2], 10);
    const mer = m[3].toLowerCase();
    const toMin = (h, pm) => {
      if (h === 12) return pm ? 12 * 60 : 0;
      return (pm ? h + 12 : h) * 60;
    };
    const endPm = mer === 'pm';
    const startPm = mer === 'pm' ? !(a === 10 || a === 11) : false;
    let start = toMin(a, startPm);
    let end = toMin(b, endPm);
    if (end <= start) end += 12 * 60;
    const pretty = schedHourNum(start) + '–' + schedFmtHour(end);
    return { start: start, end: end, label: pretty, raw: label };
  }

  function schedHourNum(min) {
    const h = Math.floor(((min % (24 * 60)) + (24 * 60)) % (24 * 60) / 60);
    if (h === 0 || h === 12) return 12;
    return h > 12 ? h - 12 : h;
  }

  function schedFmtHour(min) {
    let h = Math.floor(((min % (24 * 60)) + (24 * 60)) % (24 * 60) / 60);
    if (h === 0) return '12am';
    if (h === 12) return '12pm';
    if (h > 12) return (h - 12) + 'pm';
    return h + 'am';
  }

  function schedNowMin() {
    const n = new Date();
    return n.getHours() * 60 + n.getMinutes();
  }

  function schedTodayId() {
    return ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'][new Date().getDay()];
  }

  function schedPeopleNames() {
    return Array.from(new Set(schedShifts.flatMap(s => s.people))).sort((a, b) => a.localeCompare(b));
  }

  function schedDayMeta(id) {
    return SCHED_DAYS.find(d => d.id === id) || SCHED_DAYS[0];
  }

  function schedTone(name) {
    let h = 0;
    const s = String(name || '').toLowerCase();
    for (let i = 0; i < s.length; i++) h = (h * 33 + s.charCodeAt(i)) >>> 0;
    return SCHED_PALETTE[h % SCHED_PALETTE.length];
  }

  function schedRange(shifts) {
    let start = 10 * 60, end = 22 * 60;
    shifts.forEach(s => {
      if (s.time.start < start) start = Math.floor(s.time.start / 60) * 60;
      if (s.time.end > end) end = Math.ceil(s.time.end / 60) * 60;
    });
    return { start: start, end: end, hours: (end - start) / 60 };
  }

  function renderSchedule(shifts, names) {
    const board = document.getElementById('sched-board');
    const filters = document.getElementById('sched-filters');
    const lede = document.getElementById('sched-lede');
    if (!board) return;
    if (!schedFocusDay) schedFocusDay = schedTodayId();
    const range = schedRange(shifts);
    const today = schedTodayId();
    const ticks = [];
    for (let m = range.start; m < range.end; m += 60) ticks.push(m);

    if (lede) {
      if (schedMode === 'week') {
        lede.textContent = 'This week’s desk coverage. Tap a day to open it.';
      } else {
        const day = schedDayMeta(schedFocusDay);
        lede.textContent = (day.id === today ? 'Today’s' : (day.full + '’s'))
          + ' desk coverage. Switch to Week to see the rest of the schedule.';
      }
    }

    if (filters) {
      const chips = [{ id: 'all', label: 'Everyone' }].concat(names.map(n => ({ id: n, label: n })));
      filters.innerHTML = chips.map(c => {
        const on = (c.id === 'all' && schedFilter === 'all') || (c.id !== 'all' && schedFilter === c.id.toLowerCase());
        return '<button type="button" class="sched-chip' + (on ? ' is-active' : '') + '" data-who="' + escapePeople(c.id) + '">'
          + escapePeople(c.label) + '</button>';
      }).join('');
      filters.hidden = false;
      filters.onclick = ev => {
        const btn = ev.target.closest('.sched-chip');
        if (!btn) return;
        schedFilter = String(btn.getAttribute('data-who') || 'all').toLowerCase();
        renderSchedule(schedShifts, names);
      };
    }

    board.classList.toggle('is-day', schedMode === 'day');
    board.classList.toggle('is-week', schedMode === 'week');
    if (schedMode === 'day') renderSchedDay(board, shifts, range, today, ticks, names);
    else renderSchedWeek(board, shifts, range, today, ticks, names);
    updateSchedViewBtns();
    schedTick();
  }

  function schedGutterHtml(ticks, range) {
    return '<div class="sched-gutter">'
      + '<div class="sched-gutter-head" aria-hidden="true"></div>'
      + '<div class="sched-gutter-body">'
      + ticks.map(m => '<span class="sched-tick" style="top:' + (((m - range.start) / (range.end - range.start)) * 100) + '%">' + escapePeople(schedFmtHour(m)) + '</span>').join('')
      + '</div></div>';
  }

  function schedShiftArticle(s, day, range, today) {
    const visible = schedFilter === 'all' || s.people.some(p => p.toLowerCase() === schedFilter);
    const top = ((s.time.start - range.start) / (range.end - range.start)) * 100;
    const height = ((s.time.end - s.time.start) / (range.end - range.start)) * 100;
    const tone = schedTone(s.people[0]);
    const nowOn = day.id === today && schedNowMin() >= s.time.start && schedNowMin() < s.time.end;
    const kw = (s.people.join(' ') + ' ' + day.full + ' ' + s.time.raw + ' shift').toLowerCase();
    const namesHtml = s.people.map(p => '<b>' + escapePeople(p) + '</b>').join('<span class="sched-and">+</span>');
    return '<article class="sched-shift' + (nowOn ? ' is-now' : '') + (visible ? '' : ' is-dim') + '" data-keywords="' + escapePeople(kw) + '" '
      + 'style="top:' + top + '%;height:' + height + '%;--shift-bg:' + tone.bg + ';--shift-fg:' + tone.fg + ';--shift-bar:' + tone.bar + '" '
      + 'aria-label="' + escapePeople(s.people.join(' and ') + ', ' + day.full + ' ' + s.time.label) + '">'
      + '<span class="sched-shift-time">' + escapePeople(s.time.label) + '</span>'
      + '<span class="sched-shift-who">' + namesHtml + '</span>'
      + '</article>';
  }

  function schedColHtml(day, shifts, range, today, ticks, opts) {
    const isToday = day.id === today;
    const isWe = day.id === 'saturday' || day.id === 'sunday';
    const dayShifts = shifts.filter(s => s.day === day.id);
    const blocks = dayShifts.map(s => schedShiftArticle(s, day, range, today)).join('');
    const empty = opts && opts.empty && !dayShifts.length
      ? '<p class="sched-empty">No one is scheduled</p>'
      : '';
    return '<div class="sched-day' + (isToday ? ' is-today' : '') + (isWe ? ' is-weekend' : '') + '" data-day="' + day.id + '">'
      + '<div class="sched-day-head">'
      + '<span class="sched-day-name">' + day.short + '</span>'
      + '<span class="sched-day-full">' + day.full + '</span>'
      + (isToday ? '<span class="sched-today-pill">Today</span>' : '')
      + '</div>'
      + '<div class="sched-day-body">'
      + ticks.map(m => '<span class="sched-hourline" style="top:' + (((m - range.start) / (range.end - range.start)) * 100) + '%"></span>').join('')
      + blocks
      + empty
      + '</div></div>';
  }

  function renderSchedDay(board, shifts, range, today, ticks, names) {
    const focus = schedDayMeta(schedFocusDay);
    const jumps = SCHED_DAYS.map(day => {
      const on = day.id === focus.id;
      const isToday = day.id === today;
      return '<button type="button" class="sched-jump' + (on ? ' is-on' : '') + (isToday ? ' is-today' : '') + '" data-day="' + day.id + '" '
        + 'aria-pressed="' + (on ? 'true' : 'false') + '" aria-label="' + escapePeople(day.full) + (isToday ? ', today' : '') + '">'
        + escapePeople(day.short)
        + (isToday ? '<span class="sched-jump-dot" aria-hidden="true"></span>' : '')
        + '</button>';
    }).join('');
    board.innerHTML = '<div class="sched-daybar">'
      + '<div class="sched-jumps" role="group" aria-label="Day of week">' + jumps + '</div>'
      + '<button type="button" class="sched-back-today"' + (focus.id === today ? ' hidden' : '') + '>Today</button>'
      + '</div>'
      + '<div class="sched-dayframe">'
      + '<div class="sched-week is-day" style="--sched-hours:' + range.hours + '">'
      + schedGutterHtml(ticks, range)
      + schedColHtml(focus, shifts, range, today, ticks, { empty: true })
      + '</div></div>';
    const bar = board.querySelector('.sched-daybar');
    if (bar) {
      bar.addEventListener('click', ev => {
        const jump = ev.target.closest('.sched-jump');
        if (jump) {
          schedFocusDay = jump.getAttribute('data-day') || today;
          renderSchedule(shifts, names);
          return;
        }
        if (ev.target.closest('.sched-back-today')) {
          schedFocusDay = today;
          renderSchedule(shifts, names);
        }
      });
    }
    board.dataset.mode = 'day';
  }

  function renderSchedWeek(board, shifts, range, today, ticks, names) {
    const keepX = board.dataset.mode === 'week' ? (board.querySelector('.sched-scroll') || {}).scrollLeft : null;
    const cols = SCHED_DAYS.map(day => schedColHtml(day, shifts, range, today, ticks, null)).join('');
    board.innerHTML = '<div class="sched-hscroll">'
      + '<div class="sched-scroll" tabindex="0" role="region" aria-label="This week\'s desk coverage. Scroll sideways for other days. Tap a day name to open it.">'
      + '<div class="sched-week" style="--sched-hours:' + range.hours + '">'
      + schedGutterHtml(ticks, range) + cols + '</div></div>'
      + '<button type="button" class="sched-nudge sched-nudge--prev" aria-label="Earlier days">'
      + '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"></path></svg>'
      + '</button>'
      + '<button type="button" class="sched-nudge sched-nudge--next" aria-label="Later days">'
      + '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"></path></svg>'
      + '</button></div>';
    const scroller = board.querySelector('.sched-scroll');
    if (keepX != null && scroller) scroller.scrollLeft = keepX;
    else if (scroller) {
      pinSchedTodayVisible(scroller);
      requestAnimationFrame(function () { pinSchedTodayVisible(scroller); });
    }
    wireSchedScroll(scroller);
    if (scroller) {
      scroller.addEventListener('click', ev => {
        const head = ev.target.closest('.sched-day-head');
        if (!head) return;
        const col = head.closest('.sched-day');
        const id = col && col.getAttribute('data-day');
        if (!id) return;
        schedFocusDay = id;
        schedMode = 'day';
        renderSchedule(shifts, names);
      });
    }
    board.dataset.mode = 'week';
  }

  function pinSchedTodayVisible(scroller) {
    if (!scroller || scroller.scrollWidth <= scroller.clientWidth + 4) return;
    const todayCol = scroller.querySelector('.sched-day.is-today');
    const gutter = scroller.querySelector('.sched-gutter');
    if (!todayCol) return;
    const s = scroller.getBoundingClientRect();
    const t = todayCol.getBoundingClientRect();
    const g = gutter ? gutter.getBoundingClientRect().width : 0;
    if (t.left >= s.left + g - 2 && t.right <= s.right + 2) return;
    scroller.scrollLeft += (t.left - s.left) - g;
  }

  function bindSchedChrome() {
    const views = document.getElementById('sched-views');
    if (!views || views.dataset.bound) return;
    views.dataset.bound = '1';
    views.addEventListener('click', ev => {
      const btn = ev.target.closest('[data-view]');
      if (!btn) return;
      const next = btn.getAttribute('data-view');
      if (next !== 'day' && next !== 'week') return;
      if (next === schedMode) return;
      schedMode = next;
      updateSchedViewBtns();
      if (!schedShifts.length) return;
      if (schedMode === 'day' && !schedFocusDay) schedFocusDay = schedTodayId();
      renderSchedule(schedShifts, schedPeopleNames());
    });
  }

  function updateSchedViewBtns() {
    document.querySelectorAll('#sched-views [data-view]').forEach(btn => {
      const on = btn.getAttribute('data-view') === schedMode;
      btn.classList.toggle('is-active', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
  }

  function updateSchedScrollCue(scroller) {
    const wrap = scroller && scroller.parentElement;
    if (!scroller || !wrap || !wrap.classList.contains('sched-hscroll')) return;
    const max = Math.max(0, scroller.scrollWidth - scroller.clientWidth);
    const x = scroller.scrollLeft;
    wrap.classList.toggle('is-scrollable', max > 4);
    wrap.classList.toggle('has-left', x > 4);
    wrap.classList.toggle('has-right', x < max - 4);
  }

  function wireSchedScroll(scroller) {
    if (!scroller) return;
    const wrap = scroller.parentElement;
    const gutter = scroller.querySelector('.sched-gutter');
    if (gutter) scroller.style.scrollPaddingLeft = gutter.getBoundingClientRect().width + 'px';
    const nudge = (dir) => {
      const col = scroller.querySelector('.sched-day');
      const w = col ? col.getBoundingClientRect().width : 88;
      const smooth = !window.matchMedia || !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      scroller.scrollBy({ left: dir * w, behavior: smooth ? 'smooth' : 'auto' });
    };
    scroller.addEventListener('scroll', () => updateSchedScrollCue(scroller), { passive: true });
    if (wrap) {
      wrap.querySelectorAll('.sched-nudge--prev').forEach(btn => btn.addEventListener('click', () => nudge(-1)));
      wrap.querySelectorAll('.sched-nudge--next').forEach(btn => btn.addEventListener('click', () => nudge(1)));
    }
    if (schedScrollObs) schedScrollObs.disconnect();
    if (window.ResizeObserver) {
      schedScrollObs = new ResizeObserver(() => updateSchedScrollCue(scroller));
      schedScrollObs.observe(scroller);
      const week = scroller.querySelector('.sched-week');
      if (week) schedScrollObs.observe(week);
    }
    updateSchedScrollCue(scroller);
    requestAnimationFrame(() => updateSchedScrollCue(scroller));
  }

  function schedTick() {
    const board = document.getElementById('sched-board');
    if (!board || !schedShifts.length) return;
    const range = schedRange(schedShifts);
    const today = schedTodayId();
    const now = schedNowMin();
    const dayCol = board.querySelector('.sched-day.is-today .sched-day-body');
    if (dayCol) {
      let line = dayCol.querySelector('.sched-nowline');
      if (now >= range.start && now <= range.end) {
        if (!line) {
          line = document.createElement('div');
          line.className = 'sched-nowline';
          line.innerHTML = '<span>Now</span>';
          dayCol.appendChild(line);
        }
        line.style.top = ((now - range.start) / (range.end - range.start) * 100) + '%';
        line.hidden = false;
      } else if (line) {
        line.hidden = true;
      }
    }
    board.querySelectorAll('.sched-shift').forEach(el => {
      const col = el.closest('.sched-day');
      if (!col || !col.classList.contains('is-today')) {
        el.classList.remove('is-now');
        return;
      }
      const top = parseFloat(el.style.top) || 0;
      const h = parseFloat(el.style.height) || 0;
      const pct = (now - range.start) / (range.end - range.start) * 100;
      el.classList.toggle('is-now', pct >= top && pct < top + h);
    });
  }

  // Registry links (Equipment / Inventory tables, item detail, labels) point at the
  // web app. On the unlisted admin page, data-admin adds &admin=1 so the page opens
  // in edit mode (add / edit / delete, still passcode-gated server-side).
  document.querySelectorAll('a[data-registry]').forEach(a => {
    if (!SPACE_STATUS_URL) return;
    const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
    let href = SPACE_STATUS_URL + sep + 'registry=' + encodeURIComponent(a.dataset.registry);
    if (a.dataset.which) href += '&which=' + encodeURIComponent(a.dataset.which);
    if (a.hasAttribute('data-admin')) href += '&admin=1';
    a.href = href;
  });

  // Module links (Projects, Proposals, …) point at the same web app via ?module=.
  // data-admin links (only on the unlisted admin page) add &admin=1, which unlocks
  // the admin controls without a passcode. Regular staff links never carry it.
  // back= lets the full page show a return link to this hub.
  document.querySelectorAll('a[data-module]').forEach(a => {
    if (!SPACE_STATUS_URL) return;
    const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
    a.href = SPACE_STATUS_URL + sep + 'module=' + encodeURIComponent(a.dataset.module)
      + (a.hasAttribute('data-admin') ? '&admin=1' : '')
      + '&back=' + encodeURIComponent(location.origin + location.pathname);
  });

  // "New" flags expire on their own: set data-new-until="YYYY-MM-DD" on a tool and its
  // badge (and gold sheen) show until that date passes, so nobody has to remember to
  // take it down.
  document.querySelectorAll('[data-new-until]').forEach(el => {
    const until = new Date(el.dataset.newUntil + 'T23:59:59');
    const badge = el.querySelector('.action-badge--new');
    if (!badge || Number.isNaN(until.getTime()) || Date.now() > until.getTime()) return;
    badge.hidden = false;
    el.classList.add('is-new');
    if (!el.querySelector('.action-shine')) {
      const shine = document.createElement('span');
      shine.className = 'action-shine';
      shine.setAttribute('aria-hidden', 'true');
      el.prepend(shine);
    }
  });

  const search = document.getElementById('search');
  const categories = document.querySelectorAll('.category');
  const zones = document.querySelectorAll('.hub-zone');
  const noResults = document.getElementById('no-results');

  const adminNav = document.getElementById('admin-nav');

  // Open space issues dashboard: slide-over panel under Space (not inline in ELL).
  const issuesPanel = document.getElementById('issues-panel');
  const issuesOpenBtn = document.getElementById('issues-open');
  const issuesCloseBtn = document.getElementById('issues-panel-close');
  const issuesBackdrop = document.getElementById('issues-panel-backdrop');
  const issuesPanelExt = document.getElementById('issues-panel-ext');
  const issuesFrame = document.getElementById('issues-frame');
  const issuesEmbedBody = document.getElementById('issues-embed-body');
  const issuesFallback = document.getElementById('issues-fallback');
  const issuesFallbackLink = document.getElementById('issues-fallback-link');
  const issuesLoadingText = document.getElementById('issues-loading-text');
  const issuesLoadingHint = document.getElementById('issues-loading-hint');
  let issuesFrameStarted = false;
  let issuesLoadTimer = null;
  let issuesLoadStepTimer = null;

  // Which view the panel embeds: admin = all teams, staff = one team (set via
  // data-embed-query on #issues-panel). Defaults to the all-teams dashboard.
  const embedQuery = (issuesPanel && issuesPanel.dataset.embedQuery) || 'view=all';
  const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
  const issuesUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + embedQuery + '&embed=1' : '';
  const issuesFullUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + embedQuery : '';

  if (issuesFallbackLink && issuesFullUrl) issuesFallbackLink.href = issuesFullUrl;
  if (issuesPanelExt && issuesFullUrl) issuesPanelExt.href = issuesFullUrl;

  // Combined Tasks/Projects segmented toggle. Present only on the student panel
  // (index.html); the admin dashboard panel has no toggle, so these are no-ops there.
  // The Projects iframe is lazy-loaded the first time its tab is opened.
  const segTasks = document.getElementById('seg-tasks');
  const segProjects = document.getElementById('seg-projects');
  const projectsFrame = document.getElementById('projects-frame');
  const projectsEmbedBody = document.getElementById('projects-embed-body');
  // Admin panels (data-embed-query carries admin=1) open the Projects tab in admin
  // mode so create / approve / delete are available; staff panels stay student-mode.
  const adminSuffix = /(?:^|&)admin=1(?:&|$)/.test(embedQuery || '') ? '&admin=1' : '';
  const projectsUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'module=projects&embed=1' + adminSuffix : '';
  const projectsFullUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'module=projects' + adminSuffix : '';
  const projectsNew = document.getElementById('seg-projects-new');
  let projectsFrameStarted = false;
  let projectsLoadBound = false;

  function hideProjectsNew() {
    if (projectsNew) projectsNew.hidden = true;
  }

  function loadProjectsFrame(refresh) {
    if (!projectsFrame || !projectsUrl) return;
    if (projectsFrameStarted && !refresh) return;
    if (!projectsLoadBound) {
      projectsLoadBound = true;
      projectsFrame.addEventListener('load', () => {
        if (projectsEmbedBody) projectsEmbedBody.classList.remove('is-loading');
      });
    }
    projectsFrameStarted = true;
    if (projectsEmbedBody) projectsEmbedBody.classList.add('is-loading');
    projectsFrame.src = projectsUrl;
  }

  function showPanelView(view, opts) {
    const tasks = view !== 'projects';
    if (issuesEmbedBody) issuesEmbedBody.hidden = !tasks;
    if (projectsEmbedBody) projectsEmbedBody.hidden = tasks;
    if (segTasks) { segTasks.classList.toggle('is-active', tasks); segTasks.setAttribute('aria-selected', String(tasks)); }
    if (segProjects) { segProjects.classList.toggle('is-active', !tasks); segProjects.setAttribute('aria-selected', String(!tasks)); }
    if (issuesPanelExt) issuesPanelExt.href = tasks ? issuesFullUrl : projectsFullUrl;
    if (!tasks) {
      hideProjectsNew();
      loadProjectsFrame(opts && opts.refresh);
    }
  }

  if (segTasks) segTasks.addEventListener('click', () => showPanelView('tasks'));
  if (segProjects) segProjects.addEventListener('click', () => showPanelView('projects'));

  function setIssuesLoadingMessage(primary, hint) {
    if (issuesLoadingText && primary) issuesLoadingText.textContent = primary;
    if (issuesLoadingHint && hint) issuesLoadingHint.textContent = hint;
  }

  function startIssuesLoadingSteps() {
    setIssuesLoadingMessage('Connecting to Space Status…', 'Fetching open issues and photos from Google Sheets');
    stopIssuesLoadingSteps();
    const step2 = window.setTimeout(() => {
      setIssuesLoadingMessage('Loading issues…', 'This can take a few seconds on first open');
    }, 2800);
    const step3 = window.setTimeout(() => {
      setIssuesLoadingMessage('Almost ready…', 'Building filters and issue cards');
    }, 7000);
    issuesLoadStepTimer = [step2, step3];
  }

  function stopIssuesLoadingSteps() {
    if (!issuesLoadStepTimer) return;
    (Array.isArray(issuesLoadStepTimer) ? issuesLoadStepTimer : [issuesLoadStepTimer])
      .forEach(id => clearTimeout(id));
    issuesLoadStepTimer = null;
  }

  function showIssuesError() {
    stopIssuesLoadingSteps();
    if (issuesEmbedBody) issuesEmbedBody.classList.remove('is-loading');
    if (issuesFallback) issuesFallback.hidden = false;
  }

  function finishIssuesLoad() {
    if (issuesLoadTimer) {
      clearTimeout(issuesLoadTimer);
      issuesLoadTimer = null;
    }
    stopIssuesLoadingSteps();
    if (issuesEmbedBody) issuesEmbedBody.classList.remove('is-loading');
  }

  function loadIssuesFrame() {
    if (!issuesFrame || issuesFrameStarted) return;
    if (!issuesUrl) {
      showIssuesError();
      return;
    }

    issuesFrameStarted = true;
    if (issuesEmbedBody) issuesEmbedBody.classList.add('is-loading');
    if (issuesFallback) issuesFallback.hidden = true;
    startIssuesLoadingSteps();

    issuesFrame.addEventListener('load', finishIssuesLoad, { once: true });
    issuesFrame.addEventListener('error', showIssuesError, { once: true });

    issuesLoadTimer = window.setTimeout(() => {
      issuesLoadTimer = null;
      if (issuesEmbedBody && issuesEmbedBody.classList.contains('is-loading')) {
        showIssuesError();
      }
    }, 45000);

    issuesFrame.src = issuesUrl;
  }

  function openIssuesPanel(view) {
    if (view && typeof view !== 'string') view = 'tasks';
    if (!issuesPanel) return;
    issuesPanel.hidden = false;
    requestAnimationFrame(() => issuesPanel.classList.add('is-open'));
    document.body.classList.add('issues-panel-open');
    showPanelView(view || 'tasks');
    loadIssuesFrame();
    if (issuesCloseBtn) issuesCloseBtn.focus();
  }

  function closeIssuesPanel() {
    if (!issuesPanel) return;
    issuesPanel.classList.remove('is-open');
    document.body.classList.remove('issues-panel-open');
    window.setTimeout(() => {
      if (!issuesPanel.classList.contains('is-open')) issuesPanel.hidden = true;
    }, 320);
    if (issuesOpenBtn) issuesOpenBtn.focus();
  }

  if (issuesOpenBtn) {
    issuesOpenBtn.addEventListener('click', openIssuesPanel);
    prefetchOnIntent(issuesOpenBtn, loadIssuesFrame);   // absorb the Apps Script cold start
  }
  if (issuesCloseBtn) issuesCloseBtn.addEventListener('click', closeIssuesPanel);
  if (issuesBackdrop) issuesBackdrop.addEventListener('click', closeIssuesPanel);

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && issuesPanel && issuesPanel.classList.contains('is-open')) {
      e.preventDefault();
      closeIssuesPanel();
    }
  });

  // Open the issues dashboard straight away via a shared admin.html#issues link.
  const bootHash = location.hash.slice(1);
  if (bootHash === 'issues') openIssuesPanel();
  if (bootHash === 'projects') openIssuesPanel('projects');

  /**
   * Warm an embedded module before it is asked for: begin loading once the pointer has
   * rested on its button, or as soon as it takes keyboard focus. Only fires once, and a
   * pointer that leaves before the dwell elapses cancels it.
   * @param {Element} el
   * @param {Function} start
   */
  function prefetchOnIntent(el, start) {
    let timer = null;
    let done = false;
    const go = () => { if (done) return; done = true; start(); };
    const arm = () => { if (!done && timer === null) timer = window.setTimeout(go, 120); };
    const disarm = () => { if (timer !== null) { clearTimeout(timer); timer = null; } };
    el.addEventListener('mouseenter', arm);
    el.addEventListener('mouseleave', disarm);
    el.addEventListener('focus', go);
    el.addEventListener('touchstart', go, { passive: true });
  }

  /**
   * Generic slide-over panel: same shell and behaviour as the tasks panel above, but
   * driven by data-attributes instead of fixed ids, so any button can open any embedded
   * module. The iframe is lazy-loaded on first open and its loading copy steps forward
   * while Apps Script wakes up.
   * @param {{panelId: string, openBtnId: string, url: string, fullUrl: string, steps: Array<[string, string, number]>}} opts
   */
  function createSlideOver(opts) {
    const panel = document.getElementById(opts.panelId);
    const openBtn = document.getElementById(opts.openBtnId);
    if (!panel || !openBtn) return null;

    const q = sel => panel.querySelector(sel);
    const frame = q('[data-panel-frame]');
    const body = q('[data-panel-body]');
    const closeBtn = q('[data-panel-close]');
    const backdrop = q('[data-panel-backdrop]');
    const fallback = q('[data-panel-fallback]');
    const fallbackLink = q('[data-panel-fallback-link]');
    const extLink = q('[data-panel-ext]');
    const loadText = q('.issues-loading-text');
    const loadHint = q('.issues-loading-hint');

    if (fallbackLink && opts.fullUrl) fallbackLink.href = opts.fullUrl;
    if (extLink && opts.fullUrl) extLink.href = opts.fullUrl;

    let started = false;
    let timers = [];

    const clearTimers = () => { timers.forEach(clearTimeout); timers = []; };

    const finish = () => { clearTimers(); if (body) body.classList.remove('is-loading'); };

    const fail = () => {
      clearTimers();
      if (body) body.classList.remove('is-loading');
      if (fallback) fallback.hidden = false;
    };

    function load() {
      if (started || !frame) return;
      if (!opts.url) { fail(); return; }
      started = true;
      if (body) body.classList.add('is-loading');
      if (fallback) fallback.hidden = true;

      (opts.steps || []).forEach(([text, hint, at]) => {
        timers.push(window.setTimeout(() => {
          if (loadText) loadText.textContent = text;
          if (loadHint) loadHint.textContent = hint;
        }, at));
      });
      timers.push(window.setTimeout(() => {
        if (body && body.classList.contains('is-loading')) fail();
      }, 45000));

      frame.addEventListener('load', finish, { once: true });
      frame.addEventListener('error', fail, { once: true });
      frame.src = opts.url;
    }

    function open() {
      panel.hidden = false;
      requestAnimationFrame(() => panel.classList.add('is-open'));
      document.body.classList.add('issues-panel-open');
      load();
      if (closeBtn) closeBtn.focus();
    }

    function close() {
      panel.classList.remove('is-open');
      document.body.classList.remove('issues-panel-open');
      window.setTimeout(() => {
        if (!panel.classList.contains('is-open')) panel.hidden = true;
      }, 320);
      openBtn.focus();
    }

    openBtn.addEventListener('click', open);
    // Start fetching on intent rather than on click. An embedded Apps Script page costs
    // a cold start of a couple of seconds, and the hover before the click is enough to
    // absorb most of it. Guarded by a short dwell so sweeping the mouse across the list
    // does not fetch every module.
    prefetchOnIntent(openBtn, load);
    if (closeBtn) closeBtn.addEventListener('click', close);
    if (backdrop) backdrop.addEventListener('click', close);
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && panel.classList.contains('is-open')) { e.preventDefault(); close(); }
    });

    return { open, close };
  }

  // Proposals board. On the admin page this carries ?admin=1 so Edit / Approve /
  // Decline / Delete show. The full-page link carries ?back= so the Apps Script page
  // can offer a way back to this hub (it has no other way to know where it was opened from).
  const proposalsAdmin = document.getElementById('admin') ? '&admin=1' : '';
  const proposalsBase = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'module=proposals' : '';
  const proposalsSlide = createSlideOver({
    panelId: 'proposals-panel',
    openBtnId: 'proposals-open',
    url: proposalsBase ? proposalsBase + '&embed=1' + proposalsAdmin : '',
    fullUrl: proposalsBase ? proposalsBase + proposalsAdmin + '&back=' + encodeURIComponent(location.origin + location.pathname) : '',
    steps: [
      ['Loading the proposal board…', 'This can take a few seconds on first open', 2800],
      ['Almost ready…', 'Counting reviews and impact scores', 7000],
    ],
  });

  window.addEventListener('message', ev => {
    const d = ev.data;
    if (!d || d.source !== 'ops-hub') return;
    if (d.action === 'projects-news') {
      if (projectsNew) projectsNew.hidden = !(d.count > 0);
    }
    if (d.action === 'strip-changed') {
      if (issuesFrame && issuesFrame.contentWindow) {
        try { issuesFrame.contentWindow.postMessage({ source: 'ops-hub', action: 'refresh-strip' }, '*'); } catch (e) { /* ignore */ }
      }
    }
    if (d.action === 'open-projects') {
      const open = issuesPanel && issuesPanel.classList.contains('is-open');
      if (!open) openIssuesPanel('projects');
      else showPanelView('projects', { refresh: true });
    }
    if (d.action === 'open-proposals') {
      closeIssuesPanel();
      if (proposalsSlide) proposalsSlide.open();
    }
  });

  // Inventory catalog. Staff and admin both carry data-admin so stock counts can be
  // updated from the hub (same as the old new-tab link). embed=1 drops the Apps Script
  // chrome so the panel header is the only frame.
  const invOpen = document.getElementById('inventory-open');
  const invAdmin = (invOpen && invOpen.hasAttribute('data-admin')) || document.getElementById('admin') ? '&admin=1' : '';
  const invBase = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'registry=inventory' : '';
  createSlideOver({
    panelId: 'inventory-panel',
    openBtnId: 'inventory-open',
    url: invBase ? invBase + '&embed=1' + invAdmin : '',
    fullUrl: invBase ? invBase + invAdmin : '',
    steps: [
      ['Loading the shop catalog…', 'This can take a few seconds on first open', 2800],
      ['Almost ready…', 'Fetching items and stock counts', 7000],
    ],
  });

  // Scroll-spy: highlight the sidebar link for the section currently in view.
  if (adminNav && 'IntersectionObserver' in window) {
    const navLinks = [...adminNav.querySelectorAll('a[href^="#"]')];
    const linkById = {};
    navLinks.forEach(a => { linkById[a.getAttribute('href').slice(1)] = a; });
    const setActive = id => {
      navLinks.forEach(a => a.classList.remove('active'));
      if (linkById[id]) linkById[id].classList.add('active');
    };
    const spy = new IntersectionObserver(entries => {
      const inView = entries.filter(e => e.isIntersecting)
        .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
      if (inView.length) setActive(inView[0].target.id);
    }, { rootMargin: '-96px 0px -65% 0px', threshold: 0 });
    document.querySelectorAll('#admin .admin-block').forEach(b => spy.observe(b));
    navLinks.forEach(a => a.addEventListener('click', () => setActive(a.getAttribute('href').slice(1))));
  }

  if (!search) return;

  const clearBtn = document.getElementById('search-clear');
  const kbdHint = document.getElementById('search-kbd');
  const meta = document.getElementById('search-meta');
  const countEl = document.getElementById('search-count');

  if (kbdHint) {
    const isMac = /mac|iphone|ipad/i.test(navigator.platform || navigator.userAgent || '');
    kbdHint.textContent = isMac ? '⌘K' : 'Ctrl K';
  }

  // Cache each label's original text so highlights can be re-rendered cleanly.
  document.querySelectorAll('.action, button.action').forEach(action => {
    const title = action.querySelector('.action-title');
    const sub = action.querySelector('.action-sub');
    if (title && title.dataset.text == null) title.dataset.text = title.textContent;
    if (sub && sub.dataset.text == null) sub.dataset.text = sub.textContent;
  });

  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /**
   * fzf-style subsequence scorer. Returns -1 when `needle` is not a
   * subsequence of `hay`, otherwise a score rewarding contiguous runs and
   * matches that land on a word boundary.
   * @param {string} hay
   * @param {string} needle
   * @returns {number}
   */
  function fuzzyScore(hay, needle) {
    if (!needle) return 0;
    let score = 0, at = 0, streak = 0;
    for (let i = 0; i < needle.length; i++) {
      const c = needle[i];
      let found = -1;
      for (let k = at; k < hay.length; k++) {
        if (hay[k] === c) { found = k; break; }
      }
      if (found === -1) return -1;
      if (found === at) { streak++; score += 4 + streak; }
      else { streak = 0; score += 1; }
      const prev = found > 0 ? hay[found - 1] : ' ';
      if (prev === ' ' || prev === '-' || prev === '/' || prev === '&') score += 10;
      at = found + 1;
    }
    return score;
  }

  /**
   * Build searchable text for an action link from cached labels + keywords.
   * @param {Element} action
   * @param {Element} section
   * @returns {string}
   */
  function actionSearchText(action, section) {
    const title = action.querySelector('.action-title');
    const sub = action.querySelector('.action-sub');
    const block = action.closest('.task-block');
    const footnote = (block || section).querySelector('.action-note');
    const footnoteText = footnote
      ? `${footnote.textContent} ${footnote.dataset.keywords || ''}`
      : '';
    const titleText = title ? title.dataset.text : action.textContent;
    const subText = sub ? sub.dataset.text : '';
    return `${titleText} ${subText} ${action.dataset.keywords || ''} ${footnoteText}`.toLowerCase();
  }

  /**
   * Wrap query-token occurrences in the title/sub labels with <mark>.
   * @param {Element} action
   * @param {string[]} tokens
   */
  function highlight(action, tokens) {
    ['.action-title', '.action-sub'].forEach(sel => {
      const el = action.querySelector(sel);
      if (!el) return;
      const orig = el.dataset.text != null ? el.dataset.text : el.textContent;
      if (!tokens.length) { el.textContent = orig; return; }
      const re = new RegExp('(' + tokens.map(escapeRe).join('|') + ')', 'ig');
      el.innerHTML = escapeHtml(orig).replace(re, '<mark>$1</mark>');
    });
  }

  // Keyboard navigation across the currently visible results.
  let results = [];
  let activeIdx = -1;

  function rebuildResults() {
    results = [...document.querySelectorAll('.action:not(.hidden), button.action:not(.hidden), a.people-card:not(.hidden)')]
      .filter(a => a.offsetParent !== null);
  }

  function clearActive() {
    if (activeIdx >= 0 && results[activeIdx]) {
      results[activeIdx].classList.remove('is-active-result');
    }
    activeIdx = -1;
  }

  function setActive(i) {
    if (activeIdx >= 0 && results[activeIdx]) {
      results[activeIdx].classList.remove('is-active-result');
    }
    if (!results.length) { activeIdx = -1; return; }
    activeIdx = Math.max(0, Math.min(results.length - 1, i));
    const el = results[activeIdx];
    el.classList.add('is-active-result');
    el.scrollIntoView({ block: 'nearest' });
  }

  function runSearch() {
    const raw = search.value;
    const q = raw.toLowerCase().trim();
    const tokens = q ? q.split(/\s+/).filter(Boolean) : [];
    const isSearching = tokens.length > 0;

    document.body.classList.toggle('is-searching', isSearching);
    if (clearBtn) clearBtn.hidden = !raw;
    if (kbdHint) kbdHint.hidden = Boolean(raw);

    let matchCount = 0;

    categories.forEach(section => {
      const sectionKeywords = (section.dataset.keywords || '').toLowerCase();
      const actions = section.matches('.action, button.action')
        ? [section]
        : [...section.querySelectorAll('.action, button.action')];
      let sectionHasMatch = false;

      actions.forEach(action => {
        const hay = actionSearchText(action, section);
        let matched = true;
        let score = 0;

        if (tokens.length) {
          for (const t of tokens) {
            const best = Math.max(fuzzyScore(hay, t), fuzzyScore(sectionKeywords, t));
            if (best < 0) { matched = false; break; }
            score += best;
          }
        }

        action.classList.toggle('hidden', tokens.length ? !matched : false);
        action.style.order = (tokens.length && matched) ? String(-score) : '';
        highlight(action, tokens);

        const block = action.closest('.task-block');
        if (block && block !== action) {
          const blockVisible = [...block.querySelectorAll('.action')].some(
            a => !a.classList.contains('hidden')
          ) || (block.querySelector('.action-note') && matched);
          block.classList.toggle('hidden', tokens.length && !blockVisible);
        }

        if (!tokens.length || matched) sectionHasMatch = true;
        if (tokens.length && matched) matchCount++;
      });

      const peopleCards = [...section.querySelectorAll('.people-card')];
      if (peopleCards.length) {
        peopleCards.forEach(card => {
          const hay = (card.dataset.keywords || card.textContent || '').toLowerCase();
          const matched = !tokens.length || tokens.every(t => hay.indexOf(t) >= 0);
          card.classList.toggle('hidden', tokens.length && !matched);
          if (!tokens.length || matched) sectionHasMatch = true;
          if (tokens.length && matched) matchCount++;
        });
        section.querySelectorAll('.people-group').forEach(g => {
          const vis = [...g.querySelectorAll('.people-card')].some(c => !c.classList.contains('hidden'));
          g.hidden = tokens.length
            ? !vis
            : peopleGroupFilter !== 'all' && g.dataset.group !== peopleGroupFilter;
        });
        const peopleFilters = section.querySelector('.people-filters');
        if (peopleFilters && peopleFilters.id === 'people-filters') {
          peopleFilters.hidden = tokens.length > 0;
        }
        const fold = section.querySelector('.people-fold');
        if (fold && tokens.length && peopleCards.some(c => !c.classList.contains('hidden'))) {
          fold.open = true;
        }
      }

      const schedShiftsEl = [...section.querySelectorAll('.sched-shift')];
      if (schedShiftsEl.length) {
        let any = !tokens.length;
        schedShiftsEl.forEach(el => {
          const hay = (el.dataset.keywords || el.textContent || '').toLowerCase();
          const matched = !tokens.length || tokens.every(t => hay.indexOf(t) >= 0 || sectionKeywords.indexOf(t) >= 0);
          el.classList.toggle('is-dim', tokens.length && !matched);
          if (matched) any = true;
        });
        if (any) {
          sectionHasMatch = true;
          if (tokens.length) matchCount++;
        }
      } else if (!peopleCards.length && !actions.length && (!tokens.length || fuzzyScore(sectionKeywords, q) >= 0)) {
        sectionHasMatch = true;
        if (tokens.length) matchCount++;
      }

      section.classList.toggle('hidden', tokens.length && !sectionHasMatch);
    });

    zones.forEach(zone => {
      if (!tokens.length) {
        zone.classList.remove('hidden');
        zone.classList.add('is-active');
        return;
      }
      const hasVisible = zone.querySelector('.category:not(.hidden), .action:not(.hidden)');
      zone.classList.toggle('hidden', !hasVisible);
      if (hasVisible) zone.classList.add('is-active');
    });

    noResults.classList.toggle('visible', isSearching && matchCount === 0);

    rebuildResults();
    if (isSearching) setActive(0); else clearActive();

    if (meta) {
      meta.hidden = !isSearching;
      if (isSearching && countEl) {
        countEl.textContent = matchCount === 1 ? '1 result' : `${matchCount} results`;
      }
    }
  }

  search.addEventListener('input', runSearch);

  search.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive(activeIdx + 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(activeIdx - 1); }
    else if (e.key === 'Enter') {
      if (results[activeIdx]) {
        e.preventDefault();
        const el = results[activeIdx];
        if (el.id === 'issues-open') openIssuesPanel();
        else el.click();
      }
    } else if (e.key === 'Escape') {
      if (search.value) { e.preventDefault(); search.value = ''; runSearch(); }
      else { search.blur(); }
    }
  });

  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      search.value = '';
      runSearch();
      search.focus();
    });
  }

  // Global shortcuts: "/" or Cmd/Ctrl+K to jump to search from anywhere.
  document.addEventListener('keydown', e => {
    const tag = (e.target.tagName || '').toLowerCase();
    const typing = tag === 'input' || tag === 'textarea' || e.target.isContentEditable;
    if (!typing && e.key === '/') { e.preventDefault(); search.focus(); }
    else if ((e.metaKey || e.ctrlKey) && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      search.focus();
      search.select();
    }
  });

  // Deep-link from other pages: index.html?q=fleet
  const params = new URLSearchParams(location.search);
  const initialQ = params.get('q');
  if (initialQ) {
    search.value = initialQ;
    runSearch();
    search.focus();
    if (history.replaceState) {
      const url = location.pathname + location.hash;
      history.replaceState(null, '', url);
    }
  }
})();
