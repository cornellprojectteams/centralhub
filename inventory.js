/**
 * Tool & equipment inventory app (inventory.html).
 *
 * Data comes from the inventory Google Sheet through its web app
 * (apps-script/08_inventory.gs) at window.INVENTORY.apiUrl. Reads and changes are
 * JSONP calls: the hub is a static site and a script tag needs no CORS setup (the
 * Command Center works the same way). Photos are too big for a URL, so they go up
 * as a POST. The last good copy of the list is kept in localStorage, so the page
 * paints at once and still shows something if the sheet is slow to answer.
 *
 * Tools stay in the shops; nobody borrows them. The board shows each one on a steel
 * pegboard, like a shop shadow board: striped tape means out of service, a hang tag
 * flags a repair or a supply running low, and retired items fade.
 *
 * URL options
 *   ?item=ID   open one item (what the printed QR labels link to)
 *   ?admin=1   admin view, from the admin page: add and edit items, photos,
 *              maintenance, stock counts, test data
 *   ?embed=1   inside the hub's slide-over panel (no header or footer)
 *   ?labels=1  printable QR labels
 */
(function () {
  'use strict';

  const CFG = window.INVENTORY || {};
  const API = String(CFG.apiUrl || '').trim();
  const APP_URL = String(CFG.appUrl || '').trim() || new URL('inventory.html', location.href).href;
  const params = new URLSearchParams(location.search);
  const ADMIN = params.get('admin') === '1';
  const LABELS = params.get('labels') === '1';
  const START_ID = String(params.get('item') || '').trim().toUpperCase();
  const CACHE_KEY = 'ops-hub-inventory-v3';
  const SHARED = 'Shared';
  const QR_LIB = 'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js';
  const mq = (q) => !!(window.matchMedia && window.matchMedia(q).matches);
  const CAN_HOVER = mq('(hover: hover) and (pointer: fine)');

  const app = document.getElementById('inv-app');
  const shell = document.querySelector('.app-shell') || app;
  const drawer = document.getElementById('inv-drawer');
  const drawerPanel = drawer.querySelector('.inv-drawer-panel');
  const drawerBody = document.getElementById('inv-drawer-body');
  const toastEl = document.getElementById('inv-toast');

  const state = {
    items: [],
    lists: { categories: [], subcategories: [], locations: [], teams: [] },
    meta: { today: localToday(), testItems: 0, api: 0 },
    loaded: false,        // there is data to show (fresh or the saved copy)
    loading: false,
    error: '',
    savedAt: 0,
    filter: 'all', q: '', team: '', loc: '', cat: '',
    view: '',             // drawer: '' (closed), 'item', 'edit', 'add', 'maint'
    openId: '',
    details: {},          // id -> { at, maintenance } or { at, error }
    returnFocus: null,
    armed: '',            // a destructive button waiting for its second click
    startHandled: !START_ID,
    pushed: false,        // the open drawer added a browser history entry
    flashId: '',          // the item to highlight once the drawer closes
    localPhotos: {},      // id -> { data, photo }: a photo just uploaded (Drive thumbnails take a moment)
    busyPhoto: '',
    pendingPhoto: '',     // a photo picked in the add/edit form, sent after the item saves
  };

  // ---------------------------------------------------------------- small helpers

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function pad(n) { return String(n).padStart(2, '0'); }
  function isoOf(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function localToday() { return isoOf(new Date()); }
  function todayIso() { return state.meta.today || localToday(); }
  function parseDay(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
  }
  // Whole days from a to b (positive when b is later).
  function daysFrom(a, b) {
    const x = parseDay(a), y = parseDay(b);
    return x && y ? Math.round((y - x) / 86400000) : 0;
  }
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function fmtDay(iso) {
    const d = parseDay(iso);
    if (!d) return String(iso || '');
    return MONTHS[d.getMonth()] + ' ' + d.getDate() + (d.getFullYear() === new Date().getFullYear() ? '' : ', ' + d.getFullYear());
  }
  function clock(d) {
    let h = d.getHours();
    const ap = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    return h + ':' + pad(d.getMinutes()) + ' ' + ap;
  }
  function fmtWhen(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/.exec(s || '');
    if (!m) return String(s || '');
    const d = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
    return fmtDay(isoOf(d)) + ', ' + clock(d);
  }
  const plural = (n, one, many) => n + ' ' + (n === 1 ? one : many);
  const money = (n) => '$' + Number(n).toFixed(2).replace(/\.00$/, '');
  const byName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
  const byId = (a, b) => a.id.localeCompare(b.id, undefined, { numeric: true });
  const uniq = (arr) => arr.filter((v, i) => v && arr.indexOf(v) === i);

  function recall(key) { try { return localStorage.getItem(key); } catch (e) { return null; } }
  function store(key, val) {
    try { if (val == null || val === '') localStorage.removeItem(key); else localStorage.setItem(key, val); } catch (e) { /* private mode */ }
  }

  function toast(msg) {
    toastEl.textContent = msg;
    toastEl.hidden = false;
    requestAnimationFrame(() => toastEl.classList.add('is-on'));
    clearTimeout(toast.t);
    toast.t = setTimeout(() => {
      toastEl.classList.remove('is-on');
      setTimeout(() => { if (!toastEl.classList.contains('is-on')) toastEl.hidden = true; }, 250);
    }, 4200);
  }

  function readCache() {
    try { return JSON.parse(recall(CACHE_KEY) || 'null'); } catch (e) { return null; }
  }
  function writeCache() {
    store(CACHE_KEY, JSON.stringify({ items: state.items, lists: state.lists, meta: state.meta, savedAt: state.savedAt }));
  }

  // ---------------------------------------------------------------- talking to the sheet

  // JSONP call to the sheet's web app. Resolves with the payload when ok is true.
  let seq = 0;
  function api(query, timeoutMs) {
    return new Promise((resolve, reject) => {
      if (!API) { reject(new Error('The inventory is not connected yet.')); return; }
      const cb = '__inv' + Date.now().toString(36) + (++seq);
      const q = new URLSearchParams(query);
      if (ADMIN) q.set('admin', '1');
      q.set('callback', cb);
      const s = document.createElement('script');
      let done = false;
      const finish = (fn, v) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        window[cb] = function () {};   // a late reply after a timeout is ignored quietly
        if (s.parentNode) s.parentNode.removeChild(s);
        fn(v);
      };
      const timer = setTimeout(() => finish(reject, new Error('The inventory took too long to answer. Try again.')), timeoutMs || 30000);
      window[cb] = (data) => {
        if (data && data.ok) finish(resolve, data);
        else finish(reject, new Error((data && data.error) || 'Something went wrong. Try again.'));
      };
      s.onerror = () => finish(reject, new Error('Could not reach the inventory. Check your connection and try again.'));
      s.src = API + (API.indexOf('?') >= 0 ? '&' : '?') + q.toString();
      document.head.appendChild(s);
    });
  }

  // Photo upload: JSON sent as text/plain (no CORS preflight). If the browser will
  // not hand over the reply, the item is read back to see whether the photo landed.
  async function postPhoto(id, data, before) {
    const ctrl = typeof AbortController === 'function' ? new AbortController() : null;
    const timer = setTimeout(() => { if (ctrl) ctrl.abort(); }, 90000);
    try {
      const res = await fetch(API, {
        method: 'POST',
        body: JSON.stringify({ action: 'photo', admin: '1', id: id, data: data }),
        signal: ctrl ? ctrl.signal : undefined,
      });
      const out = await res.json();
      if (!out || !out.ok) {
        const e = new Error((out && out.error) || 'The photo did not save. Try again.');
        e.fromServer = true;
        throw e;
      }
      return out;
    } catch (err) {
      if (err.fromServer) throw err;
      const check = await api({ action: 'item', id: id }).catch(() => null);
      if (check && check.item && check.item.photo && check.item.photo !== before) return { ok: true, item: check.item };
      throw new Error('The photo did not upload. Check your connection and try again.');
    } finally {
      clearTimeout(timer);
    }
  }

  // Shrink a picked photo to at most 1600 px and re-encode it as JPEG.
  function readPhoto(file) {
    return new Promise((resolve, reject) => {
      if (!file) { reject(new Error('No photo picked.')); return; }
      if (file.type && !/^image\//.test(file.type)) { reject(new Error('That file is not a photo. Pick a JPG or PNG.')); return; }
      const url = URL.createObjectURL(file);
      const img = new Image();
      const fail = () => { URL.revokeObjectURL(url); reject(new Error('Could not read that photo. Try a JPG or PNG.')); };
      img.onload = () => {
        try {
          const k = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
          const w = Math.max(1, Math.round(img.naturalWidth * k)), h = Math.max(1, Math.round(img.naturalHeight * k));
          const c = document.createElement('canvas');
          c.width = w; c.height = h;
          const g = c.getContext('2d');
          g.fillStyle = '#fff';            // transparent PNGs get a white background
          g.fillRect(0, 0, w, h);
          g.drawImage(img, 0, 0, w, h);
          URL.revokeObjectURL(url);
          resolve(c.toDataURL('image/jpeg', 0.84));
        } catch (e) { fail(); }
      };
      img.onerror = fail;
      img.src = url;
    });
  }

  function driveId(url) {
    const s = String(url || '');
    const m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(s) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(s);
    return m ? m[1] : '';
  }
  // What to put in an <img> for an item's photo, or '' when there is none to show.
  function photoSrc(it, width) {
    const local = state.localPhotos[it.id];
    if (local && (local.photo === null || local.photo === it.photo)) return local.data;
    const p = String(it.photo || '').trim();
    if (!/^https:\/\//i.test(p)) return '';
    const id = driveId(p);
    if (id && /google\.com/i.test(p)) return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + width;
    if (/\.(jpe?g|png|gif|webp|avif)(\?|#|$)/i.test(p) || /googleusercontent\.com/i.test(p)) return p;
    return '';   // a link to some other page: offered as a link in the details instead
  }

  // ---------------------------------------------------------------- status

  // What to show for an item: a status key (colour, filters), a label, the line under
  // its name, and the hang tag on its picture.
  function look(it) {
    if (it.oos === 'Retired' || it.status === 'Retired') return { key: 'retired', label: 'Retired', line: 'Retired' };
    if (it.oos === 'Maintenance' || it.status === 'Maintenance') return { key: 'maint', label: 'Out of service', line: 'Out of service for maintenance' };
    if (it.consumable) {
      const reorder = it.min != null ? ', reorder at ' + it.min : '';
      if (it.qty == null) return { key: 'stock', label: 'Not counted', line: 'Not counted yet' };
      if (it.qty <= 0) return { key: 'empty', label: 'Out of stock', line: 'None left' + reorder, tag: 'Out of stock' };
      if (it.min != null && it.qty < it.min) return { key: 'low', label: 'Low stock', line: it.qty + ' left' + reorder, tag: 'Low stock' };
      return { key: 'stock', label: 'In stock', line: it.qty + ' left' + reorder };
    }
    const line = it.loc || 'No location set';
    if (it.cond === 'Needs Repair') return { key: 'ok', label: 'In service', line: line, tag: 'Needs repair', tagKey: 'repair' };
    return { key: 'ok', label: 'In service', line: line };
  }

  function maintDue(it) {
    return !!(it.nextMaint && parseDay(it.nextMaint) && daysFrom(todayIso(), it.nextMaint) <= 14);
  }

  const FILTERS = [
    { key: 'all', label: 'All', test: (it, l) => l.key !== 'retired' },
    { key: 'ok', label: 'In service', test: (it, l) => l.key === 'ok' },
    { key: 'repair', label: 'Needs repair', alert: true, test: (it, l) => it.cond === 'Needs Repair' && l.key !== 'retired' },
    { key: 'stock', label: 'Low stock', alert: true, test: (it, l) => l.key === 'low' || l.key === 'empty' },
    { key: 'oos', label: 'Out of service', test: (it, l) => l.key === 'maint' },
    { key: 'maintdue', label: 'Maintenance due', alert: true, admin: true, test: (it, l) => l.key !== 'retired' && maintDue(it) },
    { key: 'retired', label: 'Retired', admin: true, test: (it, l) => l.key === 'retired' },
  ].filter((f) => ADMIN || !f.admin);

  const PICKS = [
    { key: 'team', label: 'Team', all: 'All teams' },
    { key: 'loc', label: 'Location', all: 'All locations' },
    { key: 'cat', label: 'Category', all: 'All categories' },
  ];

  function findItem(id) { return state.items.find((i) => i.id === id) || null; }
  function teamOf(it) { return it.team || SHARED; }
  function photosOn() { return ADMIN && state.meta.api >= 2; }

  function upsert(item) {
    const i = state.items.findIndex((x) => x.id === item.id);
    if (i >= 0) state.items[i] = item; else state.items.push(item);
    state.meta.testItems = state.items.filter((x) => x.test).length;
    state.savedAt = Date.now();
    writeCache();
  }

  // ---------------------------------------------------------------- pictures

  // Silhouettes for items without a photo, one per category, drawn on a 48 px grid.
  // Each is a single path wound one way, so it can be filled as one shape.
  const GLYPHS = {
    drill: 'M8 10H27A4 4 0 0 1 31 14V18A4 4 0 0 1 27 22H8A4 4 0 0 1 4 18V14A4 4 0 0 1 8 10Z' +
      'M31 12.5H35.5L37.5 14.2V17.8L35.5 19.5H31Z M37.5 15.2H43L45.5 16L43 16.8H37.5Z M10.5 21H21.5L19 36H8Z' +
      'M7.5 35H20A2 2 0 0 1 22 37V40.5A2 2 0 0 1 20 42.5H7.5A2 2 0 0 1 5.5 40.5V37A2 2 0 0 1 7.5 35Z M21.5 23C23.8 23.2 24.4 25.6 23 27.5H21Z',
    wrench: 'M16 21H31A3 3 0 0 1 34 24A3 3 0 0 1 31 27H16A3 3 0 0 1 13 24A3 3 0 0 1 16 21Z' +
      'M4.08 21.2A7 7 0 1 1 4.08 26.8H10A2.8 2.8 0 0 0 10 21.2Z' +
      'M31.5 24A6.5 6.5 0 1 1 44.5 24A6.5 6.5 0 1 1 31.5 24Z M34.6 24A3.4 3.4 0 1 0 41.4 24A3.4 3.4 0 1 0 34.6 24Z',
    hat: 'M9 30C9 20 15 13 24 13C33 13 39 20 39 30Z M23 10.5H25A2 2 0 0 1 27 12.5V20H21V12.5A2 2 0 0 1 23 10.5Z' +
      'M6.5 30H41.5A2.5 2.5 0 0 1 44 32.5A2.5 2.5 0 0 1 41.5 35H6.5A2.5 2.5 0 0 1 4 32.5A2.5 2.5 0 0 1 6.5 30Z',
    scope: 'M8 11H40A3 3 0 0 1 43 14V32A3 3 0 0 1 40 35H8A3 3 0 0 1 5 32V14A3 3 0 0 1 8 11Z M9.5 15.5V29.5H29.5V15.5Z' +
      'M33 20A2.5 2.5 0 1 0 38 20A2.5 2.5 0 1 0 33 20Z M33 28A2.5 2.5 0 1 0 38 28A2.5 2.5 0 1 0 33 28Z M10 35H15V38H10Z M33 35H38V38H33Z',
    box: 'M6 11H42A1.5 1.5 0 0 1 43.5 12.5V17H4.5V12.5A1.5 1.5 0 0 1 6 11Z M7 18.5H41V40A2 2 0 0 1 39 42H9A2 2 0 0 1 7 40Z M17 24V32H31V24Z',
    tag: 'M24 6L34 14V40A2 2 0 0 1 32 42H16A2 2 0 0 1 14 40V14Z M21.4 15A2.6 2.6 0 1 0 26.6 15A2.6 2.6 0 1 0 21.4 15Z',
  };
  function glyphKey(cat) {
    const c = String(cat || '').toLowerCase();
    if (/power/.test(c)) return 'drill';
    if (/hand/.test(c)) return 'wrench';
    if (/safety/.test(c)) return 'hat';
    if (/equip|machine|instrument/.test(c)) return 'scope';
    if (/consum|suppl/.test(c)) return 'box';
    return 'tag';
  }
  function glyphSvg(key) {
    const extra = key === 'scope' ? '<path d="M11.5 24.5l3-4.5 3.2 7 3.1-7 3.2 7 3-4.5" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/>' : '';
    return '<svg viewBox="0 0 48 48" focusable="false" aria-hidden="true"><path fill="currentColor"' +
      (key === 'wrench' ? ' transform="rotate(-40 24 24)"' : '') + ' d="' + GLYPHS[key] + '"/>' + extra + '</svg>';
  }

  // Photos seen once are drawn straight in on later redraws (they only fade in the first time).
  const loadedPhotos = new Set();

  function gauge(it, l) {
    const top = Math.max(1, it.min != null ? it.min * 2 : 0, it.qty || 0);
    const pct = Math.max(0, Math.min(100, Math.round(((it.qty || 0) / top) * 100)));
    return '<span class="inv-gauge is-' + l.key + '"><i style="width:' + pct + '%"></i></span>';
  }

  // The picture area of an item: its photo, or its category silhouette on a
  // pegboard, plus the shop markings for its state.
  function bed(it, big) {
    const l = look(it);
    const src = photoSrc(it, big ? 1200 : 480);
    return '<span class="inv-bed is-' + l.key + (src ? ' has-photo' : '') + (big ? ' is-big' : '') + '" aria-hidden="true">' +
      '<span class="inv-glyph">' + glyphSvg(glyphKey(it.cat)) + '</span>' +
      (src ? '<img class="inv-photo' + (loadedPhotos.has(src) ? ' is-in' : '') + '" src="' + esc(src) + '" alt=""' + (big ? '' : ' loading="lazy"') + ' decoding="async" referrerpolicy="no-referrer">' : '') +
      (l.key === 'maint' ? '<span class="inv-hazard"></span>' : '') +
      (it.consumable && it.qty != null ? gauge(it, l) : '') +
      (l.tag ? '<span class="inv-tag is-' + (l.tagKey || l.key) + '">' + esc(l.tag) + '</span>' : '') +
    '</span>';
  }

  // ---------------------------------------------------------------- the board

  function scaffold() {
    const icon = (d) => '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + d + '</svg>';
    app.innerHTML =
      '<section class="inv-top">' +
        '<div class="inv-head">' +
          '<div class="inv-title-row"><h1 class="inv-title">Tool inventory</h1>' + (ADMIN ? '<span class="inv-admin-pill">Admin view</span>' : '') + '</div>' +
          '<p class="inv-reading" id="inv-reading"></p>' +
        '</div>' +
        (ADMIN
          ? '<div class="inv-head-tools">' +
              '<button type="button" class="inv-btn is-primary" data-act="add">' + icon('<line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line>') + 'Add item</button>' +
              '<a class="inv-btn" href="inventory.html?labels=1" target="_blank" rel="noopener">' + icon('<polyline points="6 9 6 2 18 2 18 9"></polyline><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect>') + 'Print labels</a>' +
            '</div>'
          : '') +
      '</section>' +
      '<div id="inv-notice"></div>' +
      '<div class="inv-controls" id="inv-controls">' +
        '<label class="inv-search">' + icon('<circle cx="11" cy="11" r="7"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line>') +
          '<input type="search" id="inv-q" placeholder="Search tools, IDs, teams or places" autocomplete="off" spellcheck="false" aria-label="Search the inventory">' +
          (CAN_HOVER ? '<kbd class="inv-kbd" aria-hidden="true">/</kbd>' : '') +
        '</label>' +
        '<div class="inv-chips" id="inv-chips" role="group" aria-label="Filter by status"></div>' +
      '</div>' +
      '<div class="inv-picks" id="inv-picks" hidden>' +
        PICKS.map((p) => '<label class="inv-pick" data-pick="' + p.key + '"><span class="inv-pick-k">' + p.label + '</span>' +
          '<span class="inv-pick-v">All</span><select aria-label="Filter by ' + p.label.toLowerCase() + '"></select></label>').join('') +
      '</div>' +
      '<div class="inv-countrow"><p class="inv-count" id="inv-count"></p><span class="inv-fresh" id="inv-fresh"></span></div>' +
      '<div id="inv-results"></div>' +
      (ADMIN ? '<section class="inv-admin-foot" id="inv-admin-foot"></section>' : '');

    const q = document.getElementById('inv-q');
    q.addEventListener('input', () => { state.q = q.value; renderResults(); });
    q.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      const list = visibleItems();
      if (list.length === 1) { e.preventDefault(); openItem(list[0].id, q); }
    });
    watchStuck();
  }

  function renderBoard() {
    if (!document.getElementById('inv-results')) scaffold();
    renderReading();
    renderNotice();
    renderChips();
    renderPicks();
    renderResults();
    renderAdminFoot();
  }

  function renderReading() {
    const el = document.getElementById('inv-reading');
    if (!el) return;
    if (!state.loaded) { el.textContent = state.error ? '' : 'Loading the inventory…'; return; }
    const live = state.items.filter((it) => look(it).key !== 'retired');
    const tools = live.filter((it) => !it.consumable);
    if (!live.length) { el.textContent = ''; return; }
    const parts = [];
    if (tools.length) {
      const ok = tools.filter((it) => look(it).key === 'ok').length;
      parts.push(ok === tools.length
        ? (tools.length === 1 ? 'The 1 tool is in service.' : 'All ' + tools.length + ' tools are in service.')
        : ok + ' of ' + tools.length + ' tools are in service.');
      const repair = tools.filter((it) => it.cond === 'Needs Repair').length;
      if (repair) parts.push(repair === 1 ? '1 needs repair.' : repair + ' need repair.');
    }
    const low = live.filter((it) => it.consumable && (look(it).key === 'low' || look(it).key === 'empty')).length;
    if (low) parts.push(low === 1 ? '1 supply is running low.' : low + ' supplies are running low.');
    el.textContent = parts.join(' ') || plural(live.length, 'supply', 'supplies') + ' on hand.';
  }

  function renderNotice() {
    const el = document.getElementById('inv-notice');
    if (!el) return;
    let html = '';
    if (state.error && state.loaded) {
      html += '<div class="inv-notice is-warn" role="alert"><span>' + esc(state.error) +
        (state.savedAt ? ' Showing what was loaded at ' + esc(clock(new Date(state.savedAt))) + '.' : '') +
        '</span><button type="button" class="inv-btn is-small" data-act="reload">Try again</button></div>';
    }
    if (ADMIN && state.loaded && !state.error && state.meta.api < 2) {
      html += '<div class="inv-notice is-info"><span><b>Photos and teams are switched off.</b> The inventory sheet is still running the previous script. ' +
        'Paste the new <code>08_inventory.gs</code> into the sheet’s Apps Script, run setupInventory, then Deploy &gt; Manage deployments &gt; New version.</span></div>';
    }
    el.innerHTML = html;
  }

  function counts() {
    const out = {};
    FILTERS.forEach((f) => { out[f.key] = 0; });
    state.items.forEach((it) => {
      const l = look(it);
      FILTERS.forEach((f) => { if (f.test(it, l)) out[f.key]++; });
    });
    return out;
  }

  function renderChips() {
    const el = document.getElementById('inv-chips');
    if (!el) return;
    if (!state.loaded || !state.items.length) { el.innerHTML = ''; return; }
    const n = counts();
    el.innerHTML = FILTERS.map((f) => {
      const on = state.filter === f.key;
      const cls = 'inv-chip' + (on ? ' is-on' : '') + (f.alert && n[f.key] > 0 ? ' is-alert' : '') + (!n[f.key] && f.key !== 'all' ? ' is-zero' : '');
      return '<button type="button" class="' + cls + '" data-filter="' + f.key + '" aria-pressed="' + on + '">' +
        esc(f.label) + '<b>' + n[f.key] + '</b></button>';
    }).join('');
  }

  // Values for a dropdown filter: only what the items use, with how many each.
  function pickOptions(kind) {
    const count = {};
    state.items.forEach((it) => {
      if (look(it).key === 'retired' && state.filter !== 'retired') return;
      const v = kind === 'team' ? teamOf(it) : kind === 'loc' ? it.loc : it.cat;
      if (v) count[v] = (count[v] || 0) + 1;
    });
    const vals = Object.keys(count);
    if (kind === 'team') vals.sort((a, b) => (a === SHARED ? -1 : b === SHARED ? 1 : a.localeCompare(b)));
    else if (kind === 'cat') {
      const order = state.lists.categories.map((c) => c.name);
      const rank = (v) => (order.indexOf(v) < 0 ? 99 : order.indexOf(v));
      vals.sort((a, b) => rank(a) - rank(b) || a.localeCompare(b));
    } else vals.sort((a, b) => a.localeCompare(b));
    return vals.map((v) => ({ v: v, n: count[v] }));
  }

  function renderPicks() {
    const el = document.getElementById('inv-picks');
    if (!el) return;
    el.hidden = !(state.loaded && state.items.length);
    if (el.hidden) return;
    PICKS.forEach((p) => {
      const wrap = el.querySelector('[data-pick="' + p.key + '"]');
      const sel = wrap.querySelector('select');
      const opts = pickOptions(p.key);
      if (state[p.key] && !opts.some((o) => o.v === state[p.key])) opts.push({ v: state[p.key], n: 0 });
      // Rebuild the options only when they change, so an open menu is never yanked away.
      const sig = opts.map((o) => o.v + ':' + o.n).join('|');
      if (sel.dataset.sig !== sig) {
        sel.innerHTML = '<option value="">' + esc(p.all) + '</option>' +
          opts.map((o) => '<option value="' + esc(o.v) + '">' + esc(o.v) + ' (' + o.n + ')</option>').join('');
        sel.dataset.sig = sig;
      }
      if (sel.value !== state[p.key]) sel.value = state[p.key];
      wrap.querySelector('.inv-pick-v').textContent = state[p.key] || 'All';
      wrap.classList.toggle('is-set', !!state[p.key]);
      wrap.hidden = p.key === 'team' && opts.length < 2 && !state.team;
    });
    renderFresh();
  }

  function renderFresh() {
    const el = document.getElementById('inv-fresh');
    if (!el) return;
    el.innerHTML = state.loading
      ? '<span class="inv-spin" aria-hidden="true"></span>Updating'
      : (state.savedAt ? 'Updated ' + esc(clock(new Date(state.savedAt))) : '') +
        '<button type="button" class="inv-icon-btn" data-act="reload" aria-label="Refresh" title="Refresh">' +
          '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12a9 9 0 1 1-2.64-6.36"></path><polyline points="21 3 21 9 15 9"></polyline></svg></button>';
  }

  function visibleItems() {
    const f = FILTERS.find((x) => x.key === state.filter) || FILTERS[0];
    const words = state.q.toLowerCase().split(/\s+/).filter(Boolean);
    const list = state.items.filter((it) => {
      const l = look(it);
      if (!f.test(it, l)) return false;
      if (state.team && teamOf(it) !== state.team) return false;
      if (state.loc && it.loc !== state.loc) return false;
      if (state.cat && it.cat !== state.cat) return false;
      if (!words.length) return true;
      const hay = [it.id, it.name, it.cat, it.sub, it.loc, teamOf(it), it.serial, l.label, l.tag].join(' ').toLowerCase();
      return words.every((w) => hay.indexOf(w) >= 0);
    });
    if (state.filter === 'stock') list.sort((a, b) => (a.qty || 0) - (b.qty || 0) || byName(a, b));
    else if (state.filter === 'maintdue') list.sort((a, b) => a.nextMaint.localeCompare(b.nextMaint) || byName(a, b));
    else list.sort(byName);
    return list;
  }

  function slot(it, flash) {
    const l = look(it);
    const team = teamOf(it) !== SHARED ? '<span class="inv-team">' + esc(teamOf(it)) + '</span>' : '';
    return '<li><button type="button" class="inv-slot is-' + l.key + (flash === it.id ? ' is-flash' : '') + '" data-id="' + esc(it.id) + '">' +
      bed(it, false) +
      '<span class="inv-slot-body">' +
        '<span class="inv-slot-top"><span class="inv-assettag">' + esc(it.id) + '</span>' + team +
          (it.test ? '<span class="inv-testmark">Test</span>' : '') + '</span>' +
        '<span class="inv-slot-name">' + esc(it.name) + '</span>' +
        '<span class="inv-slot-line"><span class="inv-dot is-' + l.key + '" aria-hidden="true"></span><span>' + esc(l.line) +
          (l.tag ? '<span class="sr-only">. ' + esc(l.tag) + '.</span>' : '') + '</span></span>' +
      '</span></button></li>';
  }

  function renderResults() {
    const el = document.getElementById('inv-results');
    const count = document.getElementById('inv-count');
    if (!el) return;
    renderFresh();

    if (!state.loaded) {
      count.innerHTML = '';
      el.innerHTML = state.error
        ? '<div class="inv-empty is-error"><h2 class="inv-empty-title">The inventory didn’t load</h2><p>' + esc(state.error) +
            '</p><button type="button" class="inv-btn is-primary" data-act="reload">Try again</button></div>'
        : '<ul class="inv-grid" aria-hidden="true">' + new Array(8).fill('<li><div class="inv-slot is-skeleton"><span class="inv-bed"></span><span class="inv-slot-body"><i></i><i></i></span></div></li>').join('') + '</ul>';
      return;
    }

    if (!state.items.length) {
      count.innerHTML = '';
      el.innerHTML = ADMIN
        ? '<div class="inv-empty"><h2 class="inv-empty-title">Nothing on the board yet</h2><p>Add the first tool, or load test data to see how everything works.</p>' +
            '<div class="inv-empty-actions"><button type="button" class="inv-btn is-primary" data-act="add">Add item</button>' +
            '<button type="button" class="inv-btn" data-act="test-add">Load test data</button></div></div>'
        : '<div class="inv-empty"><h2 class="inv-empty-title">Nothing on the board yet</h2><p>The inventory is empty for now. Check back soon.</p></div>';
      return;
    }

    const list = visibleItems();
    const total = state.items.filter((it) => look(it).key !== 'retired').length;
    const filtered = state.filter !== 'all' || state.q || state.team || state.loc || state.cat;
    count.innerHTML = filtered
      ? 'Showing ' + list.length + ' of ' + plural(total, 'item', 'items') + '<button type="button" class="inv-link" data-act="clear">Clear filters</button>'
      : plural(total, 'item', 'items');
    const flash = state.view ? '' : state.flashId;
    el.innerHTML = list.length
      ? '<ul class="inv-grid">' + list.map((it) => slot(it, flash)).join('') + '</ul>'
      : '<div class="inv-empty"><h2 class="inv-empty-title">Nothing matches</h2><p>Try another search, or clear the filters to see everything.</p>' +
          '<button type="button" class="inv-btn" data-act="clear">Clear filters</button></div>';
    if (flash) {
      state.flashId = '';
      setTimeout(() => { const s = el.querySelector('.inv-slot.is-flash'); if (s) s.classList.remove('is-flash'); }, 1900);
    }
  }

  function renderAdminFoot() {
    const el = document.getElementById('inv-admin-foot');
    if (!el) return;
    if (!state.loaded) { el.innerHTML = ''; return; }
    const n = state.items.filter((i) => i.test).length;
    const btn = (act, label, armedLabel, danger) =>
      '<button type="button" class="inv-btn is-small' + (state.armed === act ? ' is-danger' : danger ? ' is-quiet-danger' : '') + '" data-act="' + act + '">' +
      (state.armed === act ? armedLabel : label) + '</button>';
    el.innerHTML = '<div class="inv-admin-foot-text"><h2>Test data</h2><p>' +
      (n ? plural(n, 'test item is', 'test items are') + ' loaded. Each one is tagged “[test data]” in its notes, and removing them leaves real items alone.'
        : 'Sample tools, supplies and maintenance you can try things on, then remove in one click.') + '</p></div>' +
      '<div class="inv-admin-foot-actions">' +
        (n ? btn('test-reset', 'Reset test data', 'Click again to reset', false) + btn('test-remove', 'Remove test data', 'Click again to remove', true)
           : '<button type="button" class="inv-btn is-small" data-act="test-add">Load test data</button>') +
      '</div>';
  }

  function renderNotConnected() {
    app.innerHTML = '<section class="inv-setup"><h1 class="inv-title">The inventory isn’t connected yet</h1>' +
      (ADMIN
        ? '<p>Deploy the inventory sheet’s script as a web app (Deploy &gt; New deployment &gt; Web app, execute as Me, access Anyone) and put its URL in <code>config.js</code> as <code>INVENTORY.apiUrl</code>. The README has the steps.</p>'
        : '<p>Check back soon, or ask the Operations team.</p>') +
      '</section>';
  }

  // The search and status chips stay at the top while scrolling; this adds a hairline once they stick.
  function watchStuck() {
    const el = document.getElementById('inv-controls');
    if (!el) return;
    let raf = 0;
    const check = () => {
      raf = 0;
      const top = parseFloat(getComputedStyle(el).top) || 0;
      el.classList.toggle('is-stuck', el.getBoundingClientRect().top <= top + 0.5 && window.scrollY > 0);
    };
    window.addEventListener('scroll', () => { if (!raf) raf = requestAnimationFrame(check); }, { passive: true });
  }

  // ---------------------------------------------------------------- loading

  async function load() {
    if (state.loading) return;
    state.loading = true;
    if (state.loaded && !LABELS) renderFresh();
    let same = false;
    try {
      const res = await api({ action: 'list' });
      const lists = Object.assign({ categories: [], subcategories: [], locations: [], teams: [] }, res.lists || {});
      const meta = Object.assign({ today: localToday(), testItems: 0, api: 0 }, res.meta || {});
      same = state.loaded && !state.error &&
        JSON.stringify([state.items, state.lists, state.meta.today, state.meta.api]) === JSON.stringify([res.items || [], lists, meta.today, meta.api]);
      state.items = res.items || [];
      state.lists = lists;
      state.meta = meta;
      state.loaded = true;
      state.error = '';
      state.savedAt = Date.now();
      writeCache();
    } catch (err) {
      state.error = err.message;
    } finally {
      state.loading = false;
    }
    if (LABELS) { renderLabels(); return; }
    if (same && document.getElementById('inv-results')) { renderFresh(); handleStart(true); return; }
    renderBoard();
    if (state.view === 'item' && state.openId && findItem(state.openId)) refreshOpenItem();
    handleStart(true);
  }

  // A ?item= link (a scanned label) opens that item once the data is in.
  function handleStart(fresh) {
    if (state.startHandled || !state.loaded) return;
    const it = findItem(START_ID);
    if (it) { state.startHandled = true; openItem(it.id, null, { replace: true }); }
    else if (fresh) { state.startHandled = true; toast('There is no item ' + START_ID + ' in the inventory.'); setUrlItem('', false); }
  }

  // ---------------------------------------------------------------- the item drawer

  function setUrlItem(id, push) {
    const u = new URL(location.href);
    if (id) u.searchParams.set('item', id); else u.searchParams.delete('item');
    const url = u.pathname + u.search + u.hash;
    try {
      if (id && push) { history.pushState({ invItem: id }, '', url); state.pushed = true; }
      else history.replaceState(id ? { invItem: id } : null, '', url);
    } catch (e) { /* sandboxed frame: the address just doesn't change */ }
  }

  function showDrawer() {
    if (drawer.hidden) {
      drawer.hidden = false;
      requestAnimationFrame(() => drawer.classList.add('is-open'));
    }
    document.body.classList.add('inv-locked');
    shell.inert = true;
    drawerPanel.scrollTop = 0;
    drawerPanel.focus({ preventScroll: true });
  }

  function closeDrawer(opts) {
    opts = opts || {};
    if (!state.view) return;
    drawer.classList.remove('is-open');
    document.body.classList.remove('inv-locked');
    shell.inert = false;
    setTimeout(() => { if (!drawer.classList.contains('is-open')) { drawer.hidden = true; drawerBody.innerHTML = ''; } }, 280);
    const id = state.openId;
    state.view = '';
    state.openId = '';
    state.armed = '';
    state.pendingPhoto = '';
    if (opts.fromHistory) state.pushed = false;
    else if (state.pushed) { state.pushed = false; history.back(); }
    else setUrlItem('', false);
    if (state.flashId && state.flashId === id) renderResults();
    const back = state.returnFocus;
    state.returnFocus = null;
    const card = id && app.querySelector('.inv-slot[data-id="' + id + '"]');
    const target = back && document.contains(back) ? back : card;
    if (target) target.focus({ preventScroll: true });
  }

  function openItem(id, from, opts) {
    opts = opts || {};
    const it = findItem(id);
    if (!it) return;
    if (from) state.returnFocus = from;
    const wasOpen = !!state.view;
    state.view = 'item';
    state.openId = id;
    state.armed = '';
    drawerBody.innerHTML = itemView(it);
    showDrawer();
    if (opts.fromHistory) state.pushed = true;
    else if (!wasOpen) setUrlItem(id, !opts.replace);
    loadDetail(id);
  }

  // Re-draw the open item after its data changed, without wiping a half-filled form.
  function refreshOpenItem() {
    if (state.view !== 'item') return;
    const it = findItem(state.openId);
    if (!it) return;
    const active = document.activeElement;
    const typing = active && drawerBody.contains(active) && /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) && active.type !== 'file';
    if (typing) { paintHistory(it.id); return; }
    const scroll = drawerPanel.scrollTop;
    drawerBody.innerHTML = itemView(it);
    drawerPanel.scrollTop = scroll;
  }

  async function loadDetail(id, force) {
    const cur = state.details[id];
    if (!force && cur && (cur.loading || (!cur.error && Date.now() - cur.at < 60000))) {
      if (!cur.loading) paintHistory(id);
      return;
    }
    state.details[id] = Object.assign({}, cur || {}, { loading: true });
    try {
      const res = await api({ action: 'item', id: id });
      state.details[id] = { at: Date.now(), maintenance: res.maintenance || [] };
      if (res.item) {
        const before = findItem(id);
        const changed = !before || JSON.stringify(before) !== JSON.stringify(res.item);
        upsert(res.item);
        if (changed) {
          renderBoard();
          if (state.view === 'item' && state.openId === id) refreshOpenItem();
        }
      }
    } catch (err) {
      state.details[id] = { at: Date.now(), error: err.message };
    }
    paintHistory(id);
  }

  function paintHistory(id) {
    const el = document.getElementById('inv-hist');
    if (el && state.view === 'item' && state.openId === id) el.innerHTML = historyHtml(findItem(id));
  }

  const closeSvg = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';
  const backSvg = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"></polyline></svg>';
  // Inside the hub's panel the item fills the panel, so its control reads as a way back.
  const EMBED = document.documentElement.classList.contains('inv-embed');

  function drawerBar(inner) {
    return '<div class="inv-d-bar">' + inner + (EMBED
      ? '<button type="button" class="inv-d-close" data-close aria-label="Back to all tools">' + backSvg + 'All tools</button>'
      : '<button type="button" class="inv-d-close" data-close aria-label="Close">' + closeSvg + '</button>') + '</div>';
  }

  function fact(label, value, wide) {
    return value ? '<div class="inv-fact' + (wide ? ' is-wide' : '') + '"><dt>' + esc(label) + '</dt><dd>' + value + '</dd></div>' : '';
  }

  function itemView(it) {
    const l = look(it);
    const nextIn = it.nextMaint && parseDay(it.nextMaint) ? daysFrom(todayIso(), it.nextMaint) : null;
    const notes = String(it.notes || '').replace(/\[test data\]\s*/i, '').trim();
    const link = /^https?:\/\//i.test(it.photo) && !photoSrc(it, 100) ? '<a href="' + esc(it.photo) + '" target="_blank" rel="noopener">Open photo</a>' : '';
    const facts =
      fact('Team', teamOf(it) === SHARED ? (it.consumable ? 'Shared supplies' : 'Shared. Any team can use it.') : esc(teamOf(it)), teamOf(it) === SHARED) +
      fact('Location', esc(it.loc)) +
      fact('Category', esc([it.cat, it.sub].filter(Boolean).join(', '))) +
      fact('Condition', it.consumable ? '' : esc(it.cond)) +
      fact('Serial or asset tag', esc(it.serial)) +
      fact('Purchased', it.bought ? esc(fmtDay(it.bought)) : '') +
      fact('Next maintenance', it.nextMaint ? esc(fmtDay(it.nextMaint)) + (nextIn !== null && nextIn < 0 ? ' <span class="inv-late">overdue</span>' : '') : '') +
      fact('Notes', esc(notes), true) +
      fact('Photo', link);

    const where = [it.sub || it.cat, it.loc ? 'kept in ' + it.loc : ''].filter(Boolean).join(', ');
    const action = actionCard(it);
    return drawerBar('<span class="inv-assettag is-lg">' + esc(it.id) + '</span>' +
        '<span class="inv-d-state is-' + l.key + '"><span class="inv-dot is-' + l.key + '" aria-hidden="true"></span>' + esc(l.label) + '</span>' +
        (it.test ? '<span class="inv-testmark">Test data</span>' : '')) +
      '<div class="inv-d-photo">' + bed(it, true) + photoTools(it) + '</div>' +
      '<h2 class="inv-d-title" id="inv-d-title">' + esc(it.name) + '</h2>' +
      '<p class="inv-d-sub">' + esc(where.charAt(0).toUpperCase() + where.slice(1)) + (teamOf(it) !== SHARED ? '<span class="inv-team">' + esc(teamOf(it)) + '</span>' : '') + '</p>' +
      (action ? '<section class="inv-d-action" id="inv-action">' + action + '</section>' : '') +
      (facts ? '<section class="inv-d-sec"><h3 class="inv-d-h">Details</h3><dl class="inv-facts">' + facts + '</dl></section>' : '') +
      '<section class="inv-d-sec"><h3 class="inv-d-h">Maintenance</h3><div id="inv-hist">' + historyHtml(it) + '</div></section>' +
      (ADMIN
        ? '<div class="inv-d-admin"><button type="button" class="inv-btn" data-act="edit">Edit item</button>' +
            '<button type="button" class="inv-btn" data-act="maint">Log maintenance</button></div>'
        : '');
  }

  function photoTools(it) {
    if (!photosOn()) return '';
    if (state.busyPhoto === it.id) return '<div class="inv-photo-tools"><span class="inv-photo-busy"><span class="inv-spin" aria-hidden="true"></span>Saving the photo</span></div>';
    const has = !!String(it.photo || '').trim();
    return '<div class="inv-photo-tools">' +
      '<label class="inv-btn is-small is-on-photo"><input type="file" accept="image/*" class="inv-file" data-photo-for="' + esc(it.id) + '">' +
        (has ? 'Change photo' : 'Add photo') + '</label>' +
      (has ? '<button type="button" class="inv-btn is-small is-on-photo' + (state.armed === 'photo' ? ' is-danger' : '') + '" data-act="photo-remove">' +
        (state.armed === 'photo' ? 'Click again to remove' : 'Remove') + '</button>' : '') +
    '</div>';
  }

  // What to know before using it: out of service, retired, needs repair, or the stock
  // level of a supply. Nothing for a tool that is fine.
  function actionCard(it) {
    const l = look(it);
    if (l.key === 'retired') return '<div class="inv-callout is-retired"><b>Retired.</b> It is no longer in use.</div>';
    if (l.key === 'maint') return '<div class="inv-callout is-maint"><span class="inv-hazard" aria-hidden="true"></span><b>Out of service for maintenance.</b> Ask the Operations team before you use it.</div>';
    if (it.consumable) return stockCard(it, l);
    if (it.cond === 'Needs Repair') return '<div class="inv-callout is-repair"><b>Needs repair.</b> Check it over before you use it, or ask the Operations team.</div>';
    return '';
  }

  function stockCard(it, l) {
    return '<div class="inv-stock is-' + l.key + '">' +
        '<span class="inv-stock-n">' + (it.qty == null ? '?' : it.qty) + '</span>' +
        '<span class="inv-stock-text"><b>on hand</b>' + (it.min != null ? 'Reorder at ' + it.min : 'No reorder level set') + '</span>' +
      '</div>' +
      (ADMIN ? countForm(it) : '<p class="inv-hint">Running low? Tell the Operations team.</p>');
  }

  // ---------------------------------------------------------------- forms

  function field(label, control, cls) {
    return '<label class="inv-field' + (cls ? ' ' + cls : '') + '"><span class="inv-label">' + label + '</span>' + control + '</label>';
  }
  const formFoot = (submit) => '<p class="inv-form-err" role="alert" hidden></p>' +
    '<button type="submit" class="inv-btn is-primary is-block">' + submit + '</button>';

  function countForm(it) {
    return '<form class="inv-card inv-form" data-form="count" novalidate><h3 class="inv-card-title">Update the count</h3>' +
      '<div class="inv-stepper"><button type="button" class="inv-btn" data-step="-1" aria-label="One less">−</button>' +
      '<input type="number" name="qty" min="0" step="1" inputmode="numeric" value="' + (it.qty == null ? '' : it.qty) + '" aria-label="On hand">' +
      '<button type="button" class="inv-btn" data-step="1" aria-label="One more">+</button></div>' +
      formFoot('Save count') + '</form>';
  }

  // A dropdown fed by one of the lists. With addLabel, its last option opens a box
  // for a new value (saved to the list on the Lists tab).
  function pickField(name, label, values, current, o) {
    o = o || {};
    const list = uniq(values.slice());
    if (current && list.indexOf(current) < 0) list.push(current);
    return '<div class="inv-field' + (o.wide ? ' is-wide' : '') + '"><label class="inv-label" for="inv-f-' + name + '">' + label + '</label>' +
      '<span class="inv-select"><select id="inv-f-' + name + '" name="' + name + '"' + (o.required ? ' required' : '') + '>' +
        (o.blank != null ? '<option value="">' + esc(o.blank) + '</option>' : '') +
        list.map((v) => '<option value="' + esc(v) + '"' + (v === current ? ' selected' : '') + '>' + esc(v) + '</option>').join('') +
        (o.addLabel ? '<option value="__new">' + esc(o.addLabel) + '</option>' : '') +
      '</select></span>' +
      (o.addLabel ? '<input class="inv-newval" name="' + name + '__new" maxlength="80" placeholder="' + esc(o.placeholder || '') + '" aria-label="' + esc(o.placeholder || label) + '" hidden>' : '') +
    '</div>';
  }

  function segField(name, label, options, current) {
    return '<fieldset class="inv-field is-wide"><legend class="inv-label">' + label + '</legend><div class="inv-seg">' +
      options.map((o) => '<label class="inv-seg-opt"><input type="radio" name="' + name + '" value="' + esc(o[0]) + '"' + (o[0] === current ? ' checked' : '') + '><span>' + esc(o[1]) + '</span></label>').join('') +
    '</div></fieldset>';
  }

  function teamsList() {
    const t = state.lists.teams && state.lists.teams.length ? state.lists.teams.slice() : [SHARED];
    return uniq([SHARED].concat(t));
  }

  function photoPicker(it) {
    if (!photosOn()) return '';
    const src = it && it.id ? photoSrc(it, 640) : '';
    return '<div class="inv-field is-wide"><span class="inv-label">Photo</span>' +
      '<label class="inv-drop' + (src ? ' has-photo' : '') + '" data-drop>' +
        '<input type="file" accept="image/*" class="inv-file" data-photo-pick>' +
        '<span class="inv-drop-preview">' + (src ? '<img src="' + esc(src) + '" alt="">' : glyphSvg(glyphKey(it && it.cat))) + '</span>' +
        '<span class="inv-drop-text"><b>' + (src ? 'Change the photo' : 'Add a photo') + '</b>' +
          '<span>' + (CAN_HOVER ? 'Drop an image here, or click to choose one.' : 'Take a photo or choose one.') + '</span></span>' +
      '</label></div>';
  }

  function editView(it) {
    const isNew = !it;
    it = it || { cat: '', sub: '', loc: '', team: SHARED, cond: 'Good', oos: '' };
    const cats = state.lists.categories.map((c) => c.name);
    const consumable = String(it.cat).toLowerCase() === 'consumable';
    return drawerBar('<span class="inv-assettag is-lg' + (isNew ? ' is-new' : '') + '">' + (isNew ? 'New' : esc(it.id)) + '</span>') +
      '<h2 class="inv-d-title" id="inv-d-title">' + (isNew ? 'Add an item' : 'Edit ' + esc(it.name)) + '</h2>' +
      '<form class="inv-form inv-edit" data-form="save" novalidate>' +
        photoPicker(isNew ? null : it) +
        field('Name', '<input name="name" required maxlength="120" value="' + esc(it.name || '') + '" placeholder="DeWalt 20V cordless drill" autocomplete="off">', 'is-wide') +
        '<div class="inv-form-grid">' +
          pickField('cat', 'Category', cats, it.cat, { blank: isNew ? 'Pick one' : null, required: true }) +
          pickField('sub', 'Subcategory', state.lists.subcategories, it.sub, { blank: 'None', addLabel: 'Add a new subcategory…', placeholder: 'New subcategory' }) +
          pickField('team', 'Team', teamsList(), teamOf(it), {}) +
          pickField('loc', 'Location', state.lists.locations, it.loc, { blank: 'Not set', addLabel: 'Add a new location…', placeholder: 'New location' }) +
        '</div>' +
        segField('cond', 'Condition', [['Good', 'Good'], ['Fair', 'Fair'], ['Needs Repair', 'Needs repair']], it.cond) +
        segField('oos', 'Status', [['', 'In service'], ['Maintenance', 'Maintenance'], ['Retired', 'Retired']], it.oos || '') +
        '<div class="inv-form-grid">' +
          field('Serial or asset tag', '<input name="serial" maxlength="80" value="' + esc(it.serial || '') + '" autocomplete="off">') +
          field('Purchase date', '<input type="date" name="bought" value="' + esc(parseDay(it.bought) ? it.bought : '') + '">') +
        '</div>' +
        '<div class="inv-form-grid inv-qty-fields' + (consumable ? '' : ' is-hidden') + '">' +
          field('On hand', '<input type="number" name="qty" min="0" step="1" inputmode="numeric" value="' + (it.qty == null ? '' : it.qty) + '">') +
          field('Reorder at', '<input type="number" name="min" min="0" step="1" inputmode="numeric" value="' + (it.min == null ? '' : it.min) + '">') +
        '</div>' +
        field('Notes', '<textarea name="notes" rows="3" maxlength="1000">' + esc(it.notes || '') + '</textarea>', 'is-wide') +
        '<div class="inv-form-bar"><p class="inv-form-err" role="alert" hidden></p><div class="inv-form-actions">' +
          '<button type="submit" class="inv-btn is-primary">' + (isNew ? 'Add item' : 'Save changes') + '</button>' +
          '<button type="button" class="inv-btn" data-act="cancel">Cancel</button></div></div>' +
      '</form>';
  }

  function maintView(it) {
    const t = todayIso();
    return drawerBar('<span class="inv-assettag is-lg">' + esc(it.id) + '</span>') +
      '<h2 class="inv-d-title" id="inv-d-title">Log maintenance</h2><p class="inv-d-sub">' + esc(it.name) + '</p>' +
      '<form class="inv-form inv-edit" data-form="maint" novalidate>' +
        field('Issue or check', '<input name="issue" required maxlength="300" placeholder="Monthly inspection, broken switch, …" autocomplete="off">', 'is-wide') +
        field('What was done', '<input name="taken" maxlength="300" autocomplete="off">', 'is-wide') +
        '<div class="inv-form-grid">' +
          field('Done by', '<input name="by" maxlength="80" placeholder="NetID or name" autocomplete="off">') +
          field('Cost <i>optional</i>', '<input type="number" name="cost" min="0" step="0.01" inputmode="decimal">') +
          field('Next check due <i>optional</i>', '<input type="date" name="next" min="' + t + '">') +
        '</div>' +
        '<div class="inv-form-bar"><p class="inv-form-err" role="alert" hidden></p><div class="inv-form-actions">' +
          '<button type="submit" class="inv-btn is-primary">Save</button><button type="button" class="inv-btn" data-act="cancel">Cancel</button></div></div>' +
      '</form>';
  }

  function openForm(view, it) {
    state.view = view;
    state.openId = it ? it.id : '';
    state.pendingPhoto = '';
    state.armed = '';
    drawerBody.innerHTML = view === 'maint' ? maintView(it) : editView(view === 'add' ? null : it);
    showDrawer();
    const first = drawerBody.querySelector('input[name="name"], input[name="issue"]');
    if (first && CAN_HOVER) first.focus();
  }

  function historyHtml(it) {
    const d = it && state.details[it.id];
    if (!it || !d || (d.loading && !d.maintenance)) return '<div class="inv-hist-wait" aria-label="Loading the maintenance log"><i></i><i></i><i></i></div>';
    if (d.error) return '<p class="inv-hint">Couldn’t load the maintenance log. ' + esc(d.error) + '</p>';
    const rows = (d.maintenance || []).map((m) => ({
      when: m.when,
      html: '<b>' + esc(m.issue.replace(/\s*\[test data\]/i, '')) + '</b>' + (m.taken ? '. ' + esc(m.taken) : '') + (m.by ? ' (' + esc(m.by) + ')' : '') +
        (m.cost != null ? ', ' + esc(money(m.cost)) : '') + (m.next ? '. Next check ' + esc(fmtDay(m.next)) + '.' : ''),
    }));
    if (!rows.length) return '<p class="inv-hint">No maintenance logged yet.</p>';
    rows.sort((a, b) => (a.when < b.when ? 1 : a.when > b.when ? -1 : 0));
    return '<ol class="inv-hist">' + rows.map((r) =>
      '<li class="is-fix"><span class="inv-hist-text">' + r.html + '</span><time>' + esc(fmtWhen(r.when)) + '</time></li>').join('') + '</ol>';
  }

  function formError(form, msg) {
    const el = form.querySelector('.inv-form-err');
    if (el) { el.textContent = msg; el.hidden = !msg; }
    if (msg) {
      const bad = form.querySelector('[aria-invalid="true"]');
      if (bad) bad.focus();
    }
    return null;
  }
  function markBad(form, name) {
    form.querySelectorAll('[aria-invalid]').forEach((x) => x.removeAttribute('aria-invalid'));
    const el = name && form.elements[name];
    if (el && el.setAttribute) el.setAttribute('aria-invalid', 'true');
  }

  // Disable the form while a request runs; show the server's message on failure.
  async function submitWith(form, busyLabel, request, onDone) {
    const btn = form.querySelector('button[type="submit"]');
    const label = btn ? btn.innerHTML : '';
    const controls = Array.prototype.slice.call(form.elements);
    formError(form, '');
    controls.forEach((c) => { c.disabled = true; });
    form.classList.add('is-busy');
    if (btn) btn.innerHTML = '<span class="inv-spin" aria-hidden="true"></span>' + esc(busyLabel);
    try {
      const res = await request();
      onDone(res);
    } catch (err) {
      if (document.contains(form)) {
        controls.forEach((c) => { c.disabled = false; });
        form.classList.remove('is-busy');
        if (btn) btn.innerHTML = label;
        formError(form, err.message);
      }
    }
  }

  function selectValue(form, name) {
    const el = form.elements[name];
    if (!el) return undefined;
    if (el.value === '__new') {
      const box = form.elements[name + '__new'];
      return box ? String(box.value || '').trim() : '';
    }
    return String(el.value || '').trim();
  }

  function onSubmit(form) {
    const kind = form.dataset.form;
    const it = findItem(state.openId);
    const val = (name) => (form.elements[name] ? String(form.elements[name].value || '').trim() : '');
    const radio = (name) => { const r = form.querySelector('input[name="' + name + '"]:checked'); return r ? r.value : ''; };
    markBad(form, '');

    if (kind === 'count') {
      const qty = val('qty');
      if (!/^\d+$/.test(qty)) { markBad(form, 'qty'); return formError(form, 'Enter how many are on hand.'); }
      return submitWith(form, 'Saving', () => api({ action: 'count', id: it.id, qty: qty }), (res) => {
        afterChange(res.item, 'Count saved: ' + res.item.qty + ' on hand.');
      });
    }
    if (kind === 'save') {
      const f = {
        name: val('name'), cat: selectValue(form, 'cat'), sub: selectValue(form, 'sub'), loc: selectValue(form, 'loc'),
        team: selectValue(form, 'team'), cond: radio('cond'), oos: radio('oos'),
        serial: val('serial'), bought: val('bought'), qty: val('qty'), min: val('min'), notes: val('notes'),
      };
      if (!f.team) delete f.team;   // older scripts don't know teams
      if (!f.name) { markBad(form, 'name'); return formError(form, 'Give the item a name.'); }
      if (!f.cat) { markBad(form, 'cat'); return formError(form, 'Pick a category.'); }
      if (form.elements.loc && form.elements.loc.value === '__new' && !f.loc) { markBad(form, 'loc__new'); return formError(form, 'Type the new location, or pick one from the list.'); }
      if (form.elements.sub && form.elements.sub.value === '__new' && !f.sub) { markBad(form, 'sub__new'); return formError(form, 'Type the new subcategory, or pick one from the list.'); }
      const editing = state.view === 'edit' && it;
      if (editing) f.id = it.id;
      // The previous script rewrites every field, so send the photo link back to it unchanged.
      if (editing && state.meta.api < 2) f.photo = it.photo || '';
      const photo = state.pendingPhoto;
      return submitWith(form, 'Saving', async () => {
        const res = await api({ action: 'save', fields: JSON.stringify(f) });
        if (!photo) return res;
        try {
          const up = await postPhoto(res.item.id, photo, res.item.photo);
          state.localPhotos[res.item.id] = { data: photo, photo: up.item.photo };
          return { item: up.item };
        } catch (err) {
          return { item: res.item, photoError: err.message };
        }
      }, (res) => {
        state.pendingPhoto = '';
        if (f.loc && state.lists.locations.indexOf(f.loc) < 0) state.lists.locations.push(f.loc);
        if (f.sub && state.lists.subcategories.indexOf(f.sub) < 0) state.lists.subcategories.push(f.sub);
        afterChange(res.item, res.photoError ? 'Saved, but the photo didn’t upload. ' + res.photoError : (editing ? 'Saved.' : 'Added ' + res.item.id + '.'));
      });
    }
    if (kind === 'maint') {
      if (!val('issue')) { markBad(form, 'issue'); return formError(form, 'Say what the issue or the check was.'); }
      return submitWith(form, 'Saving', () => api({ action: 'maintenance', id: it.id, issue: val('issue'), taken: val('taken'), by: val('by'), next: val('next'), cost: val('cost') }),
        (res) => afterChange(res.item, 'Maintenance logged.'));
    }
    return null;
  }

  // After an admin change: back to the item page with a short note.
  function afterChange(item, message) {
    upsert(item);
    state.flashId = item.id;
    delete state.details[item.id];
    renderBoard();
    toast(message);
    state.view = 'item';
    state.openId = item.id;
    drawerBody.innerHTML = itemView(item);
    drawerPanel.scrollTop = 0;
    if (!state.pushed) setUrlItem(item.id, false);
    loadDetail(item.id, true);
  }

  // ---------------------------------------------------------------- photos (admin)

  async function uploadPhoto(id, file) {
    const it = findItem(id);
    if (!it) return;
    let data;
    try { data = await readPhoto(file); } catch (err) { toast(err.message); return; }
    const before = it.photo;
    state.busyPhoto = id;
    state.localPhotos[id] = { data: data, photo: null };   // show it right away
    refreshOpenItem();
    renderResults();
    try {
      const res = await postPhoto(id, data, before);
      state.localPhotos[id] = { data: data, photo: res.item.photo };
      upsert(res.item);
      toast('Photo saved.');
    } catch (err) {
      delete state.localPhotos[id];
      toast(err.message);
    } finally {
      state.busyPhoto = '';
      renderBoard();
      refreshOpenItem();
    }
  }

  async function removePhoto(id) {
    state.armed = '';
    state.busyPhoto = id;
    refreshOpenItem();
    try {
      const res = await api({ action: 'removephoto', id: id });
      delete state.localPhotos[id];
      upsert(res.item);
      toast('Photo removed.');
    } catch (err) {
      toast(err.message);
    } finally {
      state.busyPhoto = '';
      renderBoard();
      refreshOpenItem();
    }
  }

  async function pickFormPhoto(input) {
    const file = input.files && input.files[0];
    if (!file) return;
    const drop = input.closest('[data-drop]');
    try {
      state.pendingPhoto = await readPhoto(file);
      if (drop) {
        drop.classList.add('has-photo');
        drop.querySelector('.inv-drop-preview').innerHTML = '<img src="' + state.pendingPhoto + '" alt="">';
        drop.querySelector('.inv-drop-text b').textContent = 'Photo ready';
        drop.querySelector('.inv-drop-text span').textContent = 'It uploads when you save. Click to pick a different one.';
      }
    } catch (err) {
      toast(err.message);
    }
  }

  // ---------------------------------------------------------------- test data (admin)

  async function testData(mode) {
    state.armed = '';
    const btns = document.querySelectorAll('[data-act^="test-"]');
    btns.forEach((b) => { b.disabled = true; });
    const busy = document.querySelector('[data-act="test-' + mode + '"]');
    if (busy) busy.innerHTML = '<span class="inv-spin" aria-hidden="true"></span>' + (mode === 'remove' ? 'Removing' : mode === 'reset' ? 'Resetting' : 'Loading');
    try {
      const res = await api(mode === 'add' ? { action: 'testdata' } : { action: 'testdata', mode: mode }, 120000);
      toast(mode === 'remove'
        ? 'Removed ' + res.removed.items + ' test items and their maintenance.'
        : mode === 'reset'
          ? 'Test data reset: ' + res.added.items + ' fresh test items.'
          : 'Loaded ' + res.added.items + ' test items with ' + res.added.maintenance + ' maintenance entries.');
      if (state.view) closeDrawer();
      state.details = {};
      await load();
    } catch (err) {
      renderAdminFoot();
      toast(err.message);
    }
  }

  function arm(act, onSecond, render) {
    if (state.armed === act) { state.armed = ''; onSecond(); return; }
    state.armed = act;
    render();
    clearTimeout(arm.t);
    arm.t = setTimeout(() => { if (state.armed === act) { state.armed = ''; render(); } }, 5000);
  }

  // ---------------------------------------------------------------- labels (?labels=1)

  function loadQr() {
    return new Promise((resolve, reject) => {
      if (window.QRCode) { resolve(); return; }
      const s = document.createElement('script');
      s.src = QR_LIB;
      s.onload = () => (window.QRCode ? resolve() : reject(new Error('QR library missing')));
      s.onerror = () => reject(new Error('Could not load the QR library'));
      document.head.appendChild(s);
    });
  }

  function renderLabels() {
    document.title = 'Item labels | Project Teams Ops Hub';
    document.documentElement.classList.add('inv-qrlabels-mode');
    if (!state.loaded) {
      app.innerHTML = '<section class="inv-setup"><h1 class="inv-title">Item labels</h1><p>' +
        esc(state.error || 'Loading…') + '</p>' + (state.error ? '<button type="button" class="inv-btn is-primary" data-act="reload">Try again</button>' : '') + '</section>';
      return;
    }
    const items = state.items.filter((it) => !it.consumable && look(it).key !== 'retired').sort(byId);
    app.innerHTML =
      '<section class="inv-top inv-noprint"><div class="inv-head"><h1 class="inv-title">Item labels</h1>' +
        '<p class="inv-reading">Print, cut, and stick one on each tool. Scanning a label opens that tool’s page: where it lives, what shape it’s in, and its maintenance.</p></div>' +
        '<div class="inv-head-tools"><button type="button" class="inv-btn is-primary" data-act="print">Print labels</button></div></section>' +
      (items.length
        ? '<div class="inv-qrlabels">' + items.map((it) =>
            '<div class="inv-qrlabel"><div class="inv-qrlabel-code" data-url="' + esc(APP_URL + '?item=' + encodeURIComponent(it.id)) + '"></div>' +
            '<div class="inv-qrlabel-text"><div class="inv-assettag">' + esc(it.id) + '</div><div class="inv-qrlabel-name">' + esc(it.name) + '</div>' +
            (teamOf(it) !== SHARED ? '<div class="inv-qrlabel-team">' + esc(teamOf(it)) + '</div>' : '') +
            '<div class="inv-qrlabel-hint">Scan for this tool’s details</div></div></div>').join('') + '</div>'
        : '<div class="inv-empty"><h2 class="inv-empty-title">No tools to label yet</h2></div>');
    loadQr().then(() => {
      document.querySelectorAll('.inv-qrlabel-code').forEach((el) => {
        try { new window.QRCode(el, { text: el.dataset.url, width: 360, height: 360, correctLevel: window.QRCode.CorrectLevel.M }); }
        catch (e) { el.innerHTML = '<span class="inv-qr-fallback">' + esc(el.dataset.url) + '</span>'; }
      });
    }, () => {
      document.querySelectorAll('.inv-qrlabel-code').forEach((el) => { el.innerHTML = '<span class="inv-qr-fallback">' + esc(el.dataset.url) + '</span>'; });
    });
  }

  // ---------------------------------------------------------------- events

  app.addEventListener('click', (e) => {
    const chip = e.target.closest('[data-filter]');
    if (chip) { state.filter = chip.dataset.filter; renderChips(); renderPicks(); renderResults(); return; }
    const card = e.target.closest('.inv-slot[data-id]');
    if (card) { openItem(card.dataset.id, card); return; }
    const act = e.target.closest('[data-act]');
    if (!act) return;
    const a = act.dataset.act;
    if (a === 'reload') load();
    else if (a === 'add') { state.returnFocus = act; openForm('add', null); }
    else if (a === 'clear') {
      state.filter = 'all'; state.q = ''; state.team = ''; state.loc = ''; state.cat = '';
      const q = document.getElementById('inv-q');
      if (q) q.value = '';
      renderChips(); renderPicks(); renderResults();
    }
    else if (a === 'test-add') testData('add');
    else if (a === 'test-remove') arm('test-remove', () => testData('remove'), renderAdminFoot);
    else if (a === 'test-reset') arm('test-reset', () => testData('reset'), renderAdminFoot);
    else if (a === 'print') window.print();
  });

  app.addEventListener('change', (e) => {
    const pick = e.target.closest('[data-pick]');
    if (pick && e.target.tagName === 'SELECT') {
      state[pick.dataset.pick] = e.target.value;
      renderPicks();
      renderResults();
    }
  });

  // A photo that will not load (removed, or not shared) falls back to the silhouette.
  app.addEventListener('error', (e) => {
    const img = e.target;
    if (img && img.classList && img.classList.contains('inv-photo')) {
      const b = img.closest('.inv-bed');
      if (b) b.classList.remove('has-photo');
      img.remove();
    }
  }, true);
  drawer.addEventListener('error', (e) => {
    const img = e.target;
    if (img && img.classList && img.classList.contains('inv-photo')) {
      const b = img.closest('.inv-bed');
      if (b) b.classList.remove('has-photo');
      img.remove();
    }
  }, true);

  function photoLoaded(e) {
    const img = e.target;
    if (img && img.classList && img.classList.contains('inv-photo')) {
      img.classList.add('is-in');
      loadedPhotos.add(img.getAttribute('src'));
    }
  }
  app.addEventListener('load', photoLoaded, true);
  drawer.addEventListener('load', photoLoaded, true);

  // On a computer, start fetching an item's maintenance log while the pointer rests on it.
  let hoverTimer = 0;
  app.addEventListener('pointerover', (e) => {
    if (!CAN_HOVER) return;
    const s = e.target.closest('.inv-slot[data-id]');
    if (!s) return;
    clearTimeout(hoverTimer);
    hoverTimer = setTimeout(() => {
      const d = state.details[s.dataset.id];
      const busy = Object.keys(state.details).some((k) => state.details[k].loading);
      if (!busy && (!d || Date.now() - d.at > 60000)) loadDetail(s.dataset.id);
    }, 400);
  });
  app.addEventListener('pointerout', (e) => { if (e.target.closest('.inv-slot')) clearTimeout(hoverTimer); });

  drawer.addEventListener('click', (e) => {
    if (e.target.closest('[data-close]')) { closeDrawer(); return; }
    const step = e.target.closest('[data-step]');
    if (step) {
      const input = step.form && step.form.elements.qty;
      if (input) input.value = Math.max(0, (parseInt(input.value, 10) || 0) + Number(step.dataset.step));
      return;
    }
    const act = e.target.closest('[data-act]');
    if (!act) return;
    const it = findItem(state.openId);
    const a = act.dataset.act;
    if (a === 'edit' && it) openForm('edit', it);
    else if (a === 'maint' && it) openForm('maint', it);
    else if (a === 'cancel') { if (it) { state.view = 'item'; drawerBody.innerHTML = itemView(it); drawerPanel.scrollTop = 0; loadDetail(it.id); } else closeDrawer(); }
    else if (a === 'photo-remove' && it) arm('photo', () => removePhoto(it.id), refreshOpenItem);
  });

  drawer.addEventListener('submit', (e) => {
    e.preventDefault();
    onSubmit(e.target);
  });

  // Typing again clears an earlier error, so stale messages don't linger.
  drawer.addEventListener('input', (e) => {
    const form = e.target.form;
    if (!form) return;
    if (e.target.getAttribute('aria-invalid')) e.target.removeAttribute('aria-invalid');
    const err = form.querySelector('.inv-form-err');
    if (err && !err.hidden && !form.classList.contains('is-busy')) formError(form, '');
  });

  drawer.addEventListener('change', (e) => {
    const t = e.target;
    const form = t.form;
    if (t.tagName === 'SELECT' && form && form.elements[t.name + '__new']) {
      const box = form.elements[t.name + '__new'];
      box.hidden = t.value !== '__new';
      if (!box.hidden) box.focus();
    }
    if (t.name === 'cat' && form && form.dataset.form === 'save') {
      const qty = form.querySelector('.inv-qty-fields');
      if (qty) qty.classList.toggle('is-hidden', t.value.toLowerCase() !== 'consumable');
    }
    if (t.matches && t.matches('[data-photo-for]')) { const f = t.files && t.files[0]; if (f) uploadPhoto(t.dataset.photoFor, f); return; }
    if (t.matches && t.matches('[data-photo-pick]')) pickFormPhoto(t);
  });

  // Drag a photo onto the picker in the add/edit form.
  drawer.addEventListener('dragover', (e) => {
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (!drop) return;
    e.preventDefault();
    drop.classList.add('is-over');
  });
  drawer.addEventListener('dragleave', (e) => {
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (drop && !drop.contains(e.relatedTarget)) drop.classList.remove('is-over');
  });
  drawer.addEventListener('drop', (e) => {
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (!drop) return;
    e.preventDefault();
    drop.classList.remove('is-over');
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    const input = drop.querySelector('input[type="file"]');
    if (!file || !input) return;
    try { input.files = e.dataTransfer.files; } catch (err) { /* read-only in some browsers */ }
    if (!input.files || !input.files[0]) { readPhoto(file).then((d) => { state.pendingPhoto = d; drop.querySelector('.inv-drop-preview').innerHTML = '<img src="' + d + '" alt="">'; }, (er) => toast(er.message)); return; }
    pickFormPhoto(input);
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && state.view) { e.preventDefault(); closeDrawer(); return; }
    if (e.key === 'Tab' && state.view) {
      const f = Array.prototype.filter.call(
        drawerPanel.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled])'),
        (x) => x.offsetParent !== null || x.classList.contains('inv-file'));
      if (!f.length) return;
      const first = f[0], last = f[f.length - 1];
      if (e.shiftKey && (document.activeElement === first || document.activeElement === drawerPanel)) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      return;
    }
    if (e.key === '/' && !state.view && !/^(INPUT|TEXTAREA|SELECT)$/.test((document.activeElement || {}).tagName || '')) {
      const q = document.getElementById('inv-q');
      if (q) { e.preventDefault(); q.focus(); q.select(); }
    }
  });

  // Back and forward buttons open and close items.
  window.addEventListener('popstate', (e) => {
    const id = e.state && e.state.invItem;
    if (id && findItem(id)) { if (state.openId !== id) openItem(id, null, { fromHistory: true }); }
    else if (state.view) closeDrawer({ fromHistory: true });
  });

  // Coming back to the tab after a while, or leaving it open on a shop screen: refresh.
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && API && !state.loading && Date.now() - state.savedAt > 60000) load();
  });
  setInterval(() => {
    if (!document.hidden && API && !state.loading && !state.view && Date.now() - state.savedAt > 170000) load();
  }, 60000);

  // ---------------------------------------------------------------- start

  if (window.self !== window.top) {
    // Inside the hub's slide-over: "Back to the hub" makes no sense there.
    const back = document.getElementById('inv-head-back');
    if (back) back.hidden = true;
  }

  if (!API) {
    renderNotConnected();
    return;
  }

  // Left behind by the version with check outs: the NetID it remembered, and its copy of the list.
  store('ops-hub-netid', null);
  store('ops-hub-inventory-v2', null);

  const saved = readCache();
  if (saved && Array.isArray(saved.items) && Date.now() - (saved.savedAt || 0) < 7 * 86400000) {
    state.items = saved.items;
    state.lists = Object.assign(state.lists, saved.lists || {});
    state.meta = Object.assign({ testItems: 0, api: 0 }, saved.meta || {}, { today: localToday() });
    state.savedAt = saved.savedAt || 0;
    state.loaded = true;
  }

  if (LABELS) renderLabels(); else { renderBoard(); handleStart(false); }
  load();
})();
