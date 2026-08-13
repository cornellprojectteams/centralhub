/**
 * Equipment & Inventory registry - the "Ops registry" module (the ?registry=... pages
 * and ?module=registry-dash dashboard). Split out of 02_notify_on_submit.gs for
 * maintainability. Apps Script shares ONE global scope across all .gs files, so doGet
 * (in 02) routes straight into these functions with no import needed. Equipment and
 * Inventory run the same code, branched by regFields_(which).
 *
 * SPEED. Sheet round-trips dominate here, so:
 *  - Catalog and item pages paint a named loading shell first, then fetch
 *    the body so the wait is never a blank white screen.
 *  - Spreadsheet handle and tab rows are memoized per execution (REG_MEMO).
 *  - Tab rows and the rendered card HTML are cached in CacheService (~5 min),
 *    dropped on every write. A warm catalog load skips the spreadsheet.
 *  - Admin row maps load after first paint so the grid is not blocked on JSON.
 *  - Drive photos render as thumbnails (sz=w400), not full files.
 *  - After the page paints, the client warms the other tab (Inventory <-> Equipment)
 *    so the toggle is a cache hit instead of another cold read.
 *
 * Reuses engine helpers that live in 02/04: swissShell_, portalStyles_, tpDashStyles_,
 * escapeHtml_, norm_, fmtShort_, phrase_, registrySs_, icTeamNames_, tpSaveUpload_,
 * tpViewUrl_, extractFileIds_. Any change here still needs a NEW deployment version.
 */

// Read-only registry table (Equipment or Inventory) for the hub. Phase 1 CMMS.
// CMMS dashboard: one read-only glance at the shop - open work, stock to reorder,
// maintenance coming due, and the live action-item feed. Reads the registry spreadsheet.
function registryDashboardPage_(embedded) {
  var eq = readTab_('Equipment');
  var invd = readTab_('Inventory');
  var pm = readTab_('Maintenance');
  var act = readTab_('Action items');

  var col = function (H, name) { for (var i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(name)) return i; } return -1; };
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var DAY = 86400000;
  var toDate = function (v) { if (v instanceof Date) return v; if (v === '' || v == null) return null; var d = new Date(v); return isNaN(d.getTime()) ? null : d; };
  var startOf = function (d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
  var truthy = function (v) { if (v === true) return true; var s = String(v == null ? '' : v).trim().toLowerCase(); return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'active' || s === 'x'; };
  var dueLabel = function (d, diff) {
    if (!d || diff === null) return { txt: 'No date set', cls: 'is-ok' };
    if (diff < 0) return { txt: 'Overdue ' + (-diff) + 'd', cls: 'is-over' };
    if (diff === 0) return { txt: 'Due today', cls: 'is-over' };
    if (diff <= 7) return { txt: 'Due in ' + diff + 'd', cls: 'is-soon' };
    return { txt: fmtShort_(d), cls: 'is-ok' };
  };
  var statusPill = function (s) {
    var t = String(s || '').toLowerCase();
    if (t.indexOf('block') >= 0 || t.indexOf('urgent') >= 0) return 'pill-out';
    if (t.indexOf('progress') >= 0 || t.indexOf('order') >= 0 || t.indexOf('wait') >= 0) return 'pill-prog';
    return 'pill-open';
  };

  // Inventory: healthy vs low vs out, and a reorder list.
  var iOn = col(invd.headers, 'On hand'), iRe = col(invd.headers, 'Reorder point'),
      iItem = col(invd.headers, 'Item'), iUnit = col(invd.headers, 'Unit');
  var invCounted = 0, outCount = 0, lowCount = 0, healthy = 0, reorder = [];
  invd.rows.forEach(function (r) {
    if (iOn < 0 || !isNum_(r[iOn])) return;
    var on = Number(r[iOn]);
    var re = (iRe >= 0 && isNum_(r[iRe])) ? Number(r[iRe]) : 0;
    var name = iItem >= 0 ? String(r[iItem]).trim() : '';
    var unit = iUnit >= 0 ? String(r[iUnit]).trim() : '';
    invCounted++;
    var out = on <= 0;
    var low = on > 0 && re > 0 && on <= re;
    if (out) outCount++; else if (low) lowCount++; else healthy++;
    if (out || low) reorder.push({ name: name, on: on, re: re, unit: unit, out: out });
  });
  reorder.sort(function (a, b) { if (a.out !== b.out) return a.out ? -1 : 1; return (b.re - b.on) - (a.re - a.on); });

  // Maintenance: active tasks, soonest due first.
  var mAct = col(pm.headers, 'Active'), mNext = col(pm.headers, 'Next due'),
      mTask = col(pm.headers, 'Task'), mArea = col(pm.headers, 'Equipment / area');
  var pmDue = 0, pmList = [];
  pm.rows.forEach(function (r) {
    var active = mAct < 0 ? true : truthy(r[mAct]);
    if (!active) return;
    var next = mNext >= 0 ? toDate(r[mNext]) : null;
    var diff = next ? Math.round((startOf(next) - today) / DAY) : null;
    if (diff !== null && diff <= 0) pmDue++;
    pmList.push({ task: mTask >= 0 ? String(r[mTask]).trim() : '', area: mArea >= 0 ? String(r[mArea]).trim() : '', next: next, diff: diff });
  });
  pmList.sort(function (a, b) { var av = a.diff === null ? 1e9 : a.diff, bv = b.diff === null ? 1e9 : b.diff; return av - bv; });

  // Action items: everything not yet done, soonest due first.
  var aStatus = col(act.headers, 'Status'), aDone = col(act.headers, 'Completed'), aTitle = col(act.headers, 'Title'),
      aLoc = col(act.headers, 'Location'), aEquip = col(act.headers, 'Equipment / part'), aWho = col(act.headers, 'Assigned to'),
      aDue = col(act.headers, 'Due'), aCreated = col(act.headers, 'Created');
  var openActions = [];
  act.rows.forEach(function (r) {
    var st = aStatus >= 0 ? String(r[aStatus] || '').trim() : '';
    var stl = st.toLowerCase();
    var done = (aDone >= 0 && String(r[aDone]).trim() !== '') || stl === 'done' || stl === 'completed' || stl === 'closed';
    if (done) return;
    openActions.push({
      title: aTitle >= 0 ? String(r[aTitle]).trim() : '',
      loc: aLoc >= 0 ? String(r[aLoc]).trim() : '',
      equip: aEquip >= 0 ? String(r[aEquip]).trim() : '',
      who: aWho >= 0 ? String(r[aWho]).trim() : '',
      status: st || 'Open',
      due: aDue >= 0 ? toDate(r[aDue]) : null,
      created: aCreated >= 0 ? toDate(r[aCreated]) : null
    });
  });
  openActions.sort(function (a, b) {
    var ad = a.due ? startOf(a.due).getTime() : 8.64e15, bd = b.due ? startOf(b.due).getTime() : 8.64e15;
    if (ad !== bd) return ad - bd;
    return (b.created ? b.created.getTime() : 0) - (a.created ? a.created.getTime() : 0);
  });

  // Equipment out on loan, overdue first.
  var eName = col(eq.headers, 'Name'), eCoTo = col(eq.headers, 'Checked out to'), eDue = col(eq.headers, 'Due back');
  var checkedOut = 0, loans = [];
  eq.rows.forEach(function (r) {
    var who = eCoTo >= 0 ? String(r[eCoTo] || '').trim() : '';
    if (!who) return;
    checkedOut++;
    var due = eDue >= 0 ? toDate(r[eDue]) : null;
    var diff = due ? Math.round((startOf(due) - today) / DAY) : null;
    loans.push({ name: eName >= 0 ? String(r[eName]).trim() : '', who: who, due: due, diff: diff });
  });
  loans.sort(function (a, b) { var av = a.diff === null ? 1e9 : a.diff, bv = b.diff === null ? 1e9 : b.diff; return av - bv; });
  var overdueLoans = loans.filter(function (x) { return x.diff !== null && x.diff < 0; }).length;

  var assets = eq.rows.length, invItems = invd.rows.length, reorderCount = outCount + lowCount, openCount = openActions.length;

  var cmms = '<style>'
    + '.cmms-list{display:flex;flex-direction:column}'
    + '.cmms-row{display:grid;grid-template-columns:1fr auto;gap:2px 14px;align-items:center;padding:13px 0;border-bottom:1px solid #f0efe9}'
    + '.cmms-row:first-child{padding-top:2px}.cmms-row:last-child{border-bottom:0;padding-bottom:2px}'
    + '.cmms-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14px;font-weight:700;color:#26231f;line-height:1.3}'
    + '.cmms-sub{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;color:#8a857c;margin-top:2px}'
    + '.cmms-when{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:800;white-space:nowrap}'
    + '.cmms-when-wrap{display:flex;align-items:center;gap:10px;justify-self:end;white-space:nowrap}'
    + '.cmms-when.is-over{color:#b31b1b}.cmms-when.is-soon{color:#b06a00}.cmms-when.is-ok{color:#8a857c}'
    + '.pill{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;border-radius:99px;padding:3px 9px;white-space:nowrap}'
    + '.pill-open{color:#b06a00;background:#fdf3df;border:1px solid #f0dca6}'
    + '.pill-out{color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0}'
    + '.pill-prog{color:#2563c9;background:#eaf1fe;border:1px solid #cfe0fb}'
    + '@media(max-width:520px){.cmms-row{grid-template-columns:1fr}.cmms-when-wrap{justify-self:start;margin-top:6px}}'
    + '</style>';

  var inner = tpDashStyles_() + cmms + '<div class="dash">';

  inner += '<div class="dash-hero"><div class="dash-hero-glow"></div>'
    + '<div class="dash-kicker">Ops registry</div>'
    + '<h1 class="dash-h1">Equipment &amp; Inventory</h1>'
    + '<div class="dash-hero-row">'
    +   '<div><div class="dash-hero-num">' + openCount + '</div><div class="dash-hero-lbl">open action item' + (openCount === 1 ? '' : 's') + ' to fix, restock, or maintain</div></div>'
    +   '<div class="dash-hero-mini">'
    +     '<div><b>' + reorderCount + '</b>to reorder</div>'
    +     '<div><b>' + pmDue + '</b>maintenance due</div>'
    +     '<div><b>' + checkedOut + '</b>out on loan</div>'
    +   '</div>'
    + '</div></div>';

  var tile = function (n, lbl, c) { return '<div class="dash-tile" style="--c:' + c + '"><div class="dash-tile-num">' + n + '</div><div class="dash-tile-lbl">' + lbl + '</div></div>'; };
  inner += '<div class="dash-tiles">'
    + tile(assets, 'Assets tracked', '#2563c9')
    + tile(invItems, 'Inventory items', '#0d9488')
    + tile(reorderCount, 'Low / out of stock', '#b31b1b')
    + tile(pmDue, 'Maintenance due', '#e08a1e')
    + '</div>';

  if (!assets && !invItems && !pm.rows.length && !act.rows.length) {
    inner += '<div class="dash-empty">The registry is empty. Run <b>Set up / update tabs</b> in the registry sheet, then add equipment and inventory.</div></div>';
    return swissShell_(inner, 'Equipment & inventory', true, embedded);
  }

  var ring;
  if (!invCounted) {
    ring = '<div class="dash-card"><div class="dash-card-h">Inventory in stock</div><div class="dash-ringwrap"><div class="dash-ring-cap">No inventory counts yet.</div></div></div>';
  } else {
    var inStock = Math.round(healthy / invCounted * 100);
    var offset = 100 - inStock;
    ring = '<div class="dash-card"><div class="dash-card-h">Inventory in stock</div><div class="dash-ringwrap">'
      + '<svg class="dash-ring" viewBox="0 0 42 42" aria-hidden="true">'
      +   '<defs><linearGradient id="dashgrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8f1515"/><stop offset="0.6" stop-color="#d62b2b"/><stop offset="1" stop-color="#f0c050"/></linearGradient></defs>'
      +   '<circle class="dash-ring-bg" cx="21" cy="21" r="15.915"/>'
      +   '<circle class="dash-ring-fg" cx="21" cy="21" r="15.915" stroke-dasharray="100" stroke-dashoffset="100" data-off="' + offset + '" transform="rotate(-90 21 21)"/>'
      +   '<text class="dash-ring-num" x="21" y="20.4" text-anchor="middle">' + inStock + '%</text>'
      +   '<text class="dash-ring-sub" x="21" y="26" text-anchor="middle">IN STOCK</text>'
      + '</svg>'
      + '<div class="dash-ring-cap">' + healthy + ' of ' + invCounted + ' items healthy</div>'
      + '</div></div>';
  }

  var bars = '<div class="dash-card"><div class="dash-card-h">Stock to reorder</div>';
  if (!reorder.length) {
    bars += '<div class="dash-gal-who">Everything is above its reorder point.</div>';
  } else {
    bars += '<div class="dash-bars">';
    reorder.slice(0, 7).forEach(function (it) {
      var frac = it.re > 0 ? Math.max(0, Math.min(1, it.on / it.re)) : (it.on > 0 ? 1 : 0);
      var val = it.out ? '<span class="pill pill-out">Out</span>' : escapeHtml_(it.on + (it.unit ? ' ' + it.unit : '') + ' / ' + it.re);
      bars += '<div><div class="dash-bar-row"><span class="dash-bar-name">' + escapeHtml_(it.name || '(unnamed)') + '</span><span class="dash-bar-val">' + val + '</span></div>'
        + '<div class="dash-bar-track"><div class="dash-bar-fill" style="--w:' + frac.toFixed(3) + '"></div></div></div>';
    });
    bars += '</div>';
  }
  bars += '</div>';
  inner += '<div class="dash-grid">' + ring + bars + '</div>';

  if (pmList.length) {
    inner += '<div class="dash-sec">Maintenance</div><div class="dash-card"><div class="cmms-list">';
    pmList.slice(0, 8).forEach(function (m) {
      var w = dueLabel(m.next, m.diff);
      inner += '<div class="cmms-row">'
        + '<div><div class="cmms-title">' + escapeHtml_(m.task || '(untitled)') + '</div>'
        + (m.area ? '<div class="cmms-sub">' + escapeHtml_(m.area) + '</div>' : '') + '</div>'
        + '<div class="cmms-when ' + w.cls + '">' + escapeHtml_(w.txt) + '</div>'
        + '</div>';
    });
    inner += '</div></div>';
  }

  if (loans.length) {
    inner += '<div class="dash-sec">Out on loan' + (overdueLoans ? ' &middot; ' + overdueLoans + ' overdue' : '') + '</div><div class="dash-card"><div class="cmms-list">';
    loans.slice(0, 8).forEach(function (l) {
      var w = dueLabel(l.due, l.diff);
      inner += '<div class="cmms-row">'
        + '<div><div class="cmms-title">' + escapeHtml_(l.name || '(unnamed)') + '</div>'
        + '<div class="cmms-sub">' + escapeHtml_(l.who) + '</div></div>'
        + '<div class="cmms-when ' + w.cls + '">' + escapeHtml_(l.due ? w.txt : 'No due date') + '</div>'
        + '</div>';
    });
    inner += '</div></div>';
  }

  inner += '<div class="dash-sec">Open action items</div>';
  if (!openActions.length) {
    inner += '<div class="dash-empty">No open action items. The shop is caught up.</div>';
  } else {
    inner += '<div class="dash-card"><div class="cmms-list">';
    openActions.slice(0, 12).forEach(function (a) {
      var subBits = [];
      if (a.loc) subBits.push(a.loc);
      if (a.equip) subBits.push(a.equip);
      if (a.who) subBits.push(a.who);
      var due = a.due ? dueLabel(a.due, Math.round((startOf(a.due) - today) / DAY)) : null;
      inner += '<div class="cmms-row">'
        + '<div><div class="cmms-title">' + escapeHtml_(a.title || '(untitled)') + '</div>'
        + (subBits.length ? '<div class="cmms-sub">' + escapeHtml_(subBits.join(' · ')) + '</div>' : '') + '</div>'
        + '<div class="cmms-when-wrap"><span class="pill ' + statusPill(a.status) + '">' + escapeHtml_(a.status) + '</span>'
        + (due ? '<span class="cmms-when ' + due.cls + '">' + escapeHtml_(due.txt) + '</span>' : '') + '</div>'
        + '</div>';
    });
    inner += '</div></div>';
    if (openActions.length > 12) inner += '<div class="dash-gal-who" style="margin-top:10px">Showing 12 of ' + openActions.length + ' open items. Open the action items sheet for the rest.</div>';
  }

  inner += '</div>';
  inner += '<script>requestAnimationFrame(function(){var c=document.querySelector(".dash-ring-fg");if(c)c.style.strokeDashoffset=c.getAttribute("data-off");});</script>';
  return swissShell_(inner, 'Equipment & inventory', true, embedded);
}

// Canonical, sortable categories and LiPo states, shared by the form and filters.
var REG_CATEGORIES = ['Machine tools', 'Power & hand tools', 'Cutting tooling', 'Measurement & inspection', 'Electronics & test', '3D printing & prototyping', 'Materials', 'Fasteners & hardware', 'Abrasives', 'Adhesives & chemicals', 'Filters', 'PPE', 'First aid', 'Batteries', 'Facility & finishing', 'Other'];
var REG_BATTERY_STATES = ['Charged', 'Storage', 'Needs charge', 'Retire'];

var REG_MEMO = {};
function regCacheSvc_() { try { return CacheService.getScriptCache(); } catch (e) { return null; } }
function regSs_() {
  if (REG_MEMO.ss) return REG_MEMO.ss;
  REG_MEMO.ss = registrySs_();
  return REG_MEMO.ss;
}
function regWebUrl_() {
  if (REG_MEMO.url != null) return REG_MEMO.url;
  var u = '';
  try { u = ScriptApp.getService().getUrl(); } catch (e) { u = CONFIG.webAppUrl || ''; }
  REG_MEMO.url = u;
  return u;
}
function regInvalidate_() {
  REG_MEMO = {};
  var c = regCacheSvc_();
  if (!c) return;
  var keys = [];
  ['reg_t_Inventory', 'reg_t_Equipment',
    'reg_c3_inventory_1', 'reg_c3_inventory_0', 'reg_c3_equipment_1', 'reg_c3_equipment_0',
    'reg_c3_inventory_1_m', 'reg_c3_inventory_0_m', 'reg_c3_equipment_1_m', 'reg_c3_equipment_0_m'
  ].forEach(function (p) {
    keys.push(p, p + '_n');
    for (var i = 0; i < 6; i++) keys.push(p + '_' + i);
  });
  try { c.removeAll(keys); } catch (e) { /* cache is best-effort */ }
}

// Field spec drives the add/edit forms and identifies each tab's natural key.
function regFields_(which) {
  if (String(which).toLowerCase() === 'inventory') {
    return { tab: 'Inventory', key: 'Item', idPrefix: '', which: 'inventory', noun: 'item',
      fields: [
        { h: 'Item', label: 'Item', req: true },
        { h: 'Category', label: 'Category', type: 'select', opts: 'category' },
        { h: 'Owning team', label: 'Owner (team or Program)', type: 'select', opts: 'team' }, { h: 'Location', label: 'Location' },
        { h: 'On hand', label: 'On hand', type: 'number' }, { h: 'Unit', label: 'Unit' },
        { h: 'Supplier', label: 'Supplier', adv: true }, { h: 'Product link', label: 'Product link', type: 'url', adv: true },
        { h: 'Battery state', label: 'Battery state', type: 'select', opts: 'battery', group: 'battery', adv: true },
        { h: 'Battery spec', label: 'Battery spec (e.g. 6S 5000mAh)', group: 'battery', adv: true },
        { h: 'Image', label: 'Image', type: 'image' } ] };
  }
  return { tab: 'Equipment', key: 'Asset ID', idPrefix: 'EQ-', which: 'equipment', noun: 'equipment',
    fields: [
      { h: 'Asset ID', label: 'Asset ID', auto: true }, { h: 'Name', label: 'Name', req: true },
      { h: 'Category', label: 'Category', type: 'select', opts: 'category' },
      { h: 'Owning team', label: 'Owner (team or Program)', type: 'select', opts: 'team' }, { h: 'Location', label: 'Location' },
      { h: 'Owner', label: 'Owner (person)', adv: true },
      { h: 'Status', label: 'Status', adv: true }, { h: 'Installed', label: 'Installed', adv: true }, { h: 'Notes', label: 'Notes', adv: true },
      { h: 'Image', label: 'Image', type: 'image' } ] };
}

function readTabRows_(name) {
  var memoKey = 'tab_' + name;
  if (REG_MEMO[memoKey]) return REG_MEMO[memoKey];
  var ck = 'reg_t_' + name, hit = null;
  try { hit = regCacheGet_(ck); } catch (e) { hit = null; }
  if (hit) {
    try {
      var parsed = JSON.parse(hit);
      REG_MEMO[memoKey] = parsed;
      return parsed;
    } catch (e) { /* fall through to a live read */ }
  }
  const sh = regSs_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  const v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [] };
  const headers = v[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (cell) { return String(cell).trim() === ''; })) continue;
    rows.push({
      row: i + 1,
      cells: v[i].map(function (cell) {
        if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return cell;
      })
    });
  }
  var out = { headers: headers, rows: rows };
  REG_MEMO[memoKey] = out;
  try { regCachePut_(ck, JSON.stringify(out), 300); } catch (e) { /* cache is best-effort */ }
  return out;
}

function regRawVal_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (v === 0) return '0';
  return String(v == null ? '' : v);
}

// Drive file -> thumbnail (fast). Only real Drive / googleusercontent URLs (or a bare
// file id) are rewritten. Wikimedia / Amazon / product URLs must pass through as-is:
// extractFileIds_ would otherwise treat a long filename as a Drive id and the <img>
// would 404 (this is why Inventory photos vanished while Equipment still showed).
function regDriveId_(img) {
  var s = String(img || '').trim();
  var im = /^\s*=?\s*IMAGE\s*\(\s*"([^"]+)"/i.exec(s);
  if (im) s = im[1];
  if (!s) return '';
  var m = /drive\.google\.com\/file\/d\/([\w-]{25,44})/i.exec(s);
  if (m) return m[1];
  m = /lh3\.googleusercontent\.com\/d\/([\w-]{25,44})/i.exec(s);
  if (m) return m[1];
  m = /[?&]id=([\w-]{25,44})/.exec(s);
  if (m && /(?:drive|docs)\.google\.com/i.test(s)) return m[1];
  if (/^[\w-]{25,44}$/.test(s)) return s;
  return '';
}

function regThumb_(img, sz) {
  var s = String(img || '').trim();
  if (!s) return '';
  var im = /^\s*=?\s*IMAGE\s*\(\s*"([^"]+)"/i.exec(s);
  if (im) s = im[1];
  if (s.indexOf('data:') === 0) return s;
  var id = regDriveId_(s);
  if (id) return 'https://lh3.googleusercontent.com/d/' + encodeURIComponent(id) + '=w' + (sz || 400);
  if (/^https?:\/\//i.test(s)) return s;
  return '';
}

function regImgTag_(raw, sz, mono) {
  var src = regThumb_(raw, sz);
  if (!src) return '<span class="reg-card-mono" aria-hidden="true">' + escapeHtml_(mono) + '</span>';
  var id = regDriveId_(raw);
  var fb = id ? 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(id) : '';
  if (fb && fb === src) fb = '';
  return '<img class="reg-img" src="' + escapeHtml_(src) + '"'
    + (fb ? ' data-fb="' + escapeHtml_(fb) + '"' : '')
    + ' data-mono="' + escapeHtml_(mono) + '"'
    + ' width="' + (sz || 400) + '" height="' + (sz || 400) + '" loading="lazy" decoding="async" alt="" referrerpolicy="no-referrer"'
    + ' onload="this.classList.add(\'is-in\')" onerror="regImgFb(this)">';
}

function regMonogram_(name) {
  var parts = String(name || '').trim().split(/\s+/).filter(function (p) { return p; });
  var a = parts[0] ? parts[0].charAt(0) : '';
  var b = parts[1] ? parts[1].charAt(0) : (parts[0] && parts[0].length > 1 ? parts[0].charAt(1) : '');
  return (a + b).toUpperCase() || '·';
}

function regKeepHeaders_(spec) {
  const keepH = {};
  spec.fields.forEach(function (f) { keepH[f.h] = 1; });
  ['On hand', 'Reorder point', 'Checked out to', 'Due back', 'Out since', spec.key].forEach(function (h) { keepH[h] = 1; });
  return keepH;
}

function regSlimMap_(spec) {
  const data = readTabRows_(spec.tab);
  const H = data.headers, rows = data.rows;
  const idx = function (n) { for (let i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(n)) return i; } return -1; };
  const keepH = regKeepHeaders_(spec);
  const mapParts = [];
  rows.forEach(function (rr) {
    const r = rr.cells, obj = {};
    Object.keys(keepH).forEach(function (h) {
      var i = idx(h);
      obj[h] = i >= 0 ? regRawVal_(r[i]) : '';
    });
    mapParts.push(JSON.stringify(String(rr.row)) + ':' + JSON.stringify(obj));
  });
  return '{' + mapParts.join(',') + '}';
}

function regCacheGet_(key) {
  var c = regCacheSvc_();
  if (!c) return null;
  try {
    var nRaw = c.get(key + '_n');
    if (nRaw) {
      var n = parseInt(nRaw, 10);
      if (!(n > 0 && n <= 6)) return null;
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(key + '_' + i);
      var bag = c.getAll(keys) || {};
      var parts = [];
      for (var i = 0; i < n; i++) {
        if (bag[keys[i]] == null) return null;
        parts.push(bag[keys[i]]);
      }
      return parts.join('');
    }
    return c.get(key);
  } catch (e) { return null; }
}

function regCachePut_(key, value, sec) {
  var c = regCacheSvc_();
  if (!c || value == null) return;
  sec = sec || 300;
  var CHUNK = 85000;
  try {
    if (value.length < CHUNK) {
      c.put(key, value, sec);
      try { c.remove(key + '_n'); } catch (e2) { /* ok */ }
      return;
    }
    var n = Math.ceil(value.length / CHUNK);
    if (n > 6) return;
    c.put(key + '_n', String(n), sec);
    for (var i = 0; i < n; i++) {
      c.put(key + '_' + i, value.substring(i * CHUNK, (i + 1) * CHUNK), sec);
    }
  } catch (e) { /* cache is best-effort */ }
}

// Build the product cards + a {rowNumber: {header: value}} map, reused by the page and refreshes.
function regBuildCards_(which, admin) {
  const spec = regFields_(which);
  admin = !!admin;
  const inv = spec.tab === 'Inventory';
  var ck = 'reg_c3_' + spec.which + '_' + (admin ? '1' : '0');
  var hit = regCacheGet_(ck), mapHit = admin ? regCacheGet_(ck + '_m') : null;
  if (hit) {
    try {
      var cached = JSON.parse(hit);
      if (cached && cached.html != null) {
        cached.mapJson = admin ? (mapHit || regSlimMap_(spec)) : '{}';
        return cached;
      }
    } catch (e) { /* rebuild */ }
  }

  const data = readTabRows_(spec.tab);
  const H = data.headers, rows = data.rows;
  const idx = function (n) { for (let i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(n)) return i; } return -1; };
  const cName = idx(inv ? 'Item' : 'Name'), cImg = idx('Image'), cKey = idx(spec.key), cTeam = idx('Owning team'),
    cOn = idx('On hand'), cRe = idx('Reorder point'), cUnit = idx('Unit'), cLoc = idx('Location'),
    cSup = idx('Supplier'), cCat = idx('Category'), cStatus = idx('Status'), cCoTo = idx('Checked out to'), cDue = idx('Due back');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const base = regWebUrl_();
  let html = ''; const mapParts = [], owners = {}, cats = {};
  var nOut = 0, nLow = 0, nCo = 0, nOver = 0;

  const keepH = {};
  spec.fields.forEach(function (f) { keepH[f.h] = 1; });
  ['On hand', 'Reorder point', 'Checked out to', 'Due back', 'Out since', spec.key].forEach(function (h) { keepH[h] = 1; });

  rows.forEach(function (rr) {
    const r = rr.cells;
    const name = cName >= 0 ? String(r[cName]).trim() : '';
    const key = cKey >= 0 ? String(r[cKey]).trim() : '';
    const owner = cTeam >= 0 ? String(r[cTeam]).trim() : ''; if (owner) owners[owner] = 1;
    const cat = cCat >= 0 ? String(r[cCat]).trim() : ''; if (cat) cats[cat] = 1;
    const loc = cLoc >= 0 ? String(r[cLoc]).trim() : '';
    const hay = [name, cat, owner, loc, key, inv && cSup >= 0 ? String(r[cSup]).trim() : ''].join(' ').toLowerCase();
    const imgHtml = regImgTag_(cImg >= 0 ? r[cImg] : '', 400, regMonogram_(name));
    let primary = '', chip = '', coState = '', stockState = '', stepper = '', stockBadge = '';
    var unit = '', reVal = '';
    if (inv) {
      const on = cOn >= 0 && isNum_(r[cOn]) ? Number(r[cOn]) : null;
      const re = cRe >= 0 && isNum_(r[cRe]) ? Number(r[cRe]) : 0;
      unit = cUnit >= 0 ? String(r[cUnit]).trim() : '';
      reVal = re ? String(re) : '';
      if (on !== null) {
        primary = '<span class="reg-card-num">' + on + '</span><span class="reg-card-num-sub">' + escapeHtml_((unit ? unit + ' ' : '') + 'on hand') + '</span>';
        stockBadge = '<span class="reg-card-qty">' + on + (unit ? ' ' + escapeHtml_(unit) : '') + '</span>';
      }
      stockState = (on !== null && on <= 0) ? 'out' : ((on !== null && re > 0 && on > 0 && on <= re) ? 'low' : '');
      if (stockState === 'out') nOut++;
      else if (stockState === 'low') nLow++;
      chip = stockState === 'out' ? '<span class="reg-chip">Out of stock</span>' : (stockState === 'low' ? '<span class="reg-chip reg-chip-low">Low stock</span>' : '');
      if (admin && on !== null) {
        stepper = '<div class="reg-count">'
          + '<button type="button" class="reg-cbtn" onclick="regQadj(' + rr.row + ',this,-1)" aria-label="Decrease">−</button>'
          + '<span class="reg-cval"><span class="reg-cn">' + on + '</span><span class="reg-cu">' + escapeHtml_((unit ? unit + ' ' : '') + 'on hand') + '</span></span>'
          + '<button type="button" class="reg-cbtn" onclick="regQadj(' + rr.row + ',this,1)" aria-label="Increase">+</button>'
          + '<span class="reg-csaved" aria-hidden="true"></span></div>';
        primary = '';
      }
    } else {
      const outTo = cCoTo >= 0 ? String(r[cCoTo]).trim() : '';
      let overdue = false;
      if (outTo && cDue >= 0 && String(r[cDue]).trim()) { const d = new Date(r[cDue]); if (!isNaN(d.getTime())) { d.setHours(0, 0, 0, 0); overdue = d < today; } }
      if (overdue) { chip = '<span class="reg-chip">Overdue</span>'; coState = 'overdue'; nOver++; nCo++; }
      else if (outTo) { chip = '<span class="reg-chip reg-chip-low">Checked out</span>'; coState = 'out'; nCo++; }
      else { const st = cStatus >= 0 ? String(r[cStatus]).trim() : ''; chip = st && /down|out of service|repair|broken/i.test(st) ? '<span class="reg-chip">' + escapeHtml_(st) + '</span>' : ''; }
    }
    const tag = cat ? '<div class="reg-card-tag">' + escapeHtml_(cat) + '</div>' : '';
    const metaBits = [];
    if (loc) metaBits.push(loc);
    if (owner) metaBits.push(owner);
    if (inv && cSup >= 0 && String(r[cSup]).trim()) metaBits.push(String(r[cSup]).trim());
    const meta = metaBits.length ? '<div class="reg-card-meta">' + escapeHtml_(metaBits.join(' · ')) + '</div>' : '';
    const href = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'registry=item&which=' + spec.which + '&id=' + encodeURIComponent(key) + (admin ? '&admin=1' : '');
    html += '<div class="reg-card reg-row" data-hay="' + escapeHtml_(hay) + '" data-cat="' + escapeHtml_(cat) + '" data-owner="' + escapeHtml_(owner) + '" data-co="' + coState + '" data-stock="' + stockState + '" data-key="' + escapeHtml_(key) + '" data-rp="' + escapeHtml_(reVal) + '" data-unit="' + escapeHtml_(unit) + '">'
      + '<a class="reg-card-link" href="' + escapeHtml_(href) + '" title="Open ' + escapeHtml_(name || 'item') + '" onclick="return regGoItem(event,this)"><div class="reg-card-img">' + imgHtml
      + (stockBadge ? stockBadge : '')
      + '<span class="reg-card-go" aria-hidden="true"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"></path></svg></span></div>'
      + '<div class="reg-card-b"><div class="reg-card-title">' + escapeHtml_(name || '(unnamed)') + '</div>' + tag
      + (primary ? '<div class="reg-card-primary">' + primary + '</div>' : '') + (chip || '') + meta + '</div></a>'
      + stepper
      + (admin ? '<div class="reg-card-act"><button type="button" class="reg-iconbtn" title="Edit" onclick="regOpenEdit(' + rr.row + ')">Edit</button><button type="button" class="reg-iconbtn reg-del" title="Delete" onclick="regDeleteRow(' + rr.row + ',this)">Delete</button></div>' : '')
      + '</div>';
    const obj = {};
    Object.keys(keepH).forEach(function (h) {
      var i = idx(h);
      obj[h] = i >= 0 ? regRawVal_(r[i]) : '';
    });
    mapParts.push(JSON.stringify(String(rr.row)) + ':' + JSON.stringify(obj));
  });

  var mapJson = '{' + mapParts.join(',') + '}';
  var out = {
    html: html,
    owners: Object.keys(owners).sort(), cats: Object.keys(cats).sort(),
    key: spec.key, count: rows.length, nOut: nOut, nLow: nLow, nCo: nCo, nOver: nOver
  };
  regCachePut_(ck, JSON.stringify(out), 300);
  if (admin) regCachePut_(ck + '_m', mapJson, 300);
  out.mapJson = mapJson;
  return out;
}

// Client-callable: re-render the cards after a change (no page reload, which blanks the sandbox).
function regRowsHtml(which) {
  const b = regBuildCards_(which, true);
  return { ok: true, html: b.html, mapJson: b.mapJson };
}

function regRowsMap(which) {
  return regBuildCards_(which, true).mapJson || '{}';
}

function registryPage_(which, embedded, admin) {
  const spec = regFields_(which);
  const title = spec.tab === 'Inventory' ? 'Inventory' : 'Equipment registry';
  const inv = spec.tab === 'Inventory';
  const regBase = regWebUrl_();
  const other = spec.which === 'inventory' ? 'equipment' : 'inventory';
  const waitT = inv ? 'Loading the shop catalog…' : 'Loading equipment…';
  const waitH = inv ? 'Fetching items and stock counts' : 'Fetching equipment records';

  let inner = '<div id="reg-root" class="reg-page">'
    + regWaitHtml_(waitT, waitH, true)
    + '<div id="reg-swap"></div>'
    + '<div class="reg-slide" id="reg-slide" hidden><div class="reg-slide-bd" onclick="regBack(event)"></div>'
    + '<div class="reg-slide-dialog" id="reg-itempane" role="dialog" aria-modal="true" aria-label="Item"></div></div>'
    + regStyles_() + regWaitJs_() + regFilterJs_() + regSwitchJs_()
    + '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_ADMIN=' + (admin ? 'true' : 'false')
    + ';var REG_EMBED=' + (embedded ? 'true' : 'false')
    + ';var REG_BASE=' + JSON.stringify(regBase) + ';var REG_KEY="";var REG_ROWS={};var REG_SAVED="";</script>';
  if (admin) inner += regEditJs_();
  inner += '<script>regBoot(' + JSON.stringify(spec.which) + ',' + JSON.stringify(other) + ');</script></div>';
  return swissShell_(inner, title, true, embedded, '1120px');
}

// Pre-build the other catalog tab into CacheService so Inventory <-> Equipment is instant.
function regWarm(which, admin) {
  try { regBuildCards_(which, !!admin); } catch (e) { /* warm is best-effort */ }
  return { ok: true };
}

// The swappable body of the registry page - everything the Inventory/Equipment toggle
// changes. Kept separate so regSwitchHtml() can re-render it IN PLACE via
// google.script.run: no page navigation means no Apps Script cold-load white flash
// (which is what the "back and forth" between tabs was hitting), and it lets the toggle
// show a loading spinner.
function regBodyMarkup_(which, b, admin, embedded) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const title = inv ? 'Inventory' : 'Equipment';
  const base = regWebUrl_();
  const n = b.count || 0;
  const stats = inv
    ? (n + ' item' + (n === 1 ? '' : 's')
      + (b.nLow ? ' · <b class="reg-stat-low">' + b.nLow + ' low</b>' : '')
      + (b.nOut ? ' · <b class="reg-stat-out">' + b.nOut + ' out</b>' : ''))
    : (n + ' item' + (n === 1 ? '' : 's')
      + (b.nCo ? ' · ' + b.nCo + ' out' : '')
      + (b.nOver ? ' · <b class="reg-stat-out">' + b.nOver + ' overdue</b>' : ''));

  let head = '';
  if (!embedded) {
    head = '<div class="page-head"><div class="page-kicker">' + (inv ? 'Shop catalog' : 'Equipment registry') + '</div>'
      + '<div class="reg-head-row"><div class="page-title">' + escapeHtml_(title) + '</div>'
      + '<div class="reg-head-stat" id="reg-count" data-base="' + escapeHtml_(stats) + '">' + stats + '</div></div>'
      + '<div class="page-rule"></div></div>';
  } else {
    head = '<div class="reg-head-row" style="margin-bottom:4px"><div class="reg-head-stat" id="reg-count" data-base="' + escapeHtml_(stats) + '">' + stats + '</div></div>';
  }

  const toggle = '<div class="reg-toggle">'
    + '<button type="button" class="reg-toggle-btn' + (inv ? ' on' : '') + '" onclick="regSwitch(\'inventory\')">Inventory</button>'
    + '<button type="button" class="reg-toggle-btn' + (!inv ? ' on' : '') + '" onclick="regSwitch(\'equipment\')">Equipment</button>'
    + '</div>';

  let toolbar = '';
  if (admin) {
    const labelsHref = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'registry=labels&which=' + spec.which + '&admin=1';
    toolbar = '<div class="reg-tools" id="reg-tools">'
      +   '<button type="button" class="btn btn-primary" onclick="regOpenAdd()">+ Add ' + escapeHtml_(spec.noun) + '</button>'
      +   '<button type="button" class="btn btn-ghost" onclick="regScan()">Scan</button>'
      +   '<a class="btn btn-ghost" href="' + escapeHtml_(labelsHref) + '" target="_blank" rel="noopener">Print labels</a>'
      + '</div>';
  }

  let controls = '';
  const catList = [], catSeen = {};
  const pushCat = function (cat) { if (cat && !catSeen[cat]) { catSeen[cat] = 1; catList.push(cat); } };
  REG_CATEGORIES.forEach(function (cat) { if (cat !== 'Other') pushCat(cat); });
  b.cats.forEach(pushCat);
  if (REG_CATEGORIES.indexOf('Other') >= 0) pushCat('Other');
  controls += '<select id="cat" onchange="flt()"><option value="">All categories</option>'
    + catList.map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('') + '</select>';
  if (b.owners.length) {
    controls += '<select id="owner" onchange="flt()"><option value="">All owners</option>'
      + b.owners.map(function (t) { return '<option value="' + escapeHtml_(t) + '">' + escapeHtml_(t) + '</option>'; }).join('') + '</select>';
  }

  const chipsHtml = inv
    ? '<div class="reg-chips"><button type="button" class="reg-fchip" data-f="stock" onclick="regChip(this)">Low or out</button></div>'
    : '<div class="reg-chips"><button type="button" class="reg-fchip" data-f="out" onclick="regChip(this)">Checked out</button><button type="button" class="reg-fchip" data-f="overdue" onclick="regChip(this)">Overdue</button></div>';

  let out = head + '<div class="reg-toolbar">' + toggle + toolbar + '</div>'
    + '<div class="reg-bar"><div class="filters"><div class="search-wrap"><input id="q" type="search" placeholder="Search name, location, or team" oninput="flt()" autocomplete="off"></div>' + controls + '</div>'
    + chipsHtml + '</div>'
    + '<div class="reg-grid" id="reg-cards">' + b.html + '</div>'
    + '<div id="empty" class="empty reg-empty" style="display:none"><div class="reg-empty-title">No matching items</div>Try a different search, or clear the filters.</div>';
  if (!b.html) out += '<div class="empty">Nothing here yet.' + (admin ? ' Add the first ' + escapeHtml_(spec.noun) + '.' : '') + '</div>';
  if (admin) out += regFormOverlay_(which);
  return out;
}

// Client-callable: fresh body markup + row data for the other tab (drives the in-place toggle).
function regSwitchHtml(which, admin, embedded) {
  const spec = regFields_(which);
  const b = regBuildCards_(which, !!admin);
  return { ok: true, html: regBodyMarkup_(which, b, !!admin, !!embedded), which: spec.which, key: b.key, rowsJson: admin ? b.mapJson : '{}' };
}

// The in-place Inventory/Equipment toggle: swap the body via google.script.run instead
// of navigating. Named wait copy so a slow fetch never looks like a freeze.
function regSwitchJs_() {
  return '<script>'
    + 'function regFillBody(r){var sw=document.getElementById("reg-swap");if(!sw||!r||!r.ok)return false;'
    + 'sw.innerHTML=r.html;REG_WHICH=r.which;if(typeof REG_ADMIN!=="undefined"&&REG_ADMIN){REG_KEY=r.key;try{REG_ROWS=JSON.parse(r.rowsJson);}catch(e){REG_ROWS={};}}'
    + 'if(typeof REG_FILTER!=="undefined")REG_FILTER="";if(typeof flt==="function")flt();return true;}'
    + 'function regBoot(which,other){var ad=!!(typeof REG_ADMIN!=="undefined"&&REG_ADMIN);var em=!!(typeof REG_EMBED!=="undefined"&&REG_EMBED);'
    + 'regWait(true,which==="inventory"?"Loading the shop catalog\\u2026":"Loading equipment\\u2026",which==="inventory"?"Fetching items and stock counts":"Fetching equipment records");'
    + 'google.script.run.withSuccessHandler(function(r){if(!regFillBody(r)){regWaitSay("Could not load","Refresh the page and try again.");return;}'
    + 'regWait(false);try{if(other)google.script.run.regWarm(other,ad);}catch(e){}})'
    + '.withFailureHandler(function(){regWaitSay("Could not load","Refresh the page and try again.");}).regSwitchHtml(which,ad,em);}'
    + 'function regSwitch(which){if(typeof REG_WHICH!=="undefined"&&which===REG_WHICH)return;'
    + 'var inv=which==="inventory";regWait(true,inv?"Loading inventory\\u2026":"Loading equipment\\u2026",inv?"Fetching the shop catalog":"Fetching equipment records");'
    + 'var ad=!!(typeof REG_ADMIN!=="undefined"&&REG_ADMIN);var em=!!(typeof REG_EMBED!=="undefined"&&REG_EMBED);'
    + 'google.script.run.withSuccessHandler(function(r){if(!regFillBody(r)){regWait(false);return;}regWait(false);})'
    + '.withFailureHandler(function(){regWait(false);alert("Could not switch catalogs. Try again.");}).regSwitchHtml(which,ad,em);}'
    + 'function regGoItem(ev,a){if(!ev||ev.metaKey||ev.ctrlKey||ev.shiftKey||ev.altKey||(ev.button&&ev.button!==0))return true;'
    + 'ev.preventDefault();var card=a.closest?a.closest(".reg-card"):null;var key=card?card.getAttribute("data-key"):"";'
    + 'var titleEl=card?card.querySelector(".reg-card-title"):null;var name=titleEl?titleEl.textContent:"item";'
    + 'var slide=document.getElementById("reg-slide");var pane=document.getElementById("reg-itempane");'
    + 'var ad=!!(typeof REG_ADMIN!=="undefined"&&REG_ADMIN);var req=++REG_ITEM_REQ;'
    + 'if(slide&&pane){pane.innerHTML=\'<div class="reg-slide-load"><span class="reg-wait-spin" aria-hidden="true"></span><p class="reg-wait-t"></p><p class="reg-wait-h">Fetching the catalog record</p></div>\';'
    + 'var t=pane.querySelector(".reg-wait-t");if(t)t.textContent="Opening "+name;regOpenSlide();}'
    + 'else{regWait(true,"Opening "+name,"Fetching the catalog record");}'
    + 'google.script.run.withSuccessHandler(function(r){if(req!==REG_ITEM_REQ)return;'
    + 'if(!r||!r.html){if(slide)regCloseItem();else regWait(false);alert((r&&r.error)||"Could not open that item.");return;}'
    + 'if(pane&&slide){pane.innerHTML=r.html;}else{var sw=document.getElementById("reg-swap");REG_SAVED=sw.innerHTML;sw.innerHTML=r.html;}'
    + 'REG_ITEM=1;if(r.ok&&r.row!=null&&r.vals)REG_ROWS[String(r.row)]=r.vals;if(r.key)REG_KEY=r.key;'
    + 'try{window.scrollTo(0,0);}catch(e){}regWait(false);})'
    + '.withFailureHandler(function(){if(req!==REG_ITEM_REQ)return;regWait(false);if(slide)regCloseItem();try{location.href=a.href;}catch(e){}}).regItemHtml(REG_WHICH,key,ad);return false;}'
    + '</script>';
}

// Single-item view: product page. Client-callable so the catalog can open an item
// in place (named wait, no white flash) and so a QR/deep link can paint a shell first.
function regItemHtml(which, id, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  let regBase = regWebUrl_();
  const backHref = regBase + (regBase.indexOf('?') >= 0 ? '&' : '?') + 'registry=' + spec.which + (admin ? '&admin=1' : '');
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  let found = null;
  data.rows.forEach(function (rr) { if (ki >= 0 && String(rr.cells[ki]).trim() === String(id).trim()) found = rr; });

  const back = '<a class="reg-back" href="' + escapeHtml_(backHref) + '" target="_top" rel="noopener" onclick="return regBack(event,this)"><span aria-hidden="true">&#8592;</span> Back to ' + escapeHtml_(spec.tab.toLowerCase()) + '</a>';

  if (!found) {
    return { ok: false, error: 'Not found', html: '<div class="reg-item">' + back
      + '<div class="page-title" style="margin-top:18px">Not found</div>'
      + '<p class="reg-detail-sub">No ' + escapeHtml_(spec.noun) + ' matches "' + escapeHtml_(id) + '".</p></div>' };
  }

  const r = found.cells;
  const hi = function (n) { return H.map(function (h) { return norm_(h); }).indexOf(norm_(n)); };
  const val = function (n) { var i = hi(n); return i >= 0 ? r[i] : ''; };
  const txt = function (n) { return String(val(n) == null ? '' : val(n)).trim(); };
  const name = txt(inv ? 'Item' : 'Name') || String(id);
  const cat = txt('Category'), loc = txt('Location'), owner = txt('Owning team');
  const imgRaw = val('Image');

  const toi = hi('Checked out to'), osi = hi('Out since'), dbi = hi('Due back'), oni = hi('On hand');
  const outTo = toi >= 0 ? String(r[toi]).trim() : '';
  const dueTxt = dbi >= 0 && String(r[dbi]).trim() ? regCell_(r[dbi]) : '';
  const overdue = (function () {
    if (!outTo || dbi < 0 || !String(r[dbi]).trim()) return false;
    const d = r[dbi] instanceof Date ? new Date(r[dbi]) : new Date(r[dbi]);
    if (isNaN(d.getTime())) return false;
    const t = new Date(); t.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0); return d < t;
  })();

  const on = oni >= 0 && isNum_(r[oni]) ? Number(r[oni]) : null;
  const re = hi('Reorder point') >= 0 && isNum_(val('Reorder point')) ? Number(val('Reorder point')) : 0;
  const unit = txt('Unit');
  const stockState = (on !== null && on <= 0) ? 'out' : ((on !== null && re > 0 && on > 0 && on <= re) ? 'low' : '');
  const stockChip = stockState === 'out' ? '<span class="reg-chip">Out of stock</span>'
    : (stockState === 'low' ? '<span class="reg-chip reg-chip-low">Low stock</span>' : '');

  const media = '<div class="reg-detail-media">' + (regThumb_(imgRaw, 800)
    ? regImgTag_(imgRaw, 800, regMonogram_(name))
    : '<span class="reg-card-mono">' + escapeHtml_(regMonogram_(name)) + '</span>') + '</div>';

  const subBits = [];
  if (cat) subBits.push(escapeHtml_(cat));
  if (loc) subBits.push(escapeHtml_(loc));
  if (owner) subBits.push(escapeHtml_(owner));
  const sub = subBits.length ? '<div class="reg-detail-sub">' + subBits.join('<span class="reg-dotsep" aria-hidden="true"></span>') + '</div>' : '';

  let stock = '';
  if (inv && on !== null) {
    stock = '<div class="reg-stock' + (stockState ? ' is-' + stockState : '') + '">'
      + '<span class="reg-stock-n">' + on + '</span>'
      + '<span class="reg-stock-u">' + escapeHtml_((unit ? unit + ' ' : '') + 'on hand') + '</span>'
      + stockChip + '</div>'
      + (re > 0 ? '<div class="reg-stock-re">Reorder at ' + re + '</div>' : '');
  }

  let panel = '';
  if (!inv) {
    if (outTo) {
      const when = (osi >= 0 && String(r[osi]).trim() ? 'out since ' + regCell_(r[osi]) : '') + (dueTxt ? (osi >= 0 && String(r[osi]).trim() ? ' · ' : '') + 'due ' + dueTxt : '');
      panel = '<div class="reg-co' + (overdue ? ' is-over' : '') + '"><div class="reg-co-h">' + (overdue ? 'Overdue' : 'Checked out') + '</div>'
        + '<div class="reg-co-who">' + escapeHtml_(outTo) + '</div>'
        + (when ? '<div class="reg-co-sub">' + when + '</div>' : '')
        + (admin ? '<button type="button" class="btn btn-confirm" onclick="regReturnOne(' + found.row + ',this)">Mark returned</button>' : '') + '</div>';
    } else if (admin) {
      panel = '<div class="reg-co"><div class="reg-co-h">Check out</div>'
        + '<div class="reg-co-form"><input id="co-person" placeholder="Who is taking it?"><input id="co-due" type="date" title="Due back"><button type="button" class="btn btn-primary" onclick="regCheckoutOne(' + found.row + ',this)">Check out</button></div></div>';
    }
  } else if (admin && on !== null) {
    panel = '<div class="reg-co reg-co-count"><div class="reg-co-h">Update count</div>'
      + '<div class="reg-step"><button type="button" class="reg-stepbtn" onclick="regStep(-1)" aria-label="Decrease">&minus;</button>'
      + '<input id="co-count" type="number" inputmode="numeric" value="' + on + '" aria-label="On hand">'
      + '<button type="button" class="reg-stepbtn" onclick="regStep(1)" aria-label="Increase">+</button>'
      + '<button type="button" class="btn btn-confirm" onclick="regCountSave(' + found.row + ',this)">Save</button></div></div>';
  }

  const nice = { 'Owning team': 'Owner', 'Product link': 'Product', 'eShop info': 'eShop', 'Asset ID': 'ID' };
  const skip = { Image: 1, Item: 1, Name: 1, Category: 1, Location: 1, 'Owning team': 1 };
  if (inv) { skip['On hand'] = 1; skip['Unit'] = 1; if (re > 0) skip['Reorder point'] = 1; }
  if (!inv) { skip['Checked out to'] = 1; skip['Out since'] = 1; skip['Due back'] = 1; }

  const primaryNames = inv
    ? ['Supplier', 'Product link']
    : ['Status', 'Asset ID'];
  const moreNames = inv
    ? ['Reorder point', 'Reorder qty', 'eShop info', 'Assign to', 'Last restocked', 'Last counted', 'Battery state', 'Battery spec']
    : ['Owner', 'Installed', 'Notes'];

  const kv = function (h) {
    var i = hi(h);
    if (i < 0 || skip[h] || !String(r[i]).trim()) return '';
    var raw = String(r[i]).trim();
    var body = (h === 'Product link' && /^https?:\/\//i.test(raw))
      ? '<a class="reg-ext" href="' + escapeHtml_(raw) + '" target="_blank" rel="noopener">Open product page</a>'
      : regCell_(r[i]);
    return '<div class="reg-fact"><span class="reg-detail-l">' + escapeHtml_(nice[h] || h) + '</span><span class="reg-detail-v">' + body + '</span></div>';
  };
  const seen = {};
  let primary = '';
  primaryNames.forEach(function (h) { seen[h] = 1; primary += kv(h); });
  let extra = '';
  moreNames.forEach(function (h) { seen[h] = 1; extra += kv(h); });
  H.forEach(function (h) {
    if (seen[h] || skip[h]) return;
    extra += kv(h);
  });
  let facts = '';
  if (primary || extra) {
    facts = '<div class="reg-facts">' + primary
      + (extra ? '<details class="reg-facts-more"><summary>More details</summary><div class="reg-facts">' + extra + '</div></details>' : '')
      + '</div>';
  }

  let actions = '';
  if (admin) {
    actions = '<div class="reg-detail-act">'
      + '<button type="button" class="btn btn-primary" onclick="regOpenEdit(' + found.row + ')">Edit</button>'
      + '<button type="button" class="btn btn-ghost reg-del" onclick="regDeleteRow(' + found.row + ',this)">Delete</button>'
      + '</div>';
  }

  let html = '<div class="reg-item">' + back
    + '<div class="reg-detail">'
    +   media
    +   '<div class="reg-detail-main">'
    +     '<div class="reg-detail-kicker">' + (inv ? 'Shop catalog' : 'Equipment') + '</div>'
    +     '<h1 class="reg-detail-title">' + escapeHtml_(name) + '</h1>'
    +     '<div class="reg-detail-rule" aria-hidden="true"></div>'
    +     sub + stock + panel + facts + actions
    +   '</div>'
    + '</div></div>';
  if (admin) html += regFormOverlay_(which);

  const vals = {};
  H.forEach(function (h, i) { vals[h] = regRawVal_(r[i]); });
  return { ok: true, html: html, name: name, row: found.row, vals: vals, key: spec.key };
}

function regItemPage_(which, id, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const regBase = regWebUrl_();
  let inner = '<div id="reg-root" class="reg-page">'
    + regWaitHtml_('Opening item…', 'Fetching the catalog record', true)
    + '<div id="reg-swap"></div>'
    + regStyles_() + regWaitJs_()
    + '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_ADMIN=' + (admin ? 'true' : 'false')
    + ';var REG_BASE=' + JSON.stringify(regBase) + ';var REG_KEY=' + JSON.stringify(spec.key)
    + ';var REG_ROWS={};var REG_SAVED="";var REG_ITEM=1;</script>';
  if (admin) inner += regEditJs_();
  inner += '<script>regBootItem(' + JSON.stringify(spec.which) + ',' + JSON.stringify(String(id)) + ',' + (admin ? 'true' : 'false') + ');</script></div>';
  return swissShell_(inner, inv ? 'Inventory' : 'Equipment', true, false, '1040px');
}

// Printable QR labels. Each label deep-links to its item page; scan with any phone camera.
function regLabelsPage_(which) {
  const spec = regFields_(which);
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  const ni = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.tab === 'Inventory' ? 'Item' : 'Name'));
  let base = regWebUrl_();

  let cards = '';
  data.rows.forEach(function (rr) {
    const key = ki >= 0 ? String(rr.cells[ki]).trim() : '';
    if (!key) return;
    const name = ni >= 0 ? String(rr.cells[ni]).trim() : key;
    const url = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'registry=item&which=' + spec.which + '&id=' + encodeURIComponent(key) + '&admin=1';
    cards += '<div class="lab"><div class="lab-qr" data-url="' + escapeHtml_(url) + '"></div>'
      + '<div class="lab-name">' + escapeHtml_(name) + '</div><div class="lab-id">' + escapeHtml_(key) + '</div></div>';
  });

  const inner = '<div class="lab-head"><div><div class="page-kicker">Ops registry</div><div class="page-title">' + escapeHtml_(spec.tab) + ' labels</div></div>'
    + '<button class="btn btn-primary lab-print" onclick="window.print()">Print</button></div>'
    + '<p class="lab-hint">Print, cut, and stick one on each ' + escapeHtml_(spec.noun) + '. Scanning a label with any phone camera opens its record.</p>'
    + '<div class="lab-grid">' + (cards || '<div class="empty">Nothing to label yet.</div>') + '</div>'
    + '<style>'
    + '.lab-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap}'
    + '.lab-hint{color:#8a857c;font-size:13.5px;margin:6px 0 16px}'
    + '.lab-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px}'
    + '.lab{border:1.5px solid #ddd;border-radius:12px;padding:14px;text-align:center;background:#fff;break-inside:avoid}'
    + '.lab-qr{width:150px;height:150px;margin:0 auto 10px;display:flex;align-items:center;justify-content:center}'
    + '.lab-qr img,.lab-qr canvas{width:150px !important;height:150px !important}'
    + '.lab-qr .lab-fallback{font-size:9px;word-break:break-all;color:#666;line-height:1.3}'
    + '.lab-name{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-weight:800;font-size:14px;color:#111;line-height:1.2}'
    + '.lab-id{font-size:11px;color:#8a857c;margin-top:3px;font-weight:700;letter-spacing:.04em}'
    + '@media print{.lab-print{display:none}.lab-head{margin-bottom:8px}}'
    + '</style>'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>'
    + '<script>(function(){function draw(){document.querySelectorAll(".lab-qr").forEach(function(el){var u=el.getAttribute("data-url");'
    + 'if(window.QRCode){try{new QRCode(el,{text:u,width:150,height:150,correctLevel:QRCode.CorrectLevel.M});return;}catch(e){}}'
    + 'el.innerHTML="<span class=\\"lab-fallback\\">"+u+"</span>";});}'
    + 'if(window.QRCode){draw();}else{var t=0,iv=setInterval(function(){t+=200;if(window.QRCode||t>4000){clearInterval(iv);draw();}},200);}})();</script>';

  return swissShell_(inner, spec.tab + ' labels', true, false);
}

// ---- registry admin: shared styles, filter, unlock, form overlay, edit JS ----

function regWaitHtml_(title, hint, show) {
  return '<div class="reg-topbar' + (show ? ' on' : '') + '" id="reg-top" aria-hidden="true"></div>'
    + '<div class="reg-wait" id="reg-wait"' + (show ? '' : ' hidden') + ' role="status" aria-live="polite">'
    + '<div class="reg-wait-card"><span class="reg-wait-spin" aria-hidden="true"></span>'
    + '<p class="reg-wait-t" id="reg-wait-t">' + escapeHtml_(title || 'Loading…') + '</p>'
    + '<p class="reg-wait-h" id="reg-wait-h">' + escapeHtml_(hint || '') + '</p></div></div>';
}

function regWaitJs_() {
  return '<script>'
    + 'var REG_BUSY=0,REG_WAIT_IV=null,REG_ITEM_REQ=0;'
    + 'function regTop(on){var b=document.getElementById("reg-top");if(!b)return;'
    + 'REG_BUSY=Math.max(0,REG_BUSY+(on?1:-1));'
    + 'if(REG_BUSY>0)b.className="reg-topbar on";else{b.className="reg-topbar done";setTimeout(function(){if(REG_BUSY<=0)b.className="reg-topbar";},600);}}'
    + 'function regWait(on,title,hint){var w=document.getElementById("reg-wait");if(REG_WAIT_IV){clearInterval(REG_WAIT_IV);REG_WAIT_IV=null;}'
    + 'if(on){regTop(true);regWaitSay(title||"Loading\\u2026",hint||"");if(w)w.hidden=false;'
    + 'var n=0;REG_WAIT_IV=setInterval(function(){n++;var h=document.getElementById("reg-wait-h");if(!h)return;'
    + 'h.textContent=n>=2?"Almost there\\u2026":"Still working\\u2026 this can take a few seconds";},1800);}'
    + 'else{regTop(false);if(w)w.hidden=true;}}'
    + 'function regWaitSay(title,hint){var t=document.getElementById("reg-wait-t");var h=document.getElementById("reg-wait-h");if(t&&title)t.textContent=title;if(h&&hint!=null)h.textContent=hint;}'
    + 'function regImgFb(el){el.classList.remove("reg-img","is-in");var f=el.getAttribute("data-fb");if(f){el.removeAttribute("data-fb");el.src=f;return;}'
    + 'var m=el.getAttribute("data-mono")||"\\u00b7";el.outerHTML=\'<span class="reg-card-mono" aria-hidden="true">\'+m+\'</span>\';}'
    + 'function regOpenSlide(){var slide=document.getElementById("reg-slide");if(!slide)return;slide.hidden=false;requestAnimationFrame(function(){slide.classList.add("is-open");});}'
    + 'function regCloseItem(){REG_ITEM_REQ++;REG_ITEM=undefined;var slide=document.getElementById("reg-slide");var pane=document.getElementById("reg-itempane");'
    + 'if(!slide)return;slide.classList.remove("is-open");setTimeout(function(){if(!slide.classList.contains("is-open")){slide.hidden=true;if(pane)pane.innerHTML="";}},360);}'
    + 'function regBack(ev,a){var slide=document.getElementById("reg-slide");if(slide&&(slide.classList.contains("is-open")||!slide.hidden)){if(ev)ev.preventDefault();regCloseItem();return false;}'
    + 'if(typeof REG_SAVED==="string"&&REG_SAVED){if(ev)ev.preventDefault();document.getElementById("reg-swap").innerHTML=REG_SAVED;REG_SAVED="";REG_ITEM=undefined;try{window.scrollTo(0,0);}catch(e){}if(typeof flt==="function")flt();return false;}'
    + 'regWait(true,"Loading the catalog\\u2026","Taking you back");return true;}'
    + 'document.addEventListener("keydown",function(e){if(e.key==="Escape"){var slide=document.getElementById("reg-slide");if(slide&&slide.classList.contains("is-open")){e.preventDefault();regCloseItem();}}});'
    + 'function regBootItem(which,id,admin){regWait(true,"Opening item\\u2026","Fetching the catalog record");google.script.run.withSuccessHandler(function(r){var sw=document.getElementById("reg-swap");'
    + 'if(!sw||!r||!r.html){regWaitSay("Could not load","Refresh the page and try again.");return;}'
    + 'sw.innerHTML=r.html;if(r.ok){REG_ITEM=1;if(typeof REG_ROWS!=="object"||!REG_ROWS)REG_ROWS={};if(r.row!=null&&r.vals)REG_ROWS[String(r.row)]=r.vals;if(r.key)REG_KEY=r.key;if(r.name)try{document.title=r.name;}catch(e){}}'
    + 'regWait(false);}).withFailureHandler(function(){regWaitSay("Could not load","Refresh the page and try again.");}).regItemHtml(which,id,!!admin);}'
    + '</script>';
}

function regStyles_() {
  return '<style>'
    + '.reg-page .page-rule{width:56px;background:linear-gradient(90deg,#0d9488,#14b8a6,#f0c050)}'
    + '.reg-head-row{display:flex;align-items:baseline;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-top:8px}'
    + '.reg-head-row .page-title{margin-top:0}'
    + '.reg-head-stat{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:600;color:#8a857c}'
    + '.reg-head-stat b{font-weight:800}'
    + '.reg-stat-low{color:#b06a00}.reg-stat-out{color:#b31b1b}'
    + '.reg-toolbar{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin:18px 0 4px}'
    + '.reg-toggle{display:inline-flex;background:#efece6;border-radius:999px;padding:3px;gap:2px;margin:0;border:1px solid #e7e2d8}'
    + '.reg-toggle-btn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#8a857c;text-decoration:none;padding:8px 18px;border-radius:999px;border:none;background:transparent;cursor:pointer;transition:background .15s,color .15s,box-shadow .15s}'
    + '.reg-toggle-btn.on{background:#0d9488;color:#fff;box-shadow:0 2px 8px rgba(13,148,136,.28)}'
    + '.reg-tools{display:flex;gap:8px;flex-wrap:wrap;margin:0}'
    + '#reg-swap{position:relative}'
    + '#reg-swap.reg-loading{opacity:.45;pointer-events:none}'
    + '#reg-swap.reg-loading::after{content:"";position:absolute;left:50%;top:160px;width:34px;height:34px;margin-left:-17px;border:3px solid #e6e1d8;border-top-color:#0d9488;border-radius:50%;animation:regspin .7s linear infinite;z-index:4}'
    + '@keyframes regspin{to{transform:rotate(360deg)}}'
    + '.reg-topbar{position:fixed;top:0;left:0;height:3px;width:0;z-index:91;background:linear-gradient(90deg,#0d9488,#14b8a6,#f0c050);opacity:0;pointer-events:none}'
    + '.reg-topbar.on{width:88%;opacity:1;transition:width 8s cubic-bezier(.05,.8,.25,1),opacity .2s}'
    + '.reg-topbar.done{width:100%;opacity:0;transition:width .2s,opacity .45s .15s}'
    + '.reg-wait{position:fixed;inset:0;z-index:70;display:flex;align-items:center;justify-content:center;padding:24px;background:rgba(245,244,240,.78);-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px)}'
    + '.reg-wait[hidden]{display:none}'
    + '.reg-wait-card{text-align:center;background:#fff;border:1px solid #e8e4dc;border-radius:20px;padding:28px 32px 24px;box-shadow:0 18px 40px rgba(20,17,14,.1);max-width:360px}'
    + '.reg-wait-spin{display:inline-block;width:28px;height:28px;border:3px solid #e6e1d8;border-top-color:#0d9488;border-radius:50%;animation:regspin .7s linear infinite;margin-bottom:14px}'
    + '.reg-wait-t{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:17px;font-weight:800;letter-spacing:-.02em;color:#14110e;margin:0 0 6px}'
    + '.reg-wait-h{font-size:13.5px;font-weight:600;color:#8a857c;margin:0;min-height:1.35em;line-height:1.35}'
    + '.reg-img{opacity:0;transition:opacity .28s ease}.reg-img.is-in{opacity:1}'
    + '.reg-slide{position:fixed;inset:0;z-index:45;display:flex;align-items:stretch;justify-content:flex-end;pointer-events:none}'
    + '.reg-slide[hidden]{display:none}'
    + '.reg-slide.is-open{pointer-events:auto}'
    + '.reg-slide-bd{position:absolute;inset:0;background:rgba(14,14,18,.52);-webkit-backdrop-filter:blur(4px);backdrop-filter:blur(4px);opacity:0;transition:opacity .32s ease}'
    + '.reg-slide.is-open .reg-slide-bd{opacity:1}'
    + '.reg-slide-dialog{position:relative;z-index:1;width:100%;height:100%;background:#f5f4f0;box-shadow:-12px 0 48px rgba(20,20,30,.22);transform:translateX(100%);transition:transform .36s cubic-bezier(.22,1,.36,1);overflow:auto;padding:22px 24px 48px}'
    + '.reg-slide.is-open .reg-slide-dialog{transform:translateX(0)}'
    + '.reg-slide-load{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:60vh;text-align:center;gap:8px}'
    + '@media(prefers-reduced-motion:reduce){.reg-slide-bd,.reg-slide-dialog{transition:none}}'
    + '.reg-spin{display:inline-block;width:13px;height:13px;margin-right:7px;vertical-align:-2px;border:2px solid rgba(255,255,255,.45);border-top-color:#fff;border-radius:50%;animation:regspin .6s linear infinite}'
    + '.reg-bar{position:sticky;top:0;z-index:8;padding:12px 2px 14px;margin:8px -2px 2px;background:rgba(245,244,240,.9);-webkit-backdrop-filter:saturate(1.15) blur(12px);backdrop-filter:saturate(1.15) blur(12px);border-bottom:1px solid rgba(20,17,14,.06)}'
    + '.reg-page .filters{margin:0;padding:0;background:transparent;border:none;box-shadow:none;gap:8px;align-items:center}'
    + '.reg-page .filters select{min-width:0;padding:10px 12px;border-radius:11px}'
    + '.reg-page .search-wrap{min-width:200px}'
    + '.reg-page .search-wrap input:focus{border-color:#0d9488;box-shadow:0 0 0 4px rgba(13,148,136,.14)}'
    + '.reg-page .page-kicker{color:#0f766e}'
    + '.reg-page .filters select:focus{border-color:#0d9488;box-shadow:0 0 0 4px rgba(13,148,136,.14)}'
    + '.reg-page .empty{background:transparent;border:none;box-shadow:none}'
    + '.reg-empty{text-align:center;padding:48px 16px;color:#8a857c}'
    + '.reg-empty-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:16px;font-weight:800;color:#14110e;margin-bottom:6px}'
    + '.reg-page .reg-fchip.on{background:#0d9488;border-color:#0d9488;color:#fff}'
    + '.reg-page .reg-fchip:hover{border-color:#0d9488;color:#0f766e}'
    + '.reg-page .reg-fchip.on:hover{color:#fff}'
    + '.reg-page .reg-detail-media .reg-card-mono{font-size:56px;color:#cbbfa8}'
    + '.reg-chip{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0;padding:3px 9px;border-radius:999px;align-self:flex-start}'
    + '.reg-chip-low{color:#b06a00;background:#fbf3e1;border-color:#eeddb4}'
    + '.reg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:18px;margin-top:8px}'
    + '.reg-card{position:relative;display:flex;flex-direction:column;background:#fff;border:1px solid #e8e4dc;border-radius:18px;overflow:hidden;box-shadow:0 1px 2px rgba(20,17,14,.04);transition:box-shadow .2s ease,transform .2s ease,border-color .2s ease;content-visibility:auto;contain-intrinsic-size:400px}'
    + '.reg-card:hover{box-shadow:0 16px 36px rgba(20,17,14,.12);transform:translateY(-4px);border-color:#ddd6c8}'
    + '.reg-card[data-stock="out"]{border-color:#f0cfcf}'
    + '.reg-card[data-stock="low"]{border-color:#eed9a8}'
    + '.reg-card-link{display:flex;flex-direction:column;flex:1;min-height:0;text-decoration:none;color:inherit;cursor:pointer}'
    + '.reg-card-img{position:relative;aspect-ratio:1/1;background:linear-gradient(160deg,#f3efe6 0%,#e8e2d6 100%);display:flex;align-items:center;justify-content:center;overflow:hidden}'
    + '.reg-card-img img{width:100%;height:100%;object-fit:cover;display:block;transition:transform .35s ease}'
    + '.reg-card:hover .reg-card-img img{transform:scale(1.04)}'
    + '.reg-card[data-stock="out"] .reg-card-img img{filter:saturate(.35) brightness(.92)}'
    + '.reg-card-img::after{content:"";position:absolute;left:0;right:0;bottom:0;height:42%;background:linear-gradient(180deg,transparent,rgba(20,17,14,.18));pointer-events:none}'
    + '.reg-card-mono{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:28px;font-weight:800;letter-spacing:.04em;color:#c4bba8}'
    + '.reg-card-qty{position:absolute;left:10px;top:10px;z-index:1;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.02em;color:#14110e;background:rgba(255,255,255,.92);border-radius:999px;padding:4px 9px;box-shadow:0 2px 8px rgba(20,17,14,.12)}'
    + '.reg-card[data-stock="out"] .reg-card-qty{color:#fff;background:#b31b1b}'
    + '.reg-card[data-stock="low"] .reg-card-qty{color:#7a4e00;background:#f6e3b0}'
    + '.reg-card-go{position:absolute;right:10px;bottom:10px;z-index:1;width:30px;height:30px;border-radius:50%;background:#0d9488;box-shadow:0 2px 8px rgba(13,148,136,.28);display:flex;align-items:center;justify-content:center;color:#fff;transition:transform .16s ease,background .16s,box-shadow .16s}'
    + '.reg-card:hover .reg-card-go{background:#0a6d64;transform:translateX(2px);box-shadow:0 4px 12px rgba(13,148,136,.35)}'
    + '.reg-card-noimg{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#cbc4b8}'
    + '.reg-card-b{padding:13px 14px 15px;display:flex;flex-direction:column;gap:5px}'
    + '.reg-card-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:14.5px;font-weight:800;letter-spacing:-.02em;color:#14110e;line-height:1.3;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}'
    + '.reg-card:hover .reg-card-title{color:#0a6d64}'
    + '.reg-card-tag{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#0d9488}'
    + '.reg-card-primary{display:flex;align-items:baseline;gap:5px;flex-wrap:wrap}'
    + '.reg-card-num{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:19px;font-weight:800;letter-spacing:-.02em;color:#14110e}'
    + '.reg-card-num-sub{font-size:12px;font-weight:600;color:#8a857c}'
    + '.reg-card-meta{font-size:12px;color:#8a857c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    + '.reg-card-act{position:absolute;top:8px;right:8px;display:flex;gap:6px;z-index:2}'
    + '.reg-iconbtn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#57534e;background:rgba(255,255,255,.96);border:1px solid #e2ddd6;border-radius:8px;padding:5px 10px;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.1)}'
    + '.reg-iconbtn:hover{border-color:#0d9488;color:#0a6d64}.reg-iconbtn.reg-del:hover{border-color:#e6b3b3;color:#b31b1b}'
    + '@media(hover:hover) and (pointer:fine){.reg-card-act{opacity:0;transition:opacity .16s}.reg-card:hover .reg-card-act,.reg-card:focus-within .reg-card-act{opacity:1}}'
    + '.reg-count{display:flex;align-items:center;gap:8px;padding:9px 12px;border-top:1px solid #f0ebe3;background:#faf8f4}'
    + '.reg-cbtn{flex:0 0 auto;width:34px;height:34px;border-radius:10px;border:1.5px solid #e2ddd6;background:#fff;font-size:20px;font-weight:700;color:#26231f;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:border-color .14s,color .14s,background .14s}'
    + '.reg-cbtn:hover{border-color:#0d9488;color:#0d9488}.reg-cbtn:active{background:#eef8f6;transform:scale(.96)}'
    + '.reg-cval{flex:1;min-width:0;display:flex;flex-direction:column;line-height:1.15;text-align:center}'
    + '.reg-cn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;letter-spacing:-.02em;color:#14110e}'
    + '.reg-cu{font-size:10px;font-weight:700;letter-spacing:.04em;color:#a8a29e;text-transform:uppercase}'
    + '.reg-csaved{flex:0 0 auto;width:16px;text-align:center;font-size:13px;font-weight:800;transition:color .14s}'
    + '.reg-csaved.saving{color:#cbc4b8}.reg-csaved.ok{color:#0d9488}.reg-csaved.bad{color:#b31b1b}'
    + '@media(prefers-reduced-motion:reduce){.reg-card,.reg-card-img img,.reg-card-go,.reg-img{transition:none}.reg-card:hover{transform:none}.reg-card:hover .reg-card-img img{transform:none}.reg-wait-spin,.reg-topbar.on{animation:none;transition:opacity .2s}}'
    // ---- product detail ----
    + '.reg-item{padding-bottom:12px}'
    + '.reg-back{display:inline-flex;align-items:center;gap:8px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:800;letter-spacing:-.01em;color:#0f766e;text-decoration:none;padding:9px 14px 9px 12px;border-radius:10px;background:#fff;border:1.5px solid #cde8e4;box-shadow:0 1px 2px rgba(20,17,14,.05);transition:color .15s,background .15s,border-color .15s,box-shadow .15s,transform .15s}'
    + '.reg-back:hover{color:#fff;background:#0d9488;border-color:#0d9488;box-shadow:0 6px 16px rgba(13,148,136,.22);transform:translateY(-1px)}'
    + '.reg-detail{display:grid;grid-template-columns:minmax(0,420px) minmax(0,1fr);gap:40px;align-items:start;margin-top:22px}'
    + '.reg-detail-media{position:relative;aspect-ratio:1/1;background:linear-gradient(165deg,#f7f4ee 0%,#efe8da 100%);border:1px solid #e8e4dc;border-radius:22px;overflow:hidden;display:flex;align-items:center;justify-content:center;box-shadow:0 18px 40px rgba(20,17,14,.08)}'
    + '.reg-detail-media:not(:has(img.is-in)):not(:has(.reg-card-mono))::before{content:"";position:absolute;width:28px;height:28px;border:3px solid #e6e1d8;border-top-color:#0d9488;border-radius:50%;animation:regspin .7s linear infinite}'
    + '.reg-detail-media img{width:100%;height:100%;object-fit:contain;padding:16px}'
    + '.reg-detail-main{display:flex;flex-direction:column;min-width:0;padding-top:4px}'
    + '.reg-detail-kicker{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#0f766e}'
    + '.reg-detail-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:34px;font-weight:800;letter-spacing:-.04em;line-height:1.08;margin:8px 0 0;color:#14110e}'
    + '.reg-detail-rule{width:48px;height:3px;border-radius:99px;background:linear-gradient(90deg,#0d9488,#14b8a6,#f0c050);margin:14px 0 12px}'
    + '.reg-detail-sub{font-size:14px;font-weight:600;color:#8a857c;line-height:1.45}'
    + '.reg-dotsep{display:inline-block;width:4px;height:4px;margin:0 9px 2px;border-radius:50%;background:#14b8a6;vertical-align:middle;opacity:.75}'
    + '.reg-stock{display:flex;align-items:baseline;flex-wrap:wrap;gap:8px 10px;margin:18px 0 0}'
    + '.reg-stock-n{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:48px;font-weight:800;letter-spacing:-.05em;line-height:1;color:#14110e}'
    + '.reg-stock.is-out .reg-stock-n{color:#b31b1b}.reg-stock.is-low .reg-stock-n{color:#b06a00}'
    + '.reg-stock-u{font-size:14px;font-weight:700;color:#8a857c}'
    + '.reg-item .reg-chip{font-size:10.5px;padding:5px 11px;align-self:center}'
    + '.reg-stock-re{font-size:12.5px;font-weight:600;color:#8a857c;margin:6px 0 0}'
    + '.reg-facts{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:22px}'
    + '.reg-fact{background:#fff;border:1px solid #ece8e0;border-radius:14px;padding:13px 15px;display:flex;flex-direction:column;gap:5px}'
    + '.reg-detail-l{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#a8a29e}'
    + '.reg-detail-v{font-size:14.5px;font-weight:700;color:#26231f;line-height:1.35;word-break:break-word}'
    + '.reg-facts > .reg-fact:only-of-type,.reg-fact:has(.reg-ext){grid-column:1/-1}'
    + '.reg-detail-v a{color:#0d9488;text-decoration:none;border-bottom:1px solid #99f6e4}'
    + '.reg-detail-v a:hover{color:#0a6d64;border-bottom-color:#0d9488}'
    + '.reg-ext{display:inline-flex;align-items:center;color:#0f766e;background:#e6f7f5;border:1px solid #cde8e4;border-radius:999px;padding:7px 12px;font-size:13px;font-weight:800;text-decoration:none}'
    + '.reg-ext:hover{color:#fff;background:#0d9488;border-color:#0d9488}'
    + '.reg-facts-more{grid-column:1/-1;margin-top:2px}'
    + '.reg-facts-more summary{list-style:none;cursor:pointer;user-select:none;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:800;color:#0f766e;padding:8px 2px}'
    + '.reg-facts-more summary::-webkit-details-marker{display:none}'
    + '.reg-facts-more summary::after{content:" +";font-weight:700}'
    + '.reg-facts-more[open] summary::after{content:" \\2212"}'
    + '.reg-facts-more .reg-facts{margin-top:8px}'
    + '.reg-detail-act{display:flex;gap:8px;flex-wrap:wrap;margin-top:22px}'
    + '.reg-item .btn-primary{background:linear-gradient(180deg,#14b8a6 0%,#0d9488 55%,#0f766e 100%);box-shadow:0 4px 14px rgba(13,148,136,.28)}'
    + '.reg-item .btn-primary:hover{box-shadow:0 8px 20px rgba(13,148,136,.36)}'
    + '@media(max-width:760px){.reg-detail{grid-template-columns:1fr;gap:22px}.reg-detail-media{max-width:420px;margin:0 auto}.reg-detail-title{font-size:28px}.reg-stock-n{font-size:40px}.reg-facts{grid-template-columns:1fr}}'
    // quick-filter chips
    + '.reg-chips{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 2px}'
    + '.reg-fchip{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#57534e;background:#fff;border:1.5px solid #e2ddd6;border-radius:999px;padding:7px 14px;cursor:pointer}'
    + '.reg-fchip:hover{border-color:#b5b0a8}.reg-fchip.on{background:#b31b1b;border-color:#b31b1b;color:#fff}'
    // check-out / count panel
    + '.reg-co{background:#fff;border:1.5px solid #d7ebe7;border-radius:16px;padding:16px 18px;margin:18px 0 0;box-shadow:0 1px 2px rgba(13,148,136,.06)}'
    + '.reg-co.is-over{background:#fdecec;border-color:#f5d0d0;box-shadow:none}'
    + '.reg-co-h{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#0f766e;margin-bottom:10px}'
    + '.reg-co.is-over .reg-co-h{color:#b31b1b}'
    + '.reg-co-who{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;color:#14110e}'
    + '.reg-co-sub{font-size:13px;color:#8a857c;margin-top:3px;margin-bottom:12px}'
    + '.reg-co-form{display:flex;gap:8px;flex-wrap:wrap;align-items:center}'
    + '.reg-co-form input{font:inherit;font-size:14px;padding:11px 13px;border:1.5px solid #e0e0dc;border-radius:10px;outline:none;flex:1;min-width:150px}'
    + '.reg-co-form input:focus{border-color:#0d9488;box-shadow:0 0 0 3px rgba(13,148,136,.14)}'
    + '.reg-step{display:flex;align-items:center;gap:10px;flex-wrap:wrap}'
    + '.reg-stepbtn{width:46px;height:46px;border-radius:12px;border:1.5px solid #e2ddd6;background:#fff;font-size:22px;font-weight:700;color:#26231f;cursor:pointer;line-height:1;display:flex;align-items:center;justify-content:center;flex:0 0 auto}'
    + '.reg-stepbtn:hover{border-color:#0d9488;color:#0d9488}.reg-stepbtn:active{background:#eef8f6;transform:scale(.96)}'
    + '.reg-step input{width:86px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:22px;font-weight:800;text-align:center;padding:9px;border:1.5px solid #e0e0dc;border-radius:12px;outline:none}'
    + '.reg-step input:focus{border-color:#0d9488;box-shadow:0 0 0 3px rgba(13,148,136,.14)}'
    // ---- modal form ----
    + '.reg-msg{font-size:12.5px;color:#8a857c;font-weight:600}'
    + '.reg-ov{position:fixed;inset:0;z-index:80;display:flex;align-items:flex-start;justify-content:center;padding:5vh 16px}.reg-ov[hidden]{display:none}'
    + '.reg-ov-bd{position:absolute;inset:0;background:rgba(20,17,14,.5)}'
    + '.reg-ov-card{position:relative;background:#fff;border-radius:18px;max-width:600px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 24px 60px rgba(0,0,0,.32)}'
    + '.reg-ov-h{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:20px;font-weight:800;padding:24px 28px 8px}'
    + '.reg-form{display:grid;grid-template-columns:1fr 1fr;gap:18px 16px;padding:18px 28px}'
    + '.reg-f{display:flex;flex-direction:column;gap:7px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif}.reg-f.wide{grid-column:1 / -1}'
    // "More details" disclosure that folds away the rarely-used fields
    + '.reg-more{margin:2px 28px 4px;border-top:1px solid #f0efe9}'
    + '.reg-more-sum{list-style:none;cursor:pointer;user-select:none;display:inline-flex;align-items:center;gap:9px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12.5px;font-weight:800;letter-spacing:.04em;color:#8f1515;padding:15px 2px 5px}'
    + '.reg-more-sum::-webkit-details-marker{display:none}'
    + '.reg-more-ic{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;font-size:16px;font-weight:400;line-height:1;color:#8f1515;border-radius:6px;border:1.5px solid #e6cccc;transition:transform .16s ease}'
    + '.reg-more[open] .reg-more-ic{transform:rotate(45deg)}'
    + '.reg-form-more{padding:6px 0 16px}'
    + '.reg-f span{font-size:10.5px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#8a857c}.reg-f span i{color:#b31b1b;font-style:normal}'
    + '.reg-f input,.reg-f select{font:inherit;font-size:14px;padding:12px 13px;border:1.5px solid #e0e0dc;border-radius:10px;outline:none;background:#fff}.reg-f input:focus,.reg-f select:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    // image block
    + '.reg-imgblock{grid-column:1 / -1;display:flex;gap:14px;align-items:flex-start;padding:13px;background:#faf9f6;border:1px solid #ececea;border-radius:12px}'
    + '.reg-preview{flex:0 0 auto;width:96px;height:96px;border-radius:11px;background:#fff;border:1px solid #ececea;overflow:hidden;display:flex;align-items:center;justify-content:center;color:#cbc4b8;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;text-align:center}'
    + '.reg-preview img{width:100%;height:100%;object-fit:cover}'
    + '.reg-imgctl{flex:1;min-width:0;display:flex;flex-direction:column;gap:9px}'
    + '.reg-imgbtns{display:flex;gap:8px;flex-wrap:wrap}'
    + '.reg-fetch{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11.5px;font-weight:700;color:#57534e;background:#fff;border:1.5px solid #e2ddd6;border-radius:8px;padding:8px 12px;cursor:pointer;white-space:nowrap}.reg-fetch:hover{border-color:#b5b0a8;color:#292524}'
    + '.reg-urlrow{display:flex;gap:8px}.reg-urlrow input{flex:1;font:inherit;font-size:13.5px;padding:9px 11px;border:1.5px solid #e0e0dc;border-radius:8px;outline:none}'
    + '.reg-past{display:flex;gap:8px;flex-wrap:wrap;max-height:130px;overflow:auto;padding-top:2px}'
    + '.reg-past-img{width:50px;height:50px;border-radius:8px;overflow:hidden;border:2px solid transparent;cursor:pointer;background:#fff;padding:0}'
    + '.reg-past-img img{width:100%;height:100%;object-fit:cover}.reg-past-img:hover{border-color:#b31b1b}'
    + '.reg-ov-foot{display:flex;align-items:center;gap:10px;padding:16px 24px 22px;border-top:1.5px solid #f1f1f1;margin-top:6px}.reg-ov-foot>:first-child{margin-right:auto}'
    + '@media(max-width:520px){.reg-form{grid-template-columns:1fr}.reg-grid{grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px}}'
    + '</style>';
}

function regFilterJs_() {
  return '<script>'
    + 'var REG_FILTER="";'
    + 'function regChip(btn){var was=btn.classList.contains("on");document.querySelectorAll(".reg-fchip").forEach(function(b){b.classList.remove("on");});REG_FILTER=was?"":btn.dataset.f;if(!was)btn.classList.add("on");flt();}'
    + 'function flt(){var qEl=document.getElementById("q");if(!qEl)return;var q=qEl.value.toLowerCase().trim();'
    + 'var cEl=document.getElementById("cat");var cv=cEl?cEl.value:"";var oEl=document.getElementById("owner");var ov=oEl?oEl.value:"";var n=0;'
    + 'document.querySelectorAll(".reg-row").forEach(function(r){var fp=true;'
    + 'if(REG_FILTER==="stock")fp=r.dataset.stock!=="";else if(REG_FILTER==="out")fp=(r.dataset.co==="out"||r.dataset.co==="overdue");else if(REG_FILTER==="overdue")fp=r.dataset.co==="overdue";'
    + 'var ok=fp&&(!q||r.dataset.hay.indexOf(q)>=0)&&(!cv||r.dataset.cat===cv)&&(!ov||r.dataset.owner===ov);r.style.display=ok?"":"none";if(ok)n++;});'
    + 'var e=document.getElementById("empty");if(e)e.style.display=n?"none":"block";'
    + 'var tot=document.querySelectorAll(".reg-row").length;var st=document.getElementById("reg-count");'
    + 'if(st){var filtering=!!(q||cv||ov||REG_FILTER);st.innerHTML=filtering?("Showing <b>"+n+"</b> of "+tot):(st.getAttribute("data-base")||st.innerHTML);}}'
    + '</script>';
}

function regFormOverlay_(which) {
  const spec = regFields_(which);
  // Owning team is a dropdown of the project teams (+ "Shared"); regFill keeps any
  // existing value that is not in the list when editing.
  let teamOpts = [];
  if (REG_MEMO.teams) teamOpts = REG_MEMO.teams;
  else {
    try { teamOpts = icTeamNames_(); } catch (e) { teamOpts = []; }
    ['Shared', 'Program', 'Facilities'].forEach(function (x) { if (teamOpts.indexOf(x) < 0) teamOpts = teamOpts.concat([x]); });
    REG_MEMO.teams = teamOpts;
  }
  const renderField = function (f) {
    if (f.auto) return '';
    if (f.type === 'image') {
      return '<div class="reg-imgblock"><span class="reg-preview" id="reg-preview">No photo</span>'
        + '<div class="reg-imgctl"><div class="reg-imgbtns">'
        +   '<input type="file" accept="image/*" id="reg-file" style="display:none" onchange="regUpload(this)">'
        +   '<button type="button" class="reg-fetch" onclick="document.getElementById(\'reg-file\').click()">Upload photo</button>'
        +   '<button type="button" class="reg-fetch" onclick="regTogglePast()">Past photos</button>'
        +   '<button type="button" class="reg-fetch" onclick="regSearchWeb()">Search the web</button>'
        + '</div>'
        + '<div class="reg-urlrow"><input data-h="Image" id="reg-imgurl" type="text" placeholder="or paste an image URL" oninput="regPreview()"></div>'
        + '<div class="reg-past" id="reg-past" hidden></div>'
        + '</div></div>';
    }
    const grp = f.group ? ' data-group="' + f.group + '"' : '';
    if (f.type === 'select') {
      const opts = f.opts === 'battery' ? REG_BATTERY_STATES : (f.opts === 'team' ? teamOpts : REG_CATEGORIES);
      const chg = f.h === 'Category' ? ' onchange="regToggleBattery()"' : '';
      return '<label class="reg-f"' + grp + '><span>' + escapeHtml_(f.label) + (f.req ? ' <i>required</i>' : '') + '</span>'
        + '<select data-h="' + escapeHtml_(f.h) + '"' + chg + '><option value=""></option>'
        + opts.map(function (o) { return '<option value="' + escapeHtml_(o) + '">' + escapeHtml_(o) + '</option>'; }).join('') + '</select></label>';
    }
    const type = f.type === 'number' ? 'number' : (f.type === 'url' ? 'url' : 'text');
    const wide = (f.h === 'Notes' || f.h === 'eShop info' || f.h === 'Product link' || f.h === 'Battery spec') ? ' wide' : '';
    return '<label class="reg-f' + wide + '"' + grp + '><span>' + escapeHtml_(f.label) + (f.req ? ' <i>required</i>' : '') + '</span>'
      + '<input data-h="' + escapeHtml_(f.h) + '" type="' + type + '"' + (type === 'number' ? ' step="any"' : '') + '></label>';
  };
  // Essentials show up front; rarely-touched fields fold into "More details".
  const essential = spec.fields.filter(function (f) { return !f.auto && !f.adv; }).map(renderField).join('');
  const advanced = spec.fields.filter(function (f) { return !f.auto && f.adv; }).map(renderField).join('');
  return '<div class="reg-ov" id="reg-ov" hidden><div class="reg-ov-bd" onclick="regClose()"></div>'
    + '<div class="reg-ov-card"><div class="reg-ov-h" id="reg-ov-title">Add</div>'
    + '<div class="reg-form">' + essential + '</div>'
    + (advanced ? '<details class="reg-more"><summary class="reg-more-sum"><span class="reg-more-ic">+</span> More details</summary><div class="reg-form reg-form-more">' + advanced + '</div></details>' : '')
    + '<div class="reg-ov-foot"><span id="reg-msg" class="reg-msg"></span>'
    + '<button type="button" class="btn btn-ghost" onclick="regClose()">Cancel</button>'
    + '<button type="button" class="btn btn-confirm" id="reg-save" onclick="regSave()">Save</button></div>'
    + '</div></div>';
}

function regEditJs_() {
  return '<script>'
    + 'var REG_EROW=null,REG_EKEY=null,REG_REDIRKEY=null;'
    // Navigate the TOP frame using the ABSOLUTE /exec URL - the same thing the card
    // links do. A relative "location.href=?registry=..." resolves against the sandbox
    // iframe and lands on a blank page (this was the check-out / return / count white
    // screen). window.top is cross-origin but navigation (setting location) is allowed.
    + 'function regNav(qs){regWait(true,"Opening\\u2026","Loading the next page");var b=(typeof REG_BASE!=="undefined"&&REG_BASE)?REG_BASE:"";var u=b+(b.indexOf("?")>=0?"&":"?")+qs;if(typeof REG_EMBED!=="undefined"&&REG_EMBED){location.href=u;return;}try{window.top.location.href=u;}catch(e){location.href=u;}}'
    + 'function regInputs(){return [].slice.call(document.querySelectorAll("#reg-ov .reg-form input[data-h], #reg-ov .reg-form select[data-h]"));}'
    + 'function regFill(v){regInputs().forEach(function(i){var val=(v&&v[i.dataset.h]!=null)?v[i.dataset.h]:"";if(i.tagName==="SELECT"&&val){var found=false;[].forEach.call(i.options,function(o){if(o.value===val)found=true;});if(!found){var o=document.createElement("option");o.value=val;o.textContent=val;i.appendChild(o);}}i.value=val;});var g=document.getElementById("reg-past");if(g)g.hidden=true;regPreview();regToggleBattery();}'
    + 'function regToggleBattery(){var cat="";regInputs().forEach(function(i){if(i.dataset.h==="Category")cat=i.value;});var show=cat==="Batteries";[].slice.call(document.querySelectorAll("#reg-ov .reg-form [data-group=\\"battery\\"]")).forEach(function(el){el.style.display=show?"":"none";});if(show){var d=document.querySelector("#reg-ov .reg-more");if(d)d.open=true;}}'
    + 'function regPreview(){var el=document.getElementById("reg-imgurl");var p=document.getElementById("reg-preview");if(!p)return;var u=el?el.value.trim():"";p.innerHTML=/^https?:|^data:/.test(u)?("<img src=\\""+u.replace(/"/g,"%22")+"\\" onerror=\\"this.parentNode.textContent=\'No photo\'\\">"):"No photo";}'
    + 'function regSetImg(u){var el=document.getElementById("reg-imgurl");if(el)el.value=u;regPreview();}'
    + 'function regSearchWeb(){var name="";regInputs().forEach(function(i){if(i.dataset.h==="Item"||i.dataset.h==="Name"){if(i.value)name=i.value;}});if(!name){regMsg("Type the name first, then search.",true);return;}window.open("https://www.google.com/search?tbm=isch&q="+encodeURIComponent(name),"_blank");regMsg("Opened image search. Right-click a photo, Copy image address, paste it above.");}'
    + 'function regUpload(input){var f=input.files&&input.files[0];if(!f)return;if(!/^image\\//.test(f.type)){alert("Choose an image file.");input.value="";return;}if(f.size>10485760){alert("Image too large (max 10 MB).");input.value="";return;}var rd=new FileReader();rd.onload=function(){regMsg("Uploading photo...");google.script.run.withSuccessHandler(function(res){input.value="";if(res&&res.ok&&res.url){regSetImg(res.url);regMsg("Photo added.");}else{regMsg((res&&res.error)||"Upload failed.",true);}}).withFailureHandler(function(e){regMsg(String(e&&e.message||e),true);}).regUploadImage(rd.result,f.name);};rd.onerror=function(){regMsg("Could not read that file.",true);};rd.readAsDataURL(f);}'
    + 'function regTogglePast(){var g=document.getElementById("reg-past");if(!g)return;if(!g.hidden){g.hidden=true;return;}g.hidden=false;if(g.getAttribute("data-loaded"))return;g.innerHTML="<span class=\\"reg-msg\\">Loading...</span>";'
    + 'google.script.run.withSuccessHandler(function(res){g.setAttribute("data-loaded","1");var imgs=(res&&res.images)||[];if(!imgs.length){g.innerHTML="<span class=\\"reg-msg\\">No past photos yet.</span>";return;}g.innerHTML=imgs.map(function(u){var e=u.replace(/"/g,"%22");return "<button type=\\"button\\" class=\\"reg-past-img\\" data-u=\\""+e+"\\" onclick=\\"regSetImg(this.getAttribute(\'data-u\'));document.getElementById(\'reg-past\').hidden=true;\\"><img src=\\""+e+"\\" onerror=\\"this.parentNode.remove()\\"></button>";}).join("");}).withFailureHandler(function(){g.innerHTML="<span class=\\"reg-msg\\">Could not load.</span>";}).regPastImages();}'
    + 'function regCollect(){var o={};regInputs().forEach(function(i){o[i.dataset.h]=i.value;});return o;}'
    + 'function regShow(){document.getElementById("reg-ov").hidden=false;}'
    + 'function regClose(){document.getElementById("reg-ov").hidden=true;}'
    + 'function regMsg(m,err){var e=document.getElementById("reg-msg");if(e){e.textContent=m||"";e.style.color=err?"#b31b1b":"#8a857c";}}'
    + 'function regOpenAdd(){REG_EROW=null;REG_EKEY=null;document.getElementById("reg-ov-title").textContent="Add";regFill({});regMsg("");regShow();var f=document.querySelector("#reg-ov .reg-form input[data-h]");if(f)f.focus();}'
    + 'function regOpenEdit(row){function go(){REG_EROW=row;var v=REG_ROWS[row]||{};REG_EKEY=v[REG_KEY];document.getElementById("reg-ov-title").textContent="Edit";regFill(v);regMsg("");regShow();}if(REG_ROWS[row]){go();return;}'
    + 'regWait(true,"Loading item\\u2026","Fetching the full record");google.script.run.withSuccessHandler(function(j){regWait(false);try{REG_ROWS=JSON.parse(j||"{}");}catch(e){}go();}).withFailureHandler(function(){regWait(false);go();}).regRowsMap(REG_WHICH);}'
    + 'function regAfter(res){var s=document.getElementById("reg-save");if(s)s.disabled=false;if(res&&res.ok){regClose();regRefresh();}else{regMsg((res&&res.error)||"Could not save.",true);}}'
    + 'function regSave(){var vals=regCollect();REG_REDIRKEY=vals[REG_KEY]||REG_EKEY;var s=document.getElementById("reg-save");if(s)s.disabled=true;regMsg("Saving...");'
    + 'var fail=function(e){if(s)s.disabled=false;regMsg(String(e&&e.message||e),true);};'
    + 'if(REG_EROW==null){google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regAdd(REG_WHICH,vals);}'
    + 'else{google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regUpdateRow(REG_WHICH,REG_EROW,REG_EKEY,vals);}}'
    + 'function regDeleteRow(row,btn){var card=btn&&btn.closest?btn.closest(".reg-row"):null;var v=REG_ROWS[row]||{};var key=v[REG_KEY]||(card&&card.getAttribute("data-key"))||"";if(!confirm("Delete \\""+key+"\\"? This cannot be undone."))return;if(btn)btn.disabled=true;'
    + 'if(typeof REG_ITEM!=="undefined"&&REG_ITEM)regWait(true,"Deleting\\u2026","Removing this item");'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){if(typeof REG_ITEM!=="undefined"&&REG_ITEM){var slide=document.getElementById("reg-slide");if(slide){document.querySelectorAll(".reg-card").forEach(function(c){if(c.getAttribute("data-key")===key)c.remove();});delete REG_ROWS[row];regWait(false);regCloseItem();}else{REG_SAVED="";regNav("registry="+REG_WHICH+"&admin=1");}}else{var c=btn&&btn.closest?btn.closest(".reg-row"):null;if(c)c.remove();delete REG_ROWS[row];}}else{if(btn)btn.disabled=false;regWait(false);alert((res&&res.error)||"Could not delete.");}}).withFailureHandler(function(e){if(btn)btn.disabled=false;regWait(false);alert(String(e));}).regDelete(REG_WHICH,row,key);}'
    + 'function regReloadItem(key){REG_SAVED="";regWait(true,"Updating\\u2026","Refreshing this item");'
    + 'google.script.run.withSuccessHandler(function(r){if(!r||!r.html){regNav("registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(key)+"&admin=1");return;}'
    + 'var slide=document.getElementById("reg-slide");var pane=document.getElementById("reg-itempane");'
    + 'var target=(slide&&slide.classList.contains("is-open")&&pane)?pane:document.getElementById("reg-swap");'
    + 'if(target)target.innerHTML=r.html;REG_ITEM=1;if(r.row!=null&&r.vals)REG_ROWS[String(r.row)]=r.vals;regWait(false);})'
    + '.withFailureHandler(function(){regNav("registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(key)+"&admin=1");}).regItemHtml(REG_WHICH,key,true);}'
    + 'function regBtnBusy(btn,label){if(!btn)return;btn.disabled=true;if(btn.getAttribute("data-txt")==null)btn.setAttribute("data-txt",btn.innerHTML);btn.innerHTML="<span class=\\"reg-spin\\"></span>"+(label||"Working\\u2026");}'
    + 'function regBtnReset(btn){if(!btn)return;btn.disabled=false;var t=btn.getAttribute("data-txt");if(t!=null){btn.innerHTML=t;btn.removeAttribute("data-txt");}}'
    + 'function regCheckoutOne(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";var person=(document.getElementById("co-person")||{}).value||"";var due=(document.getElementById("co-due")||{}).value||"";if(!person.trim()){alert("Enter who is taking it.");return;}regBtnBusy(btn,"Checking out\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not check out.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regCheckout(REG_WHICH,row,key,person,due);}'
    + 'function regReturnOne(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";regBtnBusy(btn,"Returning\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not update.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regReturn(REG_WHICH,row,key);}'
    + 'function regStep(d){var el=document.getElementById("co-count");if(!el)return;el.value=Math.max(0,(parseInt(el.value,10)||0)+d);}'
    // Inline +/- on an inventory card: bump the number, then auto-save after a short pause
    // (debounced per row) so counting a shelf never opens the item page.
    + 'var REG_QT={};'
    + 'function regQadj(row,btn,delta){var card=btn.closest(".reg-card");if(!card)return;var span=card.querySelector(".reg-cn");var n=Math.max(0,(parseInt(span.textContent,10)||0)+delta);span.textContent=n;'
    + 'var v=REG_ROWS[row]||{};var key=v[REG_KEY]||card.getAttribute("data-key")||"";if(v)v["On hand"]=String(n);'
    + 'var rp=parseInt((v["Reorder point"]||card.getAttribute("data-rp")||"0"),10)||0;var st=n<=0?"out":(rp>0&&n<=rp?"low":"");card.dataset.stock=st;'
    + 'var qb=card.querySelector(".reg-card-qty");if(qb){var u=card.getAttribute("data-unit")||"";qb.textContent=n+(u?" "+u:"");}'
    + 'var chip=card.querySelector(".reg-chip");if(chip){if(!st){chip.parentNode.removeChild(chip);}else{chip.className="reg-chip"+(st==="low"?" reg-chip-low":"");chip.textContent=st==="out"?"Out of stock":"Low stock";}}else if(st){var b=card.querySelector(".reg-card-b");if(b){var sp=document.createElement("span");sp.className="reg-chip"+(st==="low"?" reg-chip-low":"");sp.textContent=st==="out"?"Out of stock":"Low stock";var mt=b.querySelector(".reg-card-meta");if(mt)b.insertBefore(sp,mt);else b.appendChild(sp);}}'
    + 'var sv=card.querySelector(".reg-csaved");if(sv){sv.textContent="\\u2026";sv.className="reg-csaved saving";}'
    + 'clearTimeout(REG_QT[row]);REG_QT[row]=setTimeout(function(){'
    + 'google.script.run.withSuccessHandler(function(res){if(sv){if(res&&res.ok){sv.textContent="\\u2713";sv.className="reg-csaved ok";setTimeout(function(){if(sv&&sv.className.indexOf("ok")>=0)sv.textContent="";},1400);}else{sv.textContent="!";sv.className="reg-csaved bad";}}})'
    + '.withFailureHandler(function(){if(sv){sv.textContent="!";sv.className="reg-csaved bad";}}).regCount(row,key,n);},600);}'
    + 'function regCountSave(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";var el=document.getElementById("co-count");var n=el?el.value:"";regBtnBusy(btn,"Saving\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not save.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regCount(row,key,n);}'
    + 'function regRefresh(){if(typeof REG_ITEM!=="undefined"&&REG_ITEM){regReloadItem(REG_REDIRKEY||REG_EKEY||"");return;}'
    + 'regWait(true,"Refreshing the catalog\\u2026","Updating items and stock");'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){var el=document.getElementById("reg-cards");if(el)el.innerHTML=res.html;REG_ROWS=JSON.parse(res.mapJson);if(typeof flt==="function")flt();}regWait(false);})'
    + '.withFailureHandler(function(){regWait(false);}).regRowsHtml(REG_WHICH);}'
    + 'function regScan(){if(!("BarcodeDetector" in window)){alert("Live scanning is not available in this browser. Use Print labels and scan them with your phone camera instead.");return;}'
    + 'navigator.mediaDevices.getUserMedia({video:{facingMode:"environment"}}).then(function(stream){regRunScan(stream);}).catch(function(){alert("Camera is blocked here (common in this app). Use Print labels and scan with your phone camera instead.");});}'
    + 'function regRunScan(stream){var ov=document.createElement("div");ov.style.cssText="position:fixed;inset:0;z-index:90;background:#000;display:flex;flex-direction:column;align-items:center;justify-content:center";'
    + 'var vid=document.createElement("video");vid.setAttribute("playsinline","");vid.style.cssText="max-width:100%;max-height:80vh";vid.srcObject=stream;vid.play();'
    + 'var btn=document.createElement("button");btn.textContent="Close";btn.className="btn btn-ghost";btn.style.margin="16px";btn.onclick=function(){stream.getTracks().forEach(function(t){t.stop();});ov.remove();};'
    + 'ov.appendChild(vid);ov.appendChild(btn);document.body.appendChild(ov);'
    + 'var det=new BarcodeDetector();var loop=function(){if(!ov.isConnected)return;det.detect(vid).then(function(codes){if(codes&&codes.length){var val=codes[0].rawValue||"";stream.getTracks().forEach(function(t){t.stop();});ov.remove();var q=document.getElementById("q");if(q){q.value=val;if(typeof flt==="function")flt();}}else{requestAnimationFrame(loop);}}).catch(function(){requestAnimationFrame(loop);});};requestAnimationFrame(loop);}'
    + '</script>';
}

// ---- registry server writes. The admin page is the gate (reached via ?admin=1);
// these mirror the app's existing admin actions, which are not passcode-checked. ----

function regSheet_(which) {
  const name = String(which).toLowerCase() === 'inventory' ? 'Inventory' : 'Equipment';
  const sh = regSs_().getSheetByName(name);
  if (!sh) throw new Error('No "' + name + '" tab. Run the registry setup first.');
  return sh;
}

function regHeaderMap_(sh) {
  const hs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const map = {}; hs.forEach(function (h, i) { map[h] = i; });
  return { headers: hs, map: map };
}

function regNextId_(sh, hm, key, prefix) {
  const ci = hm.map[key];
  let max = 0;
  if (ci == null) return prefix + '001';
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][ci]).trim();
    if (id.indexOf(prefix) === 0) { const n = parseInt(id.slice(prefix.length), 10); if (!isNaN(n) && n > max) max = n; }
  }
  let s = String(max + 1); while (s.length < 3) s = '0' + s;
  return prefix + s;
}

function regAdd(which, vals) {
  vals = vals || {};
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  const reqF = spec.fields.filter(function (f) { return f.req; })[0];
  if (reqF && !String(vals[reqF.h] || '').trim()) return { ok: false, error: reqF.label + ' is required.' };
  const row = hm.headers.map(function (h) { return vals[h] != null ? vals[h] : ''; });
  if (spec.idPrefix) { const ki = hm.map[spec.key]; if (ki != null && !String(row[ki]).trim()) row[ki] = regNextId_(sh, hm, spec.key, spec.idPrefix); }
  sh.appendRow(row);
  regInvalidate_();
  return { ok: true };
}

function regUpdateRow(which, rowNum, expectedKey, vals) {
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map[spec.key];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedKey).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  hm.headers.forEach(function (h, i) { if (vals[h] != null && !(spec.idPrefix && i === ki)) cur[i] = vals[h]; });
  sh.getRange(rowNum, 1, 1, hm.headers.length).setValues([cur]);
  regInvalidate_();
  return { ok: true };
}

function regDelete(which, rowNum, expectedKey) {
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map[spec.key];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedKey).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  sh.deleteRow(rowNum);
  regInvalidate_();
  return { ok: true };
}

function regRestock(rowNum, expectedItem) {
  const sh = regSheet_('inventory');
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { ok: false, error: 'Bad row.' };
  const ki = hm.map['Item'];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedItem).trim()) return { ok: false, error: 'This item moved. Refresh and try again.' };
  const oni = hm.map['On hand'], rqi = hm.map['Reorder qty'], lri = hm.map['Last restocked'];
  const qty = (rqi != null && Number(cur[rqi]) > 0) ? Number(cur[rqi]) : 0;
  if (oni != null) cur[oni] = (Number(cur[oni]) || 0) + qty;
  if (lri != null) cur[lri] = new Date();
  sh.getRange(rowNum, 1, 1, hm.headers.length).setValues([cur]);
  regInvalidate_();
  return { ok: true };
}

// ---- check out / return (equipment) + quick count (inventory) ----

function regRowGuard_(which, rowNum, expectedKey) {
  const spec = regFields_(which);
  const sh = regSheet_(which);
  const hm = regHeaderMap_(sh);
  rowNum = Number(rowNum);
  if (!(rowNum >= 2 && rowNum <= sh.getLastRow())) return { error: 'Bad row.' };
  const ki = hm.map[spec.key];
  const cur = sh.getRange(rowNum, 1, 1, hm.headers.length).getValues()[0];
  if (ki != null && String(cur[ki]).trim() !== String(expectedKey).trim()) return { error: 'This item moved. Refresh and try again.' };
  return { spec: spec, sh: sh, hm: hm, rowNum: rowNum, cur: cur };
}

function regCheckout(which, rowNum, expectedKey, person, due) {
  const g = regRowGuard_(which, rowNum, expectedKey);
  if (g.error) return { ok: false, error: g.error };
  person = String(person || '').trim();
  if (!person) return { ok: false, error: 'Enter who is taking it.' };
  const now = new Date();
  const dueDate = due ? new Date(due) : '';
  const set = function (h, v) { if (g.hm.map[h] != null) g.cur[g.hm.map[h]] = v; };
  set('Checked out to', person); set('Out since', now); set('Due back', dueDate);
  g.sh.getRange(g.rowNum, 1, 1, g.hm.headers.length).setValues([g.cur]);
  const nameI = g.hm.map[g.spec.tab === 'Inventory' ? 'Item' : 'Name'];
  regLogCheckout_(nameI != null ? g.cur[nameI] : '', expectedKey, person, now, dueDate);
  regInvalidate_();
  return { ok: true };
}

function regReturn(which, rowNum, expectedKey) {
  const g = regRowGuard_(which, rowNum, expectedKey);
  if (g.error) return { ok: false, error: g.error };
  ['Checked out to', 'Out since', 'Due back'].forEach(function (h) { if (g.hm.map[h] != null) g.cur[g.hm.map[h]] = ''; });
  g.sh.getRange(g.rowNum, 1, 1, g.hm.headers.length).setValues([g.cur]);
  regCloseCheckoutLog_(expectedKey);
  regInvalidate_();
  return { ok: true };
}

function regCount(rowNum, expectedItem, onHand) {
  const g = regRowGuard_('inventory', rowNum, expectedItem);
  if (g.error) return { ok: false, error: g.error };
  const n = Number(onHand);
  if (isNaN(n) || n < 0) return { ok: false, error: 'Enter a valid count.' };
  if (g.hm.map['On hand'] != null) g.cur[g.hm.map['On hand']] = n;
  if (g.hm.map['Last counted'] != null) g.cur[g.hm.map['Last counted']] = new Date();
  g.sh.getRange(g.rowNum, 1, 1, g.hm.headers.length).setValues([g.cur]);
  regInvalidate_();
  return { ok: true, onHand: n };
}

function regLogCheckout_(name, assetId, person, out, due) {
  const sh = regSs_().getSheetByName('Checkouts');
  if (!sh) return;
  const hm = regHeaderMap_(sh);
  const row = hm.headers.map(function () { return ''; });
  const set = function (h, v) { if (hm.map[h] != null) row[hm.map[h]] = v; };
  set('Timestamp', new Date()); set('Item', name); set('Asset ID', assetId);
  set('Checked out to', person); set('Out since', out); set('Due back', due);
  sh.appendRow(row);
}

function regCloseCheckoutLog_(assetId) {
  const sh = regSs_().getSheetByName('Checkouts');
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return;
  const hm = regHeaderMap_(sh);
  const ai = hm.map['Asset ID'], ri = hm.map['Returned'];
  if (ai == null || ri == null) return;
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][ai]).trim() === String(assetId).trim() && !String(v[i][ri]).trim()) {
      sh.getRange(i + 1, ri + 1).setValue(new Date());
      return;
    }
  }
}

// Save an uploaded photo to Drive (reusing the tasks/projects uploads folder) and
// return a hotlinkable thumbnail URL to store in the Image column.
function regUploadImage(dataUrl, filename) {
  try {
    const id = tpSaveUpload_(dataUrl, filename || ('item_' + Date.now() + '.jpg'));
    return { ok: true, url: 'https://lh3.googleusercontent.com/d/' + id + '=w600' };
  } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}

// Distinct photos already used across Equipment + Inventory, so one can be reused.
function regPastImages() {
  const out = [], seen = {};
  ['Equipment', 'Inventory'].forEach(function (tab) {
    const data = readTabRows_(tab);
    let ci = -1; data.headers.forEach(function (h, i) { if (norm_(h) === 'image') ci = i; });
    if (ci < 0) return;
    data.rows.forEach(function (rr) {
      const u = String(rr.cells[ci] || '').trim();
      if (/^https?:\/\//i.test(u) && !seen[u]) { seen[u] = 1; out.push(regThumb_(u, 200) || u); }
    });
  });
  return { ok: true, images: out.slice(0, 60) };
}

function regCell_(v) {
  if (v instanceof Date) return escapeHtml_(fmtShort_(v));
  const s = String(v == null ? '' : v).trim();
  if (/^https?:\/\//i.test(s)) return '<a href="' + escapeHtml_(s) + '" target="_blank" rel="noopener">link</a>';
  return escapeHtml_(s);
}

function isNum_(v) { return v !== '' && v != null && !isNaN(Number(v)); }

function readTab_(name) {
  var data = readTabRows_(name);
  return { headers: data.headers, rows: data.rows.map(function (rr) { return rr.cells; }) };
}
