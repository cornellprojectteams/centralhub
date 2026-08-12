/**
 * Equipment & Inventory registry - the "Ops registry" module (the ?registry=... pages
 * and ?module=registry-dash dashboard). Split out of 02_notify_on_submit.gs for
 * maintainability. Apps Script shares ONE global scope across all .gs files, so doGet
 * (in 02) routes straight into these functions with no import needed. Equipment and
 * Inventory run the same code, branched by regFields_(which).
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
  const sh = registrySs_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  const v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [] };
  const headers = v[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return String(c).trim() === ''; })) continue;
    rows.push({ row: i + 1, cells: v[i] });
  }
  return { headers: headers, rows: rows };
}

function regRawVal_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (v === 0) return '0';
  return String(v == null ? '' : v);
}

// Build the product cards + a {rowNumber: {header: value}} map, reused by the page and refreshes.
function regBuildCards_(which, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const data = readTabRows_(spec.tab);
  const H = data.headers, rows = data.rows;
  const idx = function (n) { for (let i = 0; i < H.length; i++) { if (norm_(H[i]) === norm_(n)) return i; } return -1; };
  const cName = idx(inv ? 'Item' : 'Name'), cImg = idx('Image'), cKey = idx(spec.key), cTeam = idx('Owning team'),
    cOn = idx('On hand'), cRe = idx('Reorder point'), cUnit = idx('Unit'), cLoc = idx('Location'),
    cSup = idx('Supplier'), cCat = idx('Category'), cStatus = idx('Status'), cCoTo = idx('Checked out to'), cDue = idx('Due back');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let base = ''; try { base = ScriptApp.getService().getUrl(); } catch (e) { base = ''; }
  let html = ''; const mapParts = [], owners = {}, cats = {};
  rows.forEach(function (rr) {
    const r = rr.cells;
    const name = cName >= 0 ? String(r[cName]).trim() : '';
    const key = cKey >= 0 ? String(r[cKey]).trim() : '';
    const owner = cTeam >= 0 ? String(r[cTeam]).trim() : ''; if (owner) owners[owner] = 1;
    const cat = cCat >= 0 ? String(r[cCat]).trim() : ''; if (cat) cats[cat] = 1;
    const hay = r.map(function (c) { return String(c); }).join(' ').toLowerCase();
    const img = cImg >= 0 ? String(r[cImg]).trim() : '';
    const imgHtml = (/^https?:\/\//i.test(img) || img.indexOf('data:') === 0)
      ? '<img src="' + escapeHtml_(img) + '" loading="lazy" alt="" onerror="this.remove()">'
      : '<span class="reg-card-noimg">No photo</span>';
    let primary = '', chip = '', coState = '', stockState = '', stepper = '';
    if (inv) {
      const on = cOn >= 0 && isNum_(r[cOn]) ? Number(r[cOn]) : null;
      const re = cRe >= 0 && isNum_(r[cRe]) ? Number(r[cRe]) : 0;
      const unit = cUnit >= 0 ? String(r[cUnit]).trim() : '';
      if (on !== null) primary = '<span class="reg-card-num">' + on + '</span><span class="reg-card-num-sub">' + escapeHtml_((unit ? unit + ' ' : '') + 'on hand') + '</span>';
      stockState = (on !== null && on <= 0) ? 'out' : ((on !== null && re > 0 && on > 0 && on <= re) ? 'low' : '');
      chip = stockState === 'out' ? '<span class="reg-chip">Out of stock</span>' : (stockState === 'low' ? '<span class="reg-chip reg-chip-low">Low stock</span>' : '');
      // Admins adjust stock inline on the card (auto-saves), so the number lives in the
      // stepper instead of a static line - no need to open the item just to count.
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
      if (outTo && cDue >= 0 && String(r[cDue]).trim()) { const d = r[cDue] instanceof Date ? new Date(r[cDue]) : new Date(r[cDue]); if (!isNaN(d.getTime())) { d.setHours(0, 0, 0, 0); overdue = d < today; } }
      if (overdue) { chip = '<span class="reg-chip">Overdue</span>'; coState = 'overdue'; }
      else if (outTo) { chip = '<span class="reg-chip reg-chip-low">Checked out</span>'; coState = 'out'; }
      else { const st = cStatus >= 0 ? String(r[cStatus]).trim() : ''; chip = st && /down|out of service|repair|broken/i.test(st) ? '<span class="reg-chip">' + escapeHtml_(st) + '</span>' : ''; }
    }
    const tag = cat ? '<div class="reg-card-tag">' + escapeHtml_(cat) + '</div>' : '';
    const metaBits = [];
    if (cLoc >= 0 && String(r[cLoc]).trim()) metaBits.push(String(r[cLoc]).trim());
    if (owner) metaBits.push(owner);
    if (inv && cSup >= 0 && String(r[cSup]).trim()) metaBits.push(String(r[cSup]).trim());
    const meta = metaBits.length ? '<div class="reg-card-meta">' + escapeHtml_(metaBits.join(' · ')) + '</div>' : '';
    const href = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'registry=item&which=' + spec.which + '&id=' + encodeURIComponent(key) + (admin ? '&admin=1' : '');
    html += '<div class="reg-card reg-row" data-hay="' + escapeHtml_(hay) + '" data-cat="' + escapeHtml_(cat) + '" data-owner="' + escapeHtml_(owner) + '" data-co="' + coState + '" data-stock="' + stockState + '">'
      + '<a class="reg-card-link" href="' + escapeHtml_(href) + '"><div class="reg-card-img">' + imgHtml
      + '<span class="reg-card-go" aria-hidden="true"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"></path></svg></span></div>'
      + '<div class="reg-card-b"><div class="reg-card-title">' + escapeHtml_(name || '(unnamed)') + '</div>' + tag
      + (primary ? '<div class="reg-card-primary">' + primary + '</div>' : '') + (chip || '') + meta + '</div></a>'
      + stepper
      + (admin ? '<div class="reg-card-act"><button type="button" class="reg-iconbtn" title="Edit" onclick="regOpenEdit(' + rr.row + ')">Edit</button><button type="button" class="reg-iconbtn reg-del" title="Delete" onclick="regDeleteRow(' + rr.row + ',this)">Delete</button></div>' : '')
      + '</div>';
    const obj = {}; H.forEach(function (h, i) { obj[h] = regRawVal_(r[i]); });
    mapParts.push(JSON.stringify(String(rr.row)) + ':' + JSON.stringify(obj));
  });
  return { html: html, mapJson: '{' + mapParts.join(',') + '}', owners: Object.keys(owners).sort(), cats: Object.keys(cats).sort(), key: spec.key };
}

// Client-callable: re-render the cards after a change (no page reload, which blanks the sandbox).
function regRowsHtml(which) {
  const b = regBuildCards_(which, true);
  return { ok: true, html: b.html, mapJson: b.mapJson };
}

function registryPage_(which, embedded, admin) {
  const spec = regFields_(which);
  const title = spec.tab === 'Inventory' ? 'Inventory' : 'Equipment registry';
  const b = regBuildCards_(which, admin);

  let regBase = ''; try { regBase = ScriptApp.getService().getUrl(); } catch (e) { regBase = CONFIG.webAppUrl || ''; }
  let inner = '<div id="reg-root"><div id="reg-swap">' + regBodyMarkup_(which, b, admin, embedded) + '</div>'
    + regStyles_() + regFilterJs_() + regSwitchJs_();
  if (admin) {
    inner += '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_ADMIN=true;var REG_BASE=' + JSON.stringify(regBase) + ';var REG_KEY=' + JSON.stringify(b.key)
      + ';var REG_ROWS=' + b.mapJson.replace(/<\//g, '<\\/') + ';</script>'
      + regEditJs_();
  } else {
    inner += '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_ADMIN=false;var REG_BASE=' + JSON.stringify(regBase) + ';</script>';
  }
  inner += '</div>';
  return swissShell_(inner, title, true, embedded);
}

// The swappable body of the registry page - everything the Inventory/Equipment toggle
// changes. Kept separate so regSwitchHtml() can re-render it IN PLACE via
// google.script.run: no page navigation means no Apps Script cold-load white flash
// (which is what the "back and forth" between tabs was hitting), and it lets the toggle
// show a loading spinner.
function regBodyMarkup_(which, b, admin, embedded) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  const title = inv ? 'Inventory' : 'Equipment registry';
  let base = ''; try { base = ScriptApp.getService().getUrl(); } catch (e) { base = CONFIG.webAppUrl || ''; }

  let head = '';
  if (!embedded) {
    head = '<div class="page-head"><div class="page-kicker">Ops registry</div>'
      + '<div class="page-title">' + escapeHtml_(title) + '</div><div class="page-rule"></div></div>';
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
  // Category filter shows the full canonical list (same as the add/edit form), plus any
  // stray legacy values; 'Other' last.
  const catList = [], catSeen = {};
  const pushCat = function (c) { if (c && !catSeen[c]) { catSeen[c] = 1; catList.push(c); } };
  REG_CATEGORIES.forEach(function (c) { if (c !== 'Other') pushCat(c); });
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

  let out = head + toggle + toolbar
    + '<div class="filters"><div class="search-wrap"><input id="q" type="search" placeholder="Search ' + escapeHtml_(spec.tab.toLowerCase()) + '" oninput="flt()"></div>' + controls + '</div>'
    + chipsHtml
    + '<div class="reg-grid" id="reg-cards">' + b.html + '</div>'
    + '<div id="empty" class="empty" style="display:none">Nothing matches those filters.</div>';
  if (!b.html) out += '<div class="empty">Nothing here yet.' + (admin ? ' Add the first ' + escapeHtml_(spec.noun) + '.' : '') + '</div>';
  if (admin) out += regFormOverlay_(which);
  return out;
}

// Client-callable: fresh body markup + row data for the other tab (drives the in-place toggle).
function regSwitchHtml(which, admin) {
  const spec = regFields_(which);
  const b = regBuildCards_(which, !!admin);
  return { ok: true, html: regBodyMarkup_(which, b, !!admin, false), which: spec.which, key: b.key, rowsJson: admin ? b.mapJson : '{}' };
}

// The in-place Inventory/Equipment toggle: swap the body via google.script.run instead
// of navigating, and dim + spin while it loads.
function regSwitchJs_() {
  return '<script>'
    + 'function regSwitch(which){if(typeof REG_WHICH!=="undefined"&&which===REG_WHICH)return;var sw=document.getElementById("reg-swap");if(sw)sw.classList.add("reg-loading");var ad=(typeof REG_ADMIN!=="undefined"&&REG_ADMIN);'
    + 'google.script.run.withSuccessHandler(function(r){if(sw)sw.classList.remove("reg-loading");if(!r||!r.ok)return;'
    + 'sw.innerHTML=r.html;REG_WHICH=r.which;if(ad){REG_KEY=r.key;try{REG_ROWS=JSON.parse(r.rowsJson);}catch(e){REG_ROWS={};}}'
    + 'if(typeof REG_FILTER!=="undefined")REG_FILTER="";if(typeof flt==="function")flt();})'
    + '.withFailureHandler(function(){if(sw)sw.classList.remove("reg-loading");}).regSwitchHtml(which, ad);}'
    + '</script>';
}

// Single-item view: the QR-label scan target. Focused card with quick actions.
function regItemPage_(which, id, admin) {
  const spec = regFields_(which);
  const inv = spec.tab === 'Inventory';
  // Absolute URL for the "Back" links (relative ones break in the sandbox iframe).
  let regBase = '';
  try { regBase = ScriptApp.getService().getUrl(); } catch (e) { regBase = CONFIG.webAppUrl || ''; }
  const backHref = regBase + (regBase.indexOf('?') >= 0 ? '&' : '?') + 'registry=' + spec.which + (admin ? '&admin=1' : '');
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  let found = null;
  data.rows.forEach(function (rr) { if (ki >= 0 && String(rr.cells[ki]).trim() === String(id).trim()) found = rr; });

  const head = '<div class="page-head"><div class="page-kicker">' + (inv ? 'Inventory item' : 'Equipment') + '</div>';
  if (!found) {
    return swissShell_('<div id="reg-root">' + head + '<div class="page-title">Not found</div><div class="page-rule"></div></div>'
      + '<div class="empty">No ' + escapeHtml_(spec.noun) + ' matches "' + escapeHtml_(id) + '".</div>'
      + '<p style="margin-top:16px"><a class="btn btn-ghost" href="' + escapeHtml_(backHref) + '">Back to ' + escapeHtml_(spec.tab.toLowerCase()) + '</a></p></div>', 'Not found', false, false);
  }
  const r = found.cells;
  const hi = function (n) { return H.map(function (h) { return norm_(h); }).indexOf(norm_(n)); };
  const nameI = hi(inv ? 'Item' : 'Name'), imgI = hi('Image');
  const name = nameI >= 0 ? String(r[nameI]) : String(id);

  const imgUrl = imgI >= 0 ? String(r[imgI]).trim() : '';
  const hasImg = /^https?:\/\//i.test(imgUrl) || imgUrl.indexOf('data:') === 0;
  const media = '<div class="reg-detail-media">' + (hasImg
    ? '<img src="' + escapeHtml_(imgUrl) + '" alt="" onerror="this.parentNode.innerHTML=\'<span class=&quot;reg-card-noimg&quot;>No photo</span>\'">'
    : '<span class="reg-card-noimg">No photo</span>') + '</div>';

  const toi = hi('Checked out to'), osi = hi('Out since'), dbi = hi('Due back'), oni = hi('On hand');
  const outTo = toi >= 0 ? String(r[toi]).trim() : '';
  const dueTxt = dbi >= 0 && String(r[dbi]).trim() ? regCell_(r[dbi]) : '';
  const overdue = (function () {
    if (!outTo || dbi < 0 || !String(r[dbi]).trim()) return false;
    const d = r[dbi] instanceof Date ? new Date(r[dbi]) : new Date(r[dbi]);
    if (isNaN(d.getTime())) return false;
    const t = new Date(); t.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0); return d < t;
  })();

  const skip = {}; skip[imgI] = 1; skip[nameI] = 1;
  if (!inv) { [toi, osi, dbi].forEach(function (x) { if (x >= 0) skip[x] = 1; }); }
  if (inv && oni >= 0) skip[oni] = 1;   // On hand is shown as the count stepper, not a duplicate row
  let fieldsHtml = '';
  H.forEach(function (h, i) {
    if (skip[i] || !String(r[i]).trim()) return;
    fieldsHtml += '<div class="reg-detail-kv"><span class="reg-detail-l">' + escapeHtml_(h) + '</span><span class="reg-detail-v">' + regCell_(r[i]) + '</span></div>';
  });

  let panel = '';
  if (!inv) {
    if (outTo) {
      const sub = (osi >= 0 && String(r[osi]).trim() ? 'out since ' + regCell_(r[osi]) : '') + (dueTxt ? (osi >= 0 && String(r[osi]).trim() ? ' · ' : '') + 'due ' + dueTxt : '');
      panel = '<div class="reg-co' + (overdue ? ' is-over' : '') + '"><div class="reg-co-h">' + (overdue ? 'Overdue' : 'Checked out') + '</div>'
        + '<div class="reg-co-who">' + escapeHtml_(outTo) + '</div>'
        + (sub ? '<div class="reg-co-sub">' + sub + '</div>' : '')
        + (admin ? '<button type="button" class="btn btn-confirm" onclick="regReturnOne(' + found.row + ',this)">Mark returned</button>' : '') + '</div>';
    } else if (admin) {
      panel = '<div class="reg-co"><div class="reg-co-h">Check out</div>'
        + '<div class="reg-co-form"><input id="co-person" placeholder="Who is taking it?"><input id="co-due" type="date" title="Due back"><button type="button" class="btn btn-primary" onclick="regCheckoutOne(' + found.row + ',this)">Check out</button></div></div>';
    }
  } else if (admin && oni >= 0) {
    const onNow = isNum_(r[oni]) ? Number(r[oni]) : 0;
    panel = '<div class="reg-co"><div class="reg-co-h">Count on hand</div>'
      + '<div class="reg-step"><button type="button" class="reg-stepbtn" onclick="regStep(-1)">&minus;</button>'
      + '<input id="co-count" type="number" inputmode="numeric" value="' + onNow + '">'
      + '<button type="button" class="reg-stepbtn" onclick="regStep(1)">+</button>'
      + '<button type="button" class="btn btn-confirm" onclick="regCountSave(' + found.row + ',this)">Save count</button></div></div>';
  }

  let actions = '';
  if (admin) {
    actions = '<div class="reg-detail-act">'
      + '<button type="button" class="btn btn-primary" onclick="regOpenEdit(' + found.row + ')">Edit</button>'
      + '<button type="button" class="btn btn-ghost reg-del" onclick="regDeleteRow(' + found.row + ',this)">Delete</button>'
      + '</div>';
  }

  let inner = '<div id="reg-root">' + head + '<div class="page-title">' + escapeHtml_(name) + '</div><div class="page-rule"></div></div>'
    + '<div class="reg-detail">' + media + '<div class="reg-detail-main">' + panel + fieldsHtml + actions + '</div></div>'
    + '<p style="margin-top:18px"><a class="btn btn-ghost" href="' + escapeHtml_(backHref) + '">Back to ' + escapeHtml_(spec.tab.toLowerCase()) + '</a></p>';

  inner += regStyles_();
  if (admin) {
    const one = {}; one[String(found.row)] = (function () { const o = {}; H.forEach(function (h, i) { o[h] = regRawVal_(r[i]); }); return o; })();
    inner += regFormOverlay_(which)
      + '<script>var REG_WHICH=' + JSON.stringify(spec.which) + ';var REG_ADMIN=true;var REG_BASE=' + JSON.stringify(regBase) + ';var REG_KEY=' + JSON.stringify(spec.key)
      + ';var REG_ROWS=' + JSON.stringify(one).replace(/<\//g, '<\\/') + ';var REG_ITEM=1;</script>'
      + regEditJs_();
  }
  inner += '</div>';
  return swissShell_(inner, name, true, false);
}

// Printable QR labels. Each label deep-links to its item page; scan with any phone camera.
function regLabelsPage_(which) {
  const spec = regFields_(which);
  const data = readTabRows_(spec.tab);
  const H = data.headers;
  const ki = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.key));
  const ni = H.map(function (h) { return norm_(h); }).indexOf(norm_(spec.tab === 'Inventory' ? 'Item' : 'Name'));
  let base = '';
  try { base = ScriptApp.getService().getUrl(); } catch (e) { base = ''; }

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

function regStyles_() {
  return '<style>'
    + '.reg-toggle{display:inline-flex;background:#efeee8;border-radius:999px;padding:4px;gap:2px;margin:16px 0 2px}'
    + '.reg-toggle-btn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13px;font-weight:700;color:#8a857c;text-decoration:none;padding:7px 20px;border-radius:999px;border:none;background:transparent;cursor:pointer;transition:background .15s,color .15s}'
    // in-place tab swap: dim the body + spinner while the other tab loads
    + '#reg-swap{position:relative}'
    + '#reg-swap.reg-loading{opacity:.45;pointer-events:none}'
    + '#reg-swap.reg-loading::after{content:"";position:absolute;left:50%;top:130px;width:34px;height:34px;margin-left:-17px;border:3px solid #e6e1d8;border-top-color:#b31b1b;border-radius:50%;animation:regspin .7s linear infinite}'
    + '@keyframes regspin{to{transform:rotate(360deg)}}'
    // small inline spinner shown in a button while its action runs (check out, return, save count)
    + '.reg-spin{display:inline-block;width:13px;height:13px;margin-right:7px;vertical-align:-2px;border:2px solid rgba(255,255,255,.45);border-top-color:#fff;border-radius:50%;animation:regspin .6s linear infinite}'
    + '.reg-toggle-btn.on{background:#fff;color:#14110e;box-shadow:0 1px 3px rgba(0,0,0,.12)}'
    + '.reg-tools{display:flex;gap:8px;flex-wrap:wrap;margin:16px 0 4px}'
    + '.reg-chip{display:inline-block;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#b31b1b;background:#fdecec;border:1px solid #f5d0d0;padding:3px 9px;border-radius:999px;align-self:flex-start}'
    + '.reg-chip-low{color:#b06a00;background:#fbf3e1;border-color:#eeddb4}'
    // ---- product card grid ----
    + '.reg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:16px;margin-top:18px}'
    + '.reg-card{position:relative;display:flex;flex-direction:column;background:#fff;border:1px solid #ececea;border-radius:14px;overflow:hidden;transition:box-shadow .18s ease,transform .18s ease,border-color .18s ease}'
    + '.reg-card:hover{box-shadow:0 12px 30px rgba(20,20,30,.13);transform:translateY(-3px);border-color:#e4dfd6}'
    + '.reg-card-link{display:flex;flex-direction:column;flex:1;min-height:0;text-decoration:none;color:inherit}'
    + '.reg-card-img{position:relative;aspect-ratio:1/1;background:#f6f4ef;display:flex;align-items:center;justify-content:center;overflow:hidden}'
    + '.reg-card-img img{width:100%;height:100%;object-fit:cover;display:block}'
    // persistent "opens" affordance so touch users (no hover) can tell a card is tappable
    + '.reg-card-go{position:absolute;right:8px;bottom:8px;width:26px;height:26px;border-radius:50%;background:rgba(255,255,255,.94);box-shadow:0 2px 7px rgba(20,20,30,.18);display:flex;align-items:center;justify-content:center;color:#8a857c;transition:color .16s ease,transform .16s ease}'
    + '.reg-card:hover .reg-card-go{color:#b31b1b;transform:translateX(1px)}'
    + '.reg-card-noimg{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:10px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#cbc4b8}'
    + '.reg-card-b{padding:12px 14px 14px;display:flex;flex-direction:column;gap:5px}'
    + '.reg-card-title{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:13.5px;font-weight:700;color:#1c1a17;line-height:1.3;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;min-height:2.5em}'
    + '.reg-card-tag{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:9.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#8a857c}'
    + '.reg-card-primary{display:flex;align-items:baseline;gap:5px;flex-wrap:wrap}'
    + '.reg-card-num{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:19px;font-weight:800;letter-spacing:-.02em;color:#14110e}'
    + '.reg-card-num-sub{font-size:12px;font-weight:600;color:#8a857c}'
    + '.reg-card-meta{font-size:12px;color:#8a857c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
    + '.reg-card-act{position:absolute;top:8px;right:8px;display:none;gap:6px}.reg-card:hover .reg-card-act{display:flex}'
    + '.reg-iconbtn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:700;color:#57534e;background:rgba(255,255,255,.96);border:1px solid #e2ddd6;border-radius:8px;padding:5px 10px;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.1)}'
    + '.reg-iconbtn:hover{border-color:#b5b0a8;color:#292524}.reg-iconbtn.reg-del:hover{border-color:#e6b3b3;color:#b31b1b}'
    // ---- inline stock stepper on inventory cards (adjust + auto-save, no page hop) ----
    + '.reg-count{display:flex;align-items:center;gap:8px;padding:9px 12px;border-top:1px solid #f0efe9;background:#fcfbf9}'
    + '.reg-cbtn{flex:0 0 auto;width:34px;height:34px;border-radius:9px;border:1.5px solid #e2ddd6;background:#fff;font-size:20px;font-weight:700;color:#26231f;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:border-color .14s,color .14s,background .14s}'
    + '.reg-cbtn:hover{border-color:#b31b1b;color:#b31b1b}.reg-cbtn:active{background:#f7efef;transform:scale(.96)}'
    + '.reg-cval{flex:1;min-width:0;display:flex;flex-direction:column;line-height:1.15;text-align:center}'
    + '.reg-cn{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;letter-spacing:-.02em;color:#14110e}'
    + '.reg-cu{font-size:10px;font-weight:700;letter-spacing:.04em;color:#a8a29e;text-transform:uppercase}'
    + '.reg-csaved{flex:0 0 auto;width:16px;text-align:center;font-size:13px;font-weight:800;transition:color .14s}'
    + '.reg-csaved.saving{color:#cbc4b8}.reg-csaved.ok{color:#157a47}.reg-csaved.bad{color:#b31b1b}'
    // ---- product detail ----
    + '.reg-detail{display:grid;grid-template-columns:minmax(0,340px) 1fr;gap:28px;align-items:start;margin-top:8px}'
    + '.reg-detail-media{aspect-ratio:1/1;background:#f6f4ef;border:1px solid #ececea;border-radius:16px;overflow:hidden;display:flex;align-items:center;justify-content:center}'
    + '.reg-detail-media img{width:100%;height:100%;object-fit:cover}'
    + '.reg-detail-main{display:flex;flex-direction:column}'
    + '.reg-detail-kv{display:flex;justify-content:space-between;gap:16px;padding:11px 0;border-bottom:1px solid #f0efe9}.reg-detail-kv:first-child{padding-top:0}'
    + '.reg-detail-l{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#a8a29e}'
    + '.reg-detail-v{font-size:14px;font-weight:600;color:#26231f;text-align:right}.reg-detail-v a{color:#b31b1b;text-decoration:none;border-bottom:1px solid #f0c050}'
    + '.reg-detail-act{display:flex;gap:8px;flex-wrap:wrap;margin-top:20px}'
    + '@media(max-width:640px){.reg-detail{grid-template-columns:1fr}.reg-detail-media{max-width:300px}}'
    // quick-filter chips
    + '.reg-chips{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 2px}'
    + '.reg-fchip{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:12px;font-weight:700;color:#57534e;background:#fff;border:1.5px solid #e2ddd6;border-radius:999px;padding:7px 14px;cursor:pointer}'
    + '.reg-fchip:hover{border-color:#b5b0a8}.reg-fchip.on{background:#b31b1b;border-color:#b31b1b;color:#fff}'
    // check-out / count panel
    + '.reg-co{background:#faf9f6;border:1.5px solid #ececea;border-radius:14px;padding:16px 18px;margin-bottom:16px}'
    + '.reg-co.is-over{background:#fdecec;border-color:#f5d0d0}'
    + '.reg-co-h{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:11px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#8a857c;margin-bottom:9px}'
    + '.reg-co.is-over .reg-co-h{color:#b31b1b}'
    + '.reg-co-who{font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:18px;font-weight:800;color:#14110e}'
    + '.reg-co-sub{font-size:13px;color:#8a857c;margin-top:3px;margin-bottom:12px}'
    + '.reg-co-form{display:flex;gap:8px;flex-wrap:wrap;align-items:center}'
    + '.reg-co-form input{font:inherit;font-size:14px;padding:11px 13px;border:1.5px solid #e0e0dc;border-radius:10px;outline:none;flex:1;min-width:150px}'
    + '.reg-co-form input:focus{border-color:#b31b1b;box-shadow:0 0 0 3px rgba(179,27,27,.1)}'
    + '.reg-step{display:flex;align-items:center;gap:10px;flex-wrap:wrap}'
    + '.reg-stepbtn{width:46px;height:46px;border-radius:12px;border:1.5px solid #e0e0dc;background:#fff;font-size:22px;font-weight:700;color:#26231f;cursor:pointer;line-height:1;display:flex;align-items:center;justify-content:center;flex:0 0 auto}'
    + '.reg-stepbtn:hover{border-color:#b5b0a8}.reg-stepbtn:active{background:#f0efe9}'
    + '.reg-step input{width:86px;font-family:"Plus Jakarta Sans",Helvetica,Arial,sans-serif;font-size:22px;font-weight:800;text-align:center;padding:9px;border:1.5px solid #e0e0dc;border-radius:12px;outline:none}'
    + '.reg-step input:focus{border-color:#b31b1b}'
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
    + 'var e=document.getElementById("empty");if(e)e.style.display=n?"none":"block";}'
    + '</script>';
}

function regFormOverlay_(which) {
  const spec = regFields_(which);
  // Owning team is a dropdown of the project teams (+ "Shared"); regFill keeps any
  // existing value that is not in the list when editing.
  let teamOpts = [];
  try { teamOpts = icTeamNames_(); } catch (e) { teamOpts = []; }
  ['Shared', 'Program', 'Facilities'].forEach(function (x) { if (teamOpts.indexOf(x) < 0) teamOpts = teamOpts.concat([x]); });
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
    + 'function regNav(qs){var b=(typeof REG_BASE!=="undefined"&&REG_BASE)?REG_BASE:"";var u=b+(b.indexOf("?")>=0?"&":"?")+qs;try{window.top.location.href=u;}catch(e){location.href=u;}}'
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
    + 'function regOpenEdit(row){REG_EROW=row;var v=REG_ROWS[row]||{};REG_EKEY=v[REG_KEY];document.getElementById("reg-ov-title").textContent="Edit";regFill(v);regMsg("");regShow();}'
    + 'function regAfter(res){var s=document.getElementById("reg-save");if(s)s.disabled=false;if(res&&res.ok){regClose();regRefresh();}else{regMsg((res&&res.error)||"Could not save.",true);}}'
    + 'function regSave(){var vals=regCollect();REG_REDIRKEY=vals[REG_KEY]||REG_EKEY;var s=document.getElementById("reg-save");if(s)s.disabled=true;regMsg("Saving...");'
    + 'var fail=function(e){if(s)s.disabled=false;regMsg(String(e&&e.message||e),true);};'
    + 'if(REG_EROW==null){google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regAdd(REG_WHICH,vals);}'
    + 'else{google.script.run.withSuccessHandler(regAfter).withFailureHandler(fail).regUpdateRow(REG_WHICH,REG_EROW,REG_EKEY,vals);}}'
    + 'function regDeleteRow(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";if(!confirm("Delete \\""+key+"\\"? This cannot be undone."))return;if(btn)btn.disabled=true;'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){if(typeof REG_ITEM!=="undefined"){regNav("registry="+REG_WHICH+"&admin=1");}else{var c=btn&&btn.closest?btn.closest(".reg-row"):null;if(c)c.remove();delete REG_ROWS[row];}}else{if(btn)btn.disabled=false;alert((res&&res.error)||"Could not delete.");}}).withFailureHandler(function(e){if(btn)btn.disabled=false;alert(String(e));}).regDelete(REG_WHICH,row,key);}'
    + 'function regReloadItem(key){regNav("registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(key)+"&admin=1");}'
    + 'function regBtnBusy(btn,label){if(!btn)return;btn.disabled=true;if(btn.getAttribute("data-txt")==null)btn.setAttribute("data-txt",btn.innerHTML);btn.innerHTML="<span class=\\"reg-spin\\"></span>"+(label||"Working\\u2026");}'
    + 'function regBtnReset(btn){if(!btn)return;btn.disabled=false;var t=btn.getAttribute("data-txt");if(t!=null){btn.innerHTML=t;btn.removeAttribute("data-txt");}}'
    + 'function regCheckoutOne(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";var person=(document.getElementById("co-person")||{}).value||"";var due=(document.getElementById("co-due")||{}).value||"";if(!person.trim()){alert("Enter who is taking it.");return;}regBtnBusy(btn,"Checking out\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not check out.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regCheckout(REG_WHICH,row,key,person,due);}'
    + 'function regReturnOne(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";regBtnBusy(btn,"Returning\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not update.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regReturn(REG_WHICH,row,key);}'
    + 'function regStep(d){var el=document.getElementById("co-count");if(!el)return;el.value=Math.max(0,(parseInt(el.value,10)||0)+d);}'
    // Inline +/- on an inventory card: bump the number, then auto-save after a short pause
    // (debounced per row) so counting a shelf never opens the item page.
    + 'var REG_QT={};'
    + 'function regQadj(row,btn,delta){var card=btn.closest(".reg-card");if(!card)return;var span=card.querySelector(".reg-cn");var n=Math.max(0,(parseInt(span.textContent,10)||0)+delta);span.textContent=n;'
    + 'var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";if(v)v["On hand"]=String(n);'
    // keep the Out/Low chip and the filter state in sync as the count changes
    + 'var rp=parseInt((v["Reorder point"]||"0"),10)||0;var st=n<=0?"out":(rp>0&&n<=rp?"low":"");card.dataset.stock=st;'
    + 'var chip=card.querySelector(".reg-chip");if(chip){if(!st){chip.parentNode.removeChild(chip);}else{chip.className="reg-chip"+(st==="low"?" reg-chip-low":"");chip.textContent=st==="out"?"Out of stock":"Low stock";}}else if(st){var b=card.querySelector(".reg-card-b");if(b){var sp=document.createElement("span");sp.className="reg-chip"+(st==="low"?" reg-chip-low":"");sp.textContent=st==="out"?"Out of stock":"Low stock";var mt=b.querySelector(".reg-card-meta");if(mt)b.insertBefore(sp,mt);else b.appendChild(sp);}}'
    + 'var sv=card.querySelector(".reg-csaved");if(sv){sv.textContent="\\u2026";sv.className="reg-csaved saving";}'
    + 'clearTimeout(REG_QT[row]);REG_QT[row]=setTimeout(function(){'
    + 'google.script.run.withSuccessHandler(function(res){if(sv){if(res&&res.ok){sv.textContent="\\u2713";sv.className="reg-csaved ok";setTimeout(function(){if(sv&&sv.className.indexOf("ok")>=0)sv.textContent="";},1400);}else{sv.textContent="!";sv.className="reg-csaved bad";}}})'
    + '.withFailureHandler(function(){if(sv){sv.textContent="!";sv.className="reg-csaved bad";}}).regCount(row,key,n);},600);}'
    + 'function regCountSave(row,btn){var v=REG_ROWS[row]||{};var key=v[REG_KEY]||"";var el=document.getElementById("co-count");var n=el?el.value:"";regBtnBusy(btn,"Saving\\u2026");google.script.run.withSuccessHandler(function(res){if(res&&res.ok){regReloadItem(key);}else{regBtnReset(btn);alert((res&&res.error)||"Could not save.");}}).withFailureHandler(function(e){regBtnReset(btn);alert(String(e));}).regCount(row,key,n);}'
    + 'function regRefresh(){if(typeof REG_ITEM!=="undefined"){regNav("registry=item&which="+REG_WHICH+"&id="+encodeURIComponent(REG_REDIRKEY||REG_EKEY||"")+"&admin=1");return;}'
    + 'google.script.run.withSuccessHandler(function(res){if(res&&res.ok){document.getElementById("reg-cards").innerHTML=res.html;REG_ROWS=JSON.parse(res.mapJson);if(typeof flt==="function")flt();}}).regRowsHtml(REG_WHICH);}'
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
  const sh = registrySs_().getSheetByName(name);
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
  return { ok: true };
}

function regReturn(which, rowNum, expectedKey) {
  const g = regRowGuard_(which, rowNum, expectedKey);
  if (g.error) return { ok: false, error: g.error };
  ['Checked out to', 'Out since', 'Due back'].forEach(function (h) { if (g.hm.map[h] != null) g.cur[g.hm.map[h]] = ''; });
  g.sh.getRange(g.rowNum, 1, 1, g.hm.headers.length).setValues([g.cur]);
  regCloseCheckoutLog_(expectedKey);
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
  return { ok: true, onHand: n };
}

function regLogCheckout_(name, assetId, person, out, due) {
  const sh = registrySs_().getSheetByName('Checkouts');
  if (!sh) return;
  const hm = regHeaderMap_(sh);
  const row = hm.headers.map(function () { return ''; });
  const set = function (h, v) { if (hm.map[h] != null) row[hm.map[h]] = v; };
  set('Timestamp', new Date()); set('Item', name); set('Asset ID', assetId);
  set('Checked out to', person); set('Out since', out); set('Due back', due);
  sh.appendRow(row);
}

function regCloseCheckoutLog_(assetId) {
  const sh = registrySs_().getSheetByName('Checkouts');
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
    return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w600' };
  } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}

// Distinct photos already used across Equipment + Inventory, so one can be reused.
function regPastImages() {
  const out = [], seen = {};
  ['Equipment', 'Inventory'].forEach(function (tab) {
    const sh = registrySs_().getSheetByName(tab);
    if (!sh) return;
    const v = sh.getDataRange().getValues();
    if (!v.length) return;
    let ci = -1; v[0].forEach(function (h, i) { if (norm_(h) === 'image') ci = i; });
    if (ci < 0) return;
    for (let i = 1; i < v.length; i++) {
      const u = String(v[i][ci] || '').trim();
      if (/^https?:\/\//i.test(u) && !seen[u]) { seen[u] = 1; out.push(u); }
    }
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
  const sh = registrySs_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  const v = sh.getDataRange().getValues();
  if (!v.length) return { headers: [], rows: [] };
  const headers = v[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(function (c) { return String(c).trim() === ''; })) continue;
    rows.push(v[i]);
  }
  return { headers: headers, rows: rows };
}
