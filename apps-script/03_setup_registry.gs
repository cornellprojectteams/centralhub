/**
 * Ops CMMS for the registry spreadsheet. Four connected tabs:
 *   Equipment   - assets across team areas (location, ownership, status)
 *   Inventory   - consumables and parts, with reorder points and purchasing info
 *   Maintenance - recurring preventative tasks on a schedule
 *   Action items- the operations feed: due maintenance and low stock land here,
 *                 assigned and with purchasing details, then get marked complete
 *
 * Run setupRegistry() once to build/upgrade the tabs, then installDailyTrigger()
 * to turn on the daily check that generates action items automatically.
 */
var REGISTRY_SS_ID = '1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU';
var OPS_NOTIFY = 'nhh5@cornell.edu';   // action-item notifications go here; add a team alias if you have one
var DEFAULT_ASSIGNEE = 'Noah Hamm';

var EQUIPMENT_TAB = 'Equipment';
var INVENTORY_TAB = 'Inventory';
var MAINTENANCE_TAB = 'Maintenance';
var ACTIONS_TAB = 'Action items';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Registry')
    .addItem('Add inventory item', 'showAddInventory')
    .addItem('Mark selected item restocked', 'markRestockedSelected')
    .addSeparator()
    .addItem('Set up / update tabs', 'setupRegistry')
    .addItem('Import existing tools & inventory', 'importExistingData')
    .addItem('Add photos from the web', 'seedPhotos')
    .addItem('Run maintenance check now', 'generateMaintenanceTasks')
    .addItem('Run stock check now', 'checkInventoryLevels')
    .addSeparator()
    .addItem('Turn on daily automatic checks', 'installDailyTrigger')
    .addToUi();
}

// Diagnostic: print every tab and its header row so the existing layout is visible.
function listRegistrySheet() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  Logger.log('Spreadsheet: ' + ss.getName());
  ss.getSheets().forEach(function (sh) {
    var lastCol = sh.getLastColumn();
    var headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Logger.log('Tab "' + sh.getName() + '" (' + sh.getLastRow() + ' rows): ' + headers.join(' | '));
  });
}

function setupRegistry() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);

  ensureTab_(ss, EQUIPMENT_TAB,
    ['Asset ID', 'Name', 'Category', 'Location', 'Owning team', 'Owner', 'Status', 'Installed', 'Notes', 'Image'],
    ['EX-001', 'Example: spray booth (delete this row)', 'Finishing', 'Composites lab', 'Operations Team', 'Noah Hamm', 'In service', '2024', 'Filters checked monthly', '']);

  ensureTab_(ss, INVENTORY_TAB,
    ['Item', 'Location', 'On hand', 'Unit', 'Reorder point', 'Reorder qty', 'Supplier', 'Product link', 'eShop info', 'Assign to', 'Last restocked', 'Image'],
    ['Example: spray booth filter (delete this row)', 'Composites lab', 0, 'each', 2, 6, 'Uline', '', 'eShop item #', 'Noah Hamm', '2026-05', '']);

  ensureTab_(ss, MAINTENANCE_TAB,
    ['Task ID', 'Task', 'Equipment / area', 'Frequency', 'Last done', 'Next due', 'Assign to', 'Instructions', 'Parts needed', 'Active'],
    ['PM-001', 'Check and replace spray booth filters', 'Composites lab', 'Monthly', '', new Date(), 'Noah Hamm', 'Inspect and replace if loaded', 'Spray booth filter', true]);

  ensureTab_(ss, ACTIONS_TAB,
    ['Item ID', 'Created', 'Source', 'Title', 'Location', 'Equipment / part', 'Assigned to', 'Status', 'Purchasing info', 'Due', 'Completed'],
    null);

  Logger.log('Registry ready: Equipment, Inventory, Maintenance, Action items.');
}

// Create the tab if missing, write headers + an example row when empty, and add any
// new columns to an existing tab without disturbing data. Columns are read by header
// name everywhere, so appended columns do not need to be in a particular order.
function ensureTab_(ss, name, headers, exampleRow) {
  var sh = ss.getSheetByName(name);
  var created = false;
  if (!sh) { sh = ss.insertSheet(name); created = true; }

  if (!String(sh.getRange(1, 1).getValue()).trim()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (exampleRow) sh.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
  } else {
    var lastCol = sh.getLastColumn();
    var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (c) { return String(c).trim(); });
    headers.forEach(function (h) {
      if (existing.indexOf(h) < 0) {
        lastCol += 1;
        sh.getRange(1, lastCol).setValue(h);
        existing.push(h);
      }
    });
  }

  var cols = sh.getLastColumn();
  sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#8f1515').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, cols);
  Logger.log((created ? 'Created' : 'Updated') + ' "' + name + '" tab.');
  return sh;
}

//------------------------------------------------------------
//
// AUTOMATIC CHECKS: due maintenance and low stock become action items.
//
//------------------------------------------------------------

// Daily entry point. Turn on with installDailyTrigger().
function dailyCmmsCheck() {
  generateMaintenanceTasks();
  checkInventoryLevels();
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyCmmsCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyCmmsCheck').timeBased().everyDays(1).atHour(7).create();
  Logger.log('Daily CMMS check installed (runs each morning).');
}

// Any active maintenance task whose Next due has arrived becomes an action item,
// then its Next due rolls forward by its Frequency so it repeats.
function generateMaintenanceTasks() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var pm = ss.getSheetByName(MAINTENANCE_TAB);
  if (!pm) { Logger.log('No "' + MAINTENANCE_TAB + '" tab. Run setupRegistry.'); return 0; }
  var data = pm.getDataRange().getValues();
  if (data.length < 2) return 0;
  var h = hmap_(data[0]);
  var tz = Session.getScriptTimeZone();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var made = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var task = String(row[h['Task']] || '').trim();
    if (!task) continue;
    var active = row[h['Active']];
    if (active === false || String(active).toLowerCase() === 'no') continue;
    var nd = row[h['Next due']];
    if (!nd) continue;
    var due = new Date(nd);
    if (isNaN(due.getTime())) continue;
    due.setHours(0, 0, 0, 0);
    if (due > today) continue;

    var taskId = String(row[h['Task ID']] || ('PM-' + (i + 1)));
    var key = 'pm:' + taskId + ':' + Utilities.formatDate(due, tz, 'yyyy-MM-dd');
    if (openActionExists_(ss, key)) continue;

    var details = [];
    if (row[h['Instructions']]) details.push('Instructions: ' + row[h['Instructions']]);
    if (row[h['Parts needed']]) details.push('Parts needed: ' + row[h['Parts needed']]);

    var area = row[h['Equipment / area']] || '';
    createActionItem_(ss, {
      source: key,
      title: task,
      location: area,
      equipment: area,
      assignee: row[h['Assign to']] || DEFAULT_ASSIGNEE,
      purchasing: details.join(' | '),
      due: due
    });

    pm.getRange(i + 1, h['Next due'] + 1).setValue(advanceDate_(due, row[h['Frequency']]));
    made.push(task);
  }

  if (made.length) notifyOps_(made.length + ' maintenance task(s) due', ['Due now:'].concat(made).concat(['', 'See the "' + ACTIONS_TAB + '" tab.']));
  Logger.log('Maintenance: created ' + made.length + ' action item(s).');
  return made.length;
}

// Any inventory item at or below its reorder point becomes a restock action item
// with the supplier and purchasing info attached.
function checkInventoryLevels() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var inv = ss.getSheetByName(INVENTORY_TAB);
  if (!inv) { Logger.log('No "' + INVENTORY_TAB + '" tab. Run setupRegistry.'); return 0; }
  var data = inv.getDataRange().getValues();
  if (data.length < 2) return 0;
  var h = hmap_(data[0]);
  var made = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var item = String(row[h['Item']] || '').trim();
    if (!item) continue;
    var reorder = Number(row[h['Reorder point']]);
    var onHand = Number(row[h['On hand']]);
    if (!(reorder > 0) || isNaN(onHand)) continue;
    if (onHand > reorder) continue;

    var key = 'restock:' + item;
    if (openActionExists_(ss, key)) continue;

    var parts = [];
    if (row[h['Supplier']]) parts.push('Supplier: ' + row[h['Supplier']]);
    if (row[h['Reorder qty']]) parts.push('Order qty: ' + row[h['Reorder qty']]);
    if (row[h['Product link']]) parts.push('Product: ' + row[h['Product link']]);
    if (row[h['eShop info']]) parts.push('eShop: ' + row[h['eShop info']]);

    createActionItem_(ss, {
      source: key,
      title: 'Restock ' + item + ' (on hand ' + onHand + ', reorder at ' + reorder + ')',
      location: row[h['Location']] || '',
      equipment: item,
      assignee: row[h['Assign to']] || DEFAULT_ASSIGNEE,
      purchasing: parts.join(' | ')
    });
    made.push(item);
  }

  if (made.length) notifyOps_(made.length + ' item(s) need restocking', ['Low stock:'].concat(made).concat(['', 'See the "' + ACTIONS_TAB + '" tab.']));
  Logger.log('Stock: created ' + made.length + ' action item(s).');
  return made.length;
}

//------------------------------------------------------------
//
// INVENTORY QUICK ACTIONS: add via a form, restock in one click.
//
//------------------------------------------------------------

function showAddInventory() {
  var html = HtmlService.createHtmlOutput(addInventoryHtml_()).setWidth(420).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add inventory item');
}

function addInventoryItem(data) {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var sh = ss.getSheetByName(INVENTORY_TAB);
  if (!sh) return { ok: false, error: 'No Inventory tab yet. Run "Set up / update tabs" first.' };
  var item = String(data.item || '').trim();
  if (!item) return { ok: false, error: 'Item name is required.' };

  var h = hmap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var rowArr = [];
  for (var i = 0; i < sh.getLastColumn(); i++) rowArr.push('');
  function set(name, val) { if (h[name] != null) rowArr[h[name]] = val; }
  function num(v) { return v === '' || v == null ? '' : Number(v); }
  set('Item', item);
  set('Location', data.location || '');
  set('On hand', num(data.onHand));
  set('Unit', data.unit || '');
  set('Reorder point', num(data.reorderPoint));
  set('Reorder qty', num(data.reorderQty));
  set('Supplier', data.supplier || '');
  set('Product link', data.productLink || '');
  set('eShop info', data.eshop || '');
  set('Assign to', data.assignTo || DEFAULT_ASSIGNEE);
  sh.appendRow(rowArr);
  ss.toast('Added ' + item, 'Inventory', 4);
  return { ok: true, item: item };
}

// Bumps On hand by the reorder qty, stamps Last restocked, and closes the matching
// open restock action item. Click the item's row on the Inventory tab first.
function markRestockedSelected() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (!sh || sh.getName() !== INVENTORY_TAB) { ui.alert('Open the Inventory tab and click the item row first.'); return; }
  var r = sh.getActiveRange().getRow();
  if (r < 2) { ui.alert('Click an item row, not the header.'); return; }

  var h = hmap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var item = String(sh.getRange(r, h['Item'] + 1).getValue() || '').trim();
  if (!item) { ui.alert('That row has no item.'); return; }

  var onHand = Number(sh.getRange(r, h['On hand'] + 1).getValue()) || 0;
  var qty = Number(sh.getRange(r, h['Reorder qty'] + 1).getValue()) || 0;
  sh.getRange(r, h['On hand'] + 1).setValue(onHand + qty);
  if (h['Last restocked'] != null) sh.getRange(r, h['Last restocked'] + 1).setValue(new Date());

  closeRestockAction_(sh.getParent(), item);
  ui.alert('Marked "' + item + '" restocked. On hand is now ' + (onHand + qty) + '.');
}

function closeRestockAction_(ss, item) {
  var sh = ss.getSheetByName(ACTIONS_TAB);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  var h = hmap_(data[0]);
  var key = 'restock:' + item;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][h['Source']]) === key) {
      var st = String(data[i][h['Status']] || '').toLowerCase();
      if (st !== 'done' && st !== 'completed') {
        sh.getRange(i + 1, h['Status'] + 1).setValue('Done');
        if (h['Completed'] != null) sh.getRange(i + 1, h['Completed'] + 1).setValue(new Date());
      }
    }
  }
}

function addInventoryHtml_() {
  return ''
    + '<style>body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:14px;color:#222}'
    + 'label{display:block;font-size:12px;font-weight:bold;margin:10px 0 3px}'
    + 'input{width:100%;box-sizing:border-box;padding:7px 8px;border:1px solid #ccc;border-radius:6px;font:inherit}'
    + '.row{display:flex;gap:8px}.row>div{flex:1}'
    + '.btns{margin-top:16px;display:flex;gap:8px}'
    + 'button{padding:9px 14px;border:none;border-radius:7px;font:inherit;font-weight:bold;cursor:pointer}'
    + '.go{background:#8f1515;color:#fff}.cancel{background:#eee;color:#333}'
    + '#msg{margin-top:10px;font-size:12px;color:#b31b1b}</style>'
    + '<label>Item</label><input id="item" autofocus>'
    + '<label>Location</label><input id="location">'
    + '<div class="row"><div><label>On hand</label><input id="onHand" type="number"></div>'
    + '<div><label>Unit</label><input id="unit" placeholder="each, box"></div></div>'
    + '<div class="row"><div><label>Reorder point</label><input id="reorderPoint" type="number"></div>'
    + '<div><label>Reorder qty</label><input id="reorderQty" type="number"></div></div>'
    + '<label>Supplier</label><input id="supplier">'
    + '<label>Product link</label><input id="productLink">'
    + '<label>eShop info</label><input id="eshop">'
    + '<label>Assign restocks to</label><input id="assignTo" value="' + DEFAULT_ASSIGNEE + '">'
    + '<div id="msg"></div>'
    + '<div class="btns"><button class="go" onclick="save()">Add item</button>'
    + '<button class="cancel" onclick="google.script.host.close()">Cancel</button></div>'
    + '<script>'
    + 'function val(id){return document.getElementById(id).value;}'
    + 'function save(){var item=val("item").trim();if(!item){document.getElementById("msg").textContent="Item name is required.";return;}'
    + 'var b=document.querySelector(".go");b.disabled=true;b.textContent="Adding...";'
    + 'google.script.run.withSuccessHandler(function(r){if(r&&r.ok){google.script.host.close();}else{document.getElementById("msg").textContent=(r&&r.error)||"Failed";b.disabled=false;b.textContent="Add item";}})'
    + '.withFailureHandler(function(e){document.getElementById("msg").textContent=String(e);b.disabled=false;b.textContent="Add item";})'
    + '.addInventoryItem({item:item,location:val("location"),onHand:val("onHand"),unit:val("unit"),reorderPoint:val("reorderPoint"),reorderQty:val("reorderQty"),supplier:val("supplier"),productLink:val("productLink"),eshop:val("eshop"),assignTo:val("assignTo")});}'
    + '</script>';
}

//------------------------------------------------------------
//
// HELPERS.
//
//------------------------------------------------------------

function createActionItem_(ss, o) {
  var sh = ss.getSheetByName(ACTIONS_TAB);
  if (!sh) { Logger.log('No "' + ACTIONS_TAB + '" tab. Run setupRegistry.'); return null; }
  var h = hmap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var rowArr = [];
  for (var i = 0; i < sh.getLastColumn(); i++) rowArr.push('');
  function set(name, val) { if (h[name] != null) rowArr[h[name]] = val; }
  var id = 'AI-' + Utilities.getUuid().slice(0, 8);
  set('Item ID', id);
  set('Created', new Date());
  set('Source', o.source || '');
  set('Title', o.title || '');
  set('Location', o.location || '');
  set('Equipment / part', o.equipment || '');
  set('Assigned to', o.assignee || DEFAULT_ASSIGNEE);
  set('Status', 'Open');
  set('Purchasing info', o.purchasing || '');
  set('Due', o.due || '');
  sh.appendRow(rowArr);
  return id;
}

// True if an action item with this source is already open, so we do not duplicate.
function openActionExists_(ss, sourceKey) {
  var sh = ss.getSheetByName(ACTIONS_TAB);
  if (!sh) return false;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return false;
  var h = hmap_(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][h['Source']]) === sourceKey) {
      var st = String(data[i][h['Status']] || '').toLowerCase();
      if (st !== 'done' && st !== 'completed') return true;
    }
  }
  return false;
}

function notifyOps_(subject, lines) {
  if (!OPS_NOTIFY) return;
  MailApp.sendEmail(OPS_NOTIFY, '[Ops] ' + subject, lines.join('\n'));
}

// Roll a date forward by a frequency word or "N days/weeks/months/years".
function advanceDate_(date, freq) {
  var d = new Date(date);
  var f = String(freq || '').toLowerCase().trim();
  var m = f.match(/(\d+)\s*(day|week|month|year)/);
  if (m) {
    var n = parseInt(m[1], 10);
    if (m[2] === 'day') d.setDate(d.getDate() + n);
    else if (m[2] === 'week') d.setDate(d.getDate() + 7 * n);
    else if (m[2] === 'month') d.setMonth(d.getMonth() + n);
    else d.setFullYear(d.getFullYear() + n);
    return d;
  }
  if (f.indexOf('week') >= 0) d.setDate(d.getDate() + 7);
  else if (f.indexOf('quarter') >= 0) d.setMonth(d.getMonth() + 3);
  else if (f.indexOf('semi') >= 0) d.setMonth(d.getMonth() + 6);
  else if (f.indexOf('annual') >= 0 || f.indexOf('year') >= 0) d.setFullYear(d.getFullYear() + 1);
  else if (f.indexOf('month') >= 0) d.setMonth(d.getMonth() + 1);
  else if (f.indexOf('daily') >= 0 || f.indexOf('day') >= 0) d.setDate(d.getDate() + 1);
  else d.setMonth(d.getMonth() + 1);
  return d;
}

function hmap_(row) {
  var map = {};
  row.forEach(function (c, i) { map[String(c).trim()] = i; });
  return map;
}

//------------------------------------------------------------
//
// ONE-TIME IMPORT of the existing tabs into the CMMS structure.
//
//   Team tool inventory    -> Equipment  (tools you own, In service)
//   Facility non recurring -> Equipment  (one-time purchases, Planned)
//   Mill room Tooling      -> Inventory  (procurement list: supplier, P/N, price)
//
// The Locks tab is left as-is. Safe to run more than once: items already in the
// target tab (matched by name) are skipped, so a re-run never duplicates.
//
//------------------------------------------------------------

function importExistingData() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var ui = SpreadsheetApp.getUi();
  if (!ss.getSheetByName(EQUIPMENT_TAB) || !ss.getSheetByName(INVENTORY_TAB)) {
    ui.alert('Run "Set up / update tabs" first, then import.');
    return;
  }
  ss.toast('Reading existing tabs...', 'Import', 5);

  var tools = importTools_(ss);
  var facility = importFacility_(ss);
  var mill = importMillRoom_(ss);
  var eq = tools + facility;

  var msg = 'Imported ' + eq + ' equipment (' + tools + ' tools, ' + facility + ' facility) '
    + 'and ' + mill + ' inventory item(s).';
  Logger.log(msg);
  ss.toast(msg, 'Import complete', 8);
  ui.alert(msg + '\n\nReview the Equipment and Inventory tabs, then delete any leftover example rows. '
    + 'Locks was left as-is.');
}

//------------------------------------------------------------
//
// Fill the Image column with a representative product photo (Wikimedia Commons,
// hotlinkable) for items that match a known tool/consumable type. Only rows with
// a blank Image are touched, so anything you set by hand or upload is left alone.
// Most specific keywords are matched first. Swap any photo from the item's Edit form.
//
//------------------------------------------------------------

function seedPhotos() {
  var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
  var ui = SpreadsheetApp.getUi();
  var added = 0, checked = 0;
  [EQUIPMENT_TAB, INVENTORY_TAB].forEach(function (tab) {
    var sh = ss.getSheetByName(tab);
    if (!sh) return;
    var v = sh.getDataRange().getValues();
    if (v.length < 2) return;
    var h = hmap_(v[0]);
    var ni = h[tab === INVENTORY_TAB ? 'Item' : 'Name'], ii = h['Image'];
    if (ni == null || ii == null) return;
    for (var r = 1; r < v.length; r++) {
      if (String(v[r][ii]).trim()) continue;
      checked++;
      var url = matchPhoto_(String(v[r][ni] || ''));
      if (url) { sh.getRange(r + 1, ii + 1).setValue(url); added++; }
    }
  });
  var msg = 'Added photos to ' + added + ' item(s). ' + (checked - added) + ' had no match (leave a photo by uploading one).';
  Logger.log(msg);
  ss.toast(msg, 'Photos', 8);
  ui.alert(msg + '\n\nOpen the Equipment & inventory page to see them. Swap any photo from an item\'s Edit form.');
}

// Ordered most-specific-first; the first keyword found in the name wins.
function matchPhoto_(name) {
  var n = String(name || '').toLowerCase();
  var W = 'https://upload.wikimedia.org/wikipedia/commons/';
  var MAP = [
    [['drill press'], W + 'thumb/c/c9/Drillpress.jpg/500px-Drillpress.jpg'],
    [['cold saw'], W + 'thumb/2/28/Cold_Metal_Cutting_Saw.JPG/500px-Cold_Metal_Cutting_Saw.JPG'],
    [['table saw'], W + 'thumb/e/ed/Table_saw_blade_guard.jpg/500px-Table_saw_blade_guard.jpg'],
    [['circular saw', 'circ saw'], W + 'thumb/5/55/DeWalt_circular_saw_in_use.jpg/500px-DeWalt_circular_saw_in_use.jpg'],
    [['impact driver'], W + 'thumb/f/ff/HiKOKI_cordless_impact_driver_WH36DC.jpg/500px-HiKOKI_cordless_impact_driver_WH36DC.jpg'],
    [['drill bit', 'bit set', 'driver bit'], W + 'thumb/d/d9/Drill_bits_2017_G1.jpg/500px-Drill_bits_2017_G1.jpg'],
    [['end mill', 'endmill'], W + 'thumb/e/e6/MillingCutterSlotEndMillBallnose.jpg/500px-MillingCutterSlotEndMillBallnose.jpg'],
    [['drill'], W + 'thumb/4/46/Drill-driver.JPG/500px-Drill-driver.JPG'],
    [['router'], W + 'thumb/1/13/Modern_plunge_router.jpg/500px-Modern_plunge_router.jpg'],
    [['belt sander', 'sander'], W + 'thumb/6/68/Belt_sander_bosch.jpg/500px-Belt_sander_bosch.jpg'],
    [['grinder'], W + 'thumb/4/4e/Faschina-concrete_angle_grinder-Betontrennschleifen-13ASD.jpg/500px-Faschina-concrete_angle_grinder-Betontrennschleifen-13ASD.jpg'],
    [['lathe'], W + 'thumb/9/95/Metal_lathe_used_to_produce_flywheels.JPG/500px-Metal_lathe_used_to_produce_flywheels.JPG'],
    [['welder', 'welding'], W + 'thumb/7/79/Welding_power_supply-Miller-Syncrowave350LX-front-triddle.jpg/500px-Welding_power_supply-Miller-Syncrowave350LX-front-triddle.jpg'],
    [['socket'], W + 'thumb/0/00/Socket_wrench_and_sockets.JPG/500px-Socket_wrench_and_sockets.JPG'],
    [['ratchet'], W + 'thumb/0/00/Socket_wrench_and_sockets.JPG/500px-Socket_wrench_and_sockets.JPG'],
    [['allen', 'hex key'], W + 'thumb/9/9f/2023_Bity_Imbus.jpg/500px-2023_Bity_Imbus.jpg'],
    [['adjustable wrench'], W + 'thumb/5/50/2023_Klucz_nastawny.jpg/500px-2023_Klucz_nastawny.jpg'],
    [['wrench'], W + 'thumb/1/1c/Gedore_No._7_combination_wrenches_6%E2%80%9319_mm.jpg/500px-Gedore_No._7_combination_wrenches_6%E2%80%9319_mm.jpg'],
    [['caliper', 'calliper'], W + 'thumb/9/94/Messschieber.jpg/500px-Messschieber.jpg'],
    [['vise grip', 'vice grip', 'locking plier'], W + 'thumb/c/cf/Locking_pliers.jpg/500px-Locking_pliers.jpg'],
    [['plier', 'cutter', 'flush'], W + 'thumb/2/23/05-02_combination_pliers_big.jpg/500px-05-02_combination_pliers_big.jpg'],
    [['screwdriver', 'screw driver'], W + 'thumb/5/5b/Black_screwdriver.png/500px-Black_screwdriver.png'],
    [['punch'], W + 'thumb/e/ec/1970s_center_punch_by_Swedish_company_Luna_Tools.jpg/500px-1970s_center_punch_by_Swedish_company_Luna_Tools.jpg'],
    [['shear', 'snip'], W + 'thumb/6/65/Early_20th_century_tin_snips_by_Johann_Heinrich_Braun_Ronsdorf_Germany_view_1.jpg/500px-Early_20th_century_tin_snips_by_Johann_Heinrich_Braun_Ronsdorf_Germany_view_1.jpg'],
    [['hacksaw', 'hack saw'], W + 'thumb/d/d8/Hacksaw_with_grey_handle.jpg/500px-Hacksaw_with_grey_handle.jpg'],
    [['tape measure', 'measuring tape', 'measure tape'], W + 'thumb/f/f1/B%26Q_Tape_Measure.jpg/500px-B%26Q_Tape_Measure.jpg'],
    [['calculator'], W + 'thumb/8/8f/Sharp_Scientific_Calculator.jpg/500px-Sharp_Scientific_Calculator.jpg'],
    [['nitrile', 'glove'], W + '8/8d/Disposable_nitrile_glove.jpg'],
    [['sandpaper', 'sand paper', 'sanding'], W + 'thumb/9/90/Schleifpapier_verschiedene_Sorten.jpg/500px-Schleifpapier_verschiedene_Sorten.jpg'],
    [['flammable', 'flammables'], W + 'thumb/2/2c/Flammable_cabinet_black_open.jpg/500px-Flammable_cabinet_black_open.jpg'],
    [['cord reel', 'air reel', 'reel'], W + 'thumb/a/a8/Cable_reel_extension_cord.jpg/500px-Cable_reel_extension_cord.jpg'],
    [['shop vac', 'vacuum'], W + 'thumb/3/3d/Advance_VL500_wetdry_vacuum_Yonge.jpg/500px-Advance_VL500_wetdry_vacuum_Yonge.jpg']
  ];
  for (var i = 0; i < MAP.length; i++) {
    for (var j = 0; j < MAP[i][0].length; j++) {
      if (n.indexOf(MAP[i][0][j]) >= 0) return MAP[i][1];
    }
  }
  return '';
}

// Team tool inventory: a single "Tool" column of owned tools -> Equipment.
function importTools_(ss) {
  var src = ss.getSheetByName('Team tool inventory');
  if (!src) return 0;
  var v = src.getDataRange().getValues();
  if (v.length < 2) return 0;
  var cTool = findCol_(v[0], ['tool']);
  if (cTool < 0) cTool = 0;
  var tgt = ss.getSheetByName(EQUIPMENT_TAB);
  var have = existingNames_(tgt, 'Name');
  var seq = nextSeq_(tgt, 'Asset ID', 'TL-');
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][cTool]).trim();
    if (!importable_(name) || have[name.toLowerCase()]) continue;
    have[name.toLowerCase()] = true;
    rows.push({ 'Asset ID': 'TL-' + pad_(seq++, 3), 'Name': name, 'Category': 'Tool', 'Status': 'In service' });
  }
  return appendRows_(tgt, rows);
}

// Facility non recurring: one-time facility purchases (col "Equipent/ service") ->
// Equipment, marked Planned since these are not necessarily installed yet.
function importFacility_(ss) {
  var src = ss.getSheetByName('Facility non recurring');
  if (!src) return 0;
  var v = src.getDataRange().getValues();
  if (v.length < 2) return 0;
  var H = v[0];
  var cName = findCol_(H, ['equipent', 'equipment', 'service']);
  var cUnits = findCol_(H, ['units']);
  var cCost = findCol_(H, ['cost']);
  if (cName < 0) return 0;
  var tgt = ss.getSheetByName(EQUIPMENT_TAB);
  var have = existingNames_(tgt, 'Name');
  var seq = nextSeq_(tgt, 'Asset ID', 'FA-');
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][cName]).trim();
    if (!importable_(name) || name.length < 2 || have[name.toLowerCase()]) continue;
    have[name.toLowerCase()] = true;
    var notes = [];
    if (cUnits >= 0 && String(v[i][cUnits]).trim()) notes.push('qty ' + String(v[i][cUnits]).trim());
    if (cCost >= 0 && numish_(v[i][cCost])) notes.push('est ~$' + trimNum_(v[i][cCost]));
    rows.push({ 'Asset ID': 'FA-' + pad_(seq++, 3), 'Name': name, 'Category': 'Facility', 'Status': 'Planned', 'Notes': notes.join(' · ') });
  }
  return appendRows_(tgt, rows);
}

// Mill room Tooling: a procurement list -> Inventory. On hand is left blank (stock
// unknown), so nothing is falsely flagged as out; supplier/link/P-N/price are kept.
function importMillRoom_(ss) {
  var src = ss.getSheetByName('Mill room Tooling');
  if (!src) return 0;
  var v = src.getDataRange().getValues();
  if (v.length < 2) return 0;
  var H = v[0];
  var cItem = findCol_(H, ['equipment']);
  var cMsc = findCol_(H, ['supplier link']);
  var cAlt = findCol_(H, ['non msc']);
  var cPn = findCol_(H, ['p/n', 'p/ n']);
  var cNote = findCol_(H, ['note']);
  var cPrice = findCol_(H, ['price']);
  var cPrio = findCol_(H, ['priority']);
  if (cItem < 0) return 0;
  var tgt = ss.getSheetByName(INVENTORY_TAB);
  var have = existingNames_(tgt, 'Item');
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var item = String(v[i][cItem]).trim();
    if (!importable_(item) || have[item.toLowerCase()]) continue;
    have[item.toLowerCase()] = true;
    var link = (cMsc >= 0 && String(v[i][cMsc]).trim()) ? String(v[i][cMsc]).trim()
             : (cAlt >= 0 ? String(v[i][cAlt]).trim() : '');
    var bits = [];
    if (cPn >= 0 && String(v[i][cPn]).trim()) bits.push(String(v[i][cPn]).trim());
    if (cPrice >= 0 && numish_(v[i][cPrice])) bits.push('~$' + trimNum_(v[i][cPrice]));
    if (cPrio >= 0 && String(v[i][cPrio]).trim()) bits.push(String(v[i][cPrio]).trim());
    if (cNote >= 0 && String(v[i][cNote]).trim()) bits.push(String(v[i][cNote]).trim());
    rows.push({
      'Item': item, 'On hand': '', 'Reorder point': '',
      'Supplier': supplierFromLink_(link), 'Product link': link,
      'eShop info': bits.join(' · '), 'Assign to': DEFAULT_ASSIGNEE
    });
  }
  return appendRows_(tgt, rows);
}

// ---- import helpers ----

function importable_(name) {
  if (!name) return false;
  var n = name.toLowerCase();
  return n.indexOf('delete this row') < 0 && n.indexOf('example:') < 0;
}

function findCol_(headers, needles) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toLowerCase();
    if (!h) continue;
    for (var j = 0; j < needles.length; j++) { if (h.indexOf(needles[j]) >= 0) return i; }
  }
  return -1;
}

function existingNames_(sh, nameHeader) {
  var set = {};
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return set;
  var ci = hmap_(data[0])[nameHeader];
  if (ci == null) return set;
  for (var i = 1; i < data.length; i++) {
    var v = String(data[i][ci]).trim().toLowerCase();
    if (v) set[v] = true;
  }
  return set;
}

function nextSeq_(sh, idHeader, prefix) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return 1;
  var ci = hmap_(data[0])[idHeader];
  if (ci == null) return 1;
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][ci]).trim();
    if (id.indexOf(prefix) === 0) { var n = parseInt(id.slice(prefix.length), 10); if (!isNaN(n) && n > max) max = n; }
  }
  return max + 1;
}

// Map an array of {header: value} objects onto the tab's real column order, in one write.
function appendRows_(sh, rowObjs) {
  if (!rowObjs.length) return 0;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var out = rowObjs.map(function (obj) {
    return headers.map(function (hd) { var v = obj[String(hd).trim()]; return v == null ? '' : v; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, headers.length).setValues(out);
  return out.length;
}

function supplierFromLink_(url) {
  var u = String(url).toLowerCase();
  if (u.indexOf('mscdirect') >= 0) return 'MSC';
  if (u.indexOf('mcmaster') >= 0) return 'McMaster-Carr';
  if (u.indexOf('grainger') >= 0) return 'Grainger';
  if (u.indexOf('uline') >= 0) return 'Uline';
  if (u.indexOf('amazon') >= 0) return 'Amazon';
  if (u.indexOf('homedepot') >= 0) return 'Home Depot';
  return '';
}

function numish_(v) { return v !== '' && v != null && !isNaN(Number(v)) && Number(v) > 0; }
function trimNum_(v) { var n = Number(v); return n % 1 === 0 ? String(n) : String(n); }
function pad_(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }
