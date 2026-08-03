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
    ['Asset ID', 'Name', 'Category', 'Location', 'Owning team', 'Owner', 'Status', 'Installed', 'Notes'],
    ['EX-001', 'Example: spray booth (delete this row)', 'Finishing', 'Composites lab', 'Operations Team', 'Noah Hamm', 'In service', '2024', 'Filters checked monthly']);

  ensureTab_(ss, INVENTORY_TAB,
    ['Item', 'Location', 'On hand', 'Unit', 'Reorder point', 'Reorder qty', 'Supplier', 'Product link', 'eShop info', 'Assign to', 'Last restocked'],
    ['Example: spray booth filter (delete this row)', 'Composites lab', 0, 'each', 2, 6, 'Uline', '', 'eShop item #', 'Noah Hamm', '2026-05']);

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
