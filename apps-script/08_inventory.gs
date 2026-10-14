/**
 * Project Teams Tool & Equipment Inventory: the backend of the Ops Hub's
 * inventory page (inventory.html). Bound to its own Google Sheet.
 * =============================================================================
 * The hub is the app. This spreadsheet is the database, and this script is both
 * its caretaker (tabs, formulas, Item IDs, self-test) and its web API.
 *
 * Tabs it builds
 *   Dashboard         The same numbers the app shows, for anyone in the sheet.
 *   Inventory Master  One row per tool or piece of equipment, plus consumables.
 *   Checkout Log      Every check out and return (from the app or the backup form).
 *   Maintenance Log   Repairs, calibrations and inspections.
 *   Lists             Dropdown options (categories, locations, ...) and links.
 *   Latest            Hidden helper the formulas read. Leave it alone.
 *
 * SET UP (once)
 *   1. In a new, empty Google Sheet: Extensions > Apps Script, replace the sample
 *      code with this file, Save, then run setupInventory and approve access
 *      (Sheets and Forms, plus Drive for item photos). It builds the tabs and a
 *      backup Google Form, and runs a self-test.
 *   2. Deploy > New deployment > type Web app. Execute as: Me. Who has access:
 *      Anyone. Deploy, then copy the Web app URL (it ends in /exec) into the
 *      hub's config.js as INVENTORY.apiUrl. That connects the app to this sheet.
 *   After editing this file later: Deploy > Manage deployments > pencil >
 *   Version: New version > Deploy. The URL stays the same.
 *   Running setupInventory again is safe. It repairs formulas, dropdowns and
 *   formatting, and never deletes your rows. On a sheet made before teams
 *   existed it also moves the Team column next to Location.
 *
 * PHOTOS AND TEAMS
 *   Admins add item photos in the app. They are saved in an "Inventory photos"
 *   folder next to this spreadsheet in Drive, shared view-only by link (like the
 *   Space Status uploads), and the Photo URL column links to them. Every item also
 *   has a Team: the project team that owns it, or Shared for the program's own
 *   tools. The team list is on the Lists tab, column E.
 *
 * TEST DATA
 *   Inventory menu > Add test data (or the button in the app's admin view) loads
 *   about 25 sample items with checkouts and maintenance. Everything it adds is
 *   tagged "[test data]"; Remove test data takes all of it back out.
 *
 * Triggers installed by setup (all on this spreadsheet):
 *   onInventoryEdit    fills Item IDs, stamps times on hand-typed log rows,
 *                      keeps the backup form's item list in sync
 *   onInventoryChange  re-syncs that list after rows are deleted
 *   onCheckoutSubmit   copies "Condition on return" from the backup form
 * =============================================================================
 */

var INV = {
  formTitle: 'Tool & Equipment Checkout',
  tabs: {
    dash: 'Dashboard',
    master: 'Inventory Master',
    log: 'Checkout Log',
    maint: 'Maintenance Log',
    lists: 'Lists',
    calc: 'Latest'
  },
  // Form question titles. They become the Checkout Log's column headers.
  q: {
    item: 'Item',
    action: 'Action',
    netid: 'NetID',
    notes: 'Notes',
    due: 'Due back',
    cond: 'Condition on return'
  },
  checkOut: 'Check Out',
  ret: 'Return',
  conditions: ['Good', 'Fair', 'Needs Repair'],
  outOfService: ['Maintenance', 'Retired'],
  consumable: 'Consumable',   // formulas look for this category name
  sep: ' · ',            // form choices read "PT-001 · DeWalt 20V drill"
  noItems: '(No items yet)',
  listRows: 10,               // rows shown per Dashboard list
  // The hub's inventory page. Item QR labels and the Checkout Link column point here.
  appUrl: 'https://cornellprojectteams.github.io/centralhub/inventory.html',
  testTag: '[test data]',     // marks everything Add test data creates
  shared: 'Shared',           // Team value for the program's own tools
  photoFolder: 'Inventory photos',
  red: '#B31B1B',
  gray: '#5F6368',
  calcBg: '#F3F3F3',
  tag: 'Inventory: '          // prefix on protections this script owns
};

// Inventory Master columns, in order. calc = filled by a formula or by the script.
var MASTER = [
  { key: 'id', h: 'Item ID', w: 80, note: 'Fills in by itself once the row has a name and a category (PT-001, EQ-004, ...). It never changes after that, so it is safe to print on a label.' },
  { key: 'name', h: 'Item Name', w: 230 },
  { key: 'cat', h: 'Category', w: 115, note: 'Pick from the list. Add categories (and their ID prefix) on the Lists tab.' },
  { key: 'sub', h: 'Subcategory', w: 140 },
  { key: 'loc', h: 'Location', w: 160, note: 'Where it lives when nobody has it out.' },
  { key: 'team', h: 'Team', w: 150, note: 'The project team that owns it, or Shared for the program\'s own tools that any team can borrow. The list is on the Lists tab, column E.' },
  { key: 'status', h: 'Status', w: 110, calc: true, note: 'Calculated. Available / Checked Out comes from the Checkout Log. Maintenance / Retired comes from Out of Service. Consumables show In Stock, Low Stock or Out of Stock.' },
  { key: 'holder', h: 'Current Holder', w: 120, calc: true, note: 'Calculated: the NetID on the latest check out.' },
  { key: 'due', h: 'Due Date', w: 105, calc: true, fmt: 'mmm d, yyyy', note: 'Calculated: the due date on the latest check out.' },
  { key: 'cond', h: 'Condition', w: 110, note: 'Updates by itself when someone returns the item through the form. Set it back to Good after a repair.' },
  { key: 'oos', h: 'Out of Service', w: 120, note: 'Maintenance or Retired takes the item out of circulation. Leave it blank while the item is in use.' },
  { key: 'serial', h: 'Serial / Asset Tag', w: 135 },
  { key: 'bought', h: 'Purchase Date', w: 110, fmt: 'mmm d, yyyy' },
  { key: 'qty', h: 'Current Qty', w: 95, fmt: '0', note: 'Consumables only.' },
  { key: 'min', h: 'Min Qty', w: 80, fmt: '0', note: 'Consumables only. Below this number the item shows Low Stock.' },
  { key: 'nextMaint', h: 'Next Maintenance', w: 130, calc: true, fmt: 'mmm d, yyyy', note: 'Calculated: the latest Next Due for this item on the Maintenance Log.' },
  { key: 'notes', h: 'Notes', w: 230 },
  { key: 'photo', h: 'Photo URL', w: 150, note: 'Set by the app when an admin adds a photo (it is saved in the Inventory photos folder in Drive). You can also paste a link to a photo here.' },
  { key: 'link', h: 'Checkout Link', w: 160, calc: true, note: 'Filled in by the script. Opens this item in the hub\'s inventory app, where it can be checked out or returned. The app prints QR labels from the same link.' }
];

var MAINT = [
  { key: 'ts', h: 'Timestamp', w: 160, fmt: 'mmm d, yyyy h:mm am/pm', note: 'Fills in by itself when you add a row.' },
  { key: 'item', h: 'Item', w: 240, note: 'Pick the item from the list.' },
  { key: 'issue', h: 'Issue', w: 220 },
  { key: 'action', h: 'Action Taken', w: 220 },
  { key: 'by', h: 'Performed By', w: 130 },
  { key: 'next', h: 'Next Due', w: 110, fmt: 'mmm d, yyyy', note: 'When this item should be looked at next. Feeds Next Maintenance on the Inventory Master and the Dashboard.' },
  { key: 'cost', h: 'Cost', w: 90, fmt: '$#,##0.00' }
];

// Starting dropdown options. Edit them on the Lists tab, not here.
var SEED = {
  categories: [['Power Tool', 'PT'], ['Hand Tool', 'HT'], ['Safety Tool', 'ST'], ['Equipment', 'EQ'], ['Consumable', 'CS']],
  subcategories: ['Drill', 'Impact Driver', 'Saw', 'Sander', 'Grinder', 'Router', 'Heat Gun',
    'Socket Set', 'Wrenches', 'Pliers', 'Clamps',
    'Eye Protection', 'Face Shield', 'Hearing Protection', 'Fire Extinguisher', 'First Aid Kit', 'Spill Kit', 'Lockout/Tagout Kit',
    '3D Printer (FDM)', '3D Printer (Resin)', 'Soldering Station', 'Laser Cutter', 'CNC Machine', 'Drill Press', 'Band Saw',
    'Oscilloscope', 'Power Supply', 'Multimeter',
    'Solder', 'Filament', 'Drill Bits', 'Saw Blades', 'Sandpaper', 'Gloves', 'N95 Masks', 'Other'],
  locations: ['Upson ELL', 'Upson Composites Lab', 'Upson Auto Lab', 'Upson HV Room', 'Upson LVL', 'Upson Basement',
    'High Voltage Lab', 'Rhodes Penthouse', 'Rhodes Basement', 'Tang Hall', 'Hollister Basement', 'Mill Room'],
  // Shared first, then the official team names (official-team-names.txt in the hub repo).
  teams: ['Shared', 'AguaClara', 'AppDev', 'AutoBoat', 'Autonomous Drone', 'Autonomous Sailboat', 'Autonomous Underwater Vehicle',
    'Baja', 'Biomedical Device', 'ChemE Car', 'Combat Robotics', 'Concrete Canoe', 'Cornell Assistive Technologies',
    'Custom Silicon Systems', 'Data Science', 'DEBUT', 'Design Build Fly', 'Digital Technology and Innovation',
    'Electric Vehicles', 'Engineering World Health', 'Engineers for a Sustainable World', 'Engineers in Action',
    'Engineers Without Borders', 'Extended Reality', 'FSAE Racing', 'GeoData', 'Hack4Impact', 'Hyperloop', 'iGEM',
    'Mars Rover', 'Nexus', 'Operations Team', 'Rocketry', 'Seismic Design', 'SensTech', 'Solar Boat', 'Steel Bridge',
    'Unmanned Air System']
};

// Dashboard layout. Each list block: title row, header row, listRows data rows, then a "+N more" row.
var DASH = {
  tiles: [
    { col: 1, label: 'Available', key: 'available' },
    { col: 2, label: 'Checked out', key: 'out' },
    { col: 3, label: 'Overdue', key: 'overdue' },
    { col: 4, label: 'In maintenance', key: 'maint' },
    { col: 5, label: 'Needs repair', key: 'repair' },
    { col: 7, label: 'Low or out of stock', key: 'stock' },
    { col: 8, label: 'Maintenance due (14 days)', key: 'maintDue' },
    { col: 9, label: 'Items in service', key: 'total' }
  ],
  tileRow: 4,
  blocks: [
    { key: 'overdue', row: 7, col: 1, title: 'Overdue', heads: ['Item ID', 'Item', 'Who has it', 'Due back', 'Days late'] },
    { key: 'who', row: 7, col: 7, title: 'Who has what', heads: ['Who', 'Items', 'How many'] },
    { key: 'out', row: 21, col: 1, title: 'Checked out', heads: ['Item ID', 'Item', 'Who has it', 'Due back', 'Location'] },
    { key: 'repair', row: 21, col: 7, title: 'Needs repair', heads: ['Item ID', 'Item', 'Location', 'Status'] },
    { key: 'maint', row: 35, col: 1, title: 'Maintenance due in the next 14 days', heads: ['Item ID', 'Item', 'Location', 'Next due', 'Days left'] },
    { key: 'stock', row: 35, col: 7, title: 'Low or out of stock', heads: ['Item ID', 'Item', 'On hand', 'Min'] }
  ],
  widths: [90, 210, 120, 110, 110, 24, 110, 220, 120, 100]
};

// =============================================================================
// Menu
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Inventory')
    .addItem('Show links and setup info', 'showInventoryLinks')
    .addItem('Fill in missing Item IDs', 'fillMissingIds')
    .addItem('Refresh the backup form\'s item list', 'refreshFormChoices')
    .addSeparator()
    .addItem('Add test data', 'addTestData')
    .addItem('Remove test data', 'removeTestData')
    .addSeparator()
    .addItem('Run the self-test', 'runSelfTest')
    .addItem('Set up or repair', 'setupInventory')
    .addToUi();
}

// Script properties work in every context this script runs in (menu, triggers,
// and the web app), unlike document properties.
function props_() { return PropertiesService.getScriptProperties(); }

// The inventory spreadsheet, from the editor, a trigger, or the web app.
function ss_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) { ss = null; }
  if (ss) return ss;
  var id = props_().getProperty('INV_SS_ID');
  if (!id) throw new Error('Run setupInventory once from the spreadsheet first.');
  return SpreadsheetApp.openById(id);
}

// One writer at a time (people in the app, triggers, the menu). Every change also
// drops the app's cached list so the next load is fresh. Calling it again from
// inside a locked section just runs the function (the script lock is not reentrant).
var LOCK_HELD_ = false;
function withLock_(fn) {
  if (LOCK_HELD_) return fn();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('The inventory is busy. Try again in a moment.');
  LOCK_HELD_ = true;
  try {
    return fn();
  } finally {
    LOCK_HELD_ = false;
    lock.releaseLock();
    dropCache_();
  }
}

// =============================================================================
// Setup
// =============================================================================

function setupInventory() {
  var ss = SpreadsheetApp.getActive();
  guardWrongSheet_(ss);
  props_().setProperty('INV_SS_ID', ss.getId());
  if (/^Untitled spreadsheet$/i.test(ss.getName())) ss.rename('Project Teams Tool & Equipment Inventory');

  var lists = ensureLists_(ss);
  var master = ensureHeaders_(ss, INV.tabs.master, MASTER);
  placeAfter_(master, 'team', 'loc');
  var maint = ensureHeaders_(ss, INV.tabs.maint, MAINT);
  var form = ensureForm_(ss);
  var log = ensureLog_(ss, form);
  var mc = masterCols_(master);
  var lc = logCols_(log);
  var tc = colsBySpec_(maint, MAINT);

  ensureCalc_(ss, lc, tc);
  finishMaster_(ss, master, mc, lists);
  finishLists_(lists, mc);
  finishLog_(log, lc, lists);
  finishMaint_(maint, tc, lists);
  buildDashboard_(ss, mc);
  orderTabs_(ss);
  installTriggers_(ss);
  refreshFormChoices();
  writeSettings_(ss, form);

  var report = withLock_(selfTest_);
  var info = linkInfo_(ss, form);
  Logger.log(configBlock_(info));
  Logger.log(report.summary + '\n' + report.lines.join('\n'));
  showDialog_(linksHtml_(info, report), 'Inventory is set up');
}

// Refuse to build inside the old "Tool inventory, purchasing and locks" sheet: its
// Mill room Tooling and Locks tabs feed the Command Center.
function guardWrongSheet_(ss) {
  if (ss.getSheetByName('Mill room Tooling') || ss.getSheetByName('Locks')) {
    throw new Error('This looks like the old "Tool inventory, purchasing and locks" sheet. Run this script in a new, empty Google Sheet instead.');
  }
}

function ensureLists_(ss) {
  var sh = ss.getSheetByName(INV.tabs.lists);
  if (sh) return sh;
  sh = ss.insertSheet(INV.tabs.lists);
  var heads = [['Category', 'ID prefix', 'Subcategory', 'Location', '', '', '', 'Setting', 'Value']];
  sh.getRange(1, 1, 1, heads[0].length).setValues(heads);
  sh.getRange(2, 1, SEED.categories.length, 2).setValues(SEED.categories);
  sh.getRange(2, 3, SEED.subcategories.length, 1).setValues(SEED.subcategories.map(function (v) { return [v]; }));
  sh.getRange(2, 4, SEED.locations.length, 1).setValues(SEED.locations.map(function (v) { return [v]; }));
  ensureTeamList_(sh);
  return sh;
}

// Column E of the Lists tab holds the teams. Sheets made before teams existed get
// the list the first time it is needed.
function ensureTeamList_(sh) {
  if (String(sh.getRange(1, 5).getValue()).trim()) return;
  sh.getRange(1, 5).setValue('Team');
  sh.getRange(2, 5, SEED.teams.length, 1).setValues(SEED.teams.map(function (v) { return [v]; }));
}

// Move a column so it sits right after another one (setup only; the formulas are
// rebuilt right after). Used to put Team next to Location in older sheets.
function placeAfter_(sh, key, afterKey) {
  var c = colsBySpec_(sh, MASTER);
  if (!c[key] || !c[afterKey] || c[key] === c[afterKey] + 1) return;
  sh.moveColumns(sh.getRange(1, c[key]), c[afterKey] + 1);
}

// Create the tab if needed, write the headers into an empty row 1, and append any
// missing header to the right of existing data. Never moves or clears user columns.
function ensureHeaders_(ss, name, spec) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    var blank = ss.getSheetByName('Sheet1');
    if (name === INV.tabs.master && blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) {
      sh = blank.setName(name);
    } else {
      sh = ss.insertSheet(name);
    }
  }
  var have = headerMap_(sh);
  if (!Object.keys(have).length) {
    sh.getRange(1, 1, 1, spec.length).setValues([spec.map(function (s) { return s.h; })]);
    return sh;
  }
  var next = sh.getLastColumn() + 1;
  spec.forEach(function (s) {
    if (!have[s.h.toLowerCase()]) {
      if (next > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
      sh.getRange(1, next).setValue(s.h);
      next++;
    }
  });
  return sh;
}

// =============================================================================
// The checkout form
// =============================================================================

function ensureForm_(ss) {
  var form = getForm_();
  if (form) return form;

  form = FormApp.create(INV.formTitle);
  form.setDescription('Borrowing or returning a Project Teams tool? Fill this out every time. It takes about 20 seconds.')
    .setConfirmationMessage('Logged, thanks. If you took something out, please bring it back by the date you picked.')
    .setShowLinkToRespondAgain(true)
    .setAllowResponseEdits(false)
    .setProgressBar(false);

  form.addListItem().setTitle(INV.q.item)
    .setHelpText('Find the ID on the tool\'s label (like PT-001), or pick it by name.')
    .setRequired(true)
    .setChoiceValues([INV.noItems]);
  var action = form.addMultipleChoiceItem().setTitle(INV.q.action).setRequired(true);
  form.addTextItem().setTitle(INV.q.netid)
    .setHelpText('Your Cornell NetID, like abc123')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Enter your NetID, like abc123.')
      .requireTextMatchesPattern('^[A-Za-z]{2,4}[0-9]{1,6}$')
      .build());
  form.addParagraphTextItem().setTitle(INV.q.notes).setHelpText('Optional. Anything we should know?');

  var outPage = form.addPageBreakItem().setTitle('Checking out');
  form.addDateItem().setTitle(INV.q.due).setHelpText('When will you bring it back?').setRequired(true);

  var retPage = form.addPageBreakItem().setTitle('Returning');
  form.addMultipleChoiceItem().setTitle(INV.q.cond)
    .setHelpText('What shape is it in?')
    .setRequired(true)
    .setChoiceValues(INV.conditions);

  // Finishing "Checking out" reaches this page break, so jump straight to submit.
  retPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  action.setChoices([
    action.createChoice(INV.checkOut, outPage),
    action.createChoice(INV.ret, retPage)
  ]);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props_().setProperty('INV_FORM_ID', form.getId());
  return form;
}

function getForm_() {
  var props = props_();
  var id = props.getProperty('INV_FORM_ID');
  if (id) {
    try { return FormApp.openById(id); } catch (e) { /* deleted or no access */ }
  }
  // Script copied to a new project (properties lost): recover the form from the log tab.
  var log = ss_().getSheetByName(INV.tabs.log);
  var url = log ? log.getFormUrl() : null;
  if (url) {
    try {
      var form = FormApp.openByUrl(url);
      props.setProperty('INV_FORM_ID', form.getId());
      return form;
    } catch (e) { /* fall through */ }
  }
  return null;
}

function itemQuestion_(form) {
  var lists = form.getItems(FormApp.ItemType.LIST);
  for (var i = 0; i < lists.length; i++) {
    if (lists[i].getTitle() === INV.q.item) return lists[i].asListItem();
  }
  return lists.length ? lists[0].asListItem() : null;
}

// The response tab appears a moment after the form is linked. Find it, name it.
function ensureLog_(ss, form) {
  var sh = null;
  for (var attempt = 0; attempt < 20 && !sh; attempt++) {
    if (attempt) Utilities.sleep(1000);
    SpreadsheetApp.flush();
    sh = findLogSheet_(SpreadsheetApp.openById(ss.getId()), form);
    if (sh && sh.getLastColumn() < 3) sh = null;   // headers not written yet
  }
  if (!sh) throw new Error('The form exists but its response tab has not appeared yet. Wait a minute and run setupInventory again.');
  if (sh.getName() !== INV.tabs.log && !ss.getSheetByName(INV.tabs.log)) sh.setName(INV.tabs.log);
  SpreadsheetApp.flush();
  return ss.getSheetByName(sh.getName()) || sh;
}

function findLogSheet_(ss, form) {
  var id = form.getId();
  var linked = ss.getSheets().filter(function (s) { return !!s.getFormUrl(); });
  for (var i = 0; i < linked.length; i++) {
    if (linked[i].getFormUrl().indexOf(id) >= 0) return linked[i];
  }
  var named = ss.getSheetByName(INV.tabs.log);
  if (named && named.getFormUrl()) return named;
  return linked.length === 1 ? linked[0] : null;
}

// =============================================================================
// Columns and formulas
// =============================================================================

// Header text -> column number. A formula header like ={"Status"; ...} counts as
// "Status" even when it is showing an error, so a broken column is found, not duplicated.
function headerMap_(sh) {
  var map = {};
  var lastCol = sh.getLastColumn();
  if (!lastCol) return map;
  var shown = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var formulas = sh.getRange(1, 1, 1, lastCol).getFormulas()[0];
  for (var i = 0; i < lastCol; i++) {
    var text = String(shown[i] || '').trim();
    var m = String(formulas[i] || '').match(/^=\s*\{\s*"([^"]+)"/);
    if (m) text = m[1];
    var k = text.toLowerCase();
    if (k && !map[k]) map[k] = i + 1;
  }
  return map;
}

function colsBySpec_(sh, spec) {
  var have = headerMap_(sh), out = {};
  spec.forEach(function (s) { out[s.key] = have[s.h.toLowerCase()]; });
  return out;
}

function masterCols_(sh) {
  var c = colsBySpec_(sh, MASTER);
  var missing = MASTER.filter(function (s) { return !c[s.key]; });
  if (!missing.length) return c;
  if (missing.some(function (s) { return s.calc || s.key === 'id'; })) {
    throw new Error('Inventory Master is missing: ' + missing.map(function (s) { return s.h; }).join(', ') + '. Run setupInventory to add them.');
  }
  // A typed-in column from a newer version of this script (like Team) is added on
  // the fly, so the app keeps working before anyone re-runs setupInventory.
  withLock_(function () {
    var ss = sh.getParent();
    ensureHeaders_(ss, sh.getName(), MASTER);
    var now = colsBySpec_(sh, MASTER);
    var lists = ss.getSheetByName(INV.tabs.lists);
    missing.forEach(function (s) {
      var col = now[s.key];
      if (!col || c[s.key]) return;
      sh.getRange(1, col).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(INV.red)
        .setVerticalAlignment('middle').setWrap(true).setNote(s.note || '');
      sh.setColumnWidth(col, s.w);
      if (s.key === 'team' && lists) {
        ensureTeamList_(lists);
        sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(lists.getRange('E2:E'), true).setAllowInvalid(true)
          .setHelpText('Pick the team that owns it, or Shared.').build());
      }
    });
  });
  return colsBySpec_(sh, MASTER);
}

function logCols_(sh) {
  var have = headerMap_(sh), out = { ts: have['timestamp'] };
  ['item', 'action', 'netid', 'notes', 'due', 'cond'].forEach(function (k) { out[k] = have[INV.q[k].toLowerCase()]; });
  var need = { ts: 'Timestamp', item: INV.q.item, action: INV.q.action, netid: INV.q.netid, due: INV.q.due, cond: INV.q.cond };
  var missing = Object.keys(need).filter(function (k) { return !out[k]; }).map(function (k) { return need[k]; });
  if (missing.length) throw new Error('The Checkout Log is missing the column(s): ' + missing.join(', ') + '. Keep the form\'s question titles as they were created.');
  return out;
}

// The calculated columns spill "" all the way down, which makes getLastRow() report
// the bottom of the sheet. This finds the last row with something a person typed.
function lastDataRow_(sh, mc) {
  var maxRows = sh.getMaxRows();
  if (maxRows < 2 || !sh.getLastColumn()) return 1;
  var skip = {};
  MASTER.forEach(function (s) { if (s.calc) skip[mc[s.key] - 1] = true; });
  var vals = sh.getRange(2, 1, maxRows - 1, sh.getLastColumn()).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    for (var j = 0; j < vals[i].length; j++) {
      if (!skip[j] && vals[i][j] !== '' && vals[i][j] !== null) return i + 2;
    }
  }
  return 1;
}

function letter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function sheetRef_(name) { return "'" + name.replace(/'/g, "''") + "'!"; }
function colRange_(name, col) { var L = letter_(col); return sheetRef_(name) + L + '2:' + L; }
function lit_(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }

// Hidden "Latest" tab: the newest log row per item, newest first. VLOOKUP into a
// list sorted newest-first returns the latest entry, which makes Status immune to
// duplicate or missing submissions (counting check-outs against returns is not).
function ensureCalc_(ss, lc, tc) {
  var sh = ss.getSheetByName(INV.tabs.calc) || ss.insertSheet(INV.tabs.calc);
  var L = function (k) { return colRange_(INV.tabs.log, lc[k]); };
  var T = function (k) { return colRange_(INV.tabs.maint, tc[k]); };
  var idOf = function (r) { return 'IFERROR(REGEXEXTRACT(TRIM(' + r + '&""),"^\\S+"),"")'; };

  sh.clearContents();
  sh.getRange('A1').setFormula('={"Item ID","Action","Person","Due back","When";'
    + 'ARRAYFORMULA(IFERROR(SORT(FILTER({' + idOf(L('item')) + ',' + L('action') + ',' + L('netid') + ','
    + 'IF(' + L('due') + '="","",' + L('due') + '),' + L('ts') + '},' + L('item') + '<>""),5,FALSE),{"","","","",""}))}');
  sh.getRange('G1').setFormula('={"Item ID","Next due","When";'
    + 'ARRAYFORMULA(IFERROR(SORT(FILTER({' + idOf(T('item')) + ',' + T('next') + ',' + T('ts') + '},'
    + T('item') + '<>"",ISNUMBER(' + T('next') + ')),3,FALSE),{"","",""}))}');
  sh.getRange('A1:I1').setFontWeight('bold');
  sh.getRange('D2:D').setNumberFormat('mmm d, yyyy');
  sh.getRange('E2:E').setNumberFormat('mmm d, yyyy h:mm am/pm');
  sh.getRange('H2:H').setNumberFormat('mmm d, yyyy');
  sh.getRange('I2:I').setNumberFormat('mmm d, yyyy h:mm am/pm');
  resetProtections_(sh);
  sh.protect().setDescription(INV.tag + 'helper tab').setWarningOnly(true);
  if (!sh.isSheetHidden()) sh.hideSheet();
  return sh;
}

function masterFormulas_(mc) {
  var R = function (k) { var L = letter_(mc[k]); return L + '2:' + L; };
  var latest = sheetRef_(INV.tabs.calc) + '$A$2:$E';
  var latestM = sheetRef_(INV.tabs.calc) + '$G$2:$I';
  var hdr = {};
  MASTER.forEach(function (s) { hdr[s.key] = s.h; });
  return {
    status: '={' + lit_(hdr.status) + ';ARRAYFORMULA(IF(' + R('id') + '="","",'
      + 'IF(' + R('oos') + '<>"",' + R('oos') + ','
      + 'IF(' + R('cat') + '=' + lit_(INV.consumable) + ','
      + 'IF(' + R('qty') + '="","",IF(' + R('qty') + '<=0,"Out of Stock",'
      + 'IF((' + R('min') + '<>"")*(' + R('qty') + '<' + R('min') + '),"Low Stock","In Stock"))),'
      + 'IF(IFERROR(VLOOKUP(' + R('id') + ',' + latest + ',2,FALSE),"")=' + lit_(INV.checkOut) + ',"Checked Out","Available")))))}',
    holder: '={' + lit_(hdr.holder) + ';ARRAYFORMULA(IF(' + R('status') + '="Checked Out",'
      + 'IFERROR(VLOOKUP(' + R('id') + ',' + latest + ',3,FALSE),""),""))}',
    due: '={' + lit_(hdr.due) + ';ARRAYFORMULA(IF(' + R('status') + '="Checked Out",'
      + 'IFERROR(VLOOKUP(' + R('id') + ',' + latest + ',4,FALSE),""),""))}',
    nextMaint: '={' + lit_(hdr.nextMaint) + ';ARRAYFORMULA(IF(' + R('id') + '="","",'
      + 'IFERROR(VLOOKUP(' + R('id') + ',' + latestM + ',2,FALSE),"")))}'
  };
}

function finishMaster_(ss, sh, mc, lists) {
  var rows = sh.getMaxRows() - 1;
  var f = masterFormulas_(mc);
  ['status', 'holder', 'due', 'nextMaint'].forEach(function (k) {
    // The column must be empty below the header or the formula cannot spill.
    sh.getRange(2, mc[k], rows, 1).clearContent();
    sh.getRange(1, mc[k]).setFormula(f[k]);
  });
  sh.getRange(1, mc.link).setValue('Checkout Link');

  MASTER.forEach(function (s) {
    var col = mc[s.key];
    var head = sh.getRange(1, col);
    head.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(s.calc ? INV.gray : INV.red)
      .setVerticalAlignment('middle').setWrap(true);
    head.setNote(s.note || '');
    sh.setColumnWidth(col, s.w);
    var body = sh.getRange(2, col, rows, 1);
    if (s.fmt) body.setNumberFormat(s.fmt);
    body.setBackground(s.calc ? INV.calcBg : null);
  });
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.setRowHeight(1, 36);
  sh.setTabColor(INV.red);

  var dv = function (key, rule) { sh.getRange(2, mc[key], rows, 1).setDataValidation(rule); };
  var V = function () { return SpreadsheetApp.newDataValidation(); };
  dv('cat', V().requireValueInRange(lists.getRange('A2:A'), true).setAllowInvalid(false)
    .setHelpText('Pick a category. Add new ones on the Lists tab.').build());
  dv('sub', V().requireValueInRange(lists.getRange('C2:C'), true).setAllowInvalid(true)
    .setHelpText('Pick a subcategory, or add one on the Lists tab.').build());
  dv('loc', V().requireValueInRange(lists.getRange('D2:D'), true).setAllowInvalid(true)
    .setHelpText('Pick a location, or add one on the Lists tab.').build());
  dv('team', V().requireValueInRange(lists.getRange('E2:E'), true).setAllowInvalid(true)
    .setHelpText('Pick the team that owns it, or Shared.').build());
  dv('cond', V().requireValueInList(INV.conditions, true).setAllowInvalid(false).build());
  dv('oos', V().requireValueInList(INV.outOfService, true).setAllowInvalid(false)
    .setHelpText('Maintenance or Retired. Leave blank while the item is in use.').build());
  dv('bought', V().requireDate().setAllowInvalid(true).build());
  dv('qty', V().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  dv('min', V().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());

  // Conditional formatting. Order matters: the first matching rule wins for a cell.
  var L = function (k) { return letter_(mc[k]); };
  var all = sh.getRange(2, 1, rows, sh.getMaxColumns());
  var col = function (k) { return sh.getRange(2, mc[k], rows, 1); };
  var span = function (a, b) { return sh.getRange(2, Math.min(mc[a], mc[b]), rows, Math.abs(mc[b] - mc[a]) + 1); };
  var R = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  var st = '$' + L('status') + '2', due = '$' + L('due') + '2';
  var rules = [
    R().whenFormulaSatisfied('=AND(' + st + '="Checked Out",ISNUMBER(' + due + '),' + due + '<TODAY())')
      .setBackground('#F4CCCC').setFontColor('#990000').setRanges([span('status', 'due')]).build(),
    R().whenTextEqualTo('Available').setBackground('#D9EAD3').setFontColor('#274E13').setRanges([col('status')]).build(),
    R().whenTextEqualTo('Checked Out').setBackground('#CFE2F3').setFontColor('#073763').setRanges([col('status')]).build(),
    R().whenTextEqualTo('Maintenance').setBackground('#FCE5CD').setFontColor('#783F04').setRanges([col('status')]).build(),
    R().whenFormulaSatisfied('=OR(' + st + '="Low Stock",' + st + '="Out of Stock")')
      .setBackground('#F4CCCC').setFontColor('#990000').setRanges([col('status')]).build(),
    R().whenTextEqualTo('Needs Repair').setBackground('#FFF2CC').setFontColor('#7F6000').setRanges([col('cond')]).build(),
    R().whenFormulaSatisfied('=AND(ISNUMBER($' + L('qty') + '2),ISNUMBER($' + L('min') + '2),$' + L('qty') + '2<$' + L('min') + '2)')
      .setBackground('#F4CCCC').setFontColor('#990000').setRanges([col('qty')]).build(),
    R().whenFormulaSatisfied('=' + st + '="Retired"').setFontColor('#9E9E9E').setRanges([all]).build()
  ];
  sh.setConditionalFormatRules(rules);

  resetProtections_(sh);
  MASTER.filter(function (s) { return s.calc; }).forEach(function (s) {
    sh.getRange(1, mc[s.key], rows + 1, 1).protect()
      .setDescription(INV.tag + s.h + ' is calculated').setWarningOnly(true);
  });
}

function finishLists_(sh, mc) {
  ensureTeamList_(sh);
  var M = function (k) { return colRange_(INV.tabs.master, mc[k]); };
  var rows = sh.getMaxRows() - 1;
  sh.getRange(2, 6, rows, 1).clearContent();
  sh.getRange('F1').setFormula('={"Item choices (automatic)";ARRAYFORMULA(IFERROR(FILTER('
    + M('id') + '&' + lit_(INV.sep) + '&TRIM(' + M('name') + '),'
    + M('id') + '<>"",' + M('name') + '<>"",' + M('cat') + '<>' + lit_(INV.consumable) + '),""))}');
  sh.getRange('A1:I1').setFontWeight('bold');
  [1, 2, 3, 4, 5, 6, 8, 9].forEach(function (c) {
    sh.getRange(1, c).setBackground(c === 6 ? INV.gray : '#202124').setFontColor('#FFFFFF');
  });
  sh.getRange('A1').setNote('Add a category and a short ID prefix (2 letters) on the same row. Do not rename "' + INV.consumable + '": the formulas look for it.');
  sh.getRange('C1').setNote('Options for the Subcategory dropdown. Add or remove freely. A new subcategory typed in the app is added here.');
  sh.getRange('D1').setNote('Options for the Location dropdown. Add or remove freely. A new location typed in the app is added here.');
  sh.getRange('E1').setNote('Options for the Team dropdown. Shared means the program\'s own tools that any team can borrow.');
  sh.getRange('F1').setNote('Calculated from the Inventory Master. Feeds the Item dropdowns on the logs.');
  sh.getRange('H1').setNote('Filled in by setup. The first link is the one to share.');
  sh.setColumnWidths(1, 4, 150);
  sh.setColumnWidth(2, 80);
  sh.setColumnWidth(5, 200);
  sh.setColumnWidth(6, 260);
  sh.setColumnWidth(7, 24);
  sh.setColumnWidth(8, 210);
  sh.setColumnWidth(9, 420);
  sh.setFrozenRows(1);
  resetProtections_(sh);
  sh.getRange(1, 6, rows + 1, 1).protect().setDescription(INV.tag + 'item choices are calculated').setWarningOnly(true);
  sh.getRange(1, 8, rows + 1, 2).protect().setDescription(INV.tag + 'settings are filled in by setup').setWarningOnly(true);
}

function finishLog_(sh, lc, lists) {
  var rows = sh.getMaxRows() - 1;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn());
  head.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(INV.red).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
  sh.setTabColor('#1F6FB2');
  sh.getRange(2, lc.ts, rows, 1).setNumberFormat('mmm d, yyyy h:mm am/pm');
  sh.getRange(2, lc.due, rows, 1).setNumberFormat('mmm d, yyyy');
  var widths = { ts: 170, item: 260, action: 100, netid: 90, notes: 220, due: 110, cond: 150 };
  Object.keys(widths).forEach(function (k) { if (lc[k]) sh.setColumnWidth(lc[k], widths[k]); });

  sh.getRange(1, lc.ts).setNote('The form stamps this. Typing a row by hand? The time fills in by itself when you pick the item.');
  sh.getRange(1, lc.item).setNote('The form fills this in. By hand, pick from the list so the Item ID at the start is exact.');

  // Warnings only: the form writes here, and old rows must stay valid after an item is renamed.
  var V = function () { return SpreadsheetApp.newDataValidation(); };
  sh.getRange(2, lc.item, rows, 1).setDataValidation(V().requireValueInRange(lists.getRange('F2:F'), true).setAllowInvalid(true).build());
  sh.getRange(2, lc.action, rows, 1).setDataValidation(V().requireValueInList([INV.checkOut, INV.ret], true).setAllowInvalid(true).build());
  sh.getRange(2, lc.due, rows, 1).setDataValidation(V().requireDate().setAllowInvalid(true).build());
  sh.getRange(2, lc.cond, rows, 1).setDataValidation(V().requireValueInList(INV.conditions, true).setAllowInvalid(true).build());
}

function finishMaint_(sh, tc, lists) {
  var rows = sh.getMaxRows() - 1;
  MAINT.forEach(function (s) {
    var col = tc[s.key];
    var head = sh.getRange(1, col);
    head.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(INV.red).setVerticalAlignment('middle');
    head.setNote(s.note || '');
    sh.setColumnWidth(col, s.w);
    if (s.fmt) sh.getRange(2, col, rows, 1).setNumberFormat(s.fmt);
  });
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
  sh.setTabColor('#B45F06');
  var V = function () { return SpreadsheetApp.newDataValidation(); };
  sh.getRange(2, tc.item, rows, 1).setDataValidation(V().requireValueInRange(lists.getRange('F2:F'), true).setAllowInvalid(true).build());
  sh.getRange(2, tc.next, rows, 1).setDataValidation(V().requireDate().setAllowInvalid(true).build());
  sh.getRange(2, tc.cost, rows, 1).setDataValidation(V().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(true).build());
}

// =============================================================================
// Dashboard
// =============================================================================

function dashFormulas_(mc) {
  var M = function (k) { return colRange_(INV.tabs.master, mc[k]); };
  var n = INV.listRows;
  var CO = lit_('Checked Out');
  var tileCell = {};
  DASH.tiles.forEach(function (t) { tileCell[t.key] = letter_(t.col) + (DASH.tileRow + 1); });
  var more = function (cell, words) { return '=IF(' + cell + '>' + n + ',"+"&(' + cell + '-' + n + ')&' + lit_(words) + ',"")'; };
  var who = 'UNIQUE(FILTER(' + M('holder') + ',' + M('status') + '=' + CO + ',' + M('holder') + '<>""))';

  return {
    tiles: {
      available: '=COUNTIF(' + M('status') + ',"Available")',
      out: '=COUNTIF(' + M('status') + ',' + CO + ')',
      overdue: '=COUNTIFS(' + M('status') + ',' + CO + ',' + M('due') + ',"<"&TODAY())',
      maint: '=COUNTIF(' + M('status') + ',"Maintenance")',
      repair: '=COUNTIFS(' + M('cond') + ',"Needs Repair",' + M('oos') + ',"<>Retired")',
      stock: '=COUNTIF(' + M('status') + ',"Low Stock")+COUNTIF(' + M('status') + ',"Out of Stock")',
      maintDue: '=COUNTIFS(' + M('nextMaint') + ',"<="&(TODAY()+14),' + M('oos') + ',"<>Retired")',
      total: '=COUNTIFS(' + M('id') + ',"<>",' + M('oos') + ',"<>Retired")'
    },
    lists: {
      overdue: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('holder') + ',' + M('due') + ','
        + 'IFERROR(TODAY()-' + M('due') + ',"")},' + M('status') + '=' + CO + ',ISNUMBER(' + M('due') + '),' + M('due') + '<TODAY()),5,FALSE),'
        + n + ',5),"Nothing overdue"))',
      who: '=IFERROR(ARRAYFORMULA(LET(p,' + who + ',ARRAY_CONSTRAIN(SORT({p,'
        + 'BYROW(p,LAMBDA(x,TEXTJOIN(", ",TRUE,FILTER(' + M('name') + ',' + M('holder') + '=x,' + M('status') + '=' + CO + ')))),'
        + 'COUNTIFS(' + M('holder') + ',p,' + M('status') + ',' + CO + ')},3,FALSE),' + n + ',3))),"Nobody has anything out")',
      out: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('holder') + ',' + M('due') + ',' + M('loc') + '},'
        + M('status') + '=' + CO + '),4,TRUE),' + n + ',5),"Nothing checked out"))',
      repair: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(FILTER({' + M('id') + ',' + M('name') + ',' + M('loc') + ',' + M('status') + '},'
        + M('cond') + '="Needs Repair",' + M('oos') + '<>"Retired"),' + n + ',4),"Nothing flagged for repair"))',
      maint: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('loc') + ',' + M('nextMaint') + ','
        + 'IFERROR(' + M('nextMaint') + '-TODAY(),"")},ISNUMBER(' + M('nextMaint') + '),' + M('nextMaint') + '<=TODAY()+14,'
        + M('oos') + '<>"Retired"),4,TRUE),' + n + ',5),"Nothing due in the next 14 days"))',
      stock: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('qty') + ',' + M('min') + '},'
        + '(' + M('status') + '="Low Stock")+(' + M('status') + '="Out of Stock")),3,TRUE),' + n + ',4),"All stocked"))'
    },
    more: {
      overdue: more(tileCell.overdue, ' more overdue. Filter the Inventory Master by Status to see them all.'),
      who: '=IFERROR(IF(COUNTA(' + who + ')>' + n + ',"+"&(COUNTA(' + who + ')-' + n + ')&" more people",""),"")',
      out: more(tileCell.out, ' more checked out'),
      repair: more(tileCell.repair, ' more flagged for repair'),
      maint: more(tileCell.maintDue, ' more due'),
      stock: more(tileCell.stock, ' more low on stock')
    }
  };
}

function buildDashboard_(ss, mc) {
  var sh = ss.getSheetByName(INV.tabs.dash) || ss.insertSheet(INV.tabs.dash, 0);
  resetProtections_(sh);
  sh.clear();
  sh.clearNotes();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  if (sh.getMaxRows() < 50) sh.insertRowsAfter(sh.getMaxRows(), 50 - sh.getMaxRows());
  DASH.widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setHiddenGridlines(true);
  sh.setTabColor(INV.red);

  var f = dashFormulas_(mc);
  sh.getRange('A1').setValue('Tool & equipment inventory').setFontSize(18).setFontWeight('bold').setFontColor(INV.red);
  sh.getRange('A2').setValue('Live from the Inventory Master, Checkout Log and Maintenance Log. Nothing to type on this tab.')
    .setFontColor(INV.gray).setFontStyle('italic');

  var tr = DASH.tileRow;
  DASH.tiles.forEach(function (t) {
    sh.getRange(tr, t.col).setValue(t.label).setFontSize(9).setFontWeight('bold').setFontColor(INV.gray).setWrap(true).setVerticalAlignment('bottom');
    sh.getRange(tr + 1, t.col).setFormula(f.tiles[t.key]).setFontSize(22).setFontWeight('bold').setFontColor('#202124').setHorizontalAlignment('left');
  });
  sh.getRange(tr, 1, 2, 5).setBackground('#F8F9FA');
  sh.getRange(tr, 7, 2, 3).setBackground('#F8F9FA');
  sh.setRowHeight(tr, 30);
  sh.setRowHeight(tr + 1, 44);

  var n = INV.listRows;
  DASH.blocks.forEach(function (b) {
    var w = b.heads.length;
    sh.getRange(b.row, b.col).setValue(b.title).setFontSize(12).setFontWeight('bold').setFontColor(INV.red);
    sh.getRange(b.row, b.col, 1, w).setBorder(null, null, true, null, null, null, INV.red, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(b.row + 1, b.col, 1, w).setValues([b.heads]).setFontWeight('bold').setFontSize(9)
      .setFontColor(INV.gray).setBackground('#F3F3F3');
    sh.getRange(b.row + 2, b.col).setFormula(f.lists[b.key]);
    sh.getRange(b.row + 2, b.col, n, w).setVerticalAlignment('top').setFontSize(10);
    sh.getRange(b.row + 2 + n, b.col).setFormula(f.more[b.key]).setFontStyle('italic').setFontColor(INV.gray).setFontSize(9);
  });

  // Number formats per column of each list.
  var fmt = function (key, offset, pattern) {
    var b = DASH.blocks.filter(function (x) { return x.key === key; })[0];
    sh.getRange(b.row + 2, b.col + offset, n, 1).setNumberFormat(pattern);
  };
  fmt('overdue', 3, 'mmm d, yyyy'); fmt('overdue', 4, '0');
  fmt('out', 3, 'mmm d, yyyy');
  fmt('maint', 3, 'mmm d, yyyy'); fmt('maint', 4, '0');
  fmt('who', 2, '0');
  fmt('stock', 2, '0'); fmt('stock', 3, '0');
  var whoBlock = DASH.blocks.filter(function (x) { return x.key === 'who'; })[0];
  sh.getRange(whoBlock.row + 2, whoBlock.col + 1, n, 1).setWrap(true);

  var tile = function (key) { var t = DASH.tiles.filter(function (x) { return x.key === key; })[0]; return sh.getRange(tr + 1, t.col); };
  var block = function (key, offset) { var b = DASH.blocks.filter(function (x) { return x.key === key; })[0]; return sh.getRange(b.row + 2, b.col + offset, n, 1); };
  var R = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  sh.setConditionalFormatRules([
    R().whenNumberGreaterThan(0).setFontColor(INV.red).setRanges([tile('overdue'), tile('stock')]).build(),
    R().whenNumberGreaterThan(0).setFontColor('#B45F06').setRanges([tile('repair'), tile('maintDue')]).build(),
    R().whenNumberGreaterThan(0).setFontColor(INV.red).setBold(true).setRanges([block('overdue', 4)]).build(),
    R().whenNumberLessThan(0).setFontColor(INV.red).setBold(true).setRanges([block('maint', 4)]).build()
  ]);

  sh.protect().setDescription(INV.tag + 'the Dashboard is calculated').setWarningOnly(true);
}

// Dashboard, Inventory Master, Checkout Log, Maintenance Log, Lists, then the hidden helper.
function orderTabs_(ss) {
  var order = [INV.tabs.dash, INV.tabs.master, INV.tabs.log, INV.tabs.maint, INV.tabs.lists];
  order.forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  // Drop the blank starter tab if it is still around and empty.
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  var calc = ss.getSheetByName(INV.tabs.calc);
  if (calc && !calc.isSheetHidden()) calc.hideSheet();
  var dash = ss.getSheetByName(INV.tabs.dash);
  if (dash) ss.setActiveSheet(dash);
}

function resetProtections_(sh) {
  [SpreadsheetApp.ProtectionType.RANGE, SpreadsheetApp.ProtectionType.SHEET].forEach(function (type) {
    sh.getProtections(type).forEach(function (p) {
      if (String(p.getDescription()).indexOf(INV.tag) === 0 && p.canEdit()) p.remove();
    });
  });
}

function installTriggers_(ss) {
  var mine = { onInventoryEdit: 1, onInventoryChange: 1, onCheckoutSubmit: 1 };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (mine[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onInventoryEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onInventoryChange').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('onCheckoutSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

// =============================================================================
// Keeping things in sync
// =============================================================================

// Installable onEdit (runs for every editor).
function onInventoryEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var name = sh.getName();
  try {
    if (name === INV.tabs.master) {
      if (e.range.getLastRow() < 2) { dropCache_(); return; }
      withLock_(function () {
        var mc = masterCols_(sh);
        var filled = assignIds_(sh, mc, e.range.getRow(), e.range.getLastRow());
        var c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
        var touched = ['id', 'name', 'cat', 'oos'].some(function (k) { return mc[k] >= c1 && mc[k] <= c2; });
        if (filled || touched) refreshFormChoices();
      });
    } else if (name === INV.tabs.log || name === INV.tabs.maint) {
      stampTimes_(sh, e.range);
      dropCache_();
    } else {
      dropCache_();
    }
  } catch (err) {
    console.error('onInventoryEdit: ' + (err && err.stack || err));
  }
}

// Installable onChange: deleting item rows does not fire onEdit.
function onInventoryChange(e) {
  dropCache_();
  if (e && e.changeType === 'REMOVE_ROW') {
    try { refreshFormChoices(); } catch (err) { console.error('onInventoryChange: ' + err); }
  }
}

// Installable onFormSubmit: a return copies its condition onto the Inventory Master.
function onCheckoutSubmit(e) {
  if (!e || !e.range) return;
  try {
    var sh = e.range.getSheet();
    var lc = logCols_(sh);
    var row = sh.getRange(e.range.getRow(), 1, 1, sh.getLastColumn()).getValues()[0];
    dropCache_();
    if (String(row[lc.action - 1]).trim() !== INV.ret) return;
    var cond = String(row[lc.cond - 1]).trim();
    if (cond) setMasterCondition_(itemIdOf_(row[lc.item - 1]), cond);
  } catch (err) {
    console.error('onCheckoutSubmit: ' + (err && err.stack || err));
  }
}

function itemIdOf_(text) { return String(text == null ? '' : text).trim().split(/\s+/)[0] || ''; }

function setMasterCondition_(id, cond) {
  if (!id) return false;
  var sh = ss_().getSheetByName(INV.tabs.master);
  var mc = masterCols_(sh);
  var last = lastDataRow_(sh, mc);
  if (last < 2) return false;
  var ids = sh.getRange(2, mc.id, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim().toUpperCase() === id.toUpperCase()) {
      sh.getRange(i + 2, mc.cond).setValue(cond);
      return true;
    }
  }
  return false;
}

// Give every row that has a name and a known category an ID like PT-007.
// Returns true when it filled anything in. Callers hold withLock_.
function assignIds_(sh, mc, fromRow, toRow) {
  var last = lastDataRow_(sh, mc);
  fromRow = Math.max(fromRow, 2);
  toRow = Math.min(toRow, last);
  if (toRow < fromRow) return false;
  var prefixes = categoryPrefixes_();
  var ids = sh.getRange(2, mc.id, last - 1, 1).getValues();
  var names = sh.getRange(2, mc.name, last - 1, 1).getValues();
  var cats = sh.getRange(2, mc.cat, last - 1, 1).getValues();
  var top = {};
  ids.forEach(function (r) {
    var m = String(r[0]).trim().match(/^([A-Za-z]+)-(\d+)$/);
    if (m) { var p = m[1].toUpperCase(); top[p] = Math.max(top[p] || 0, parseInt(m[2], 10)); }
  });
  var changed = false;
  for (var i = fromRow - 2; i <= toRow - 2; i++) {
    if (String(ids[i][0]).trim() || !String(names[i][0]).trim()) continue;
    var p = prefixes[String(cats[i][0]).trim().toLowerCase()];
    if (!p) continue;
    top[p] = (top[p] || 0) + 1;
    ids[i][0] = p + '-' + pad3_(top[p]);
    changed = true;
  }
  if (changed) sh.getRange(fromRow, mc.id, toRow - fromRow + 1, 1).setValues(ids.slice(fromRow - 2, toRow - 1));
  return changed;
}

function pad3_(n) { var s = String(n); while (s.length < 3) s = '0' + s; return s; }

function categoryPrefixes_() {
  var sh = ss_().getSheetByName(INV.tabs.lists);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var cat = String(r[0]).trim(), p = String(r[1]).trim().toUpperCase().replace(/[^A-Z]/g, '');
    if (cat && p) out[cat.toLowerCase()] = p;
  });
  return out;
}

// Hand-typed rows on the logs get a real timestamp (the plan's =NOW() would keep
// changing). Form submissions are stamped by Google and never reach onEdit.
function stampTimes_(sh, range) {
  var tsCol = headerMap_(sh)['timestamp'];
  if (!tsCol) return;
  if (range.getNumColumns() === 1 && range.getColumn() === tsCol) return;
  var r1 = Math.max(range.getRow(), 2), r2 = range.getLastRow();
  if (r2 < r1) return;
  var vals = sh.getRange(r1, 1, r2 - r1 + 1, sh.getLastColumn()).getValues();
  var now = new Date(), changed = false;
  var out = vals.map(function (row) {
    var hasData = row.some(function (v, j) { return j !== tsCol - 1 && String(v).trim() !== ''; });
    if (!row[tsCol - 1] && hasData) { changed = true; return [now]; }
    return [row[tsCol - 1]];
  });
  if (changed) sh.getRange(r1, tsCol, out.length, 1).setValues(out);
}

function fillMissingIds() {
  var sh = ss_().getSheetByName(INV.tabs.master);
  if (!sh) throw new Error('No Inventory Master tab. Run setupInventory first.');
  withLock_(function () {
    var mc = masterCols_(sh);
    assignIds_(sh, mc, 2, lastDataRow_(sh, mc));
    refreshFormChoices();
  });
}

// Rebuild the backup form's Item dropdown from the Inventory Master: everything
// with an ID and a name, except consumables and retired items. Also writes each
// item's Checkout Link (its page in the hub app).
function refreshFormChoices() {
  var ss = ss_();
  var form = getForm_();
  var sh = ss.getSheetByName(INV.tabs.master);
  if (!form || !sh) return 0;
  var item = itemQuestion_(form);
  if (!item) return 0;
  var mc = masterCols_(sh);
  var last = lastDataRow_(sh, mc);
  var data = last > 1 ? sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues() : [];

  var seen = {}, choices = [];
  var rowChoice = data.map(function (r) {
    var id = String(r[mc.id - 1]).trim(), name = String(r[mc.name - 1]).trim();
    var cat = String(r[mc.cat - 1]).trim().toLowerCase(), oos = String(r[mc.oos - 1]).trim().toLowerCase();
    if (!id || !name || cat === INV.consumable.toLowerCase() || oos === 'retired') return '';
    var choice = id + INV.sep + name;
    if (!seen[choice]) { seen[choice] = true; choices.push(choice); }
    return choice;
  });
  choices.sort();
  item.setChoiceValues(choices.length ? choices : [INV.noItems]);

  if (data.length) {
    sh.getRange(2, mc.link, data.length, 1).setValues(rowChoice.map(function (c) {
      return [c ? appLink_(itemIdOf_(c)) : ''];
    }));
  }
  // Clear links left below the data (rows that were emptied rather than deleted).
  if (last < sh.getMaxRows()) sh.getRange(last + 1, mc.link, sh.getMaxRows() - last, 1).clearContent();
  return choices.length;
}

function appLink_(id) { return INV.appUrl + '?item=' + encodeURIComponent(id); }

// =============================================================================
// Links, dialogs, config.js block
// =============================================================================

function linkInfo_(ss, form) {
  var published = form ? form.getPublishedUrl() : '';
  var short = published;
  if (form) { try { short = form.shortenFormUrl(published); } catch (e) { /* keep the long link */ } }
  var gid = function (name) { var s = ss.getSheetByName(name); return s ? String(s.getSheetId()) : ''; };
  return {
    apiUrl: webAppUrl_(),
    appUrl: INV.appUrl,
    sheetUrl: ss.getUrl(),
    formUrl: short,
    formEdit: form ? form.getEditUrl() : '',
    gids: { dash: gid(INV.tabs.dash), master: gid(INV.tabs.master), maint: gid(INV.tabs.maint) }
  };
}

// The deployed web app's /exec URL, or '' before the first deployment.
function webAppUrl_() {
  try {
    var u = ScriptApp.getService().getUrl() || '';
    return /\/exec$/.test(u) ? u : '';
  } catch (e) {
    return '';
  }
}

function writeSettings_(ss, form) {
  var sh = ss.getSheetByName(INV.tabs.lists);
  var info = linkInfo_(ss, form);
  sh.getRange(2, 8, 4, 2).setValues([
    ['Inventory app (share this)', info.appUrl],
    ['Backup form, if the app is down', info.formUrl],
    ['Backup form editor', info.formEdit],
    ['Set up on', new Date()]
  ]);
  sh.getRange(5, 9).setNumberFormat('mmm d, yyyy h:mm am/pm').setHorizontalAlignment('left');
}

function configBlock_(info) {
  return 'window.INVENTORY = {\n'
    + "  apiUrl: '" + info.apiUrl + "',\n"
    + "  appUrl: '" + info.appUrl + "',\n"
    + "  sheetUrl: '" + info.sheetUrl + "',\n"
    + "  dashboardGid: '" + info.gids.dash + "',\n"
    + "  masterGid: '" + info.gids.master + "',\n"
    + "  maintenanceGid: '" + info.gids.maint + "',\n"
    + "  formUrl: '" + info.formUrl + "',\n"
    + '};';
}

function showInventoryLinks() {
  var ss = ss_();
  showDialog_(linksHtml_(linkInfo_(ss, getForm_()), null), 'Inventory links and setup');
}

function runSelfTest() {
  var report = withLock_(selfTest_);
  showDialog_(linksHtml_(null, report), 'Self-test');
}

function addTestData() {
  var out = withLock_(addTestData_);
  ss_().toast('Added ' + out.added.items + ' test items, ' + out.added.log + ' check outs and returns, and '
    + out.added.maintenance + ' maintenance entries.', 'Test data loaded', 8);
}

function removeTestData() {
  var out = withLock_(removeTestData_);
  ss_().toast('Removed ' + out.removed.items + ' test items, ' + out.removed.log + ' log rows and '
    + out.removed.maintenance + ' maintenance rows.', 'Test data removed', 8);
}

function showDialog_(html, title) {
  try {
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(560), title);
  } catch (e) {
    // Run from the editor with the sheet closed: the Execution log has the same details.
    Logger.log('Could not open the dialog (' + e + '). Details are in this log.');
  }
}

function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function linksHtml_(info, report) {
  var css = '<style>body{font:14px/1.5 Arial,sans-serif;color:#202124;margin:0 4px}h3{font-size:15px;margin:16px 0 6px}'
    + 'a{color:#1F6FB2;word-break:break-all}textarea{width:100%;height:132px;font:12px/1.45 Menlo,Consolas,monospace;'
    + 'border:1px solid #ccc;border-radius:6px;padding:8px;box-sizing:border-box}.ok{color:#1E7A3C}.bad{color:#B31B1B}'
    + 'ul{margin:4px 0 0;padding-left:18px}li{margin:2px 0}button{margin-top:6px;padding:7px 14px;border:0;border-radius:6px;'
    + 'background:#B31B1B;color:#fff;font-weight:bold;cursor:pointer}.muted{color:#5F6368;font-size:12.5px}</style>';
  var out = css;
  if (report) {
    out += '<h3 class="' + (report.ok ? 'ok' : 'bad') + '">' + esc_(report.summary) + '</h3><ul>'
      + report.lines.map(function (l) { return '<li>' + esc_(l) + '</li>'; }).join('') + '</ul>';
  }
  if (info) {
    out += info.apiUrl
      ? '<h3 class="ok">Web app is deployed</h3><p class="muted">The hub talks to this sheet through ' + esc_(info.apiUrl) + '</p>'
      : '<h3 class="bad">One step left: deploy the web app</h3><p>Deploy &gt; New deployment &gt; type <b>Web app</b>. Execute as: <b>Me</b>. Who has access: <b>Anyone</b>. Deploy, copy the URL that ends in /exec, and put it in config.js as <b>apiUrl</b>. Then open this dialog again (Inventory menu) to get the full block.</p>';
    out += '<h3>Paste this into the hub\'s config.js</h3><textarea id="cfg" readonly>' + esc_(configBlock_(info)) + '</textarea>'
      + '<button onclick="var t=document.getElementById(\'cfg\');t.select();try{document.execCommand(\'copy\');this.textContent=\'Copied\';}catch(e){}">Copy</button>'
      + '<p class="muted">It replaces the window.INVENTORY block. The hub\'s inventory app, labels and QR poster read it.</p>'
      + (info.formUrl ? '<h3>Backup form</h3><p class="muted">If the app is ever down, people can still log check outs at <a href="' + esc_(info.formUrl) + '" target="_blank">' + esc_(info.formUrl) + '</a>. It writes to the same Checkout Log.</p>' : '');
  }
  return out;
}

// =============================================================================
// Self-test: runs the real formulas in this spreadsheet with a throwaway item,
// then deletes every row it added.
// =============================================================================

function selfTest_() {
  var ss = ss_();
  var master = ss.getSheetByName(INV.tabs.master);
  var log = ss.getSheetByName(INV.tabs.log);
  var maint = ss.getSheetByName(INV.tabs.maint);
  var dash = ss.getSheetByName(INV.tabs.dash);
  var lines = [], fails = 0;
  var check = function (label, pass, got) {
    if (!pass) fails++;
    lines.push((pass ? 'PASS  ' : 'FAIL  ') + label + (pass ? '' : ' (got: ' + got + ')'));
  };
  if (!master || !log || !maint || !dash) {
    return { ok: false, summary: 'Self-test could not run: a tab is missing. Run setupInventory.', lines: [] };
  }

  var mc = masterCols_(master), lc = logCols_(log), tc = colsBySpec_(maint, MAINT);
  var id = 'ZZ-999', name = 'Self-test item (removed automatically)', who = 'zz999';
  var choice = id + INV.sep + name;
  var day = 86400000;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var yesterday = new Date(today.getTime() - day), inThree = new Date(today.getTime() + 3 * day);
  var sameDay = function (v, d) { return v instanceof Date && v.getFullYear() === d.getFullYear() && v.getMonth() === d.getMonth() && v.getDate() === d.getDate(); };
  var added = { master: 0, log: [], maint: 0 };
  var cell = function (sh, row, col) { return sh.getRange(row, col).getValue(); };
  var dashIds = function (key, offset) {
    var b = DASH.blocks.filter(function (x) { return x.key === key; })[0];
    return dash.getRange(b.row + 2, b.col + (offset || 0), INV.listRows, 1).getValues().map(function (r) { return String(r[0]); });
  };
  var addLog = function (action, extra) {
    var r = log.getLastRow() + 1;
    log.getRange(r, lc.ts).setValue(new Date(Date.now() - (action === INV.checkOut ? 120000 : 60000)));
    log.getRange(r, lc.item).setValue(choice);
    log.getRange(r, lc.action).setValue(action);
    log.getRange(r, lc.netid).setValue(who);
    Object.keys(extra || {}).forEach(function (k) { log.getRange(r, lc[k]).setValue(extra[k]); });
    added.log.push(r);
  };

  try {
    var mr = lastDataRow_(master, mc) + 1;
    if (mr > master.getMaxRows()) master.insertRowsAfter(master.getMaxRows(), 1);
    master.getRange(mr, mc.id).setValue(id);
    master.getRange(mr, mc.name).setValue(name);
    master.getRange(mr, mc.cat).setValue('Equipment');
    added.master = mr;
    SpreadsheetApp.flush();
    check('A new item shows Available', cell(master, mr, mc.status) === 'Available', cell(master, mr, mc.status));

    addLog(INV.checkOut, { due: yesterday });
    SpreadsheetApp.flush();
    check('Checking out sets Status to Checked Out', cell(master, mr, mc.status) === 'Checked Out', cell(master, mr, mc.status));
    check('Current Holder comes from the log', cell(master, mr, mc.holder) === who, cell(master, mr, mc.holder));
    check('Due Date comes from the log', sameDay(cell(master, mr, mc.due), yesterday), cell(master, mr, mc.due));
    var tileOverdue = DASH.tiles.filter(function (t) { return t.key === 'overdue'; })[0];
    check('Dashboard counts it as overdue', Number(cell(dash, DASH.tileRow + 1, tileOverdue.col)) >= 1, cell(dash, DASH.tileRow + 1, tileOverdue.col));
    check('Dashboard lists it under Overdue', dashIds('overdue').indexOf(id) >= 0, dashIds('overdue')[0]);
    check('Dashboard lists the borrower under Who has what', dashIds('who').indexOf(who) >= 0, dashIds('who')[0]);
    check('Dashboard lists it under Checked out', dashIds('out').indexOf(id) >= 0, dashIds('out')[0]);

    addLog(INV.ret, { cond: 'Needs Repair' });
    setMasterCondition_(id, 'Needs Repair');   // what onCheckoutSubmit does for a real form return
    SpreadsheetApp.flush();
    check('Returning sets Status back to Available', cell(master, mr, mc.status) === 'Available', cell(master, mr, mc.status));
    check('Current Holder clears on return', cell(master, mr, mc.holder) === '', cell(master, mr, mc.holder));
    check('The return condition lands on the Inventory Master', cell(master, mr, mc.cond) === 'Needs Repair', cell(master, mr, mc.cond));
    check('Dashboard lists it under Needs repair', dashIds('repair').indexOf(id) >= 0, dashIds('repair')[0]);

    var tr = maint.getLastRow() + 1;
    maint.getRange(tr, tc.ts).setValue(new Date());
    maint.getRange(tr, tc.item).setValue(choice);
    maint.getRange(tr, tc.issue).setValue('Self-test');
    maint.getRange(tr, tc.next).setValue(inThree);
    added.maint = tr;
    SpreadsheetApp.flush();
    check('Next Maintenance comes from the Maintenance Log', sameDay(cell(master, mr, mc.nextMaint), inThree), cell(master, mr, mc.nextMaint));
    check('Dashboard lists it under Maintenance due', dashIds('maint').indexOf(id) >= 0, dashIds('maint')[0]);

    master.getRange(mr, mc.cat).setValue(INV.consumable);
    master.getRange(mr, mc.qty).setValue(1);
    master.getRange(mr, mc.min).setValue(5);
    SpreadsheetApp.flush();
    check('A consumable below Min Qty shows Low Stock', cell(master, mr, mc.status) === 'Low Stock', cell(master, mr, mc.status));
    check('Dashboard lists it under Low or out of stock', dashIds('stock').indexOf(id) >= 0, dashIds('stock')[0]);

    master.getRange(mr, mc.oos).setValue('Retired');
    SpreadsheetApp.flush();
    check('Out of Service = Retired shows Retired', cell(master, mr, mc.status) === 'Retired', cell(master, mr, mc.status));

    var shown = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    var broken = shown.filter(function (v) { return /^#/.test(v); });
    check('No formula errors in the Inventory Master headers', broken.length === 0, broken.join(' '));
  } catch (err) {
    fails++;
    lines.push('ERROR ' + (err && err.message || err));
  } finally {
    // Remove only the rows this test wrote, bottom-up, after checking they are ours.
    if (added.maint && String(maint.getRange(added.maint, tc.item).getValue()) === choice) maint.deleteRow(added.maint);
    added.log.sort(function (a, b) { return b - a; }).forEach(function (r) {
      if (String(log.getRange(r, lc.item).getValue()) === choice) log.deleteRow(r);
    });
    // The Inventory Master row is cleared rather than deleted so the sheet keeps its size.
    if (added.master && String(master.getRange(added.master, mc.id).getValue()) === id) {
      ['id', 'name', 'cat', 'cond', 'oos', 'qty', 'min', 'link'].forEach(function (k) {
        master.getRange(added.master, mc[k]).clearContent();
      });
    }
    SpreadsheetApp.flush();
  }

  var ok = fails === 0;
  return {
    ok: ok,
    summary: ok ? 'Self-test passed (' + lines.length + ' checks)' : 'Self-test: ' + fails + ' problem(s). Send this list to whoever maintains the script.',
    lines: lines
  };
}

// =============================================================================
// Web API for the hub's inventory page (inventory.html)
// The page calls it JSONP-style: <script src=".../exec?action=list&callback=fn">,
// which works from the static hub without any CORS setup. Reads and check out /
// return are open to anyone with the page; catalog changes need admin=1, which
// only the (unlisted) admin page sends, the same convention as the other modules.
// Photos are too big for a URL, so they come in as a POST (doPost).
// =============================================================================

var CACHE_KEY = 'inv_list_v2';
var API_VERSION = 2;   // the app turns on photos and teams from version 2

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    out = api_(p);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  var json = JSON.stringify(out);
  var cb = String(p.callback || '');
  if (/^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function api_(p) {
  var action = String(p.action || 'list');
  var adminOnly = function () {
    if (p.admin !== '1') throw new Error('That change needs the admin view of the inventory.');
  };
  switch (action) {
    case 'list': return apiList_();
    case 'item': return apiItem_(p.id);
    case 'checkout': return withLock_(function () { return apiCheckout_(p); });
    case 'return': return withLock_(function () { return apiReturn_(p); });
    case 'save': adminOnly(); return withLock_(function () { return apiSave_(p); });
    case 'maintenance': adminOnly(); return withLock_(function () { return apiMaintenance_(p); });
    case 'count': adminOnly(); return withLock_(function () { return apiCount_(p); });
    case 'removephoto': adminOnly(); return withLock_(function () { return apiRemovePhoto_(p); });
    case 'testdata':
      adminOnly();
      if (p.mode === 'remove') return withLock_(removeTestData_);
      if (p.mode === 'reset') return withLock_(resetTestData_);
      return withLock_(addTestData_);
    default: throw new Error('Unknown request.');
  }
}

// Photo uploads. The app sends JSON as text/plain (a "simple" request, so the
// browser needs no preflight) and reads the JSON reply.
function doPost(e) {
  var out;
  try {
    var body;
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { throw new Error('Could not read the upload.'); }
    if (String(body.action || '') !== 'photo') throw new Error('Unknown request.');
    if (String(body.admin || '') !== '1') throw new Error('Adding photos needs the admin view of the inventory.');
    out = apiPhoto_(body);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function dropCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) { /* best effort */ }
}

function apiList_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);
  var out = listPayload_();
  var json = JSON.stringify(out);
  if (json.length < 90000) {
    try { cache.put(CACHE_KEY, json, 300); } catch (e) { /* too big or busy: skip the cache */ }
  }
  return out;
}

function listPayload_() {
  var ss = ss_();
  var tz = ss.getSpreadsheetTimeZone();
  var sh = ss.getSheetByName(INV.tabs.master);
  if (!sh) throw new Error('The inventory sheet is not set up yet. Run setupInventory.');
  var mc = masterCols_(sh);
  var last = lastDataRow_(sh, mc);
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues() : [];
  var items = [];
  rows.forEach(function (r) {
    var it = itemFromRow_(r, mc, tz);
    if (it.id) items.push(it);
  });
  return {
    ok: true,
    items: items,
    lists: readLists_(ss),
    meta: {
      api: API_VERSION,
      today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
      updated: new Date().toISOString(),
      testItems: items.filter(function (i) { return i.test; }).length
    }
  };
}

function itemFromRow_(r, mc, tz) {
  var get = function (k) { var v = r[mc[k] - 1]; return v == null ? '' : v; };
  var str = function (k) { return String(get(k)).trim(); };
  var day = function (k) { var v = get(k); return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim(); };
  var num = function (k) { var v = get(k); return v === '' || isNaN(Number(v)) ? null : Number(v); };
  var notes = str('notes');
  return {
    id: str('id').toUpperCase(), name: str('name'), cat: str('cat'), sub: str('sub'), loc: str('loc'),
    team: (mc.team ? str('team') : '') || INV.shared,
    status: str('status'), holder: str('holder'), due: day('due'), cond: str('cond'), oos: str('oos'),
    serial: str('serial'), bought: day('bought'), qty: num('qty'), min: num('min'), nextMaint: day('nextMaint'),
    notes: notes, photo: str('photo'),
    consumable: str('cat').toLowerCase() === INV.consumable.toLowerCase(),
    test: notes.toLowerCase().indexOf(INV.testTag) >= 0
  };
}

function readLists_(ss) {
  var out = { categories: [], subcategories: [], locations: [], teams: [], conditions: INV.conditions, outOfService: INV.outOfService };
  var sh = ss.getSheetByName(INV.tabs.lists);
  var teamCol = false;
  if (sh && sh.getLastRow() >= 2) {
    teamCol = String(sh.getRange(1, 5).getValue()).trim().toLowerCase() === 'team';
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
      var cat = String(r[0]).trim(), sub = String(r[2]).trim(), loc = String(r[3]).trim(), team = String(r[4]).trim();
      if (cat) out.categories.push({ name: cat, prefix: String(r[1]).trim().toUpperCase() });
      if (sub) out.subcategories.push(sub);
      if (loc) out.locations.push(loc);
      if (team && teamCol) out.teams.push(team);
    });
  }
  if (!teamCol) out.teams = SEED.teams.slice();
  if (out.teams.indexOf(INV.shared) < 0) out.teams.unshift(INV.shared);
  return out;
}

// A new location or subcategory typed in the app joins its dropdown on the Lists
// tab, so the sheet and the app offer it from then on. col: 3 or 4.
function addToList_(ss, col, value) {
  var v = String(value == null ? '' : value).replace(/^'/, '').trim();
  if (!v) return false;
  var sh = ss.getSheetByName(INV.tabs.lists);
  if (!sh) return false;
  var last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, col, last - 1, 1).getValues() : [];
  var end = 0;
  for (var i = 0; i < vals.length; i++) {
    var have = String(vals[i][0]).trim();
    if (have.toLowerCase() === v.toLowerCase()) return false;
    if (have) end = i + 1;
  }
  var row = end + 2;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  sh.getRange(row, col).setValue(value);
  return true;
}

function findItem_(id) {
  var ss = ss_();
  var sh = ss.getSheetByName(INV.tabs.master);
  var mc = masterCols_(sh);
  var last = lastDataRow_(sh, mc);
  if (last >= 2) {
    var ids = sh.getRange(2, mc.id, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim().toUpperCase() === id) {
        return { ss: ss, sh: sh, mc: mc, row: i + 2, item: readRow_(sh, mc, i + 2, ss.getSpreadsheetTimeZone()) };
      }
    }
  }
  throw new Error('There is no item ' + id + ' in the inventory.');
}

function readRow_(sh, mc, row, tz) {
  return itemFromRow_(sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0], mc, tz);
}

function apiItem_(rawId) {
  var ctx = findItem_(reqId_(rawId));
  var tz = ctx.ss.getSpreadsheetTimeZone();
  return { ok: true, item: ctx.item, history: history_(ctx.ss, ctx.item.id, tz), maintenance: maintHistory_(ctx.ss, ctx.item.id, tz) };
}

// Newest first: exact time, then sheet order (later rows are newer) for ties.
function newestFirst_(list) {
  list.sort(function (a, b) { return (b._t - a._t) || (b._i - a._i); });
  return list.map(function (x) { delete x._t; delete x._i; return x; });
}
function ms_(v) { return v instanceof Date ? v.getTime() : 0; }

function stamp_(v, tz) { return v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm") : String(v == null ? '' : v).trim(); }
function dayStr_(v, tz) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v == null ? '' : v).trim(); }

function history_(ss, id, tz) {
  var log = ss.getSheetByName(INV.tabs.log);
  if (!log || log.getLastRow() < 2) return [];
  var lc = logCols_(log);
  var out = [];
  log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues().forEach(function (r, i) {
    if (itemIdOf_(r[lc.item - 1]).toUpperCase() !== id) return;
    out.push({
      when: stamp_(r[lc.ts - 1], tz), action: String(r[lc.action - 1]).trim(), netid: String(r[lc.netid - 1]).trim(),
      due: dayStr_(r[lc.due - 1], tz), cond: String(r[lc.cond - 1]).trim(), notes: lc.notes ? String(r[lc.notes - 1]).trim() : '',
      _t: ms_(r[lc.ts - 1]), _i: i
    });
  });
  return newestFirst_(out).slice(0, 15);
}

function maintHistory_(ss, id, tz) {
  var sh = ss.getSheetByName(INV.tabs.maint);
  if (!sh || sh.getLastRow() < 2) return [];
  var tc = colsBySpec_(sh, MAINT);
  var out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function (r, i) {
    if (itemIdOf_(r[tc.item - 1]).toUpperCase() !== id) return;
    var cost = r[tc.cost - 1];
    out.push({
      when: stamp_(r[tc.ts - 1], tz), issue: String(r[tc.issue - 1]).trim(), taken: String(r[tc.action - 1]).trim(),
      by: String(r[tc.by - 1]).trim(), next: dayStr_(r[tc.next - 1], tz), cost: cost === '' || isNaN(Number(cost)) ? null : Number(cost),
      _t: ms_(r[tc.ts - 1]), _i: i
    });
  });
  return newestFirst_(out).slice(0, 10);
}

function apiCheckout_(p) {
  var id = reqId_(p.id), netid = reqNetid_(p.netid), due = reqDate_(p.due, 'Due back'), notes = text_(p.notes, 500);
  var ctx = findItem_(id), it = ctx.item, tz = ctx.ss.getSpreadsheetTimeZone();
  if (it.consumable) throw new Error('Consumables are not checked out. Ask an admin to update the count instead.');
  if (it.oos) throw new Error(it.name + ' is out of service (' + it.oos + ').');
  if (it.status === 'Checked Out') throw new Error(it.name + ' is already checked out to ' + (it.holder || 'someone') + '. It has to be returned first.');
  if (due < Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')) throw new Error('The due date is in the past.');
  appendLogRows_(ctx.ss, [{ item: id + INV.sep + it.name, action: INV.checkOut, netid: netid, due: due, notes: notes }]);
  SpreadsheetApp.flush();
  return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, tz) };
}

function apiReturn_(p) {
  var id = reqId_(p.id), netid = reqNetid_(p.netid), notes = text_(p.notes, 500);
  var cond = oneOf_(p.cond, INV.conditions, 'condition');
  var ctx = findItem_(id), it = ctx.item;
  if (it.status !== 'Checked Out') throw new Error(it.name + ' is not checked out.');
  appendLogRows_(ctx.ss, [{ item: id + INV.sep + it.name, action: INV.ret, netid: netid, cond: cond, notes: notes }]);
  ctx.sh.getRange(ctx.row, ctx.mc.cond).setValue(cond);
  SpreadsheetApp.flush();
  return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, ctx.ss.getSpreadsheetTimeZone()) };
}

// Add (no id) or edit (id) an item. fields is a JSON object from the app's form.
function apiSave_(p) {
  var f;
  try { f = JSON.parse(String(p.fields || '{}')); } catch (e) { throw new Error('Could not read the item details.'); }
  var ss = ss_(), sh = ss.getSheetByName(INV.tabs.master), mc = masterCols_(sh);
  var lists = readLists_(ss);
  var name = text_(f.name, 120);
  if (!name) throw new Error('Give the item a name.');
  var cat = oneOf_(f.cat, lists.categories.map(function (c) { return c.name; }), 'category');
  var vals = {
    name: name, cat: cat, sub: text_(f.sub, 80), loc: text_(f.loc, 80), team: text_(f.team, 60) || INV.shared,
    cond: f.cond ? oneOf_(f.cond, INV.conditions, 'condition') : '',
    oos: f.oos ? oneOf_(f.oos, INV.outOfService, 'out of service option') : '',
    serial: text_(f.serial, 80), bought: optDate_(f.bought, 'Purchase date'),
    qty: optNum_(f.qty, 'Current qty'), min: optNum_(f.min, 'Min qty'),
    notes: text_(f.notes, 1000), photo: optUrl_(f.photo)
  };
  var id = f.id ? reqId_(f.id) : '', row;
  if (id) {
    row = findItem_(id).row;
    // Fields the app did not send keep what the sheet has (the photo, for one,
    // is changed separately).
    var cur = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    Object.keys(vals).forEach(function (k) {
      if (f[k] === undefined && mc[k]) vals[k] = cur[mc[k] - 1];
    });
  } else {
    if (!categoryPrefixes_()[cat.toLowerCase()]) throw new Error('The ' + cat + ' category has no ID prefix on the Lists tab.');
    row = lastDataRow_(sh, mc) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
  }
  writeItems_(sh, mc, row, [vals]);
  if (f.loc !== undefined) addToList_(ss, 4, vals.loc);
  if (f.sub !== undefined) addToList_(ss, 3, vals.sub);
  SpreadsheetApp.flush();
  if (!id) assignIds_(sh, mc, row, row);
  SpreadsheetApp.flush();
  try { refreshFormChoices(); } catch (e) { console.error('form sync: ' + e); }
  return { ok: true, item: readRow_(sh, mc, row, ss.getSpreadsheetTimeZone()) };
}

function apiMaintenance_(p) {
  var id = reqId_(p.id);
  var issue = text_(p.issue, 300);
  if (!issue) throw new Error('Say what the issue or the work was.');
  var row = { issue: issue, taken: text_(p.taken, 300), by: text_(p.by, 80), next: optDate_(p.next, 'Next due'), cost: optNum_(p.cost, 'Cost') };
  var ctx = findItem_(id);
  row.item = id + INV.sep + ctx.item.name;
  appendMaintRows_(ctx.ss, [row]);
  SpreadsheetApp.flush();
  var tz = ctx.ss.getSpreadsheetTimeZone();
  return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, tz), maintenance: maintHistory_(ctx.ss, id, tz) };
}

function apiCount_(p) {
  var id = reqId_(p.id);
  var qty = optNum_(p.qty, 'The count');
  if (qty === '') throw new Error('Enter how many are on hand.');
  var ctx = findItem_(id);
  if (!ctx.item.consumable) throw new Error('Only consumables have a count.');
  ctx.sh.getRange(ctx.row, ctx.mc.qty).setValue(Math.round(Number(qty)));
  SpreadsheetApp.flush();
  return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, ctx.ss.getSpreadsheetTimeZone()) };
}

// ---- photos -------------------------------------------------------------------

// body: { id, data: 'data:image/jpeg;base64,...' }. The app shrinks photos to about
// 1600 px before sending, so they are usually well under 1 MB.
function apiPhoto_(body) {
  var id = reqId_(body.id);
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=\s]+)$/.exec(String(body.data || ''));
  if (!m) throw new Error('That file is not a photo. Use a JPG, PNG or WebP image.');
  var bytes = Utilities.base64Decode(m[2].replace(/\s+/g, ''));
  if (bytes.length > 8 * 1024 * 1024) throw new Error('That photo is too big. Try one under 8 MB.');
  var item = findItem_(id).item;   // a bad ID fails before anything is saved
  var ext = m[1] === 'jpeg' ? 'jpg' : m[1];
  var name = (id + ' ' + item.name).replace(/[\\\/:*?"<>|]+/g, ' ').slice(0, 90) + '.' + ext;
  var folder = withLock_(photosFolder_);
  var file = folder.createFile(Utilities.newBlob(bytes, 'image/' + m[1], name));
  try { file.setDescription('Inventory photo of ' + id + ' (' + item.name + ')'); } catch (e) { /* cosmetic */ }
  try {
    return withLock_(function () {
      var ctx = findItem_(id);
      var old = ctx.item.photo;
      ctx.sh.getRange(ctx.row, ctx.mc.photo).setValue(driveViewUrl_(file.getId()));
      SpreadsheetApp.flush();
      trashIfOurs_(old);
      return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, ctx.ss.getSpreadsheetTimeZone()) };
    });
  } catch (err) {
    try { file.setTrashed(true); } catch (e) { /* nothing more to undo */ }
    throw err;
  }
}

function apiRemovePhoto_(p) {
  var ctx = findItem_(reqId_(p.id));
  var old = ctx.item.photo;
  ctx.sh.getRange(ctx.row, ctx.mc.photo).clearContent();
  SpreadsheetApp.flush();
  trashIfOurs_(old);
  return { ok: true, item: readRow_(ctx.sh, ctx.mc, ctx.row, ctx.ss.getSpreadsheetTimeZone()) };
}

function driveViewUrl_(fileId) { return 'https://drive.google.com/file/d/' + fileId + '/view'; }

// The Drive file ID inside a Drive link, or ''.
function driveId_(url) {
  var s = String(url || '');
  var m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(s) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(s);
  return m ? m[1] : '';
}

// The "Inventory photos" folder, next to this spreadsheet. Link-shared once, so
// every photo saved in it can be shown in the app without a per-file call.
function photosFolder_() {
  var props = props_();
  var folder = null;
  var id = props.getProperty('INV_PHOTO_FOLDER');
  if (id) {
    try { folder = DriveApp.getFolderById(id); if (folder.isTrashed()) folder = null; } catch (e) { folder = null; }
  }
  if (!folder) {
    var parent = null;
    try {
      var parents = DriveApp.getFileById(ss_().getId()).getParents();
      if (parents.hasNext()) parent = parents.next();
    } catch (e) { parent = null; }
    var found = (parent || DriveApp.getRootFolder()).getFoldersByName(INV.photoFolder);
    if (found.hasNext()) folder = found.next();
    else {
      try { folder = (parent || DriveApp.getRootFolder()).createFolder(INV.photoFolder); }
      catch (e) { folder = DriveApp.createFolder(INV.photoFolder); }
    }
    props.setProperty('INV_PHOTO_FOLDER', folder.getId());
    props.deleteProperty('INV_PHOTO_SHARED');
  }
  if (props.getProperty('INV_PHOTO_SHARED') !== folder.getId()) {
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      props.setProperty('INV_PHOTO_SHARED', folder.getId());
    } catch (e) {
      console.warn('Could not link-share the photos folder: ' + e);
    }
  }
  return folder;
}

// Trash a photo this script saved (only ones in the Inventory photos folder, so a
// link someone pasted to their own file is never touched).
function trashIfOurs_(url) {
  var fid = driveId_(url);
  var folderId = props_().getProperty('INV_PHOTO_FOLDER');
  if (!fid || !folderId) return false;
  try {
    var file = DriveApp.getFileById(fid);
    var parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folderId) { file.setTrashed(true); return true; }
    }
  } catch (e) {
    console.warn('Could not trash the old photo: ' + e);
  }
  return false;
}

// ---- writing rows -----------------------------------------------------------

// Write item fields into consecutive rows. Only typed-in columns are written: never
// the calculated ones (that would stop their formulas from filling the column) and
// never Item ID (new rows get one from assignIds_, existing IDs never change).
function writeItems_(sh, mc, startRow, items) {
  var keys = MASTER.filter(function (s) { return !s.calc && s.key !== 'id'; })
    .map(function (s) { return s.key; })
    .sort(function (a, b) { return mc[a] - mc[b]; });
  var runs = [];
  keys.forEach(function (k) {
    var run = runs[runs.length - 1];
    if (run && mc[k] === run.end + 1) { run.keys.push(k); run.end = mc[k]; }
    else runs.push({ start: mc[k], end: mc[k], keys: [k] });
  });
  runs.forEach(function (run) {
    sh.getRange(startRow, run.start, items.length, run.keys.length).setValues(items.map(function (o) {
      return run.keys.map(function (k) { return o[k] == null ? '' : o[k]; });
    }));
  });
}

function blankRow_(sh) {
  var row = [];
  for (var i = 0; i < sh.getLastColumn(); i++) row.push('');
  return row;
}

// Rows: { item, action, netid, due, cond, notes, when? }
function appendLogRows_(ss, rows) {
  var log = ss.getSheetByName(INV.tabs.log);
  var lc = logCols_(log);
  var out = rows.map(function (o) {
    var row = blankRow_(log);
    var set = function (k, v) { if (lc[k]) row[lc[k] - 1] = v == null ? '' : v; };
    set('ts', o.when || new Date()); set('item', o.item); set('action', o.action); set('netid', o.netid);
    set('notes', o.notes); set('due', o.due); set('cond', o.cond);
    return row;
  });
  var start = log.getLastRow() + 1;
  var need = start + out.length - 1 - log.getMaxRows();
  if (need > 0) log.insertRowsAfter(log.getMaxRows(), need);
  log.getRange(start, 1, out.length, out[0].length).setValues(out);
}

// Rows: { item, issue, taken, by, next, cost, when? }
function appendMaintRows_(ss, rows) {
  var sh = ss.getSheetByName(INV.tabs.maint);
  var tc = colsBySpec_(sh, MAINT);
  var out = rows.map(function (o) {
    var row = blankRow_(sh);
    var set = function (k, v) { if (tc[k]) row[tc[k] - 1] = v == null ? '' : v; };
    set('ts', o.when || new Date()); set('item', o.item); set('issue', o.issue); set('action', o.taken);
    set('by', o.by); set('next', o.next); set('cost', o.cost);
    return row;
  });
  var start = sh.getLastRow() + 1;
  var need = start + out.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  sh.getRange(start, 1, out.length, out[0].length).setValues(out);
}

// Delete the given rows, bottom-up, in as few calls as possible. A sheet must keep
// one row under its header, so a block reaching the very bottom is cleared instead.
function deleteRows_(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < rows.length) {
    var end = rows[i], start = end, j = i + 1;
    while (j < rows.length && rows[j] === start - 1) { start = rows[j]; j++; }
    if (start <= 2 && end >= sh.getMaxRows()) sh.getRange(start, 1, end - start + 1, sh.getMaxColumns()).clearContent();
    else sh.deleteRows(start, end - start + 1);
    i = j;
  }
}

// ---- input checks -----------------------------------------------------------

function reqId_(v) {
  var s = String(v == null ? '' : v).trim().toUpperCase();
  if (!/^[A-Z]+-\d+$/.test(s)) throw new Error('That is not an item ID.');
  return s;
}
function reqNetid_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (!/^[a-z]{2,4}[0-9]{1,6}$/.test(s)) throw new Error('Enter a NetID like abc123.');
  return s;
}
function optDate_(v, label) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  var d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  if (!d || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) throw new Error(label + ' must be a date.');
  return s;
}
function reqDate_(v, label) {
  var s = optDate_(v, label);
  if (!s) throw new Error(label + ' is required.');
  return s;
}
function optNum_(v, label) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var n = Number(s);
  if (!isFinite(n) || n < 0) throw new Error(label + ' must be a number, 0 or more.');
  return n;
}
function optUrl_(v) {
  var s = String(v == null ? '' : v).trim().slice(0, 500);
  if (s && !/^https?:\/\//i.test(s)) throw new Error('The photo link must start with https://');
  return s;
}
function oneOf_(v, list, label) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  for (var i = 0; i < list.length; i++) if (String(list[i]).toLowerCase() === s) return list[i];
  throw new Error('Pick a ' + label + ' from the list.');
}
// Free text, one line, capped. A leading = + - @ would make Sheets treat it as a
// formula, so it gets a text marker.
function text_(v, max) {
  var s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// =============================================================================
// Test data: realistic sample items, check outs and maintenance, all tagged
// "[test data]" so removeTestData_ can take back exactly what it added.
// =============================================================================

var TEST_ITEMS = [
  { key: 'drill1', name: 'DeWalt 20V cordless drill', cat: 'Power Tool', sub: 'Drill', loc: 'Upson ELL', serial: 'DCD771-1042', bought: -420 },
  { key: 'drill2', name: 'DeWalt 20V cordless drill (second)', cat: 'Power Tool', sub: 'Drill', loc: 'Upson ELL', serial: 'DCD771-1043', bought: -420 },
  { key: 'impact', name: 'Milwaukee M18 impact driver', cat: 'Power Tool', sub: 'Impact Driver', loc: 'Upson ELL', serial: '2853-20-7719', bought: -300 },
  { key: 'circ', name: 'Makita cordless circular saw', cat: 'Power Tool', sub: 'Saw', loc: 'Upson Auto Lab' },
  { key: 'sander', name: 'Bosch random orbit sander', cat: 'Power Tool', sub: 'Sander', loc: 'Upson Composites Lab', cond: 'Needs Repair', notes: 'Hook-and-loop pad is worn.' },
  { key: 'cutoff', name: 'Milwaukee cutoff saw', cat: 'Power Tool', sub: 'Saw', loc: 'Upson Auto Lab', team: 'FSAE Racing' },
  { key: 'track', name: 'Festool track saw', cat: 'Power Tool', sub: 'Saw', loc: 'Upson Composites Lab', serial: 'TS55-88213', bought: -700, team: 'Solar Boat' },
  { key: 'heatgun', name: 'Heat gun', cat: 'Power Tool', sub: 'Heat Gun', loc: 'Upson LVL' },
  { key: 'olddrill', name: 'Corded drill (old)', cat: 'Power Tool', sub: 'Drill', loc: 'Upson Basement', cond: 'Fair', oos: 'Retired', notes: 'Chuck slips. Kept for parts.' },
  { key: 'sockets', name: 'Metric socket set', cat: 'Hand Tool', sub: 'Socket Set', loc: 'Upson Auto Lab', team: 'Baja' },
  { key: 'torque', name: 'Torque wrench, 3/8 in drive', cat: 'Hand Tool', sub: 'Wrenches', loc: 'Upson Auto Lab', team: 'Baja' },
  { key: 'clamps', name: 'Bar clamps (set of 4)', cat: 'Hand Tool', sub: 'Clamps', loc: 'Upson Composites Lab', team: 'Concrete Canoe' },
  { key: 'shield', name: 'Face shield', cat: 'Safety Tool', sub: 'Face Shield', loc: 'Upson ELL' },
  { key: 'extinguisher', name: 'Fire extinguisher, ELL east wall', cat: 'Safety Tool', sub: 'Fire Extinguisher', loc: 'Upson ELL', serial: 'FE-0412' },
  { key: 'firstaid', name: 'First aid kit, Auto Lab', cat: 'Safety Tool', sub: 'First Aid Kit', loc: 'Upson Auto Lab' },
  { key: 'prusa', name: 'Prusa MK4 3D printer', cat: 'Equipment', sub: '3D Printer (FDM)', loc: 'Upson LVL', serial: 'MK4-55210', bought: -500, team: 'Design Build Fly' },
  { key: 'hakko', name: 'Hakko FX-888D soldering station', cat: 'Equipment', sub: 'Soldering Station', loc: 'High Voltage Lab', team: 'Mars Rover' },
  { key: 'scope', name: 'Rigol DS1054Z oscilloscope', cat: 'Equipment', sub: 'Oscilloscope', loc: 'High Voltage Lab', serial: 'DS1ZA2024', bought: -600, team: 'Rocketry' },
  { key: 'fluke', name: 'Fluke 117 multimeter', cat: 'Equipment', sub: 'Multimeter', loc: 'High Voltage Lab' },
  { key: 'tablesaw', name: 'Craftsman contractor table saw', cat: 'Equipment', sub: 'Saw', loc: 'Mill Room', cond: 'Fair', oos: 'Maintenance', notes: 'Needs a rip fence before use.', team: 'FSAE Racing' },
  { key: 'press', name: 'Floor drill press', cat: 'Equipment', sub: 'Drill Press', loc: 'Mill Room' },
  { key: 'n95', name: 'N95 masks (box of 20)', cat: 'Consumable', sub: 'N95 Masks', loc: 'Upson ELL', qty: 4, min: 10 },
  { key: 'gloves', name: 'Nitrile gloves, size L (box)', cat: 'Consumable', sub: 'Gloves', loc: 'Upson Composites Lab', qty: 12, min: 5 },
  { key: 'solder', name: 'Solder, 60/40 rosin core', cat: 'Consumable', sub: 'Solder', loc: 'High Voltage Lab', qty: 6, min: 2 },
  { key: 'pla', name: 'PLA filament, 1 kg spool', cat: 'Consumable', sub: 'Filament', loc: 'Upson LVL', qty: 0, min: 3 },
  { key: 'sandpaper', name: 'Sandpaper, 120 grit (pack)', cat: 'Consumable', sub: 'Sandpaper', loc: 'Upson Composites Lab', qty: 8, min: 4 }
];

// [item key, Check Out / Return, days ago, NetID, due (days from today), condition, note]
var TEST_LOG = [
  ['fluke', 'out', 20, 'tst103', -14],
  ['fluke', 'in', 15, 'tst103', null, 'Good'],
  ['prusa', 'out', 12, 'tst101', -5],
  ['prusa', 'in', 5, 'tst101', null, 'Good'],
  ['drill1', 'out', 9, 'tst101', -2],
  ['sander', 'out', 6, 'tst104', -1],
  ['hakko', 'out', 4, 'tst102', -1],
  ['scope', 'out', 3, 'tst103', 4],
  ['sander', 'in', 2, 'tst104', null, 'Needs Repair', 'Pad is worn through.'],
  ['torque', 'out', 2, 'tst105', 0],
  ['impact', 'out', 1, 'tst102', 5],
  ['clamps', 'out', 0, 'tst104', 7]
];

// [item key, days ago, issue, action taken, by, next due (days from today), cost]
var TEST_MAINT = [
  ['prusa', 30, 'Clogged nozzle', 'Replaced the nozzle and recalibrated the first layer', 'tst202', -2, 12],
  ['extinguisher', 25, 'Monthly inspection', 'Gauge in the green, pin and seal intact', 'tst201', 5, null],
  ['firstaid', 10, 'Monthly restock', 'Restocked bandages and gauze', 'tst203', 20, 18.5],
  ['tablesaw', 3, 'Rip fence missing', 'Replacement fence ordered', 'tst201', 10, 89.99]
];

function countTestItems_(sh, mc) {
  var last = lastDataRow_(sh, mc);
  if (last < 2) return 0;
  return sh.getRange(2, mc.notes, last - 1, 1).getValues()
    .filter(function (r) { return String(r[0]).toLowerCase().indexOf(INV.testTag) >= 0; }).length;
}

function addTestData_() {
  var ss = ss_(), sh = ss.getSheetByName(INV.tabs.master), mc = masterCols_(sh);
  var tz = ss.getSpreadsheetTimeZone();
  var already = countTestItems_(sh, mc);
  if (already) throw new Error('Test data is already loaded (' + already + ' items). Remove it first.');
  var prefixes = categoryPrefixes_();
  var missing = TEST_ITEMS.filter(function (t) { return !prefixes[t.cat.toLowerCase()]; }).map(function (t) { return t.cat; });
  if (missing.length) throw new Error('These categories are missing from the Lists tab: ' + missing.filter(function (c, i) { return missing.indexOf(c) === i; }).join(', '));

  var dayFrom = function (offset) { return Utilities.formatDate(new Date(Date.now() + offset * 86400000), tz, 'yyyy-MM-dd'); };
  var at = function (daysAgo, hour) { var d = new Date(Date.now() - daysAgo * 86400000); d.setHours(hour, 15, 0, 0); return d; };

  var start = lastDataRow_(sh, mc) + 1;
  var need = start + TEST_ITEMS.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  writeItems_(sh, mc, start, TEST_ITEMS.map(function (t) {
    return {
      name: t.name, cat: t.cat, sub: t.sub || '', loc: t.loc || '', team: t.team || INV.shared, cond: t.cond || 'Good', oos: t.oos || '',
      serial: t.serial || '', bought: t.bought ? dayFrom(t.bought) : '',
      qty: t.qty == null ? '' : t.qty, min: t.min == null ? '' : t.min,
      notes: INV.testTag + (t.notes ? ' ' + t.notes : ''), photo: ''
    };
  }));
  SpreadsheetApp.flush();
  assignIds_(sh, mc, start, start + TEST_ITEMS.length - 1);
  SpreadsheetApp.flush();

  var ids = sh.getRange(start, mc.id, TEST_ITEMS.length, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  var label = {};
  TEST_ITEMS.forEach(function (t, i) { label[t.key] = ids[i] + INV.sep + t.name; });

  var logRows = TEST_LOG.map(function (e, i) {
    var out = e[1] === 'out';
    return {
      when: at(e[2], out ? 10 + (i % 6) : 15 + (i % 3)), item: label[e[0]], action: out ? INV.checkOut : INV.ret, netid: e[3],
      due: out ? dayFrom(e[4]) : '', cond: out ? '' : e[5], notes: INV.testTag + (e[6] ? ' ' + e[6] : '')
    };
  });
  appendLogRows_(ss, logRows);
  appendMaintRows_(ss, TEST_MAINT.map(function (m) {
    return { when: at(m[1], 11), item: label[m[0]], issue: m[2] + ' ' + INV.testTag, taken: m[3], by: m[4], next: dayFrom(m[5]), cost: m[6] == null ? '' : m[6] };
  }));
  SpreadsheetApp.flush();
  try { refreshFormChoices(); } catch (e) { console.error('form sync: ' + e); }
  return { ok: true, added: { items: ids.length, log: logRows.length, maintenance: TEST_MAINT.length } };
}

function removeTestData_() {
  var ss = ss_(), sh = ss.getSheetByName(INV.tabs.master), mc = masterCols_(sh);
  var tag = INV.testTag;
  var ids = {}, masterRows = [], photos = [];
  var last = lastDataRow_(sh, mc);
  if (last >= 2) {
    var notes = sh.getRange(2, mc.notes, last - 1, 1).getValues();
    var idv = sh.getRange(2, mc.id, last - 1, 1).getValues();
    var pics = sh.getRange(2, mc.photo, last - 1, 1).getValues();
    for (var i = 0; i < notes.length; i++) {
      if (String(notes[i][0]).toLowerCase().indexOf(tag) < 0) continue;
      masterRows.push(i + 2);
      var id = String(idv[i][0]).trim().toUpperCase();
      if (id) ids[id] = true;
      if (String(pics[i][0]).trim()) photos.push(String(pics[i][0]).trim());
    }
  }
  var pick = function (tab, itemCol, textCol) {
    var t = ss.getSheetByName(tab), rows = [];
    if (!t || t.getLastRow() < 2) return { sheet: t, rows: rows };
    t.getRange(2, 1, t.getLastRow() - 1, t.getLastColumn()).getValues().forEach(function (r, i) {
      var hit = ids[itemIdOf_(r[itemCol - 1]).toUpperCase()] || (textCol && String(r[textCol - 1]).toLowerCase().indexOf(tag) >= 0);
      if (hit) rows.push(i + 2);
    });
    return { sheet: t, rows: rows };
  };
  var lc = logCols_(ss.getSheetByName(INV.tabs.log));
  var tc = colsBySpec_(ss.getSheetByName(INV.tabs.maint), MAINT);
  var logHits = pick(INV.tabs.log, lc.item, lc.notes);
  var maintHits = pick(INV.tabs.maint, tc.item, tc.issue);
  deleteRows_(maintHits.sheet, maintHits.rows);
  deleteRows_(logHits.sheet, logHits.rows);
  deleteRows_(sh, masterRows);
  SpreadsheetApp.flush();
  photos.forEach(trashIfOurs_);
  try { refreshFormChoices(); } catch (e) { console.error('form sync: ' + e); }
  return { ok: true, removed: { items: masterRows.length, log: logHits.rows.length, maintenance: maintHits.rows.length, photos: photos.length } };
}

// Remove, then add again: a fresh copy of the sample data (its dates are relative
// to today, so overdue and due-soon items stay that way).
function resetTestData_() {
  var removed = removeTestData_();
  var added = addTestData_();
  return { ok: true, removed: removed.removed, added: added.added };
}
