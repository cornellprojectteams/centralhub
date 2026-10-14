/**
 * Project Teams Tool & Equipment Inventory: the backend of the Ops Hub's
 * inventory page (inventory.html). Bound to its own Google Sheet.
 * =============================================================================
 * The hub is the app. This spreadsheet is the database, and this script is both
 * its caretaker (tabs, formulas, Item IDs, self-test) and its web API.
 *
 * Tools stay in the shops; nobody borrows them. The inventory records what there
 * is, where it lives, what shape it is in, its maintenance, and supply counts.
 *
 * Tabs it builds
 *   Dashboard         The same numbers the app shows, for anyone in the sheet.
 *   Inventory Master  One row per tool or piece of equipment, plus consumables.
 *   Maintenance Log   Repairs, calibrations and inspections.
 *   Lists             Dropdown options (categories, locations, ...) and links.
 *   Latest            Hidden helper the formulas read. Leave it alone.
 *
 * SET UP (once)
 *   1. In a new, empty Google Sheet: Extensions > Apps Script, replace the sample
 *      code with this file, Save, then run setupInventory and approve access
 *      (Sheets, plus Drive for item photos). It builds the tabs and runs a
 *      self-test.
 *   2. Deploy > New deployment > type Web app. Execute as: Me. Who has access:
 *      Anyone. Deploy, then copy the Web app URL (it ends in /exec) into the
 *      hub's config.js as INVENTORY.apiUrl. That connects the app to this sheet.
 *   After editing this file later: Deploy > Manage deployments > pencil >
 *   Version: New version > Deploy. The URL stays the same.
 *   Running setupInventory again is safe. It repairs formulas, dropdowns and
 *   formatting, and never deletes your rows. On a sheet made before teams
 *   existed it also moves the Team column next to Location.
 *
 * SHEETS MADE WHEN THE INVENTORY HAD CHECK OUTS
 *   An earlier version tracked check outs and returns. Running setupInventory on
 *   a sheet from then takes that out: the backup Google Form is unlinked and moved
 *   to the Drive trash, the Checkout Log tab is deleted (or kept as "Old checkout
 *   log" if it holds rows that are not test data), and the Current Holder and Due
 *   Date columns go. Checkout Link becomes Item Link. Items and maintenance stay.
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
 *   about 25 sample items with maintenance. Everything it adds is tagged
 *   "[test data]"; Remove test data takes all of it back out.
 *
 * Triggers installed by setup (all on this spreadsheet):
 *   onInventoryEdit    fills Item IDs and Item Links, stamps times on hand-typed
 *                      Maintenance Log rows
 *   onInventoryChange  refreshes the app's copy after rows are deleted
 * =============================================================================
 */

var INV = {
  tabs: {
    dash: 'Dashboard',
    master: 'Inventory Master',
    maint: 'Maintenance Log',
    lists: 'Lists',
    calc: 'Latest'
  },
  inService: 'In service',    // Status of a tool that is not out of service
  conditions: ['Good', 'Fair', 'Needs Repair'],
  outOfService: ['Maintenance', 'Retired'],
  consumable: 'Consumable',   // formulas look for this category name
  sep: ' · ',            // item choices read "PT-001 · DeWalt 20V drill"
  listRows: 10,               // rows shown per Dashboard list
  // The hub's inventory page. Item QR labels and the Item Link column point here.
  appUrl: 'https://cornellprojectteams.github.io/centralhub/inventory.html',
  // Left over from when the inventory had check outs; setup removes them (retireBorrowing_).
  old: { log: 'Checkout Log', keptLog: 'Old checkout log', holder: 'Current Holder', due: 'Due Date', link: 'Checkout Link' },
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
  { key: 'loc', h: 'Location', w: 160, note: 'Where it lives.' },
  { key: 'team', h: 'Team', w: 150, note: 'The project team that owns it, or Shared for the program\'s own tools that any team can use. The list is on the Lists tab, column E.' },
  { key: 'status', h: 'Status', w: 110, calc: true, note: 'Calculated. In service, unless Out of Service says Maintenance or Retired. Consumables show In Stock, Low Stock or Out of Stock.' },
  { key: 'cond', h: 'Condition', w: 110, note: 'Good, Fair or Needs Repair. Set it back to Good after a repair.' },
  { key: 'oos', h: 'Out of Service', w: 120, note: 'Maintenance or Retired takes the item out of service. Leave it blank while the item is in use.' },
  { key: 'serial', h: 'Serial / Asset Tag', w: 135 },
  { key: 'bought', h: 'Purchase Date', w: 110, fmt: 'mmm d, yyyy' },
  { key: 'qty', h: 'Current Qty', w: 95, fmt: '0', note: 'Consumables only.' },
  { key: 'min', h: 'Min Qty', w: 80, fmt: '0', note: 'Consumables only. Below this number the item shows Low Stock.' },
  { key: 'nextMaint', h: 'Next Maintenance', w: 130, calc: true, fmt: 'mmm d, yyyy', note: 'Calculated: the latest Next Due for this item on the Maintenance Log.' },
  { key: 'notes', h: 'Notes', w: 230 },
  { key: 'photo', h: 'Photo URL', w: 150, note: 'Set by the app when an admin adds a photo (it is saved in the Inventory photos folder in Drive). You can also paste a link to a photo here.' },
  { key: 'link', h: 'Item Link', alias: 'Checkout Link', w: 160, calc: true, note: 'Filled in by the script. Opens this item in the hub\'s inventory app. The app prints QR labels from the same link.' }
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
    { col: 1, label: 'In service', key: 'inService' },
    { col: 2, label: 'In maintenance', key: 'maint' },
    { col: 3, label: 'Needs repair', key: 'repair' },
    { col: 4, label: 'Maintenance due (14 days)', key: 'maintDue' },
    { col: 7, label: 'Low or out of stock', key: 'stock' },
    { col: 8, label: 'All items (not retired)', key: 'total' }
  ],
  tileRow: 4,
  shade: [[1, 4], [7, 2]],   // tile groups: [first column, how many]
  blocks: [
    { key: 'maint', row: 7, col: 1, title: 'Maintenance due in the next 14 days', heads: ['Item ID', 'Item', 'Location', 'Next due', 'Days left'] },
    { key: 'repair', row: 7, col: 7, title: 'Needs repair', heads: ['Item ID', 'Item', 'Location', 'Status'] },
    { key: 'oos', row: 21, col: 1, title: 'Out of service for maintenance', heads: ['Item ID', 'Item', 'Location', 'Condition', 'Next due'] },
    { key: 'stock', row: 21, col: 7, title: 'Low or out of stock', heads: ['Item ID', 'Item', 'On hand', 'Min'] }
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

  var removed = retireBorrowing_(ss);
  var lists = ensureLists_(ss);
  var master = ensureHeaders_(ss, INV.tabs.master, MASTER);
  placeAfter_(master, 'team', 'loc');
  var maint = ensureHeaders_(ss, INV.tabs.maint, MAINT);
  var mc = masterCols_(master);
  var tc = colsBySpec_(maint, MAINT);

  ensureCalc_(ss, tc);
  finishMaster_(ss, master, mc, lists);
  finishLists_(lists, mc);
  finishMaint_(maint, tc, lists);
  buildDashboard_(ss, mc);
  orderTabs_(ss);
  installTriggers_(ss);
  writeLinks_(master, mc);
  writeSettings_(ss);

  var report = withLock_(selfTest_);
  var info = linkInfo_(ss);
  Logger.log(configBlock_(info));
  if (removed.length) Logger.log('Check outs taken out of this sheet: ' + removed.join('; '));
  Logger.log(report.summary + '\n' + report.lines.join('\n'));
  showDialog_(linksHtml_(info, report, removed), 'Inventory is set up');
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
    if (!have[s.h.toLowerCase()] && !(s.alias && have[s.alias.toLowerCase()])) {
      if (next > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
      sh.getRange(1, next).setValue(s.h);
      next++;
    }
  });
  return sh;
}

// =============================================================================
// Sheets from when the inventory had check outs
// =============================================================================

// The inventory used to log check outs and returns: a backup Google Form wrote to a
// Checkout Log tab, and Current Holder and Due Date came from it. Nobody borrows the
// tools, so setup takes all of that out of an older sheet. Items, maintenance and
// photos are untouched. Returns what it changed (nothing on a current sheet).
function retireBorrowing_(ss) {
  var done = [];
  var props = props_();
  var log = ss.getSheetByName(INV.old.log);

  // The backup form: unlink it from this sheet, then move it to the Drive trash.
  var form = null;
  var formId = props.getProperty('INV_FORM_ID');
  if (formId) { try { form = FormApp.openById(formId); } catch (e) { form = null; } }
  if (!form && log && log.getFormUrl()) { try { form = FormApp.openByUrl(log.getFormUrl()); } catch (e) { form = null; } }
  if (form) {
    try { form.removeDestination(); } catch (e) { console.warn('Could not unlink the backup form: ' + e); }
    try {
      DriveApp.getFileById(form.getId()).setTrashed(true);
      done.push('moved the backup check out form to the Drive trash');
    } catch (e) {
      done.push('unlinked the backup check out form (delete it from Drive yourself)');
    }
  }
  props.deleteProperty('INV_FORM_ID');

  // The Checkout Log goes when it holds nothing but test data. Real rows are kept on
  // a renamed tab, so no record is lost.
  if (log) {
    var real = 0;
    if (log.getLastRow() > 1) {
      log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues().forEach(function (r) {
        var text = r.map(function (v) { return String(v == null ? '' : v); }).join(' ').toLowerCase();
        if (text.trim() && text.indexOf(INV.testTag) < 0) real++;
      });
    }
    var linked = false;
    try { linked = !!log.getFormUrl(); } catch (e) { linked = false; }
    if (!real && !linked) {
      ss.deleteSheet(log);
      done.push('deleted the Checkout Log tab');
    } else if (!ss.getSheetByName(INV.old.keptLog)) {
      log.setName(INV.old.keptLog);
      done.push('kept the old check out records on a tab named "' + INV.old.keptLog + '" (delete it whenever you like)');
    }
  }

  // Inventory Master: drop the two columns the log fed, and rename the link column.
  var master = ss.getSheetByName(INV.tabs.master);
  if (master) {
    var have = headerMap_(master);
    var gone = [INV.old.holder, INV.old.due].filter(function (h) { return have[h.toLowerCase()]; });
    if (gone.length) {
      resetProtections_(master);
      gone.map(function (h) { return have[h.toLowerCase()]; })
        .sort(function (a, b) { return b - a; })
        .forEach(function (c) { master.deleteColumn(c); });
      done.push('removed the ' + gone.join(' and ') + ' columns');
    }
    have = headerMap_(master);
    var linkCol = have[INV.old.link.toLowerCase()];
    if (linkCol && !have['item link']) {
      master.getRange(1, linkCol).setValue('Item Link');
      done.push('renamed Checkout Link to Item Link');
    }
  }
  if (done.length) dropCache_();
  return done;
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

// alias: the header an older version used, found until setup renames it.
function colsBySpec_(sh, spec) {
  var have = headerMap_(sh), out = {};
  spec.forEach(function (s) { out[s.key] = have[s.h.toLowerCase()] || (s.alias ? have[s.alias.toLowerCase()] : undefined); });
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

// Hidden "Latest" tab: each item's Maintenance Log rows that have a Next Due, newest
// first. VLOOKUP into a list sorted newest-first returns the latest one, so Next
// Maintenance follows the most recent entry even when an older row has a later date.
function ensureCalc_(ss, tc) {
  var sh = ss.getSheetByName(INV.tabs.calc) || ss.insertSheet(INV.tabs.calc);
  var T = function (k) { return colRange_(INV.tabs.maint, tc[k]); };
  var idOf = function (r) { return 'IFERROR(REGEXEXTRACT(TRIM(' + r + '&""),"^\\S+"),"")'; };

  sh.clear();
  sh.getRange('A1').setFormula('={"Item ID","Next due","When";'
    + 'ARRAYFORMULA(IFERROR(SORT(FILTER({' + idOf(T('item')) + ',' + T('next') + ',' + T('ts') + '},'
    + T('item') + '<>"",ISNUMBER(' + T('next') + ')),3,FALSE),{"","",""}))}');
  sh.getRange('A1:C1').setFontWeight('bold');
  sh.getRange('B2:B').setNumberFormat('mmm d, yyyy');
  sh.getRange('C2:C').setNumberFormat('mmm d, yyyy h:mm am/pm');
  resetProtections_(sh);
  sh.protect().setDescription(INV.tag + 'helper tab').setWarningOnly(true);
  if (!sh.isSheetHidden()) sh.hideSheet();
  return sh;
}

function masterFormulas_(mc) {
  var R = function (k) { var L = letter_(mc[k]); return L + '2:' + L; };
  var latestM = sheetRef_(INV.tabs.calc) + '$A$2:$C';
  var hdr = {};
  MASTER.forEach(function (s) { hdr[s.key] = s.h; });
  return {
    status: '={' + lit_(hdr.status) + ';ARRAYFORMULA(IF(' + R('id') + '="","",'
      + 'IF(' + R('oos') + '<>"",' + R('oos') + ','
      + 'IF(' + R('cat') + '=' + lit_(INV.consumable) + ','
      + 'IF(' + R('qty') + '="","",IF(' + R('qty') + '<=0,"Out of Stock",'
      + 'IF((' + R('min') + '<>"")*(' + R('qty') + '<' + R('min') + '),"Low Stock","In Stock"))),'
      + lit_(INV.inService) + '))))}',
    nextMaint: '={' + lit_(hdr.nextMaint) + ';ARRAYFORMULA(IF(' + R('id') + '="","",'
      + 'IFERROR(VLOOKUP(' + R('id') + ',' + latestM + ',2,FALSE),"")))}'
  };
}

function finishMaster_(ss, sh, mc, lists) {
  var rows = sh.getMaxRows() - 1;
  var f = masterFormulas_(mc);
  ['status', 'nextMaint'].forEach(function (k) {
    // The column must be empty below the header or the formula cannot spill.
    sh.getRange(2, mc[k], rows, 1).clearContent();
    sh.getRange(1, mc[k]).setFormula(f[k]);
  });
  sh.getRange(1, mc.link).setValue(MASTER.filter(function (s) { return s.key === 'link'; })[0].h);

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
  var R = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  var st = '$' + L('status') + '2';
  var rules = [
    R().whenTextEqualTo(INV.inService).setBackground('#D9EAD3').setFontColor('#274E13').setRanges([col('status')]).build(),
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
  sh.getRange('E1').setNote('Options for the Team dropdown. Shared means the program\'s own tools that any team can use.');
  sh.getRange('F1').setNote('Calculated from the Inventory Master. Feeds the Item dropdown on the Maintenance Log.');
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
  var tileCell = {};
  DASH.tiles.forEach(function (t) { tileCell[t.key] = letter_(t.col) + (DASH.tileRow + 1); });
  var more = function (cell, words) { return '=IF(' + cell + '>' + n + ',"+"&(' + cell + '-' + n + ')&' + lit_(words) + ',"")'; };

  return {
    tiles: {
      inService: '=COUNTIF(' + M('status') + ',' + lit_(INV.inService) + ')',
      maint: '=COUNTIF(' + M('status') + ',"Maintenance")',
      repair: '=COUNTIFS(' + M('cond') + ',"Needs Repair",' + M('oos') + ',"<>Retired")',
      maintDue: '=COUNTIFS(' + M('nextMaint') + ',"<="&(TODAY()+14),' + M('oos') + ',"<>Retired")',
      stock: '=COUNTIF(' + M('status') + ',"Low Stock")+COUNTIF(' + M('status') + ',"Out of Stock")',
      total: '=COUNTIFS(' + M('id') + ',"<>",' + M('oos') + ',"<>Retired")'
    },
    lists: {
      maint: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('loc') + ',' + M('nextMaint') + ','
        + 'IFERROR(' + M('nextMaint') + '-TODAY(),"")},ISNUMBER(' + M('nextMaint') + '),' + M('nextMaint') + '<=TODAY()+14,'
        + M('oos') + '<>"Retired"),4,TRUE),' + n + ',5),"Nothing due in the next 14 days"))',
      repair: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(FILTER({' + M('id') + ',' + M('name') + ',' + M('loc') + ',' + M('status') + '},'
        + M('cond') + '="Needs Repair",' + M('oos') + '<>"Retired"),' + n + ',4),"Nothing flagged for repair"))',
      oos: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('loc') + ',' + M('cond') + ',' + M('nextMaint') + '},'
        + M('status') + '="Maintenance"),1,TRUE),' + n + ',5),"Nothing out of service"))',
      stock: '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + M('id') + ',' + M('name') + ',' + M('qty') + ',' + M('min') + '},'
        + '(' + M('status') + '="Low Stock")+(' + M('status') + '="Out of Stock")),3,TRUE),' + n + ',4),"All stocked"))'
    },
    more: {
      maint: more(tileCell.maintDue, ' more due'),
      repair: more(tileCell.repair, ' more flagged for repair'),
      oos: more(tileCell.maint, ' more out of service'),
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
  sh.getRange('A2').setValue('Live from the Inventory Master and the Maintenance Log. Nothing to type on this tab.')
    .setFontColor(INV.gray).setFontStyle('italic');

  var tr = DASH.tileRow;
  DASH.tiles.forEach(function (t) {
    sh.getRange(tr, t.col).setValue(t.label).setFontSize(9).setFontWeight('bold').setFontColor(INV.gray).setWrap(true).setVerticalAlignment('bottom');
    sh.getRange(tr + 1, t.col).setFormula(f.tiles[t.key]).setFontSize(22).setFontWeight('bold').setFontColor('#202124').setHorizontalAlignment('left');
  });
  DASH.shade.forEach(function (g) { sh.getRange(tr, g[0], 2, g[1]).setBackground('#F8F9FA'); });
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
  fmt('maint', 3, 'mmm d, yyyy'); fmt('maint', 4, '0');
  fmt('oos', 4, 'mmm d, yyyy');
  fmt('stock', 2, '0'); fmt('stock', 3, '0');

  var tile = function (key) { var t = DASH.tiles.filter(function (x) { return x.key === key; })[0]; return sh.getRange(tr + 1, t.col); };
  var block = function (key, offset) { var b = DASH.blocks.filter(function (x) { return x.key === key; })[0]; return sh.getRange(b.row + 2, b.col + offset, n, 1); };
  var R = function () { return SpreadsheetApp.newConditionalFormatRule(); };
  sh.setConditionalFormatRules([
    R().whenNumberGreaterThan(0).setFontColor(INV.red).setRanges([tile('stock')]).build(),
    R().whenNumberGreaterThan(0).setFontColor('#B45F06').setRanges([tile('maint'), tile('repair'), tile('maintDue')]).build(),
    R().whenNumberLessThan(0).setFontColor(INV.red).setBold(true).setRanges([block('maint', 4)]).build()
  ]);

  sh.protect().setDescription(INV.tag + 'the Dashboard is calculated').setWarningOnly(true);
}

// Dashboard, Inventory Master, Maintenance Log, Lists, then the hidden helper.
function orderTabs_(ss) {
  var order = [INV.tabs.dash, INV.tabs.master, INV.tabs.maint, INV.tabs.lists];
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

// onCheckoutSubmit is in the list so that a sheet from the check out days loses it.
function installTriggers_(ss) {
  var mine = { onInventoryEdit: 1, onInventoryChange: 1, onCheckoutSubmit: 1 };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (mine[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onInventoryEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onInventoryChange').forSpreadsheet(ss).onChange().create();
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
        var touched = ['id', 'name'].some(function (k) { return mc[k] >= c1 && mc[k] <= c2; });
        if (filled || touched) writeLinks_(sh, mc);
      });
    } else if (name === INV.tabs.maint) {
      stampTimes_(sh, e.range);
      dropCache_();
    } else {
      dropCache_();
    }
  } catch (err) {
    console.error('onInventoryEdit: ' + (err && err.stack || err));
  }
}

// Installable onChange: deleting rows does not fire onEdit.
function onInventoryChange(e) {
  dropCache_();
}

function itemIdOf_(text) { return String(text == null ? '' : text).trim().split(/\s+/)[0] || ''; }

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

// Hand-typed rows on the Maintenance Log get a real timestamp (=NOW() would keep
// changing). Rows the app adds come with their own.
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
    writeLinks_(sh, mc);
  });
}

// Every item with an ID and a name gets its Item Link: its page in the hub app,
// which is also what its QR label opens. Rows emptied by hand lose theirs.
function writeLinks_(sh, mc) {
  var last = lastDataRow_(sh, mc);
  if (last > 1) {
    var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    sh.getRange(2, mc.link, data.length, 1).setValues(data.map(function (r) {
      var id = String(r[mc.id - 1]).trim();
      return [id && String(r[mc.name - 1]).trim() ? appLink_(id) : ''];
    }));
  }
  if (last < sh.getMaxRows()) sh.getRange(last + 1, mc.link, sh.getMaxRows() - last, 1).clearContent();
}

function appLink_(id) { return INV.appUrl + '?item=' + encodeURIComponent(id); }

// =============================================================================
// Links, dialogs, config.js block
// =============================================================================

function linkInfo_(ss) {
  var gid = function (name) { var s = ss.getSheetByName(name); return s ? String(s.getSheetId()) : ''; };
  return {
    apiUrl: webAppUrl_(),
    appUrl: INV.appUrl,
    sheetUrl: ss.getUrl(),
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

// Settings on the Lists tab (H2:I5). Rows 4 and 5 held the backup form's links when
// the inventory had check outs; they are cleared.
function writeSettings_(ss) {
  var sh = ss.getSheetByName(INV.tabs.lists);
  var info = linkInfo_(ss);
  sh.getRange(2, 8, 4, 2).setValues([
    ['Inventory app (share this)', info.appUrl],
    ['Set up on', new Date()],
    ['', ''],
    ['', '']
  ]);
  sh.getRange(3, 9).setNumberFormat('mmm d, yyyy h:mm am/pm').setHorizontalAlignment('left');
  sh.getRange(5, 9).setNumberFormat('General');
}

function configBlock_(info) {
  return 'window.INVENTORY = {\n'
    + "  apiUrl: '" + info.apiUrl + "',\n"
    + "  appUrl: '" + info.appUrl + "',\n"
    + "  sheetUrl: '" + info.sheetUrl + "',\n"
    + "  dashboardGid: '" + info.gids.dash + "',\n"
    + "  masterGid: '" + info.gids.master + "',\n"
    + "  maintenanceGid: '" + info.gids.maint + "',\n"
    + '};';
}

function showInventoryLinks() {
  showDialog_(linksHtml_(linkInfo_(ss_()), null), 'Inventory links and setup');
}

function runSelfTest() {
  var report = withLock_(selfTest_);
  showDialog_(linksHtml_(null, report), 'Self-test');
}

function addTestData() {
  var out = withLock_(addTestData_);
  ss_().toast('Added ' + out.added.items + ' test items and ' + out.added.maintenance + ' maintenance entries.', 'Test data loaded', 8);
}

function removeTestData() {
  var out = withLock_(removeTestData_);
  ss_().toast('Removed ' + out.removed.items + ' test items and ' + out.removed.maintenance + ' maintenance rows.', 'Test data removed', 8);
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

// removed: what retireBorrowing_ took out of an older sheet, if anything.
function linksHtml_(info, report, removed) {
  var css = '<style>body{font:14px/1.5 Arial,sans-serif;color:#202124;margin:0 4px}h3{font-size:15px;margin:16px 0 6px}'
    + 'a{color:#1F6FB2;word-break:break-all}textarea{width:100%;height:132px;font:12px/1.45 Menlo,Consolas,monospace;'
    + 'border:1px solid #ccc;border-radius:6px;padding:8px;box-sizing:border-box}.ok{color:#1E7A3C}.bad{color:#B31B1B}'
    + 'ul{margin:4px 0 0;padding-left:18px}li{margin:2px 0}button{margin-top:6px;padding:7px 14px;border:0;border-radius:6px;'
    + 'background:#B31B1B;color:#fff;font-weight:bold;cursor:pointer}.muted{color:#5F6368;font-size:12.5px}</style>';
  var out = css;
  if (removed && removed.length) {
    out += '<h3>Check outs are gone from this sheet</h3><ul>'
      + removed.map(function (l) { return '<li>' + esc_(l.charAt(0).toUpperCase() + l.slice(1)) + '.</li>'; }).join('') + '</ul>'
      + '<p class="muted">Items, maintenance and photos are as they were.</p>';
  }
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
      + '<p class="muted">It replaces the window.INVENTORY block. The hub\'s inventory app and its labels read it.</p>';
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
  var maint = ss.getSheetByName(INV.tabs.maint);
  var dash = ss.getSheetByName(INV.tabs.dash);
  var lines = [], fails = 0;
  var check = function (label, pass, got) {
    if (!pass) fails++;
    lines.push((pass ? 'PASS  ' : 'FAIL  ') + label + (pass ? '' : ' (got: ' + got + ')'));
  };
  if (!master || !maint || !dash) {
    return { ok: false, summary: 'Self-test could not run: a tab is missing. Run setupInventory.', lines: [] };
  }

  var mc = masterCols_(master), tc = colsBySpec_(maint, MAINT);
  var id = 'ZZ-999', name = 'Self-test item (removed automatically)';
  var choice = id + INV.sep + name;
  var day = 86400000;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var inThree = new Date(today.getTime() + 3 * day);
  var sameDay = function (v, d) { return v instanceof Date && v.getFullYear() === d.getFullYear() && v.getMonth() === d.getMonth() && v.getDate() === d.getDate(); };
  var added = { master: 0, maint: 0 };
  var cell = function (sh, row, col) { return sh.getRange(row, col).getValue(); };
  var dashIds = function (key, offset) {
    var b = DASH.blocks.filter(function (x) { return x.key === key; })[0];
    return dash.getRange(b.row + 2, b.col + (offset || 0), INV.listRows, 1).getValues().map(function (r) { return String(r[0]); });
  };
  var set = function (key, value) { master.getRange(added.master, mc[key]).setValue(value); SpreadsheetApp.flush(); };

  try {
    var mr = lastDataRow_(master, mc) + 1;
    if (mr > master.getMaxRows()) master.insertRowsAfter(master.getMaxRows(), 1);
    master.getRange(mr, mc.id).setValue(id);
    master.getRange(mr, mc.name).setValue(name);
    master.getRange(mr, mc.cat).setValue('Equipment');
    added.master = mr;
    SpreadsheetApp.flush();
    check('A new item shows ' + INV.inService, cell(master, mr, mc.status) === INV.inService, cell(master, mr, mc.status));

    set('cond', 'Needs Repair');
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

    set('oos', 'Maintenance');
    check('Out of Service = Maintenance shows Maintenance', cell(master, mr, mc.status) === 'Maintenance', cell(master, mr, mc.status));
    check('Dashboard lists it under Out of service', dashIds('oos').indexOf(id) >= 0, dashIds('oos')[0]);

    set('oos', '');
    master.getRange(mr, mc.cat).setValue(INV.consumable);
    master.getRange(mr, mc.qty).setValue(1);
    master.getRange(mr, mc.min).setValue(5);
    SpreadsheetApp.flush();
    check('A consumable below Min Qty shows Low Stock', cell(master, mr, mc.status) === 'Low Stock', cell(master, mr, mc.status));
    check('Dashboard lists it under Low or out of stock', dashIds('stock').indexOf(id) >= 0, dashIds('stock')[0]);

    set('oos', 'Retired');
    check('Out of Service = Retired shows Retired', cell(master, mr, mc.status) === 'Retired', cell(master, mr, mc.status));

    var shown = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
    var broken = shown.filter(function (v) { return /^#/.test(v); });
    check('No formula errors in the Inventory Master headers', broken.length === 0, broken.join(' '));
  } catch (err) {
    fails++;
    lines.push('ERROR ' + (err && err.message || err));
  } finally {
    // Remove only the rows this test wrote, after checking they are ours.
    if (added.maint && String(maint.getRange(added.maint, tc.item).getValue()) === choice) maint.deleteRow(added.maint);
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
// which works from the static hub without any CORS setup. Reads are open to anyone
// with the page; changes need admin=1, which only the (unlisted) admin page sends,
// the same convention as the other modules. Photos are too big for a URL, so they
// come in as a POST (doPost).
// =============================================================================

var CACHE_KEY = 'inv_list_v3';
var API_VERSION = 3;   // 2 added photos and teams; 3 took out check outs

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
    status: str('status'), cond: str('cond'), oos: str('oos'),
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
  return { ok: true, item: ctx.item, maintenance: maintHistory_(ctx.ss, ctx.item.id, tz) };
}

// Newest first: exact time, then sheet order (later rows are newer) for ties.
function newestFirst_(list) {
  list.sort(function (a, b) { return (b._t - a._t) || (b._i - a._i); });
  return list.map(function (x) { delete x._t; delete x._i; return x; });
}
function ms_(v) { return v instanceof Date ? v.getTime() : 0; }

function stamp_(v, tz) { return v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm") : String(v == null ? '' : v).trim(); }
function dayStr_(v, tz) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v == null ? '' : v).trim(); }

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
  writeLinks_(sh, mc);
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
function optDate_(v, label) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  var d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  if (!d || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) throw new Error(label + ' must be a date.');
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
// Test data: realistic sample items and maintenance, all tagged "[test data]" so
// removeTestData_ can take back exactly what it added.
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

  appendMaintRows_(ss, TEST_MAINT.map(function (m) {
    return { when: at(m[1], 11), item: label[m[0]], issue: m[2] + ' ' + INV.testTag, taken: m[3], by: m[4], next: dayFrom(m[5]), cost: m[6] == null ? '' : m[6] };
  }));
  SpreadsheetApp.flush();
  writeLinks_(sh, mc);
  return { ok: true, added: { items: ids.length, maintenance: TEST_MAINT.length } };
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
  var tc = colsBySpec_(ss.getSheetByName(INV.tabs.maint), MAINT);
  var maintHits = pick(INV.tabs.maint, tc.item, tc.issue);
  deleteRows_(maintHits.sheet, maintHits.rows);
  deleteRows_(sh, masterRows);
  SpreadsheetApp.flush();
  photos.forEach(trashIfOurs_);
  return { ok: true, removed: { items: masterRows.length, maintenance: maintHits.rows.length, photos: photos.length } };
}

// Remove, then add again: a fresh copy of the sample data (its dates are relative
// to today, so maintenance that is overdue or due soon stays that way).
function resetTestData_() {
  var removed = removeTestData_();
  var added = addTestData_();
  return { ok: true, removed: removed.removed, added: added.added };
}
