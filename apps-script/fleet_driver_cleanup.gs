/**
 * Bulk-remove graduated student drivers from the vehicle reservation sheet.
 *
 * Everything happens in the spreadsheet. No code or logs to read.
 *
 * Paste this once (Extensions > Apps Script), Save, reload the sheet. A "Driver
 * cleanup" menu appears. Then:
 *
 *   1. List graduated drivers   -> writes everyone graduating this year or earlier
 *                                  into a "Drivers to remove" tab (Year, NetID, Name).
 *   2. Review that tab           -> delete the row of anyone you want to KEEP.
 *   3. Remove them               -> acts on whatever is left in that tab:
 *        Delete listed drivers (rows)    removes their row from the drivers list
 *        Remove listed drivers (access)  removes their edit access to the sheet
 *
 * "Graduated" means graduation year <= this year. Change CFG.gradYearThrough to
 * adjust. The drivers tab and its NetID / year columns are found automatically.
 */

var FLEET_SS_ID = '1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M';
var PREVIEW_TAB = 'Drivers to remove';

var CFG = {
  driversTab: '',                                // tab with one row per driver. '' = auto-detect.
  gradYearThrough: new Date().getFullYear(),     // remove this graduation year and earlier.
};

var NETID_RE = /^[A-Za-z]{2,4}[0-9]{1,4}$/;       // apk67, rd496, sk2682, Lc2226

function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
function report_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); } }

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Driver cleanup')
    .addItem('1. List graduated drivers', 'previewGraduated')
    .addSeparator()
    .addItem('2. Delete listed drivers (rows)', 'removeDriverRows')
    .addItem('2. Remove listed drivers (access)', 'removeDriverAccess')
    .addSeparator()
    .addItem('Show sheet info', 'showSheetInfo')
    .addToUi();
}

// ---- finding the drivers tab and its columns ----

function columnByHeader_(sh, needles) {
  var lc = sh.getLastColumn();
  if (!lc) return -1;
  var headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(norm_);
  for (var i = 0; i < headers.length; i++) {
    for (var k = 0; k < needles.length; k++) { if (headers[i].indexOf(needles[k]) >= 0) return i; }
  }
  return -1;
}

function netIdColumn_(sh) {
  var byHeader = columnByHeader_(sh, ['netid', 'net id']);
  if (byHeader >= 0) return byHeader;
  var data = sh.getDataRange().getValues(), best = -1, bestCount = 0, lc = sh.getLastColumn();
  for (var c = 0; c < lc; c++) {
    var cnt = 0;
    for (var r = 1; r < data.length; r++) { if (NETID_RE.test(String(data[r][c]).trim())) cnt++; }
    if (cnt > bestCount) { bestCount = cnt; best = c; }
  }
  return bestCount >= 3 ? best : -1;
}

function parseYear_(v) {
  if (v instanceof Date) return v.getFullYear();
  var m = String(v).match(/\b(20\d{2})\b/);
  return m ? Number(m[1]) : null;
}

function yearColumn_(sh) {
  var byHeader = columnByHeader_(sh, ['grad', 'class', 'year']);
  if (byHeader >= 0) return byHeader;
  var data = sh.getDataRange().getValues(), best = -1, bestCount = 0, lc = sh.getLastColumn();
  for (var c = 0; c < lc; c++) {
    var cnt = 0;
    for (var r = 1; r < data.length; r++) { if (parseYear_(data[r][c])) cnt++; }
    if (cnt > bestCount) { bestCount = cnt; best = c; }
  }
  return bestCount >= 3 ? best : -1;
}

function driversSheet_(ss) {
  if (CFG.driversTab) return ss.getSheetByName(CFG.driversTab);
  var best = null, bestCount = 0;
  ss.getSheets().forEach(function (sh) {
    if (sh.getName() === PREVIEW_TAB) return;
    var col = netIdColumn_(sh);
    if (col < 0) return;
    var data = sh.getDataRange().getValues(), cnt = 0;
    for (var r = 1; r < data.length; r++) { if (NETID_RE.test(String(data[r][col]).trim())) cnt++; }
    if (cnt > bestCount) { bestCount = cnt; best = sh; }
  });
  return best;
}

// ---- step 1: write the candidate list into a tab ----

function previewGraduated() {
  var ss = SpreadsheetApp.openById(FLEET_SS_ID);
  var sh = driversSheet_(ss);
  if (!sh) { report_('Could not find the drivers list. Run "Show sheet info" and send it over.'); return; }
  var cNet = netIdColumn_(sh), cYear = yearColumn_(sh), cName = columnByHeader_(sh, ['name']);
  if (cNet < 0) { report_('Could not find a NetID column in "' + sh.getName() + '".'); return; }
  if (cYear < 0) { report_('Could not find a graduation-year column in "' + sh.getName() + '". Send "Show sheet info".'); return; }

  var data = sh.getDataRange().getValues(), out = [['Graduation year', 'NetID', 'Name']], years = {};
  for (var r = 1; r < data.length; r++) {
    var year = parseYear_(data[r][cYear]), netid = String(data[r][cNet]).trim();
    if (year != null && year <= CFG.gradYearThrough && NETID_RE.test(netid)) {
      out.push([year, netid, cName >= 0 ? String(data[r][cName]).trim() : '']);
      years[year] = 1;
    }
  }

  var pv = ss.getSheetByName(PREVIEW_TAB) || ss.insertSheet(PREVIEW_TAB);
  pv.clear();
  pv.getRange(1, 1, out.length, 3).setValues(out);
  pv.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#8f1515').setFontColor('#ffffff');
  pv.setFrozenRows(1);
  pv.autoResizeColumns(1, 3);
  ss.setActiveSheet(pv);

  report_('Listed ' + (out.length - 1) + ' graduated drivers (years ' + Object.keys(years).sort().join(', ') + ') '
    + 'on the "' + PREVIEW_TAB + '" tab.\n\nReview it and delete the row of anyone you want to keep. '
    + 'Then run "Delete listed drivers (rows)" or "Remove listed drivers (access)".');
}

// NetIDs left on the review tab = the final list to act on.
function reviewedNetIds_(ss) {
  var sh = ss.getSheetByName(PREVIEW_TAB);
  if (!sh) throw new Error('Run "1. List graduated drivers" first.');
  var seen = {}, list = [];
  sh.getDataRange().getValues().forEach(function (row) {
    row.forEach(function (cell) {
      var s = String(cell).trim();
      if (NETID_RE.test(s) && !seen[s.toLowerCase()]) { seen[s.toLowerCase()] = 1; list.push(s.toLowerCase()); }
    });
  });
  return list;
}

// ---- step 3a: delete their rows ----

function removeDriverRows() {
  var ss = SpreadsheetApp.openById(FLEET_SS_ID);
  var want = {}; reviewedNetIds_(ss).forEach(function (n) { want[n] = 1; });
  if (!Object.keys(want).length) { report_('The "' + PREVIEW_TAB + '" tab is empty. Run step 1 first.'); return; }
  var sh = driversSheet_(ss);
  var col = netIdColumn_(sh);
  var data = sh.getDataRange().getValues(), rows = [];
  for (var r = 1; r < data.length; r++) { if (want[norm_(data[r][col])]) rows.push(r + 1); }
  rows.sort(function (a, b) { return b - a; }).forEach(function (rn) { sh.deleteRow(rn); });
  report_('Deleted ' + rows.length + ' rows from "' + sh.getName() + '", using the "' + PREVIEW_TAB + '" list.');
}

// ---- step 3b: remove their edit access ----

function removeDriverAccess() {
  var ss = SpreadsheetApp.openById(FLEET_SS_ID);
  var want = {}; reviewedNetIds_(ss).forEach(function (n) { want[n] = 1; });
  var total = Object.keys(want).length;
  if (!total) { report_('The "' + PREVIEW_TAB + '" tab is empty. Run step 1 first.'); return; }
  var removed = 0;
  ss.getEditors().map(function (u) { return u.getEmail(); }).forEach(function (email) {
    if (!want[norm_(email.split('@')[0])]) return;
    try { ss.removeEditor(email); removed++; } catch (e) { Logger.log('Could not remove ' + email + ': ' + e.message); }
  });
  report_('Removed edit access for ' + removed + ' people, using the "' + PREVIEW_TAB + '" list.'
    + (removed < total ? '\n' + (total - removed) + ' were not individual editors of this sheet.' : ''));
}

// ---- setup diagnostic ----

function showSheetInfo() {
  var ss = SpreadsheetApp.openById(FLEET_SS_ID);
  var lines = ss.getSheets().map(function (sh) {
    var lc = sh.getLastColumn();
    var headers = lc ? sh.getRange(1, 1, 1, lc).getValues()[0] : [];
    return '- ' + sh.getName() + ' (' + sh.getLastRow() + ' rows): ' + headers.join(', ');
  });
  report_('Tabs and columns:\n' + lines.join('\n') + '\n\nEditors: ' + ss.getEditors().length);
}
