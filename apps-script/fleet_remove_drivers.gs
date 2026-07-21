/**
 * Bulk-remove graduated drivers from the vehicle reservation spreadsheet.
 *
 * Two things this can do (use whichever matches how the drivers are stored):
 *   A. Revoke their EDITING ACCESS to the sheet (they were shared as editors).
 *   B. Delete their ROWS from a roster tab (if the drivers are listed as rows).
 *
 * Setup:
 *   1. Open the reservation Google Sheet, Extensions > Apps Script, paste this, Save.
 *   2. Make a tab that holds the people to remove (default name "Remove list").
 *      Paste the list one per row: Name in column A, NetID in column B, Year in C.
 *   3. Run previewRemoval() first and read the log (View > Logs). It changes nothing.
 *   4. If it looks right, run removeDriverAccess() (option A) or removeDriverRows() (option B).
 *
 * IMPORTANT: option A must be run by the sheet OWNER (Noah), since it changes who the
 * sheet is shared with. Nothing is deleted or unshared until you run one of the do-it
 * functions; preview is always safe.
 */

// ---- settings ----
var SHEET_ID    = '';            // leave blank if pasted inside the sheet; else the sheet id
var LIST_TAB    = 'Remove list'; // tab holding the people to remove
var NETID_COL   = 2;             // column with the NetID (A=1, B=2, C=3 ...)
var HAS_HEADER  = true;          // true if row 1 of the list tab is a header
var DOMAIN      = 'cornell.edu'; // NetID becomes NetID@cornell.edu

// Option B only: where the driver roster lives, if you delete rows instead of access.
var ROSTER_TAB       = 'Drivers';
var ROSTER_NETID_COL = 2;

function sheet_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// Unique NetIDs from the list tab, lowercased. NetIDs are case-insensitive.
function removalNetIds_() {
  var sh = sheet_().getSheetByName(LIST_TAB);
  if (!sh) throw new Error('No tab named "' + LIST_TAB + '". Create it and paste the list, or change LIST_TAB.');
  var rows = sh.getDataRange().getValues();
  var seen = {}, out = [];
  for (var i = (HAS_HEADER ? 1 : 0); i < rows.length; i++) {
    var netid = String(rows[i][NETID_COL - 1] || '').trim().toLowerCase();
    if (!netid || seen[netid]) continue;
    seen[netid] = true;
    out.push(netid);
  }
  return out;
}

// ---- Preview (safe, changes nothing) ----
function previewRemoval() {
  var ss = sheet_();
  var netids = removalNetIds_();
  var access = {};
  ss.getEditors().forEach(function (u) { access[u.getEmail().toLowerCase()] = 'editor'; });
  ss.getViewers().forEach(function (u) { var e = u.getEmail().toLowerCase(); if (!access[e]) access[e] = 'viewer'; });
  var owner = ss.getOwner() ? ss.getOwner().getEmail().toLowerCase() : '';

  var willRemove = [], notShared = [];
  netids.forEach(function (n) {
    var email = n + '@' + DOMAIN;
    if (email === owner) return;
    if (access[email]) willRemove.push(email + ' (' + access[email] + ')');
    else notShared.push(email);
  });

  Logger.log('List "' + LIST_TAB + '": ' + netids.length + ' unique NetIDs.');
  Logger.log('WOULD LOSE ACCESS (' + willRemove.length + '): ' + willRemove.join(', '));
  Logger.log('Not currently shared, nothing to do (' + notShared.length + '): ' + notShared.join(', '));
  Logger.log('Nothing was changed. Run removeDriverAccess() to apply, or removeDriverRows() to delete roster rows.');
}

// ---- Option A: revoke editing access ----
function removeDriverAccess() {
  var ss = sheet_();
  var netids = removalNetIds_();
  var isEditor = {}, isViewer = {};
  ss.getEditors().forEach(function (u) { isEditor[u.getEmail().toLowerCase()] = true; });
  ss.getViewers().forEach(function (u) { isViewer[u.getEmail().toLowerCase()] = true; });
  var owner = ss.getOwner() ? ss.getOwner().getEmail().toLowerCase() : '';

  var removed = [], notShared = [], failed = [];
  netids.forEach(function (n) {
    var email = n + '@' + DOMAIN;
    if (email === owner) return;
    if (!isEditor[email] && !isViewer[email]) { notShared.push(email); return; }
    try {
      if (isEditor[email]) ss.removeEditor(email);
      if (isViewer[email]) ss.removeViewer(email);
      removed.push(email);
    } catch (err) {
      failed.push(email + ' (' + err + ')');
    }
  });

  Logger.log('REMOVED ACCESS (' + removed.length + '): ' + removed.join(', '));
  Logger.log('Was not shared (' + notShared.length + '): ' + notShared.join(', '));
  if (failed.length) Logger.log('FAILED (' + failed.length + '): ' + failed.join(', '));
}

// ---- Option B: delete roster rows ----
function removeDriverRows() {
  var ss = sheet_();
  var remove = {};
  removalNetIds_().forEach(function (n) { remove[n] = true; });
  var sh = ss.getSheetByName(ROSTER_TAB);
  if (!sh) throw new Error('No tab named "' + ROSTER_TAB + '". Set ROSTER_TAB to your driver roster tab.');
  var rows = sh.getDataRange().getValues();
  var deleted = 0;
  // delete bottom-up so row numbers stay valid as we go
  for (var i = rows.length - 1; i >= (HAS_HEADER ? 1 : 0); i--) {
    var netid = String(rows[i][ROSTER_NETID_COL - 1] || '').trim().toLowerCase();
    if (netid && remove[netid]) { sh.deleteRow(i + 1); deleted++; }
  }
  Logger.log('Deleted ' + deleted + ' matching driver rows from "' + ROSTER_TAB + '".');
}
