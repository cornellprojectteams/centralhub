/**
 * Ops CMMS, Phase 1 setup. Creates the Equipment and Inventory tabs in the
 * INVENTORY spreadsheet (not the space-issues sheet). Set REGISTRY_SS_ID to your
 * inventory spreadsheet id. Run setupRegistry() once from the editor.
 *
 * Safe to re-run: it only creates a tab if missing and only writes headers when
 * the tab is empty, so it never overwrites data you have entered.
 *
 * If your inventory sheet already has data, run listRegistrySheet() first and send
 * the log so the data can be mapped in rather than re-entered.
 */
var REGISTRY_SS_ID = '1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU';

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

  ensureRegistryTab_(ss, 'Equipment',
    ['Asset ID', 'Name', 'Category', 'Location', 'Owning team', 'Status', 'Installed', 'Notes'],
    ['EX-001', 'Example: spray booth (delete this row)', 'Finishing', 'Composites lab', 'Operations Team', 'In service', '2024', 'Filters checked monthly']);

  ensureRegistryTab_(ss, 'Inventory',
    ['Item', 'Location', 'On hand', 'Unit', 'Supplier', 'Product link', 'eShop info', 'Last restocked'],
    ['Example: spray booth filter (delete this row)', 'Composites lab', 0, 'each', 'Uline', '', 'eShop item #', '2026-05']);

  Logger.log('Registry ready. Fill in the Equipment and Inventory tabs, then delete the example rows.');
}

function ensureRegistryTab_(ss, name, headers, exampleRow) {
  var sh = ss.getSheetByName(name);
  var created = false;
  if (!sh) { sh = ss.insertSheet(name); created = true; }

  // Write headers (+ one example row) only when the tab is empty.
  if (!String(sh.getRange(1, 1).getValue()).trim()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (exampleRow) sh.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
  }

  var hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setFontWeight('bold').setBackground('#8f1515').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  Logger.log((created ? 'Created' : 'Found') + ' "' + name + '" tab.');
}
