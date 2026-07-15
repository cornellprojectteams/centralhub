/**
 * Ops CMMS, Phase 1 setup: creates the Equipment and Inventory tabs in the
 * Space Status Tracking spreadsheet. Run setupRegistry() once from the editor.
 *
 * Safe to re-run: it only creates a tab if missing and only writes headers when
 * the tab is empty, so it never overwrites data you have entered.
 */
function setupRegistry() {
  var ss = SpreadsheetApp.openById('1mZrlnA-GiVKB4_Um21aMH9jsSkf6TUJhauDWivnf7-I');

  ensureRegistryTab_(ss, 'Equipment',
    ['Asset ID', 'Name', 'Category', 'Location', 'Owning team', 'Status', 'Installed', 'Notes'],
    ['EX-001', 'Example: spray booth (delete this row)', 'Finishing', 'Composites lab', 'Operations Team', 'In service', '2024', 'Filters checked monthly']);

  ensureRegistryTab_(ss, 'Inventory',
    ['Item', 'Location', 'On hand', 'Reorder point', 'Unit', 'Supplier', 'Product link', 'eShop info', 'Last restocked'],
    ['Example: spray booth filter (delete this row)', 'Composites lab', 0, 2, 'each', 'Uline', '', 'eShop item #', '2026-05']);

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
