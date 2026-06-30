/**
 * Space Status — Step 1: build the Team Contacts tab.
 *
 * Run setupTeamContacts() ONCE, from inside the Space Status Tracking
 * spreadsheet (Extensions -> Apps Script). It creates a "Team Contacts" tab,
 * seeds one row per team, and generates a private token for each team's portal
 * link. You then fill in the email columns by hand.
 *
 * Safe to re-run: it only adds teams/headers that are missing and never
 * overwrites contact info you've already entered.
 */

// These team values MUST match the form's "Responsible Team" dropdown exactly,
// because the notifier looks teams up by this string. Edit this list to match
// the live form if the names differ.
const TEAMS = [
  'AutoBoat', 'Baja', 'CUAir', 'Concrete Canoe', 'CU Sail', 'ChemE Car', 'CEV',
  'Combat Robotics', 'CUAUV', 'CUBMD', 'DEBUT', 'DBF', 'ESW', 'EWB', 'EWH',
  'Formula', 'Geodata', 'Hyperloop', 'iGEM', 'Mars Rover', 'Nexus', 'Rocketry',
  'Seismic Design', 'Steel Bridge', 'Operations Team', 'Unknown/unsure'
];

const CONTACTS_SHEET = 'Team Contacts';
const CONTACTS_HEADERS = [
  'Team', 'Liaison name', 'Liaison email', 'Team general email',
  'Safety lead name', 'Safety lead email', 'Extra CC', 'Team token'
];

function setupTeamContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONTACTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONTACTS_SHEET);

  // Header row (only write if not already there)
  if (sheet.getRange(1, 1).getValue() !== CONTACTS_HEADERS[0]) {
    sheet.getRange(1, 1, 1, CONTACTS_HEADERS.length).setValues([CONTACTS_HEADERS]);
    sheet.getRange(1, 1, 1, CONTACTS_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#b31b1b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, CONTACTS_HEADERS.length, 160);
    sheet.setColumnWidth(CONTACTS_HEADERS.length, 230); // token column wider
  }

  // Teams that already have a row
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String)
    : [];
  const have = {};
  existing.forEach(function (t) { have[t] = true; });

  // Append any missing teams with a fresh token
  const toAdd = TEAMS.filter(function (t) { return !have[t]; }).map(function (t) {
    const row = new Array(CONTACTS_HEADERS.length).fill('');
    row[0] = t;                                   // Team
    row[CONTACTS_HEADERS.length - 1] = makeToken(); // Team token
    return row;
  });

  if (toAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, CONTACTS_HEADERS.length)
      .setValues(toAdd);
  }

  ss.toast(
    toAdd.length + ' team(s) added, ' + existing.length + ' already present. ' +
    'Now fill in the email columns.',
    'Team Contacts ready', 8
  );
}

/** Short, URL-safe, hard-to-guess token for each team's portal link. */
function makeToken() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}
