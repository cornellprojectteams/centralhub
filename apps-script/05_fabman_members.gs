// The Fabman API key (the "SPT Key"). A FABMAN_API_KEY script property (Project Settings >
// Script properties), if set, takes its place, so a new key needs no code change.
const FABMAN_API_KEY = scriptProperty_('FABMAN_API_KEY') || '6387c077-6950-4ffd-aa96-23db92606dc3';
const FABMAN_MEMBERS_ENDPOINT = 'https://fabman.io/api/v1/members';
const FABMAN_ACCOUNT_ID = 2366;

const SPACES = 2946;
const LAB_EQUIPMENT = 2947;
const MACHINES = 2948;
const MILLS = 2949;
const SUPERVISOR = 2950;
const FLEET = 2954;

const SUPERVISOR_BRIDGE = 4770;
const VAN = 5684;

const STANDARD_PACKAGE = 11773;
const ACTIVE_VEHICLE = 11786;


//------------------------------------------------------------
//
// ACTION FUNCTIONS to read commands from the Google Sheet.
//
//------------------------------------------------------------
function fetchFabmanMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  sheet.clear();
  const headers =
    ['ID', 'First Name', 'Last Name', 'Team', 'Graduation Year', 'NetID', 'Email', 'State', 'Apron',
      'Spaces', 'Lab Equipment', 'Machines', 'Mills', 'Supervisor'];
  sheet.appendRow(headers);

  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  members.forEach(member => {
    const trainings = getMemberTrainings(member.id);
    const spaces = trainings.some(training => training.trainingCourse === SPACES);
    const lab_equipment = trainings.some(training => training.trainingCourse === LAB_EQUIPMENT);
    const machines = trainings.some(training => training.trainingCourse === MACHINES);
    const mills = trainings.some(training => training.trainingCourse === MILLS);
    const supervisor = trainings.some(training => training.trainingCourse === SUPERVISOR);

    sheet.appendRow([
      member.id || '',
      member.firstName || '',
      member.lastName || '',
      member.company || '',
      member.metadata.graduationYear || '',
      member.metadata.netid || '',
      member.emailAddress || '',
      member.state || '',
      member.metadata.apron || '',
      spaces, lab_equipment, machines, mills, supervisor
    ]);
  });
}

function updateFabmanMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  rows.forEach(row => {
    const rowObj = {};
    headers.forEach((header, i) => {
      rowObj[header] = row[i];
    });

    const memberId = rowObj['ID'];
    if (!memberId) return;
    const member = fetchMember(memberId);

    const trainings = getMemberTrainings(member.id);
    const spaces = trainings.some(training => training.trainingCourse === SPACES);
    const lab_equipment = trainings.some(training => training.trainingCourse === LAB_EQUIPMENT);
    const machines = trainings.some(training => training.trainingCourse === MACHINES);
    const mills = trainings.some(training => training.trainingCourse === MILLS);
    const supervisor = trainings.some(training => training.trainingCourse === SUPERVISOR);

    const updates = {};
    if (rowObj['First Name'] !== member.firstName) updates.firstName = rowObj['First Name'];
    if (rowObj['Last Name'] !== member.lastName) updates.lastName = rowObj['Last Name'];
    if (rowObj['Team'] !== member.company) updates.company = rowObj['Team'];
    if (rowObj['Email'] !== member.emailAddress) updates.emailAddress = rowObj['Email'];
    if (rowObj['State'] !== member.state) updates.state = rowObj['State'];
    if (rowObj['Graduation Year'] !== member.metadata.graduationYear || rowObj['Apron'] !== member.metadata.apron) {
      updates.metadata = {
        'graduationYear': rowObj['Graduation Year'],
        'apron': rowObj['Apron']
      };
    }

    if (Object.keys(updates).length > 0) {
      updates.lockVersion = member.lockVersion;
      putMember(member, updates);
    }

    if (rowObj['Spaces'] !== spaces) {
      if (rowObj['Spaces'] == true) postTraining(member, SPACES);
      else deleteTraining(member, SPACES);
    }

    if (rowObj['Lab Equipment'] !== lab_equipment) {
      if (rowObj['Lab Equipment'] == true) postTraining(member, LAB_EQUIPMENT);
      else deleteTraining(member, LAB_EQUIPMENT);
    }

    if (rowObj['Machines'] !== machines) {
      if (rowObj['Machines'] == true) postTraining(member, MACHINES);
      else deleteTraining(member, MACHINES);
    }

    if (rowObj['Mills'] !== mills) {
      if (rowObj['Mills'] == true) postTraining(member, MILLS);
      else deleteTraining(member, MILLS);
    }

    if (rowObj['Supervisor'] !== supervisor) {
      if (rowObj['Supervisor'] == true) postTraining(member, SUPERVISOR);
      else deleteTraining(member, SUPERVISOR);
    }
  });
}

function postMills() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);
    postTraining(member, MILLS);
  });
}

function postMachines() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);
    postTraining(member, MACHINES);
  });
}

function postRed() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);
    postTraining(member, MACHINES);
    postTraining(member, MILLS);

    const payload = {
      metadata: {
        "graduationYear": member.metadata.graduationYear,
        "apron": "Red Apron"
      },
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function postGreen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);

    const payload = {
      metadata: {
        "graduationYear": member.metadata.graduationYear,
        "apron": "Green Apron"
      },
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function postBlue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);

    const payload = {
      metadata: {
        "graduationYear": member.metadata.graduationYear,
        "apron": "Blue Apron"
      },
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function postSupervisor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);
    postTraining(member, SUPERVISOR);
  });
}

function deleteSupervisor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);
    deleteTraining(member, SUPERVISOR);
  });
}

function lockMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);

    const payload = {
      state: 'locked',
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function activateMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);

    const payload = {
      state: 'active',
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function deleteMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  emails.forEach(email => {
    if (!email) return;
    const member = findMemberByEmail(email, members);

    const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}`;
    const options = {
      method: 'delete',
      headers: {
        'Authorization': `Bearer ${FABMAN_API_KEY}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, options);
  });
}

// Run once when setting up a fresh account: creates an empty metadata.graduationYear
// field for each member, since putMember throws if that field does not already exist.
function addMetadata() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();

  ids.forEach(id => {
    if (!id) return;
    const member = fetchMember(id);

    var atIndex = member.emailAddress.indexOf("@");
    var netid = member.emailAddress.substring(0, atIndex);

    const payload = {
      metadata: {
        "graduationYear": member.metadata.graduationYear,
        "apron": member.metadata.apron,
        "netid": netid
      },
      lockVersion: member.lockVersion
    };
    putMember(member, payload);
  });
}

function invitation() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Fabman invitations?', 'Every address in the Actions tab gets an invitation email from Fabman.', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) { return; }

  Logger.log("Sending invitation emails.")
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    },
  };

  emails.forEach(email => {
    if (!email) return;
    var member = findMemberByEmail(email, members);
    var url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/invitation`;

    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    Logger.log('Status: ' + statusCode);
  });
}

function disclaimer() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Actions');
  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();

  emails.forEach(email => {
    if (!email) return;
    var mailto = email;
    var subject = "Fabman Account Reset";
    var message = `In preparation for the beginning of the semester, all existing Fabman accounts have been wiped. Going forward, new users will first need to visit the Safety & Training Hub Canvas page and complete the Machine Room Usage Agreements: https://canvas.cornell.edu/enroll/AMW39G`;
    MailApp.sendEmail(mailto, subject, message);
  });
}

//------------------------------------------------------------
//
// GRADUATES: flag, review, then process (lock or delete).
//
//------------------------------------------------------------

// Builds a reviewable "Graduates" sheet from grad year in metadata. Everyone is
// pre-approved and set to Lock, so the common case is a glance and one Process click.
// The checkbox in row 1 approves or clears everyone at once. Changes nothing in Fabman.
function flagGraduates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Fetching members from Fabman...', 'Graduates', -1);
  const cutoff = new Date().getFullYear();
  const sheet = ss.getSheetByName('Graduates') || ss.insertSheet('Graduates');
  sheet.clear();

  sheet.getRange('A1').insertCheckboxes().setValue(true);
  sheet.getRange('B1').setValue('Select all: tick to approve everyone, untick to clear. Uncheck any to skip, then Process graduates.');

  const headers = ['Approve', 'Action', 'ID', 'Name', 'Team', 'Grad Year', 'NetID', 'Email', 'State', 'Access', 'Result'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(2);

  const trainingName = { [SPACES]: 'Spaces', [LAB_EQUIPMENT]: 'Lab', [MACHINES]: 'Machines', [MILLS]: 'Mills', [SUPERVISOR]: 'Supervisor' };
  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);
  const grads = members.filter(member => {
    const gy = parseInt(member.metadata && member.metadata.graduationYear, 10);
    return gy && gy <= cutoff;
  });

  const rows = [];
  grads.forEach((member, i) => {
    if (i % 10 === 0) ss.toast(`Loading access, ${i + 1} of ${grads.length}...`, 'Graduates', -1);
    const access = getMemberTrainings(member.id).map(t => trainingName[t.trainingCourse]).filter(Boolean).join(', ');
    rows.push([true, 'Lock', member.id,
      `${member.firstName || ''} ${member.lastName || ''}`.trim(),
      member.company || '', member.metadata.graduationYear || '',
      (member.metadata && member.metadata.netid) || '', member.emailAddress || '', member.state || '', access, '']);
  });

  if (rows.length) {
    sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(3, 1, rows.length, 1).insertCheckboxes();
    sheet.getRange(3, 2, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Lock', 'Delete', 'Skip'], true).build());
    sheet.getRange(3, 1, rows.length, headers.length).setBackground('#fff6df');
  }

  ss.toast(`Ready: ${rows.length} graduate(s) listed.`, 'Graduates', 4);
  SpreadsheetApp.getUi().alert(
    `${rows.length} graduate(s) found (grad year <= ${cutoff}). All are pre-approved and set to Lock.\n\n` +
    `Uncheck anyone to skip (or use the Select all box in row 1), switch any Action to Delete if needed, ` +
    `then click Process graduates.`);
}

// Flips every Approve checkbox to match the "select all" box in row 1 of the
// Graduates sheet. Simple trigger: sheet-only edit, no authorization needed.
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Graduates') return;
  if (range.getRow() !== 1 || range.getColumn() !== 1) return;
  const last = sheet.getLastRow();
  if (last < 3) return;
  sheet.getRange(3, 1, last - 2, 1).setValue(range.getValue() === true);
}

// Applies the reviewed actions. Only Approve-checked rows are touched, after a confirm.
function processGraduates() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Graduates');
  if (!sheet) { ui.alert('Run "Flag graduates for review" first.'); return; }

  const data = sheet.getDataRange().getValues();
  const C = { approve: 0, action: 1, id: 2, result: 10 };
  const todo = [];
  for (let i = 2; i < data.length; i++) {   // row 1 is the select-all control, row 2 is headers
    if (data[i][C.approve] === true && data[i][C.id] && data[i][C.action] !== 'Skip') {
      todo.push({ rowNum: i + 1, id: data[i][C.id], action: data[i][C.action] || 'Lock' });
    }
  }
  if (!todo.length) { ui.alert('Nothing approved. Tick "Approve" and choose Lock or Delete.'); return; }

  const counts = todo.reduce((acc, t) => (acc[t.action] = (acc[t.action] || 0) + 1, acc), {});
  const summary = Object.keys(counts).map(k => `${counts[k]} x ${k}`).join(', ');
  const ok = ui.alert('Process graduates?',
    `Applying: ${summary}. "Delete" permanently removes the Fabman account and cannot be undone.`,
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  todo.forEach((t, i) => {
    ss.toast(`Processing ${i + 1} of ${todo.length}...`, 'Graduates', -1);
    let result = '';
    try {
      const member = fetchMember(t.id);
      if (t.action === 'Delete') {
        const url = `${FABMAN_MEMBERS_ENDPOINT}/${t.id}`;
        const res = UrlFetchApp.fetch(url, {
          method: 'delete',
          headers: { 'Authorization': `Bearer ${FABMAN_API_KEY}` },
          muteHttpExceptions: true
        });
        result = res.getResponseCode() < 300 ? 'Deleted' : `Delete failed (${res.getResponseCode()})`;
      } else {
        const done = putMember(member, { state: 'locked', lockVersion: member.lockVersion });
        result = done ? 'Locked (access revoked)' : 'Lock failed';
      }
    } catch (err) {
      result = 'Error: ' + (err.message || err);
    }
    sheet.getRange(t.rowNum, C.result + 1).setValue(result);
  });

  ss.toast('Done.', 'Graduates', 3);
  ui.alert('Done. See the Result column on the Graduates sheet.');
}

//------------------------------------------------------------
//
// HELPER FUNCTIONS to interface with the Fabman API.
//
//------------------------------------------------------------

function fetchAll(baseUrl) {
  const API_BASE = 'https://fabman.io';
  const headers = {
    'Authorization': `Bearer ${FABMAN_API_KEY}`
  };

  let results = [];
  let url = baseUrl;

  while (url) {
    const response = UrlFetchApp.fetch(url, { method: 'get', headers });
    const data = JSON.parse(response.getContentText());
    const responseHeaders = response.getAllHeaders();

    if (Array.isArray(data.items)) {
      results = results.concat(data.items);
    } else if (Array.isArray(data)) {
      results = results.concat(data);
    }

    const linkHeader = responseHeaders['Link'] || responseHeaders['link'];
    if (linkHeader) {
      const match = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
      if (match) {
        const nextPage = match[1];
        url = nextPage.startsWith('http') ? nextPage : `${API_BASE}${nextPage}`;
      } else {
        url = null;
      }
    } else {
      url = null;
    }
  }

  Logger.log(`Fetched ${results.length} records.`);
  return results;
}

function fetchMember(memberId) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${memberId}`;
  const options = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${FABMAN_API_KEY}` }
  };
  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function putMember(member, payload) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}`;
  const options = {
    method: 'put',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${FABMAN_API_KEY}`
    },
    payload: JSON.stringify(payload)
  };
  const response = UrlFetchApp.fetch(url, options);
  return response.getResponseCode() === 200;
}

function postTraining(member, trainingId) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/trainings`;

  // Skip if the member already has this training, to avoid duplicates.
  const trainings = getMemberTrainings(member.id);
  const hasTraining = trainings.some(training => training.trainingCourse === trainingId);
  if (hasTraining) return;

  const payload = {
    trainingCourse: trainingId,
    fromDate: new Date().toISOString().slice(0, 10)
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getResponseCode() === 201;
}

function postPackage(member) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/packages`;

  const payload = {
    package: STANDARD_PACKAGE,
    fromDate: new Date().toISOString().slice(0, 10)
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getResponseCode() === 201;
}

function deleteTraining(member, trainingId) {
  const trainings = getMemberTrainings(member.id);
  const targetTraining = trainings.find(t => t.trainingCourse === trainingId);
  if (!targetTraining) {
    Logger.log(`Training ${trainingId} not found for member ${member.id}`);
    return;
  }

  const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/trainings/${targetTraining.id}`;
  const options = {
    method: 'delete',
    headers: {
      'Authorization': `Bearer ${FABMAN_API_KEY}`
    },
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(url, options);
}

function deletePackage(member, packageId) {
  const packages = getMemberPackages(member.id);
  const targetPackage = packages.find(p => p.package === packageId);
  if (!targetPackage) {
    Logger.log(`Package ${packageId} not found for member ${member.id}`);
    return;
  }

  const url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/packages/${targetPackage.id}`;
  const options = {
    method: 'delete',
    headers: {
      'Authorization': `Bearer ${FABMAN_API_KEY}`
    },
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(url, options);
}

function getMemberTrainings(memberId) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${memberId}/trainings`;
  const options = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function getMemberPackages(memberId) {
  const url = `${FABMAN_MEMBERS_ENDPOINT}/${memberId}/packages`;
  const options = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function findMemberByEmail(email, members) {
  const targetEmail = email.trim().toLowerCase();
  const targetMember = members.find(m => (m.emailAddress || '').toLowerCase() === targetEmail) || null;
  if (!targetMember) Logger.log(`No member found with email: ${targetEmail}`);
  return targetMember;
}

function scriptProperty_(name) {
  try { return PropertiesService.getScriptProperties().getProperty(name) || ''; }
  catch (e) { return ''; }   // simple triggers (onEdit) may not read properties; they do not need them
}

//------------------------------------------------------------
//
// SIGN-UPS from the FabMan Agreement in the Safety & Training Hub Canvas course. Each
// agreement becomes a row on one of two tabs:
//   FabMan Agreement   written by the Qualtrics survey's workflow (Google Sheets task)
//   Sign-up Form       the older Google Form, until it is retired
// The script makes each row a Fabman account (or adds the trainings to the account the
// person already has) and writes what happened in the row's "Fabman status" cell.
//
//   onSheetChange       runs on every change to the spreadsheet (change trigger), so a
//                       row Qualtrics adds is done within seconds
//   onFormSubmit        runs on each Google Form response (form-submit trigger)
//   processNewSignups   runs every hour and does any row still without a status, so a
//                       missed trigger or a Fabman outage catches up on its own
//   setupSignups        run once, signed in as the account that should own the triggers
//   checkSignups        changes nothing: lists recent sign-ups that have no account
// The Fabman sign-ups menu in the sheet runs the last three.
//
// To redo a row (a mistyped NetID, say), fix the cell and clear its Fabman status.
//
//------------------------------------------------------------

const SIGNUPS = {
  // Where agreements arrive. "since" tabs only count rows from SIGNUPS_FROM on, since they
  // hold years of history; the Qualtrics tab holds nothing else, so every new row counts.
  tabs: [
    { name: 'FabMan Agreement', since: false },
    { name: 'Sign-up Form', since: true },
  ],
  // The Qualtrics tab's columns. setupSignups makes the tab with these; the survey's
  // Google Sheets task fills one column from each matching survey field.
  qualtricsHeaders: ['Recorded Date', 'Response ID', 'NetID', 'First Name', 'Last Name', 'Project Team', 'Graduation Year', 'Apron', 'Attestation'],
  status: 'Fabman status',
  report: 'Sign-up check',
  everyone: [SPACES, LAB_EQUIPMENT],
  // Extra trainings for the apron someone reports on the form. These match the postRed,
  // postGreen and postBlue buttons: only the red apron adds machines and mills.
  apron: { 'red apron': [MACHINES, MILLS] },
  retry: 'Will retry',   // rows whose status starts with this are tried again every hour
};
const TRAINING_NAMES = { [SPACES]: 'Spaces', [LAB_EQUIPMENT]: 'Lab', [MACHINES]: 'Machines', [MILLS]: 'Mills', [SUPERVISOR]: 'Supervisor' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Fabman sign-ups')
    .addItem('Check sign-ups (changes nothing)', 'checkSignups')
    .addItem('Process waiting sign-ups now', 'processNewSignups')
    .addItem('Process sign-ups since a date...', 'processSignupsSince')
    .addSeparator()
    .addItem('Set up the sign-up triggers', 'setupSignups')
    .addToUi();
}

// Form-submit trigger. The trigger must be installed on this spreadsheet (setupSignups
// does that), so the event carries the new row.
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const run = signupRun_();
    if (e && e.range && signupTab_(e.range.getSheet().getName())) signupRow_(e.range.getSheet(), e.range.getRow(), run);
    else signupPass_(run);
    signupAlerts_(run);
  } finally {
    lock.releaseLock();
  }
}

// Change trigger: a row added by Qualtrics (or a status cleared by hand to retry it) is
// picked up at once. Any other change finds nothing waiting and ends straight away.
function onSheetChange(e) {
  processNewSignups();
}

// Hourly trigger, the change trigger, and the menu's "Process waiting sign-ups now".
function processNewSignups() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'Another run is busy; the next one picks these up.';
  try {
    const run = signupRun_();
    const pass = signupPass_(run);
    signupAlerts_(run);
    const again = 'choose Process waiting sign-ups now again (the hourly run also finishes them).';
    const msg = pass.done && pass.left ? `${pass.done} sign-up(s) processed. ${pass.left} more ${pass.left === 1 ? 'is' : 'are'} still waiting: ${again}`
      : pass.done ? `${pass.done} sign-up(s) processed. See the Fabman status column.`
      : pass.left ? `${pass.left} sign-up(s) are waiting, but this run ran out of time: ${again}`
      : 'No sign-ups were waiting.';
    if (pass.done || pass.left) toast_(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

// Menu: process every sign-up from a date on, for catching up after an outage. Rows
// that already have a status are left alone.
function processSignupsSince() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Process sign-ups since a date',
    'Sign-ups from this date on that have no Fabman status yet get accounts (and invitation emails). Date as YYYY-MM-DD:',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const from = parseDay_(res.getResponseText());
  if (!from) { ui.alert('That is not a date. Use the form 2026-08-24.'); return; }
  PropertiesService.getScriptProperties().setProperty('SIGNUPS_FROM', from.toISOString());
  ui.alert(processNewSignups());
}

// Run once (from the editor or the menu) signed in as the account that should own the
// triggers: a shared or staff account that will be around next year, not a student's.
function setupSignups() {
  fabman_('get', `${FABMAN_MEMBERS_ENDPOINT}?limit=1`);   // throws with a clear message if Fabman refuses the key

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // The tab the Qualtrics workflow writes to, with the headers its Google Sheets task maps to.
  const qt = SIGNUPS.tabs[0].name;
  if (!ss.getSheetByName(qt)) {
    const sh = ss.insertSheet(qt);
    const H = SIGNUPS.qualtricsHeaders.concat([SIGNUPS.status]);
    sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onFormSubmit', 'onSheetChange', 'processNewSignups', 'updateQualtrics'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSheetChange').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('processNewSignups').timeBased().everyHours(1).create();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SIGNUPS_FROM')) props.setProperty('SIGNUPS_FROM', new Date().toISOString());
  signupSheets_().forEach(sh => statusColumn_(sh));

  const msg = 'Sign-ups are set up for ' + Session.getEffectiveUser().getEmail() + '. New agreements get accounts right away, '
    + 'and an hourly run catches anything missed (from ' + Utilities.formatDate(new Date(props.getProperty('SIGNUPS_FROM')), Session.getScriptTimeZone(), 'MMM d, yyyy') + '). '
    + 'Problems are emailed to ' + alertAddresses_().join(', ') + '. Other people\'s old triggers for this sheet can be removed on the Triggers page.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor */ }
}

// Changes nothing in Fabman. Writes a "Sign-up check" tab listing the sign-ups since a
// date whose NetID has no Fabman account, and what the script would do with them.
function checkSignups() {
  let from = null;
  try {
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt('Check sign-ups', 'Check sign-ups since (YYYY-MM-DD). Leave it empty for the last 120 days.', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    from = res.getResponseText().trim() ? parseDay_(res.getResponseText()) : null;
    if (res.getResponseText().trim() && !from) { ui.alert('That is not a date. Use the form 2026-08-24.'); return; }
  } catch (e) { /* run from the editor */ }
  from = from || new Date(Date.now() - 120 * 86400000);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const run = signupRun_();
  const rows = [['Tab', 'Row', 'Submitted', 'Name', 'NetID as typed', 'Fabman account', 'Fabman status', 'What processing would do']];
  let missing = 0, total = 0;
  signupSheets_().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const H = data[0].map(String);
    const stCol = headerIndex_(H, new RegExp('^' + SIGNUPS.status + '$', 'i'));
    for (let r = 1; r < data.length; r++) {
      if (!data[r].join('').trim()) continue;
      const ts = rowTime_(H, data[r]);
      if (ts && ts < from) continue;
      total++;
      let person = null, found = null, plan;
      try { person = signupPerson_(H, data[r]); found = memberByEmail_(run, person.email); }
      catch (err) { plan = 'Cannot process: ' + err.message; }
      if (person && !plan) {
        plan = !found ? 'Create the account, add ' + trainingList_(person.trainings) + ', send the invitation'
          : found.state !== 'active' ? 'Nothing: the account is ' + found.state
          : 'Add any missing trainings';
      }
      if (found && found.state === 'active') continue;   // only list the ones that need attention
      if (!found) missing++;
      rows.push([sheet.getName(), r + 1, ts || '', person ? (person.firstName + ' ' + person.lastName).trim() : '', String(data[r][headerIndex_(H, /netid/i)] || ''),
        found ? '#' + found.id + ' (' + found.state + ')' : 'none', stCol >= 0 ? String(data[r][stCol] || '') : '', plan]);
    }
  });
  const out = ss.getSheetByName(SIGNUPS.report) || ss.insertSheet(SIGNUPS.report);
  out.clear();
  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  out.setFrozenRows(1);
  const msg = total + ' sign-up(s) since ' + Utilities.formatDate(from, Session.getScriptTimeZone(), 'MMM d, yyyy') + '; '
    + missing + ' have no Fabman account. Details are on the "' + SIGNUPS.report + '" tab. Nothing was changed in Fabman.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor */ }
  return msg;
}

// ---- the work behind them ----

// One run's shared state: the member list is fetched once and reused.
function signupRun_() { return { members: null, problems: [], started: Date.now() }; }

// Every waiting row (status empty or "Will retry") on each sign-up tab; on the Google
// Form tab only rows from SIGNUPS_FROM on. Stops starting new rows after four minutes to
// stay inside Apps Script's six-minute limit, and counts the rest as left for next time.
function signupPass_(run) {
  const from = new Date(scriptProperty_('SIGNUPS_FROM') || Date.now() - 2 * 86400000);
  let done = 0, left = 0;
  SIGNUPS.tabs.forEach(tab => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab.name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    const H = data[0].map(String);
    const stCol = statusColumn_(sheet);
    for (let r = 1; r < data.length; r++) {
      if (!data[r].join('').trim()) continue;
      if (tab.since) { const ts = rowTime_(H, data[r]); if (!ts || ts < from) continue; }
      const status = String(data[r][stCol - 1] || '');
      if (status && status.indexOf(SIGNUPS.retry) !== 0) continue;
      if (Date.now() - run.started > 4 * 60 * 1000) { left++; continue; }
      signupRow_(sheet, r + 1, run);
      done++;
    }
  });
  return { done: done, left: left };
}

function signupTab_(name) { return SIGNUPS.tabs.some(t => t.name === name); }
function signupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SIGNUPS.tabs.map(t => ss.getSheetByName(t.name)).filter(Boolean);
}

// When a row was submitted: a date cell, or text such as "2026-10-26 11:45:10". Text
// comes from Qualtrics, which writes times in Mountain Standard Time (UTC-7) whatever
// the account's time zone. Null if there is no date column or it cannot be read.
function rowTime_(H, row) {
  const i = headerIndex_(H, /^(timestamp|recorded ?date|end ?date|submitted( at)?|date)$/i);
  if (i < 0) return null;
  const v = row[i];
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v || '').trim();
  const mst = (y, mo, d, h, mi, sec) => new Date(Date.UTC(+y, +mo - 1, +d, +(h || 0) + 7, +(mi || 0), +(sec || 0)));
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) return mst(m[1], m[2], m[3], m[4], m[5], m[6]);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?)?$/i.exec(s);
  if (m) return mst(m[3], m[1], m[2], m[7] ? +m[4] % 12 + (/p/i.test(m[7]) ? 12 : 0) : m[4], m[5], m[6]);
  const d = new Date(s);   // text that names its own zone, such as 2026-10-26T18:45:10Z
  return s && !isNaN(d.getTime()) ? d : null;
}

function signupRow_(sheet, rowNum, run) {
  const stCol = statusColumn_(sheet);
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const v = sheet.getRange(rowNum, 1, 1, H.length).getValues()[0];
  const before = String(v[stCol - 1] || '');
  if (before && before.indexOf(SIGNUPS.retry) !== 0) return before;   // done already
  let status;
  try {
    const result = signupUpsert_(signupPerson_(H, v), run, before);
    status = result.text;
    if (result.attention) run.problems.push({ tab: sheet.getName(), row: rowNum, status: status, retry: false, key: false });
  } catch (err) {
    status = (err.retry ? SIGNUPS.retry + ': ' : 'Error: ') + err.message;
    run.problems.push({ tab: sheet.getName(), row: rowNum, status: status, retry: !!err.retry, key: !!err.key });
  }
  sheet.getRange(rowNum, stCol).setValue(status + ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a') + ')');
  return status;
}

// Reads a response by question title, so reordering or adding questions on the form
// does not break it.
function signupPerson_(H, v) {
  const get = re => { const i = headerIndex_(H, re); return i < 0 ? '' : String(v[i] == null ? '' : v[i]).trim(); };
  const typed = get(/netid/i).toLowerCase().replace(/\s+/g, '');
  const collected = get(/^(email address|email|e-mail)$/i).toLowerCase();
  // The agreement answer: blank, or a plain "No", means it was not accepted.
  const agreed = headerIndex_(H, /attestation|i acknowledge/i);
  const answer = agreed < 0 ? 'n/a' : String(v[agreed] == null ? '' : v[agreed]).trim();
  if (!answer || /^(no|not|false|0|i do not|i don't|i disagree|disagree|i decline|decline)\b/i.test(answer)) {
    throw new Error('The agreement was not accepted on this response, so no account was made.');
  }
  let email = '';
  if (/^[a-z]+\d+@cornell\.edu$/.test(collected)) email = collected;
  else if (/^[a-z]+\d+$/.test(typed)) email = typed + '@cornell.edu';
  else if (/^[a-z]+\d+@cornell\.edu$/.test(typed)) email = typed;
  if (!email) throw new Error(`"${typed || 'blank'}" is not a Cornell NetID. Fix the NetID cell, then clear this status to try again.`);
  const apron = get(/apron/i);
  const year = (get(/graduation/i).match(/20\d\d/) || [])[0];
  return {
    email: email, netid: email.split('@')[0],
    firstName: get(/^first name/i), lastName: get(/^last name/i), team: get(/^project team|^team/i),
    year: year ? Number(year) : null, apron: apron || 'None',
    trainings: SIGNUPS.everyone.concat(SIGNUPS.apron[apron.toLowerCase()] || []),
  };
}

// Make sure the person has an active account with the right trainings. Safe to run
// twice: nothing is added twice. "before" is the row's previous status, if any.
// Returns { text: what happened, attention: true if an admin should look }.
function signupUpsert_(p, run, before) {
  let member = memberByEmail_(run, p.email);
  const createdEarlier = /created #(\d+)/i.exec(before || '');
  if (member && member.state !== 'active') {
    return { text: `Has an account (#${member.id}) but it is ${member.state}, so nothing was changed. Activate it in Fabman if they should have access.`, attention: true };
  }
  let created = false;
  if (!member) {
    member = fabman_('post', FABMAN_MEMBERS_ENDPOINT, {
      account: FABMAN_ACCOUNT_ID, firstName: p.firstName, lastName: p.lastName, company: p.team,
      emailAddress: p.email, state: 'active',
      metadata: { graduationYear: p.year, apron: p.apron, netid: p.netid },
    });
    run.members.push(member);
    created = true;
  }
  const added = [];
  try {
    const have = fabman_('get', `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/trainings`).map(t => t.trainingCourse);
    p.trainings.forEach(id => {
      if (have.indexOf(id) >= 0) return;
      fabman_('post', `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/trainings`, { trainingCourse: id, fromDate: today_() });
      added.push(id);
    });
    const packages = fabman_('get', `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/packages`);
    if (!packages.some(pk => pk.package === STANDARD_PACKAGE)) {
      fabman_('post', `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/packages`, { package: STANDARD_PACKAGE, fromDate: today_() });
    }
    if (created || createdEarlier) fabman_('post', `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/invitation`);
  } catch (err) {
    // The account exists now; the retry finishes the trainings and the invitation.
    if (created || createdEarlier) err.message = `created #${member.id}, then: ${err.message}`;
    throw err;
  }
  if (created || createdEarlier) return { text: `Created #${member.id} with ${trainingList_(p.trainings)}; invitation sent` };
  return { text: `Already had an account (#${member.id})` + (added.length ? `; added ${trainingList_(added)}` : '; nothing to add') };
}

function memberByEmail_(run, email) {
  if (!run.members) run.members = fabmanAll_(FABMAN_MEMBERS_ENDPOINT);
  const target = email.toLowerCase();
  return run.members.find(m => String(m.emailAddress || '').toLowerCase() === target) || null;
}

// One Fabman request. Returns the parsed body; throws an error whose message says what
// went wrong. err.retry marks problems worth trying again later (Fabman busy or down, or
// the key refused), err.key marks a refused key.
function fabman_(method, url, payload) {
  const options = { method: method, headers: { Authorization: `Bearer ${FABMAN_API_KEY}` }, muteHttpExceptions: true };
  if (payload !== undefined) { options.contentType = 'application/json'; options.payload = JSON.stringify(payload); }
  let res;
  try { res = UrlFetchApp.fetch(url, options); }
  catch (err) { const e = new Error('Fabman did not answer (' + (err.message || err) + ')'); e.retry = true; throw e; }
  const code = res.getResponseCode(), text = res.getContentText();
  if (code >= 200 && code < 300) { try { return text ? JSON.parse(text) : {}; } catch (e) { return {}; } }
  let detail = text;
  try { const j = JSON.parse(text); detail = j.message || j.error || j.title || text; } catch (e) { /* not JSON */ }
  detail = String(detail).slice(0, 200);
  const e = new Error(code === 401 || code === 403
    ? `Fabman refused the API key (${code}). Check the key at the top of the script, or the FABMAN_API_KEY script property.`
    : code === 429 || code >= 500 ? `Fabman is busy or down (${code})` : `Fabman said no (${code}): ${detail}`);
  e.retry = code === 401 || code === 403 || code === 429 || code >= 500;
  e.key = code === 401 || code === 403;
  throw e;
}

// Every page of a Fabman list.
function fabmanAll_(url) {
  let results = [], next = url;
  for (let page = 0; next && page < 200; page++) {
    const options = { method: 'get', headers: { Authorization: `Bearer ${FABMAN_API_KEY}` }, muteHttpExceptions: true };
    let res;
    try { res = UrlFetchApp.fetch(next, options); }
    catch (err) { const e = new Error('Fabman did not answer (' + (err.message || err) + ')'); e.retry = true; throw e; }
    if (res.getResponseCode() !== 200) fabman_('get', next);   // let fabman_ explain the error
    const data = JSON.parse(res.getContentText());
    results = results.concat(Array.isArray(data) ? data : (data.items || []));
    const headers = res.getAllHeaders();
    const link = String(headers.Link || headers.link || '');
    const m = link.match(/<([^>]+)>;\s*rel="next"/);
    next = m ? (m[1].indexOf('http') === 0 ? m[1] : 'https://fabman.io' + m[1]) : null;
  }
  return results;
}

// At most one email per run, to the addresses in the ALERT_EMAILS script property. A
// refused key is reported at most once a day, since every row would hit it.
function signupAlerts_(run) {
  const errors = run.problems.filter(p => !p.retry);
  const keyProblem = run.problems.some(p => p.key);
  const props = PropertiesService.getScriptProperties();
  const lastKeyAlert = Number(props.getProperty('KEY_ALERT_AT') || 0);
  const sendKey = keyProblem && Date.now() - lastKeyAlert > 86400000;
  if (!errors.length && !sendKey) return;
  if (sendKey) props.setProperty('KEY_ALERT_AT', String(Date.now()));
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const lines = [];
  if (sendKey) lines.push('Fabman is refusing the API key, so no new accounts are being made. Put a working key in the FABMAN_API_KEY script property (or at the top of the script). Waiting sign-ups are done automatically once it works.', '');
  if (errors.length) {
    lines.push(errors.length + ' sign-up(s) need a look:');
    errors.forEach(p => lines.push('  ' + p.tab + ', row ' + p.row + ': ' + p.status));
    lines.push('', 'Fix the row and clear its Fabman status cell; the script tries it again straight away.');
  }
  lines.push('', url);
  MailApp.sendEmail(alertAddresses_().join(','), 'Fabman sign-ups need attention', lines.join('\n'));
}

function alertAddresses_() {
  const list = scriptProperty_('ALERT_EMAILS').split(/[,\s]+/).filter(Boolean);
  return list.length ? list : ['nhh5@cornell.edu'];
}

// The 1-based column of "Fabman status", added after the last column if missing.
function statusColumn_(sheet) {
  const lc = sheet.getLastColumn();
  const H = sheet.getRange(1, 1, 1, lc).getValues()[0].map(String);
  const i = headerIndex_(H, new RegExp('^' + SIGNUPS.status + '$', 'i'));
  if (i >= 0) return i + 1;
  sheet.getRange(1, lc + 1).setValue(SIGNUPS.status).setFontWeight('bold');
  return lc + 1;
}

function headerIndex_(H, re) { return H.findIndex(h => re.test(String(h).trim())); }
function trainingList_(ids) { return ids.map(id => TRAINING_NAMES[id] || id).join(', '); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function parseDay_(text) {
  const m = /^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$/.exec(String(text || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}
function toast_(msg) { try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Fabman sign-ups', 6); } catch (e) { /* no sheet open */ } }
