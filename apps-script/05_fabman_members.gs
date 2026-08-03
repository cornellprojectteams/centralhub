const FABMAN_API_KEY = '6387c077-6950-4ffd-aa96-23db92606dc3'; // SPT Key
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

const QUALTRICS_API_KEY = 'VcfwRwqsz0Zo4GdujqzP2bTX2wTptaYFoIaaI9Rd';
const QUALTRICS_ENDPOINT = 'https://ca1.qualtrics.com/API/v3/surveys';
const SPACES_QUALTRICS = 'SV_3CAIAyTPzwIxAFg'; // TODO: UPDATE
const LAB_EQUIPMENT_QUALTRICS = 'SV_3CAIAyTPzwIxAFg'; // TODO: UPDATE

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

function updateQualtrics() {
  exportQualtrics(SPACES_QUALTRICS, SPACES);
  exportQualtrics(LAB_EQUIPMENT_QUALTRICS, LAB_EQUIPMENT);
}

function setup() {
  const setup_value = 4;
  PropertiesService.getScriptProperties().setProperty(`last_${SPACES}`, setup_value);
  PropertiesService.getScriptProperties().setProperty(`last_${LAB_EQUIPMENT}`, setup_value);
  Logger.log(`Updated last_${SPACES} and last_${LAB_EQUIPMENT}: ${setup_value}`);
}

function invitation() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) { return; }

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

function onFormSubmit(e) {
  const row = e.values;
  const firstName = row[2];
  const lastName = row[3];
  const team = row[4];
  const year = row[5];
  var netid = row[6];
  const training = row[7];

  var email;
  if (netid.includes("@cornell.edu")) {
    email = netid;
    var atIndex = netid.indexOf("@");
    netid = netid.substring(0, atIndex);
  }
  else {
    email = `${netid}@cornell.edu`
  }

  const payload = {
    account: FABMAN_ACCOUNT_ID,
    firstName: firstName,
    lastName: lastName,
    company: team,
    metadata: {
      graduationYear: parseInt(year),
      apron: training,
      netid: netid
    },
    emailAddress: email,
    state: 'active'
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": `Bearer ${FABMAN_API_KEY}`,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(FABMAN_MEMBERS_ENDPOINT, options);
  const status = response.getResponseCode();
  const responseBody = response.getContentText();

  if (status !== 201) {
    var subject = "Failed to create Fabman Member";
    var message = `Failed to create member for ${firstName} ${lastName}: ${status} - ${responseBody}`;
    MailApp.sendEmail('ej289@cornell.edu', subject, message);
    MailApp.sendEmail('nhh5@cornell.edu', subject, message);
    return;
  }

  const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);
  const member = findMemberByEmail(email, members)

  postTraining(member, SPACES);
  postTraining(member, LAB_EQUIPMENT);

  if (training.includes("Apron")) {
    postTraining(member, MACHINES);
    postTraining(member, MILLS);
  }
  postPackage(member);

  const invitation_options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${FABMAN_API_KEY}`
    },
  };
  const invitation_url = `${FABMAN_MEMBERS_ENDPOINT}/${member.id}/invitation`;
  UrlFetchApp.fetch(invitation_url, invitation_options);
}

function exportQualtrics(form, training) {
  const start_url = `${QUALTRICS_ENDPOINT}/${form}/export-responses`;
  const start_payload = { format: 'csv' };
  const start_options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'X-API-TOKEN': QUALTRICS_API_KEY
    },
    payload: JSON.stringify(start_payload)
  };

  const start_response = UrlFetchApp.fetch(start_url, start_options);
  const start_data = JSON.parse(start_response.getContentText());
  var progressId = start_data.result.progressId;

  const poll_url = `${QUALTRICS_ENDPOINT}/${form}/export-responses/${progressId}`;
  const options = {
    method: 'GET',
    headers: {
      'X-API-TOKEN': QUALTRICS_API_KEY
    }
  };

  var status = '';
  while (status !== 'complete') {
    var poll_response = UrlFetchApp.fetch(poll_url, options);
    var poll_data = JSON.parse(poll_response.getContentText());
    status = poll_data.result.status;

    if (status === 'complete') {
      const fileId = poll_data.result.fileId;
      const export_url = `${QUALTRICS_ENDPOINT}/${form}/export-responses/${fileId}/file`;
      const export_response = UrlFetchApp.fetch(export_url, options);

      const zipBlob = export_response.getBlob();
      const unzippedFiles = Utilities.unzip(zipBlob);

      let export_data = '';
      unzippedFiles.forEach(function (file) {
        if (file.getContentType() === 'application/zip' || file.getName().endsWith('.csv')) {
          export_data = file.getDataAsString();
        }
      });

      const rows = export_data.split('\n').map(function (row) {
        return row.split(',');
      });

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Qualtrics_${training}`);
      sheet.clear();

      let maxColumns = 0;
      rows.forEach(function (row) {
        maxColumns = Math.max(maxColumns, row.length);
      });

      rows.forEach(function (row) {
        while (row.length < maxColumns) {
          row.push('');
        }
      });
      sheet.getRange(1, 1, rows.length, maxColumns).setValues(rows);

      var last_upload = PropertiesService.getScriptProperties().getProperty(`last_${training}`);
      const start = parseInt(last_upload + 1, 10);
      const netids = sheet.getRange(start, 18, sheet.getLastRow() - 1).getValues().flat();
      const members = fetchAll(FABMAN_MEMBERS_ENDPOINT);
      netids.forEach(netid => {
        if (!netid) return;
        Logger.log(`NetID: ${netid}`);
        const member = findMemberByEmail(`${netid}@cornell.edu`, members);
        postTraining(member, training);
      });

      PropertiesService.getScriptProperties().setProperty(`last_${training}`, rows.length);
      return;
    }
    Utilities.sleep(5000);
  }
}
