'use strict';
// Stand-ins for the Google Forms that feed the platform. Submitting one does what
// Google does: a new row on the response tab, then the form-submit trigger.
//
// The questions mirror the real forms closely enough for the scripts, which read
// answers by question title. To change a question here, edit this file.

const { displayValue, toSigned } = require('./gas');

const WITHIN = ['Same day', '2 to 3 business days', '7 to 10 days', '14 to 21 days', 'When time allows'];
const ISSUE_TYPES = ['Fire code', 'Egress / access', 'Work Practices', 'Chemical storage', 'Shared space', 'Housekeeping', 'Maintenance / facilities', 'Informational'];
const ACTIONS = ['Clean up and organize the area', 'Remove or relocate items', 'Store chemicals correctly', 'Clear the path or exit', 'Correct an unsafe work practice', 'Submit a facilities work order', 'No action needed (informational)', 'Other'];

const ELL = {
  shift: ['10-12', '12-2', '2-4', '4-6', '6-8', '8-10'],
  activity: ['Intense - 80-100%', 'High - 60-80%', 'Moderate - 40-60%', 'Low - 20-40%', 'Quiet - 0-20%'],
  main: ['Computer-Work, Study', 'Team Meetings', 'Assembly/ Build Work', 'Tours/ Interviews', 'Social gatherings', 'Empty'],
  composites: ['Sanding / Mold-prep', 'Active Layup Work', 'Paint Booth use', 'Construction', 'Testing', 'Cleaning/ Maintenance', 'Empty'],
  auto: ['Fabrication', 'Assembly', 'Testing', 'Empty'],
  conference: ['Meetings in progress', 'Interviews', 'Work Sessions', 'Empty'],
  supply: ['Gloves', 'First Aid Supplies', 'N95 masks', 'Safety Glasses', 'Paper Towels', 'Cleaning Supplies', 'All supplies stocked'],
  safety: ['Blocked Exits/ Extinguishers', 'Equipment left on/ Unsafe conditions', 'Facility issues (lights, leaks, etc.)', 'Chemical waste needs attention', 'No issues observed', 'Other - (add note)'],
  teams: ['Rocketry', 'CUAir', 'FSAE', 'Baja', 'CEV', 'AUV', 'Mars Rover', 'CRC', 'DBF', 'Nexus', 'Seismic', 'C2S2', 'CAT', 'Hyperloop', 'CUSail', 'CUXR', 'AutoBoat', 'EWH', 'GeoData', 'iGEM', 'DEBUT', 'EWB', 'CUBMD', 'ESW'],
};
const APRONS = ['None', 'Blue Apron', 'Green Apron', 'Red Apron'];

// The Space Status report form. Team choices come from the Team Contacts tab here; the
// real form keeps its own list, which has to match Team Contacts.
function spaceForm(ids, teams) {
  return {
    key: 'space-status', title: 'Space Status Report',
    intro: 'Use this form to record stewardship and safety issues, upload photos, and notify the responsible team of required action.',
    sheet: ids.space, tab: 'Form Responses', collectEmail: true,
    trigger: { project: 'space', fn: 'sendEmails', style: 'namedValues' },
    uploadFolder: 'Space Status Report (File responses)',
    questions: [
      { title: 'NetID', kind: 'text', required: true },
      { title: 'Responsible Team', kind: 'select', required: true, choices: teams },
      { title: 'Respond within', kind: 'radio', required: true, choices: WITHIN },
      { title: 'Photo Upload', kind: 'file', required: false },
      { title: 'Issue Type', kind: 'radio', required: true, choices: ISSUE_TYPES },
      { title: 'Required action', kind: 'radio', required: true, choices: ACTIONS },
      { title: 'Details / comments', kind: 'paragraph', required: true },
    ],
  };
}

function ellForm(ids) {
  return {
    key: 'ell', title: 'ELL Shift Summary',
    intro: 'To be completed by student employees during their shift in the ELL.',
    sheet: ids.ell, tab: 'Form Responses 1', collectEmail: true, trigger: null,
    uploadFolder: 'ELL Shift Summary (File responses)',
    questions: [
      { title: 'NetID', kind: 'text', required: true },
      { title: 'Shift', kind: 'radio', required: true, choices: ELL.shift },
      { title: 'Overall Activity Level and Space Use', kind: 'radio', required: true, choices: ELL.activity },
      { title: 'Main ELL use (check all that apply)', kind: 'checks', required: true, choices: ELL.main },
      { title: 'Composites Lab Use (check all that apply)', kind: 'checks', required: true, choices: ELL.composites },
      { title: 'Auto Lab Use (check all that apply)', kind: 'checks', required: true, choices: ELL.auto },
      { title: 'Conference Room Use (check all that apply)', kind: 'checks', required: true, choices: ELL.conference },
      { title: 'Supply Status (mark if we need to purchase more)', kind: 'checks', required: true, choices: ELL.supply },
      { title: 'Safety and Maintenance Issues (check all that apply)', kind: 'checks', required: true, choices: ELL.safety },
      { title: 'Most active teams today', kind: 'checks', required: true, choices: ELL.teams },
      { title: 'Productivity notes (progress on assigned tasks)', kind: 'paragraph', required: false },
      { title: 'Additional Notes', kind: 'paragraph', required: false },
      { title: 'photos', kind: 'file', required: false },
    ],
  };
}

function fabmanForm(ids, teams) {
  return {
    key: 'fabman-signup', title: 'FabMan Agreement',
    intro: 'The agreement students sign in the Safety & Training Hub Canvas course. The Fabman script makes their Fabman account, adds the trainings and sends the invitation (here, in the fake Fabman), then writes what it did in the Fabman status column.',
    sheet: ids.fabman, tab: 'Sign-up Form', collectEmail: false,
    trigger: { project: 'fabman', fn: 'onFormSubmit', style: 'values' },
    questions: [
      { title: 'I acknowledge that I have read, understand, and agree to adhere to the Fabman and Machine Shop usage policies outlined above.', kind: 'radio', required: true, choices: ['I agree'] },
      { title: 'First Name', kind: 'text', required: true },
      { title: 'Last Name', kind: 'text', required: true },
      { title: 'Project Team', kind: 'select', required: true, choices: teams },
      { title: 'Graduation Year', kind: 'select', required: true, choices: ['2026', '2027', '2028', '2029', '2030'] },
      { title: 'NetID', kind: 'text', required: true },
      { title: 'Previously Completed Training (Emerson Apron Status)', kind: 'radio', required: true, choices: APRONS },
    ],
  };
}

// The FabMan Agreement as a Qualtrics survey whose workflow adds a row to the Fabman
// sheet's "FabMan Agreement" tab (the Google Sheets task). Like Qualtrics, it writes the
// recorded date as text in Mountain Standard Time (UTC-7) and adds a response ID, and
// the sheet's change trigger runs.
function fabmanQualtrics(ids, teams) {
  return {
    key: 'fabman-agreement', title: 'FabMan Agreement (Qualtrics)',
    intro: 'The Qualtrics version of the agreement, embedded in the Safety & Training Hub. Its workflow adds a row to the FabMan Agreement tab, and the Fabman script makes the account within seconds.',
    sheet: ids.fabman, tab: 'FabMan Agreement', collectEmail: false,
    timestampHeader: 'Recorded Date', textTimestamp: true, responseId: true,
    trigger: { project: 'fabman', style: 'change' },
    questions: [
      { title: 'NetID', kind: 'text', required: true },
      { title: 'First Name', kind: 'text', required: true },
      { title: 'Last Name', kind: 'text', required: true },
      { title: 'Project Team', kind: 'select', required: true, choices: teams },
      { title: 'Graduation Year', kind: 'select', required: true, choices: ['2026', '2027', '2028', '2029', '2030'] },
      { title: 'Apron', kind: 'radio', required: true, choices: APRONS },
      { title: 'Attestation', kind: 'radio', required: true, choices: ['I have read the Fabman and Machine Shop usage policies and agree to follow them'] },
    ],
  };
}

function teamChoices(account, ids) {
  const ss = account.ss(ids.space);
  const sh = ss && ss.getSheetByName('Team Contacts');
  if (!sh) return [];
  return sh.rows().slice(1).map((r) => String(r[0] || '').trim()).filter(Boolean);
}

function definitions(account, ids) {
  const teams = teamChoices(account, ids);
  return {
    'space-status': spaceForm(ids, teams),
    ell: ellForm(ids),
    'fabman-signup': fabmanForm(ids, teams.filter((t) => !/operations|unknown/i.test(t))),
    'fabman-agreement': fabmanQualtrics(ids, teams.filter((t) => !/operations|unknown/i.test(t))),
  };
}

// answers: { 'Question title': 'value' | ['a', 'b'] }, files: { 'Question title': [{ name, mime, buffer }] }
function submit(runner, account, form, answers, files, when) {
  const ss = account.ss(form.sheet);
  if (!ss) throw new Error('The sandbox has no response sheet for ' + form.title);
  const sh = ss.getSheetByName(form.tab) || ss.insertSheet(form.tab);
  const now = when || new Date();
  const p2 = (n) => String(n).padStart(2, '0');
  const record = {};
  const mst = new Date(now.getTime() - 7 * 3600000);   // Qualtrics's clock, whatever the account's zone
  record[form.timestampHeader || 'Timestamp'] = form.textTimestamp
    ? mst.getUTCFullYear() + '-' + p2(mst.getUTCMonth() + 1) + '-' + p2(mst.getUTCDate()) + ' ' + p2(mst.getUTCHours()) + ':' + p2(mst.getUTCMinutes()) + ':' + p2(mst.getUTCSeconds())
    : now;
  if (form.responseId) record['Response ID'] = 'R_' + Array.from({ length: 15 }, () => 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'[Math.floor(Math.random() * 62)]).join('');
  if (form.collectEmail) record['Email Address'] = answers['Email Address'] || account.user;
  form.questions.forEach((q) => {
    if (q.kind === 'file') {
      const list = (files && files[q.title]) || [];
      const folder = account.folderNamed(q.title + ' (File responses)', account.folderNamed(form.uploadFolder || form.title + ' (File responses)'));
      record[q.title] = list.map((f) => 'https://drive.google.com/open?id=' + account.addFile(f.name, f.mime, f.buffer, folder).id).join(', ');
      return;
    }
    const v = answers[q.title];
    record[q.title] = Array.isArray(v) ? v.join(', ') : (v === undefined ? '' : String(v));
  });

  // Write by header, adding a column for any question the tab has not seen, the way a
  // linked Google Form does.
  let lc = sh.getLastColumn();
  let headers = lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(String) : [];
  Object.keys(record).forEach((k) => {
    if (headers.indexOf(k) < 0) {
      if (lc + 1 > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
      sh.getRange(1, ++lc).setValue(k);
      headers.push(k);
    }
  });
  const row = headers.map((h) => (h in record ? record[h] : ''));
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  // Text answers are written as text, the way a form writes them (no date guessing).
  sh.getRange(r, 1, 1, row.length).setValues([row.map((v) => (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v) ? "'" + v : v))]);

  let trigger = null;
  if (form.trigger && form.trigger.style === 'change') {
    // A row added through the Sheets API (as Qualtrics does) fires the change trigger.
    const fired = runner.fire(ss.id, 'change', { changeType: 'INSERT_ROW', source: ss, authMode: 'FULL', triggerUid: 'sandbox-qualtrics' });
    trigger = fired[0] || null;
  } else if (form.trigger) {
    const range = sh.getRange(r, 1, 1, row.length);
    const values = row.map(displayValue);
    const namedValues = {};
    headers.forEach((h, i) => { namedValues[h] = [values[i]]; });
    const e = { range, values, namedValues, source: ss, authMode: 'FULL', triggerUid: 'sandbox-form' };
    trigger = runner.exec(form.trigger.project, form.trigger.fn, [e]);
  }
  return { row: r, trigger };
}

// A browser upload ("data:image/png;base64,...") as a file for submit().
function fileFromDataUrl(name, dataUrl) {
  const m = /^data:([^;]+);base64,(.*)$/.exec(String(dataUrl || ''));
  if (!m) return null;
  return { name: name || 'upload', mime: m[1], buffer: Buffer.from(m[2], 'base64') };
}

module.exports = { definitions, submit, fileFromDataUrl, WITHIN, ISSUE_TYPES, ACTIONS, ELL, APRONS, toSigned };
