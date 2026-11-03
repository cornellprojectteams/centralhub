'use strict';
// The fake world the sandbox starts from (and returns to on "Reset").
//
// Every person, email address, NetID and phone number here is made up, and every email
// address ends in sandbox.test, a domain that cannot receive mail. Team names are the
// real, public ones so the pages look familiar. Dates are worked out from the day of the
// reset, so an overdue issue is overdue whenever you reset.

const { Account } = require('./gas');
const images = require('./images');
const forms = require('./forms');

// ---------------------------------------------------------------- helpers

function rng(seed) {   // small seeded generator: the same reset gives the same data
  let a = seed >>> 0;
  return () => { a |= 0; a = (a + 0x6D2B79F5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}
function at(daysAgo, hour, minute) { const d = new Date(); d.setDate(d.getDate() - daysAgo); d.setHours(hour === undefined ? 10 : hour, minute || 0, 0, 0); return d; }
function day(daysAgo) { return at(daysAgo, 0, 0); }
const token = (r, n) => Array.from({ length: n || 32 }, () => '0123456789abcdef'[Math.floor(r() * 16)]).join('');
const slug = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, '.').replace(/^\.|\.$/g, '');
const email = (local) => local + '@sandbox.test';

function fill(sheet, rows) {
  const width = Math.max.apply(null, rows.map((r) => r.length));
  if (rows.length > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rows.length - sheet.getMaxRows() + 50);
  if (width > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  sheet.getRange(1, 1, rows.length, width).setValues(rows.map((r) => r.concat(Array(width - r.length).fill(''))));
  return sheet;
}
function header(sheet, cols) { sheet.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#8f1515').setFontColor('#ffffff'); sheet.setFrozenRows(1); }
function tab(ss, name, rows, opts) {
  opts = opts || {};
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const first = ss.getSheets()[0];
    if (first && first.getName() === 'Sheet1' && first.getLastRow() === 0 && name !== 'Sheet1') { first.setName(name); sh = first; if (opts.gid) sh.gid = opts.gid; }
    else sh = ss.insertSheet(name, undefined, opts.gid);
  }
  if (rows && rows.length) { fill(sh, rows); if (opts.header !== false) header(sh, rows[0].length); }
  return sh;
}

// ---------------------------------------------------------------- made-up people

const FIRST = ['Ana', 'Ben', 'Chloe', 'Dev', 'Elena', 'Femi', 'Grace', 'Hugo', 'Imani', 'Jonah', 'Kira', 'Leo', 'Maya', 'Nikhil', 'Olivia', 'Pablo', 'Quinn', 'Rosa', 'Sam', 'Tara', 'Uri', 'Vera', 'Wes', 'Xin', 'Yara', 'Zane', 'Amara', 'Bilal', 'Carmen', 'Diego', 'Esi', 'Farah', 'Gabe', 'Hana', 'Ivan', 'Jada', 'Kofi', 'Lina', 'Mateo', 'Nora', 'Omar', 'Priya', 'Reza', 'Sofia', 'Theo', 'Uma', 'Vik', 'Wren'];
const LAST = ['Brooks', 'Ortiz', 'Nakamura', 'Patel', 'Petrova', 'Adeyemi', 'Liu', 'Martins', 'Okafor', 'Weiss', 'Sato', 'Haddad', 'Fernandes', 'Rao', 'Chen', 'Reyes', 'Murphy', 'Alvarez', 'Kowalski', 'Singh', 'Cohen', 'Novak', 'Thompson', 'Zhao', 'Mansour', 'Carter', 'Diallo', 'Qureshi', 'Ruiz', 'Santos', 'Mensah', 'Khan', 'Lindqvist', 'Kim', 'Petrov', 'Brown', 'Asante', 'Moreau', 'Silva', 'Novak'];

function people(r, n) {
  const used = {};
  const out = [];
  while (out.length < n) {
    const f = FIRST[Math.floor(r() * FIRST.length)], l = LAST[Math.floor(r() * LAST.length)];
    if (used[f + l]) continue;
    used[f + l] = 1;
    const netid = (f[0] + l[0]).toLowerCase() + (100 + Math.floor(r() * 899));
    out.push({ first: f, last: l, name: f + ' ' + l, netid, email: email(netid) });
  }
  return out;
}

const TEAMS = ['AutoBoat', 'Baja', 'CUAir', 'Concrete Canoe', 'CU Sail', 'ChemE Car', 'CEV', 'Combat Robotics', 'CUAUV', 'CUBMD', 'DEBUT', 'Design Build Fly', 'ESW', 'EWB', 'EWH', 'Formula', 'Geodata', 'Hyperloop', 'iGEM', 'Mars Rover', 'Nexus', 'Rocketry', 'Seismic Design', 'Steel Bridge', 'Operations Team', 'Unknown / unsure'];

// ---------------------------------------------------------------- Space Status

const RESPONSE_HEADERS = ['Timestamp', 'Email Address', 'NetID', 'Responsible Team', 'Respond within', 'Photo Upload', 'Issue Type', 'Required action', 'Details / comments',
  'Notified at', 'Issue token', 'Addressed at', 'Last reminder', 'Completion photo', 'Completed at', 'Sent back reason', 'Completion note'];

// One line per issue. ago: days since the report. completed/addressed: days ago.
const ISSUES = [
  { team: 'Baja', within: '2 to 3 business days', type: 'Housekeeping', action: 'Clean up and organize the area', details: 'Metal shavings and oily rags under the mill bench in the Baja bay.', ago: 6, photo: 'bench', reminded: 1 },
  { team: 'Rocketry', within: 'Same day', type: 'Chemical storage', action: 'Store chemicals correctly', details: 'Two cans of epoxy hardener left open on the layup table overnight.', ago: 2, photo: 'cabinet', reminded: 1 },
  { team: 'iGEM', within: '14 to 21 days', type: 'Work Practices', action: 'Correct an unsafe work practice', details: 'Used gloves left on the bench by the centrifuge. Please remind members to bin them.', ago: 25, reminded: 2 },
  { team: 'ChemE Car', within: '2 to 3 business days', type: 'Chemical storage', action: 'Store chemicals correctly', details: 'Unlabeled squeeze bottles in the flammables cabinet.', ago: 8, photo: 'cabinet', reminded: 3 },
  { team: 'CUAir', within: '7 to 10 days', type: 'Egress / access', action: 'Clear the path or exit', details: 'Foam boxes stacked in front of the east door of the hangar.', ago: 3, photo: 'floor' },
  { team: 'Concrete Canoe', within: '14 to 21 days', type: 'Shared space', action: 'Remove or relocate items', details: 'Canoe molds are taking up the shared bench next to the paint booth.', ago: 5 },
  { team: 'Operations Team', within: '14 to 21 days', type: 'Maintenance / facilities', action: 'Submit a facilities work order', details: 'Flickering light over the Upson shop entrance.', ago: 4 },
  { team: 'Nexus', within: 'When time allows', type: 'Shared space', action: 'Other', details: 'A pegboard would keep the hand tools off the bench.', ago: 10 },
  { team: 'Mars Rover', within: '2 to 3 business days', type: 'Work Practices', action: 'Clean up and organize the area', details: 'Battery packs charging unattended on the floor.', ago: 2, photo: 'floor', completed: 1, after: 'After: packs on the charging shelf', note: 'Moved the packs to the charging shelf and put up a sign.' },
  { team: 'Steel Bridge', within: '7 to 10 days', type: 'Maintenance / facilities', action: 'Submit a facilities work order', details: 'Bandsaw blade guard is loose.', ago: 9, completed: 0, after: 'After: bandsaw guard tightened', note: 'Tightened the guard and checked the blade tension.' },
  { team: 'Hyperloop', within: '7 to 10 days', type: 'Housekeeping', action: 'Clean up and organize the area', details: 'Cardboard and packing foam piled next to the pod frame.', ago: 12, after: 'After: corner half cleared', note: 'Took the cardboard to recycling.', sentBack: 'The photo only shows half the corner. Please send one of the whole area, foam included.', reminded: 2 },
  { team: 'CU Sail', within: '2 to 3 business days', type: 'Housekeeping', action: 'Clean up and organize the area', details: 'Sail cloth scraps on the floor by the sewing table.', ago: 20, completed: 18, addressed: 17, after: 'After: sewing table swept' },
  { team: 'Baja', within: 'Same day', type: 'Fire code', action: 'Clear the path or exit', details: 'Extinguisher by the welding bay blocked by a parts cart.', ago: 30, completed: 30, addressed: 30, after: 'After: extinguisher clear' },
  { team: 'Design Build Fly', within: '7 to 10 days', type: 'Egress / access', action: 'Clear the path or exit', details: 'Wing ribs drying on the stairs to the mezzanine.', ago: 15, completed: 11, addressed: 10, after: 'After: stairs clear' },
  { team: 'EWB', within: '14 to 21 days', type: 'Shared space', action: 'Remove or relocate items', details: 'Water filter prototypes stored in the shared sink.', ago: 28, completed: 20, addressed: 19, after: 'After: sink cleared' },
];
// Filed through the sandbox form during the reset, so the real script emails the teams
// and those emails are waiting in the outbox.
const FRESH = [
  { team: 'Combat Robotics', within: '7 to 10 days', type: 'Housekeeping', action: 'Clean up and organize the area', details: 'Arena polycarbonate offcuts leaning against the wall.', hour: 8, photo: 'floor' },
  { team: 'AutoBoat', within: 'When time allows', type: 'Informational', action: 'No action needed (informational)', details: 'Nice work on the new tool shadow board. Just letting you know it was noticed.', hour: 9 },
  { team: 'Formula', within: 'Same day', type: 'Fire code', action: 'Clear the path or exit', details: 'Fuel cans stored under the stairwell by the Formula bay.', hour: 9, photo: 'cabinet' },
];

const PROJECTS = [
  { title: 'Rebuild the tool crib shelving', desc: 'Swap the sagging particleboard shelves for steel and label every bin.', status: 'In Progress', who: [0, 1], before: 'Before: sagging crib shelves', created: 8, started: 5 },
  { title: 'Label every cabinet in the Composites Lab', desc: 'One label per door saying what lives inside and who to ask.', status: 'Assigned', who: [], created: 2 },
  { title: 'Paint the mezzanine railing safety yellow', desc: 'Sand, prime and repaint the north railing.', status: 'Pending', who: [2], before: 'Before: chipped railing', after: 'After: railing painted', hours: 6, created: 14, started: 10 },
  { title: 'Set up a sharps disposal station', desc: 'Wall-mounted sharps box and a sign by the layup tables.', status: 'Completed', who: [3, 4], before: 'Before: blades on the bench', after: 'After: sharps station', hours: 3, created: 21, started: 18, done: 12 },
  { title: 'Reorganize the Auto Lab hardware drawers', desc: 'Sort fasteners by size and print drawer labels.', status: 'Completed', who: [5], before: 'Before: mixed hardware', after: 'After: sorted drawers', hours: 8.5, created: 30, started: 26, done: 20 },
  { title: 'Fix the sticking door to the Upson shop', desc: 'Plane the door edge or call Facilities if the frame has moved.', status: 'In Progress', who: [6], before: 'Before: door rubbing the frame', after: 'After: door planed', hours: 2, created: 11, started: 9, sentBack: 'It still sticks when it is humid. Can you check the top hinge too?' },
];
const PROPOSALS = [
  { key: 'eyewash', title: 'Add a second eyewash station near the paint booth', desc: 'The only one is across the room. Anyone at the paint booth has to cross the lab with paint in their eyes.', area: 'Composites Lab', who: 7, status: 'Real', created: 9, promoted: 3, photo: 'Spot by the paint booth' },
  { key: 'reels', title: 'Wall-mounted cord reels over the Auto Lab benches', desc: 'Extension cords across the floor are a trip hazard every build night.', area: 'Auto Lab', who: 8, status: 'Provisional', created: 6, photo: 'Cords across the Auto Lab floor' },
  { key: 'bins', title: 'Color-coded bins for scrap metal', desc: 'Aluminum, steel and brass all end up in the same bin, so nothing gets recycled.', area: 'Machine shop', who: 9, status: 'Provisional', created: 4 },
  { key: 'lights', title: 'Motion-sensor lights in the storage cage', desc: 'People leave the cage lights on overnight, and the switch is behind the door.', area: 'Storage cage', who: 10, status: 'Provisional', created: 1 },
  { key: 'whiteboard', title: 'Replace the broken whiteboard in the ELL conference room', desc: 'It ghosts so badly nobody can read it.', area: 'ELL', who: 11, status: 'Declined', created: 20 },
];
// Two reviewers who have rated three proposals (so their ratings count) and one who has not.
const VOTES = [
  ['A', 'eyewash', 5, 2], ['A', 'reels', 4, 3], ['A', 'bins', 3, 2],
  ['B', 'eyewash', 4, 3], ['B', 'reels', 5, 2], ['B', 'bins', 2, 3],
  ['C', 'reels', 5, 1],
];

function spaceStatus(account, ids, runner, r) {
  const ss = account.addSpreadsheet({ id: ids.space, name: 'Space Status Tracking' });
  const staff = people(r, 24);
  const uploads = account.folderNamed('Ops Hub completion and project uploads');
  const formFiles = account.folderNamed('Photo Upload (File responses)', account.folderNamed('Space Status Report (File responses)'));
  const pic = (caption, scene, folder) => account.addFile(caption.replace(/[^\w]+/g, '-').toLowerCase() + '.svg', 'image/svg+xml', images.photo(caption, scene), folder || uploads).id;
  const view = (id) => 'https://drive.google.com/file/d/' + id + '/view';

  // Team Contacts: one row per team on the form. Operations Team mail goes to an admin.
  const contacts = [['Team', 'Liaison name', 'Liaison email', 'Team general email', 'Safety lead name', 'Safety lead email', 'Extra CC', 'Team token']];
  TEAMS.forEach((t, i) => {
    if (/unknown/i.test(t)) { contacts.push([t, '', '', '', '', '', '', token(r, 16)]); return; }
    const liaison = staff[i % staff.length], safety = staff[(i + 7) % staff.length];
    const s = slug(t);
    contacts.push([t, liaison.name, email(s + '.liaison'), email(s), safety.name, email(s + '.safety'), '', token(r, 16)]);
  });
  tab(ss, 'Form Responses', [RESPONSE_HEADERS]);
  tab(ss, 'Team Contacts', contacts);

  const rows = ISSUES.map((it, i) => {
    const reporter = staff[(i * 5) % staff.length];
    const ts = at(it.ago, 9 + (i % 7), (i * 13) % 60);
    const photo = it.photo ? 'https://drive.google.com/open?id=' + pic(it.team + ': ' + it.details.split(/[.,]/)[0], it.photo, formFiles) : '';
    const notified = new Date(ts.getTime() + 60000);
    const after = it.after ? view(pic(it.after, it.photo || 'bench')) : '';
    return [ts, reporter.email, reporter.netid, it.team, it.within, photo, it.type, it.action, it.details,
      notified, token(r), it.addressed !== undefined ? at(it.addressed, 15) : '', it.reminded ? at(it.reminded, 7) : '',
      after, it.completed !== undefined ? at(it.completed, 16) : '', it.sentBack || '', it.note || ''];
  }).sort((a, b) => a[0] - b[0]);
  fill(ss.getSheetByName('Form Responses'), [RESPONSE_HEADERS].concat(rows));

  // Projects
  const PH = ['Project ID', 'Title', 'Description', 'Status', 'Assignees', 'Before photo', 'After photo', 'Hours', 'Started at', 'Completed at', 'Created at', 'Sent back reason', 'Attachment', 'Link', 'Proposal ID', 'Area'];
  const projectRows = PROJECTS.map((p) => [token(r), p.title, p.desc, p.status, p.who.map((i) => staff[i].name).join(', '),
    p.before ? view(pic(p.before, 'shelves')) : '', p.after ? view(pic(p.after, 'shelves')) : '', p.hours || '',
    p.started !== undefined ? at(p.started, 14) : '', p.done !== undefined ? at(p.done, 17) : '', at(p.created, 11), p.sentBack || '', '', '', '', '']);

  // Proposals and ratings
  const PRH = ['Proposal ID', 'Title', 'Description', 'Area', 'Proposed by', 'NetID', 'Status', 'Created at', 'Promoted at', 'Photos'];
  const VH = ['Vote ID', 'Proposal ID', 'Voter NetID', 'Impact', 'Effort', 'Comment', 'Voted at'];
  const pid = {};
  const proposalRows = PROPOSALS.map((p) => {
    pid[p.key] = token(r);
    return [pid[p.key], p.title, p.desc, p.area, staff[p.who].name, staff[p.who].netid, p.status, at(p.created, 13), p.promoted !== undefined ? at(p.promoted, 18) : '', p.photo ? pic(p.photo, 'floor') : ''];
  });
  const voters = { A: staff[12], B: staff[13], C: staff[14] };
  const voteRows = VOTES.map((v, i) => [token(r), pid[v[1]], voters[v[0]].netid, v[2], v[3], '', at(8 - i % 5, 12)]);
  // The promoted proposal is already a project, as the script would have made it.
  const eyewash = PROPOSALS[0];
  projectRows.push([token(r), eyewash.title, eyewash.desc, 'Assigned', '', '', '', '', '', '', at(eyewash.promoted, 18), '', '', '', pid.eyewash, eyewash.area]);
  tab(ss, 'Projects', [PH].concat(projectRows));
  tab(ss, 'Proposals', [PRH].concat(proposalRows));
  tab(ss, 'Proposal votes', [VH].concat(voteRows));

  // Three reports filed through the form during the reset: the real script sends the
  // team emails and stamps the rows.
  const form = forms.definitions(account, ids)['space-status'];
  FRESH.forEach((it, i) => {
    const reporter = staff[(i * 3 + 2) % staff.length];
    const files = it.photo ? { 'Photo Upload': [{ name: slug(it.team) + '.svg', mime: 'image/svg+xml', buffer: images.photo(it.team + ': ' + it.details.split(/[.,]/)[0], it.photo) }] } : {};
    let when = at(0, it.hour, 10 + i * 7);
    if (when > new Date()) when = new Date(Date.now() - (FRESH.length - i) * 25 * 60000);   // reset before 9am: earlier today
    const res = forms.submit(runner, account, form, { 'Email Address': reporter.email, NetID: reporter.netid, 'Responsible Team': it.team, 'Respond within': it.within, 'Issue Type': it.type, 'Required action': it.action, 'Details / comments': it.details }, files, when);
    if (res.trigger && !res.trigger.ok) throw new Error('The Space Status script failed on a test report: ' + res.trigger.error);
  });
  return staff;
}

// ---------------------------------------------------------------- inventory

const INVENTORY_PHOTOS = { 'drill': 'tool', 'impact': 'tool', 'soldering': 'bench', 'band saw': 'bench', 'mill': 'bench', 'lathe': 'bench', 'grinder': 'tool', 'router': 'tool', 'socket': 'shelves', 'gloves': 'shelves', 'glasses': 'shelves', 'extinguisher': 'cabinet', 'eyewash': 'cabinet' };

function inventory(account, ids, runner) {
  account.addSpreadsheet({ id: ids.inventory, name: 'Untitled spreadsheet' });
  const setup = runner.exec('inventory', 'setupInventory', []);
  if (!setup.ok) throw new Error('setupInventory failed in the sandbox: ' + setup.error);
  const data = runner.exec('inventory', 'addTestData', []);
  if (!data.ok) throw new Error('addTestData failed in the sandbox: ' + data.error);
  // Photos for some of the sample tools, as if someone had added them in the app.
  const ss = account.ss(ids.inventory);
  const sh = ss.getSheetByName('Inventory Master');
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cName = head.indexOf('Item Name') + 1, cPhoto = head.indexOf('Photo URL') + 1;
  const folder = account.folderNamed('Inventory photos');
  if (cName && cPhoto) {
    for (let row = 2; row <= sh.getLastRow(); row++) {
      const name = String(sh.getRange(row, cName).getValue() || '');
      const key = Object.keys(INVENTORY_PHOTOS).find((k) => name.toLowerCase().indexOf(k) >= 0);
      if (!key) continue;
      const id = account.addFile(slug(name) + '.svg', 'image/svg+xml', images.photo(name, INVENTORY_PHOTOS[key]), folder).id;
      sh.getRange(row, cPhoto).setValue('https://drive.google.com/file/d/' + id + '/view');
    }
  }
  // The tabs keep the gids the hub's links use (config.js INVENTORY.*Gid).
  const want = ids.inventoryGids || {};
  [['Dashboard', want.dash], ['Inventory Master', want.master], ['Maintenance Log', want.maint]].forEach(([name, gid]) => {
    const t = ss.getSheetByName(name);
    if (t && gid !== undefined && gid !== '' && !ss.getSheetById(gid)) t.gid = Number(gid);
  });
  account.projectState('inventory').cache = {};
}

// ---------------------------------------------------------------- ELL shift log

function ell(account, ids, r, staff) {
  const ss = account.addSpreadsheet({ id: ids.ell, name: 'ELL Shift Summary (Responses)' });
  const E = forms.ELL;
  const H = ['Timestamp', 'Email Address', 'NetID', 'Shift', 'Overall Activity Level and Space Use', 'Main ELL use (check all that apply)', 'Composites Lab Use (check all that apply)', 'Auto Lab Use (check all that apply)', 'Conference Room Use (check all that apply)', 'Supply Status (mark if we need to purchase more)', 'Safety and Maintenance Issues (check all that apply)', 'Most active teams today', 'Productivity notes (progress on assigned tasks)', 'Additional Notes', 'photos'];
  const pick = (list, n) => { const out = []; while (out.length < n) { const x = list[Math.floor(r() * list.length)]; if (out.indexOf(x) < 0) out.push(x); } return out.join(', '); };
  const rows = [];
  const shifts = [['6-8', 19], ['8-10', 21], ['2-4', 15]];
  for (let d = 14; d >= 1; d--) {
    shifts.forEach(([shift, hour], k) => {
      if (k === 2 && d % 3) return;
      const who = staff[(d + k) % staff.length];
      const latest = d === 1 && shift === '8-10';
      rows.push([at(d, hour, 40), who.email, who.netid, shift, latest ? 'High - 60-80%' : E.activity[Math.floor(r() * 5)],
        pick(E.main.slice(0, 5), 2), pick(E.composites.slice(0, 6), 1), pick(E.auto.slice(0, 3), 1), pick(E.conference.slice(0, 3), 1),
        latest ? 'Gloves, N95 masks' : (r() < 0.8 ? 'All supplies stocked' : pick(E.supply.slice(0, 6), 1)),
        latest ? 'Blocked Exits/ Extinguishers' : (r() < 0.85 ? 'No issues observed' : pick(E.safety.slice(0, 4), 1)),
        pick(E.teams, 3), r() < 0.5 ? 'Restocked the glove dispensers.' : '', '', '']);
    });
  }
  tab(ss, 'Form Responses 1', [H].concat(rows));
  const byDay = {};
  rows.forEach((row) => { const k = (row[0].getMonth() + 1) + '/' + row[0].getDate(); byDay[k] = (byDay[k] || 0) + 1; });
  tab(ss, 'Sheet1', [['Date', 'Shifts logged']].concat(Object.keys(byDay).map((k) => [k, byDay[k]])));
}

// ---------------------------------------------------------------- Fabman (sheet and fake API)

const AGREEMENT = 'I acknowledge that I have read, understand, and agree to adhere to the Fabman and Machine Shop usage policies outlined above.';
const FABMAN = { SPACES: 2946, LAB: 2947, MACHINES: 2948, MILLS: 2949, SUPERVISOR: 2950, STANDARD_PACKAGE: 11773 };

function fabman(account, ids, r) {
  const year = new Date().getFullYear();
  const members = people(r, 42).map((p, i) => {
    const grad = [year - 1, year, year, year + 1, year + 1, year + 2, year + 2, year + 3][i % 8];
    const apron = i % 4 === 3 ? '' : ['Red Apron', 'Green Apron', 'Blue Apron'][i % 3];
    const team = TEAMS[(i * 3) % 24];
    return {
      // NetID@cornell.edu, like real Fabman accounts, since the Fabman script finds people by
      // that address. Nothing is ever sent to these: the fake Fabman only records invitations.
      id: 500001 + i, account: 2366, firstName: p.first, lastName: p.last, company: team, emailAddress: p.netid + '@cornell.edu',
      state: i === 2 || i === 9 || i === 17 ? 'locked' : 'active', lockVersion: 1 + (i % 3),
      metadata: { graduationYear: grad, apron, netid: p.netid }, createdAt: at(200 - i * 3, 12).toISOString(),
    };
  });
  const trainings = {}, packages = {};
  let tid = 800001;
  members.forEach((m, i) => {
    const list = [];
    const add = (course) => list.push({ id: tid++, trainingCourse: course, fromDate: at(150 - i, 9).toISOString().slice(0, 10) });
    if (m.metadata.apron) { add(FABMAN.SPACES); add(FABMAN.LAB); }
    if (m.metadata.apron === 'Red Apron') { add(FABMAN.MACHINES); add(FABMAN.MILLS); }
    if (i % 11 === 5) add(FABMAN.SUPERVISOR);
    trainings[m.id] = list;
    packages[m.id] = [{ id: tid++, package: FABMAN.STANDARD_PACKAGE, fromDate: '2026-01-15' }];
  });
  account.fabman.members = members;
  account.fabman.trainings = trainings;
  account.fabman.packages = packages;
  account.fabman.nextId = 900001;

  const ss = account.addSpreadsheet({ id: ids.fabman, name: 'Project Teams Fabman Admin' });
  const has = (m, c) => trainings[m.id].some((t) => t.trainingCourse === c);
  const DH = ['ID', 'First Name', 'Last Name', 'Team', 'Graduation Year', 'NetID', 'Email', 'State', 'Apron', 'Spaces', 'Lab Equipment', 'Machines', 'Mills', 'Supervisor'];
  tab(ss, 'Dashboard', [DH].concat(members.map((m) => [m.id, m.firstName, m.lastName, m.company, m.metadata.graduationYear, m.metadata.netid, m.emailAddress, m.state, m.metadata.apron,
    has(m, FABMAN.SPACES), has(m, FABMAN.LAB), has(m, FABMAN.MACHINES), has(m, FABMAN.MILLS), has(m, FABMAN.SUPERVISOR)])));
  tab(ss, 'Actions', [['Email'], [members[5].emailAddress], [members[6].emailAddress]]);
  // The FabMan Agreement form's responses, with the real form's question titles. The
  // older rows are people who got accounts. The last few came in after the form trigger
  // stopped, so they have none yet: "Check sign-ups" in the sheet's Fabman sign-ups
  // menu finds them, and "Process sign-ups since a date" catches them up.
  const SH = ['Timestamp', AGREEMENT, 'First Name', 'Last Name', 'Project Team', 'Graduation Year', 'NetID', 'Previously Completed Training (Emerson Apron Status)'];
  const answered = members.slice(30).map((m, i) => [at(40 + i * 6, 15), 'I agree', m.firstName, m.lastName, m.company, m.metadata.graduationYear, m.metadata.netid + (i % 2 ? '@cornell.edu' : ''), m.metadata.apron || 'None']);
  const locked = members.find((m) => m.state === 'locked');
  const taken = new Set(members.map((m) => m.metadata.netid));
  const waiting = people(r, 12).filter((p) => !taken.has(p.netid)).slice(0, 6).map((p, i) => [at([18, 12, 9, 6, 3, 1][i], 10 + i, 7 * i), 'I agree', p.first, p.last, TEAMS[(i * 5 + 2) % 23], [year + 1, year + 2, 'Dec ' + (year + 1), year + 3, year + 1, year + 2][i],
    [p.netid, p.netid.toUpperCase() + '@Cornell.edu', p.first.toLowerCase() + '.' + p.last.toLowerCase() + '@gmail.com', p.netid, p.netid + '@cornell.edu', p.netid][i],
    ['None', 'Red Apron', 'None', 'Green Apron', 'None', 'Blue Apron'][i]]);
  waiting.splice(3, 0, [at(5, 16, 40), 'I agree', locked.firstName, locked.lastName, locked.company, locked.metadata.graduationYear, locked.metadata.netid, locked.metadata.apron || 'None']);
  tab(ss, 'Sign-up Form', [SH].concat(answered, waiting));
  // Where the Flag graduates and Process graduates buttons live. Flag graduates fills it.
  const grads = tab(ss, 'Graduates', [['', 'Press Flag graduates to list everyone whose graduation year has passed. Nothing changes in Fabman until Process graduates.'],
    ['Approve', 'Action', 'ID', 'Name', 'Team', 'Grad Year', 'NetID', 'Email', 'State', 'Access', 'Result']], { header: false });
  grads.getRange(2, 1, 1, 11).setFontWeight('bold');
  // A FABMAN_API_KEY script property replaces the key written in the script, so the
  // sandbox sends the fake Fabman a made-up key (it takes any). SIGNUPS_FROM is when the
  // sign-up triggers were set up: the hourly run only looks at sign-ups from then on.
  account.projectState('fabman').props = { FABMAN_API_KEY: 'sandbox-key-any-text-works', SIGNUPS_FROM: new Date().toISOString(), ALERT_EMAILS: email('ops.admin') };
  account.projectState('fabman').triggers = [
    { fn: 'onSheetChange', type: 'change', id: 't-fab-0', source: ids.fabman },
    { fn: 'onFormSubmit', type: 'submit', id: 't-fab-1', source: ids.fabman },
    { fn: 'processNewSignups', type: 'time', every: '1 hours', id: 't-fab-2', source: ids.fabman },
  ];
  // The tab the Qualtrics survey's workflow writes to, as setupSignups makes it. Empty
  // until someone signs the Qualtrics agreement under "Make something happen".
  tab(ss, 'FabMan Agreement', [['Recorded Date', 'Response ID', 'NetID', 'First Name', 'Last Name', 'Project Team', 'Graduation Year', 'Apron', 'Attestation', 'Fabman status']]);
}

// ---------------------------------------------------------------- Master Sheet (compliance)

function master(account, ids, r) {
  const ss = account.addSpreadsheet({ id: ids.master, name: 'Master Sheet - Building Access and Workday Learning' });
  tab(ss, 'Read Me', [['Compilation of student usage agreement waiver and training completions from Qualtrics surveys embedded in the Safety & Training Hub Canvas course (sandbox copy, made-up names).']], { header: false });
  const H = ['DATE', 'TIME', 'NAME', 'netID', 'CUID', 'Grad Year', 'TEAM'];
  const year = new Date().getFullYear();
  const log = (n) => people(r, n).map((p, i) => [day(3 + i * 9), '2:' + String(10 + i).padStart(2, '0') + ' PM', p.name, p.netid, '', year + (i % 4), TEAMS[(i * 5) % 24]]);
  [['ELL', 18], ['HVL', 9], ['RHODES PENTHOUSE', 6], ['FabMan and Machine Shop', 2], ['Aquatic Center', 5], ['Risk Waivers', 7], ['Private Vehicle Risk Waiver', 4], ['Canoe waiver', 3], ['Hollister room B55', 0], ['Tang Room 403', 3]]
    .forEach(([name, n]) => tab(ss, name, [H].concat(log(n))));
  const courses = ['Chemical Waste (EHS 2716)', 'Respiratory Protection (EHS 2381)', 'Hot Works (EHS 2398)', 'Compressed Gas Safety (EHS 2335)', 'Electrical Safe Work Practices (EHS 4336)'];
  tab(ss, 'Workday Trainings', [['Name', 'Team', 'Completed courses']].concat(people(r, 10).map((p, i) => [p.name, TEAMS[(i * 7) % 24], courses.slice(0, 1 + (i % 4)).join(', ')])));
  tab(ss, 'Workday Training Backend', [['Learner', 'Course', 'Completed on']]);
  tab(ss, 'Combined', [['Location', 'DATE', 'NAME', 'netID', 'TEAM']]);
}

// ---------------------------------------------------------------- Fleet (student drivers)

function drivers(account, ids, r) {
  const year = new Date().getFullYear();
  const pool = people(r, 26);
  const ss = account.addSpreadsheet({ id: ids.drivers, name: 'Project Teams Student Drivers', editors: [] });
  // Approval: row 1 is instructions, row 2 the headers (column P "requested" = filed).
  const AH = ['Name', 'NetID', 'Grad Year', 'Team', 'Justification', 'Driver safety (RMI2100)', 'Fleet application', 'Sign release', 'Trailer safety', 'Trailer test', 'Email', 'Phone', 'Approved by', 'Date approved', 'Notes', 'requested'];
  const approval = pool.map((p, i) => {
    const grad = [year - 1, year, year, year + 1, year + 2, year + 1, year + 3, year + 2][i % 8];
    const kind = i % 6;   // 0-1 filed, 2 ready, 3 blocked, 4 missing info, 5 filed
    const ini = (p.first[0] + p.last[0]).toUpperCase();
    return [p.name, p.netid, kind === 4 && i % 2 ? '' : grad, kind === 4 && !(i % 2) ? '' : TEAMS[(i * 3) % 24], 'Trailer runs to competition and material pickups',
      kind === 3 ? '' : ini, ini, kind === 3 && i % 2 ? '' : ini, i % 3 ? ini : '', i % 3 ? ini : '', p.email, '', kind === 2 || kind === 3 || kind === 4 ? '' : 'Ops', '', '', kind === 2 || kind === 3 || kind === 4 ? '' : 'yes'];
  });
  tab(ss, 'Approval', [['>> Complete all fields (application not processed with missing info)   >> Initial each upon completion (not processed without safety, fleet, and release)'], AH].concat(approval), { header: false });
  ss.getSheetByName('Approval').getRange(2, 1, 1, AH.length).setFontWeight('bold');
  // Log: row 1 is a banner, the headers are on row 3.
  const LH = ['Team', 'Name', 'NetID', 'Phone Number', 'Date (doubleclick)', 'Checkout Time', 'Return Time', 'Destination', 'Parking Location'];
  const trip = (i, daysAhead, start, end, dest) => { const p = pool[i]; return [TEAMS[(i * 3) % 24], p.name, p.netid, '', day(-daysAhead), start, end, dest, 'Upson lot']; };
  const log = [
    trip(1, -12, '9:00 AM', '1:00 PM', 'Home Depot'), trip(4, -9, '2:00 PM', '5:00 PM', 'Game Farm Road test track'), trip(6, -6, '8:00 AM', '6:00 PM', 'Syracuse airport freight'),
    trip(8, -3, '10:00 AM', '12:00 PM', 'Lowe\'s'), trip(9, -1, '3:00 PM', '6:00 PM', 'Cayuga Lake launch'),
    trip(2, 0, '9:00 AM', '1:00 PM', 'Metal supplier, Cortland'), trip(5, 0, '12:00 PM', '4:00 PM', 'Competition shipping depot'),
    trip(10, 2, '1:00 PM', '3:00 PM', 'Tractor Supply'), trip(11, 4, '9:00 AM', '', ''), trip(12, 6, '7:00 AM', '7:00 PM', 'Watkins Glen test day'),
    trip(13, 30, '8:00 AM', '8:00 PM', 'Regional competition'),
  ];
  tab(ss, 'Log', [['Project Teams Vehicle Reservations Toyota Tacoma (MBN-4258). Double-click a date cell to pick a date.'], [''], LH].concat(log), { header: false });
  ss.getSheetByName('Log').getRange(3, 1, 1, LH.length).setFontWeight('bold');
  tab(ss, 'Reference', [['Team', 'Abbreviation']].concat(TEAMS.slice(0, 24).map((t) => [t, t.replace(/[^A-Z]/g, '') || t.slice(0, 3).toUpperCase()])));
  tab(ss, 'Process', [['Step', 'What to do'], [1, 'Applicant fills in the Approval tab and initials each training as it is completed.'], [2, 'Check the row is complete, then email it to Fleet Services for approval.'], [3, 'Put "yes" in column P (requested) once it is filed.'], [4, 'Fleet confirms; the driver can book the Tacoma on the Log tab.']]);
  // Drivers were shared on the sheet as editors so they could book.
  pool.slice(0, 16).forEach((p) => ss.addEditor(p.email));
}

// ---------------------------------------------------------------- the old tool sheet (mill room orders, locks)

function oldTools(account, ids, r) {
  const ss = account.addSpreadsheet({ id: ids.oldInventory, name: 'Tool inventory, purchasing and locks' });
  tab(ss, 'Team tool inventory', [['Item', 'Team', 'Qty', 'Location'], ['Cordless drill', 'Baja', 2, 'Upson bay 3'], ['Jigsaw', 'Formula', 1, 'Upson bay 1'], ['Router', 'Rocketry', 1, 'Composites Lab'], ['Socket set', 'CUAir', 1, 'Hangar']]);
  const MH = ['Equipment', 'Supplier link', 'Priority', 'P/N', 'Price', 'Units', 'Ordered'];
  tab(ss, 'Mill room Tooling', [MH,
    ['1/2" carbide end mill', 'https://example.com/endmill', 'High', 'EM-500', 38.5, 4, 'yes'], ['3/8" carbide end mill', 'https://example.com/endmill38', 'High', 'EM-375', 31, 4, ''],
    ['Collet set ER32', 'https://example.com/collets', 'Medium', 'ER32-SET', 129, 1, 'yes'], ['Parallels set', 'https://example.com/parallels', 'Medium', 'PAR-6', 64, 1, ''],
    ['Edge finder', 'https://example.com/edgefinder', 'Low', 'EF-02', 18, 2, 'yes'], ['Center drill set', 'https://example.com/centerdrills', 'Medium', 'CD-5', 22, 1, ''],
    ['Fly cutter', 'https://example.com/flycutter', 'Low', 'FC-3', 75, 1, 'yes'], ['Way oil (1 gal)', 'https://example.com/wayoil', 'High', 'WO-1', 29, 2, ''],
    ['Chip brush (pack)', 'https://example.com/brushes', 'Low', 'CB-12', 12, 3, 'yes'], ['Tapping fluid', 'https://example.com/tapfluid', 'Medium', 'TF-16', 15, 2, '']]);
  tab(ss, 'Facility non recurring', [['Item', 'Cost', 'Bought'], ['Cord reel, 50 ft', 89, 'yes'], ['Shop vac', 249, ''], ['Stackable totes (10)', 120, 'yes']]);
  const holders = people(r, 8);
  const LH = ['Team', 'Equipment', 'Lock #', 'Key assigned to', 'NetID', 'Date out', 'Returned'];
  tab(ss, 'Locks', [LH].concat(holders.map((p, i) => [TEAMS[(i * 4) % 24], ['Welder cage', 'Paint booth', 'Battery cabinet', 'Tool crib', 'Bay 2 locker'][i % 5], 100 + i, p.name, p.netid, day(40 - i * 4), i % 3 === 1 ? '' : day(20 - i * 2)])));
}

// ---------------------------------------------------------------- workspaces, summer contacts, staff directory

function workspaces(account, ids, r) {
  const ss = account.addSpreadsheet({ id: ids.workspaces, name: 'Project Teams Workspaces' });
  const H = ['Team', 'Space Category', 'Region/Building', 'Location', 'Room/area', 'SPT Access Form', 'Activity Waiver', 'Usage Agreement', 'Safety plan Required', 'UG last updated', 'Contact', 'Notes'];
  const cats = ['Shop', 'Lab', 'Storage', 'Fabrication', 'Bench Work', 'Computing'];
  const places = [['Upson', 'Upson Auto Lab'], ['Rhodes', 'Rhodes penthouse'], ['ELL', 'Experiential Learning Lab'], ['Tang', 'Tang Hall 403'], ['Hollister', 'Hollister B55']];
  const contacts = people(r, 16);
  tab(ss, 'Team Spaces Master List', [H].concat(TEAMS.slice(0, 16).map((t, i) => { const pl = places[i % places.length]; return [t, cats[i % cats.length], pl[0], pl[1], 'Bay ' + (i + 1), 'yes', i % 3 ? 'yes' : '', 'yes', i % 4 === 0 ? 'Yes' : 'No', day(30 + i * 9), contacts[i].name, '']; })));
  tab(ss, 'Renewal survey data', [['Team', 'Still need the space?', 'Changes requested'], ['Baja', 'Yes', 'Second bay for the frame jig'], ['CUAir', 'Yes', ''], ['Nexus', 'No', 'Moving to shared bench space']]);
}

function summer(account, ids, r) {
  const ss = account.addSpreadsheet({ id: ids.summer, name: 'Summer Contact List ' + new Date().getFullYear() });
  const p = people(r, 14);
  tab(ss, 'Sheet1', [['Team', 'Summer contact', 'Email', 'On campus?']].concat(TEAMS.slice(0, 14).map((t, i) => [t, p[i].name, p[i].email, i % 3 ? 'Yes' : 'No'])));
}

const ROLES = ['', '', 'Ops', 'Ops-fab', 'Ops-chem', 'Prj', 'Prj-composites', 'Apps- dev', 'Digi-coms', 'Lead', 'Coord', 'Ops-canvas'];
function staffLog(account, ids, r) {
  const ss = account.addSpreadsheet({ id: ids.people, name: 'Student Employee Log' });
  const p = people(r, 12);
  const resources = ['Space Status, scheduling', 'Fabman access, aprons', 'Shop safety, tooling', 'Composites Lab', 'Chemical storage', 'Projects and proposals', 'Composites layup', 'The hub and its scripts', 'Newsletter, socials', 'Shift leads', 'Purchasing, events', 'Canvas trainings'];
  const H = ['Name', 'Phone', 'NetID', 'CU ID', 'Email', 'Pronouns', 'Year', 'Major', 'Start', 'Hours', 'Shift pref', 'Notes', 'Can be a Resource for…', 'Supervisor', 'Roles'];
  tab(ss, 'Contact Info', [H].concat(p.map((x, i) => [x.name, '', x.netid, '', x.email, '', '', '', '', '', '', '', resources[i], '', ROLES[i]])), { gid: ids.peopleGid });
  // The schedule tab's first row is Monday's first shift (the hub reads it as the header).
  const S = [['Monday', 'Shift 4-6pm', 'Employee ' + p[2].first + ' + ' + p[3].first]];
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const slots = ['10-12pm', '12-2pm', '2-4pm', '4-6pm', '6-8pm', '8-10pm'];
  days.forEach((d, di) => slots.forEach((s, si) => {
    if (di === 0 && si <= 3) return;
    if (di >= 5 && si < 2) return;
    const a = p[(di * 3 + si) % 12], b = p[(di * 5 + si + 4) % 12];
    S.push([si === 0 || (di === 0 && si === 4) ? d : '', s, si % 2 ? a.first + ' + ' + b.first : a.first]);
  }));
  tab(ss, 'FA25 Schedule', S, { gid: ids.scheduleGid, header: false });
}

// ---------------------------------------------------------------- build it all

function build(ids, makeRunner) {
  const account = new Account();
  const runner = makeRunner(account);
  const r = rng(20261021);
  const staff = spaceStatus(account, ids, runner, r);
  inventory(account, ids, runner);
  ell(account, ids, r, staff);
  fabman(account, ids, r);
  master(account, ids, r);
  drivers(account, ids, r);
  oldTools(account, ids, r);
  workspaces(account, ids, r);
  summer(account, ids, r);
  staffLog(account, ids, r);
  // A fresh start: setup logs and the reset's own runs are not "problems", and every
  // spreadsheet opens on its first tab.
  Object.values(account.state).forEach((s) => { s.logs = []; });
  Object.values(account.spreadsheets).forEach((ss) => { ss.active = ss.sheets[0]; });
  runner.problems.length = 0;
  account.builtAt = new Date().toISOString();
  return { account, runner };
}

module.exports = { build, TEAMS };
