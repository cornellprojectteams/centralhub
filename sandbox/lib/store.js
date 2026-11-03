'use strict';
// Saves the sandbox's fake Google account to sandbox/.state/state.json, so sheets,
// photos and emails survive a restart. Delete that folder, or press Reset in the
// control panel, to start over.

const fs = require('fs');
const path = require('path');
const { Account, Spreadsheet, Sheet, Range, DFolder, DFile, FakeFabman } = require('./gas');

const VERSION = 1;

const enc = (v) => (v instanceof Date ? { $d: v.toISOString() } : v);
const dec = (v) => (v && typeof v === 'object' && v.$d ? new Date(v.$d) : v);

function toJSON(account) {
  return {
    version: VERSION, builtAt: account.builtAt || null, savedAt: new Date().toISOString(), user: account.user,
    spreadsheets: Object.values(account.spreadsheets).map((ss) => ({
      id: ss.id, title: ss.title, owner: ss.owner, editors: ss.editors, note: ss.note,
      active: ss.sheets.indexOf(ss.active),
      sheets: ss.sheets.map((sh) => ({
        name: sh.name, gid: sh.gid, maxRows: sh.maxRows, maxCols: sh.maxCols, hidden: sh.hidden, frozenRows: sh.frozenRows, frozenCols: sh.frozenCols,
        widths: sh.widths, rules: sh.rules, formUrl: sh.formUrl, tab: sh.tab || null,
        protections: sh.protections.map((p) => ({ type: p.type, desc: p.desc, warn: p.warn, range: p.range ? [p.range.r, p.range.c, p.range.nr, p.range.nc] : null })),
        cells: Array.from(sh.cells.entries()).filter(([, cl]) => sh.hasContent(cl) || Object.keys(cl).length > 2).map(([k, cl]) => {
          const out = { k };
          Object.keys(cl).forEach((p) => { if (cl[p] !== '' && cl[p] !== undefined) out[p] = p === 'v' ? enc(cl[p]) : cl[p]; });
          return out;
        }),
      })),
    })),
    drive: {
      root: account.root.id,
      folders: Object.values(account.drive.folders).map((f) => ({ id: f.id, name: f.name, parentId: f.parentId, trashed: f.trashed, access: f.access, perm: f.perm })),
      files: Object.values(account.drive.files).map((f) => ({ id: f.id, name: f.name, mime: f.mime, parentId: f.parentId, trashed: f.trashed, desc: f.desc, access: f.access, created: enc(f.created), data: Buffer.from(f.bytes.map((b) => b & 255)).toString('base64') })),
    },
    outbox: account.outbox,
    fabman: account.fabman.toJSON(),
    state: Object.fromEntries(Object.entries(account.state).map(([k, s]) => [k, { props: s.props, triggers: s.triggers, logs: s.logs.slice(-200), cache: s.cache }])),
  };
}

function fromJSON(data) {
  if (!data || data.version !== VERSION) return null;
  const account = new Account();
  account.user = data.user || account.user;
  account.builtAt = data.builtAt;
  account.drive = { folders: {}, files: {} };
  data.drive.folders.forEach((f) => { const d = new DFolder(account, f.name, f.parentId, f.id); Object.assign(d, { trashed: f.trashed, access: f.access, perm: f.perm }); });
  account.root = account.drive.folders[data.drive.root] || new DFolder(account, 'My Drive', null);
  data.drive.files.forEach((f) => {
    const bytes = Array.from(Buffer.from(f.data, 'base64')).map((b) => (b > 127 ? b - 256 : b));
    const file = new DFile(account, f.name, f.mime, bytes, f.parentId, f.id);
    Object.assign(file, { trashed: f.trashed, desc: f.desc || '', access: f.access, created: dec(f.created) || new Date() });
  });
  data.spreadsheets.forEach((s) => {
    const ss = new Spreadsheet(account, { id: s.id, name: s.title, editors: s.editors, owner: s.owner, note: s.note, empty: true });
    s.sheets.forEach((t) => {
      const sh = new Sheet(ss, t.name, t.gid);
      Object.assign(sh, { maxRows: t.maxRows, maxCols: t.maxCols, hidden: t.hidden, frozenRows: t.frozenRows, frozenCols: t.frozenCols, widths: t.widths || {}, rules: t.rules || [], formUrl: t.formUrl, tab: t.tab || undefined });
      t.cells.forEach((c) => { const cl = { v: '', f: '' }; Object.keys(c).forEach((p) => { if (p !== 'k') cl[p] = p === 'v' ? dec(c[p]) : c[p]; }); sh.cells.set(c.k, cl); });
      t.protections.forEach((p) => {
        const pr = p.range ? sh.getRange(p.range[0], p.range[1], p.range[2], p.range[3]).protect() : sh.protect();
        pr.setDescription(p.desc || '').setWarningOnly(!!p.warn);
      });
      ss.sheets.push(sh);
    });
    ss.active = ss.sheets[s.active] || ss.sheets[0];
    account.spreadsheets[ss.id] = ss;
  });
  account.outbox = data.outbox || [];
  account.fabman = new FakeFabman(data.fabman);
  account.state = data.state || {};
  account.dirty = false;
  return account;
}

function load(file) {
  try {
    if (!fs.existsSync(file)) return null;
    return fromJSON(JSON.parse(fs.readFileSync(file, 'utf8')));
  } catch (err) {
    console.warn('  Could not read the saved sandbox (' + err.message + '); starting fresh.');
    return null;
  }
}

function save(file, account) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + '.tmp';
  // A script can stash a live Range or Sheet somewhere odd (a rule, a note); save it as
  // text rather than failing to save at all.
  const flat = (k, v) => (v instanceof Range ? v.getA1Notation() : v instanceof Sheet ? v.getName() : v instanceof Spreadsheet ? v.getId() : v);
  fs.writeFileSync(tmp, JSON.stringify(toJSON(account), flat));
  fs.renameSync(tmp, file);
  account.dirty = false;
}

module.exports = { load, save, toJSON, fromJSON };
