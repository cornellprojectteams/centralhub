'use strict';
// A stand-in Google account for the sandbox: spreadsheets, Drive, a Gmail outbox, a
// fake Fabman API, and the Apps Script services the hub's scripts call. Method names
// match the real ones, so a typo in a script fails here the way it would on Google.
// Nothing in this file talks to the internet.
//
// Formulas are stored, not calculated. The one exception is the inventory sheet's
// Status and Next Maintenance columns, worked out in Spreadsheet.engine() below.

const crypto = require('crypto');

let seq = 1;
const DRIVE_CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-';
// Drive-style ids: long enough for the scripts' /[-\w]{25,}/ id finder.
function driveId() {
  const bytes = crypto.randomBytes(33);
  let s = '1';
  for (let i = 0; i < 32; i++) s += DRIVE_CHARS[bytes[i] % 64];
  return s;
}

// ---------------------------------------------------------------- cells and ranges

function colNumber(letters) { return letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
function colLetters(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function a1ToRC(a1) {
  const m = /^(?:'?[^!]+'?!)?\$?([A-Z]+)\$?(\d*)(?::\$?([A-Z]+)\$?(\d*))?$/.exec(String(a1).trim());
  if (!m) throw new Error('Exception: Range not found: ' + a1);
  const c1 = colNumber(m[1]), r1 = m[2] ? +m[2] : 1;
  if (!m[3]) return { r: r1, c: c1, nr: 1, nc: 1 };
  return { r: r1, c: c1, nr: m[4] ? +m[4] - r1 + 1 : null, nc: colNumber(m[3]) - c1 + 1 };
}

// What Sheets does with a typed-in string: an ISO date becomes a date, a number-looking
// string stays as typed by a script (Sheets would convert it; scripts here rarely rely on
// that), and a leading apostrophe forces text and is hidden.
function sheetsParse(x) {
  if (x === null || x === undefined) return '';
  if (typeof x !== 'string') return x;
  if (x.charAt(0) === "'") return x.slice(1);
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(x);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return x;
}

function displayValue(v) {
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, '0');
    const time = v.getHours() || v.getMinutes() || v.getSeconds() ? ' ' + v.getHours() + ':' + p(v.getMinutes()) + ':' + p(v.getSeconds()) : '';
    return (v.getMonth() + 1) + '/' + v.getDate() + '/' + v.getFullYear() + time;
  }
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  return v === null || v === undefined ? '' : String(v);
}

function headerKey(sheet, c) {
  const cl = sheet.cells.get('1,' + c);
  if (!cl) return '';
  if (cl.f) { const m = /^=\s*\{\s*"([^"]+)"/.exec(cl.f); return m ? m[1] : ''; }
  return String(cl.v == null ? '' : cl.v);
}
function colOf(sheet, header) {
  for (let c = 1; c <= sheet.maxCols; c++) if (headerKey(sheet, c).toLowerCase() === header.toLowerCase()) return c;
  return 0;
}
const CALC_HEADERS = ['Status', 'Next Maintenance'];

class Range {
  constructor(sheet, r, c, nr, nc) {
    // Like Sheets: a range past the sheet's last row or column is an error, so a script
    // that forgets insertRowsAfter / insertColumnsAfter fails here too.
    if (r < 1 || c < 1 || nr < 1 || nc < 1 || r + nr - 1 > sheet.maxRows || c + nc - 1 > sheet.maxCols) {
      throw new Error(`Exception: The coordinates of the range are outside the dimensions of the sheet. (${sheet.name}: row ${r}, column ${c}, ${nr} x ${nc}; the sheet is ${sheet.maxRows} x ${sheet.maxCols})`);
    }
    Object.assign(this, { sheet, r, c, nr, nc });
  }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); return this; }
  cell(r, c) { return this.sheet.cell(r, c); }
  getSheet() { return this.sheet; }
  getRow() { return this.r; } getColumn() { return this.c; }
  getRowIndex() { return this.r; }
  getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getLastRow() { return this.r + this.nr - 1; } getLastColumn() { return this.c + this.nc - 1; }
  getA1Notation() {
    const tl = colLetters(this.c) + this.r;
    return this.nr === 1 && this.nc === 1 ? tl : tl + ':' + colLetters(this.c + this.nc - 1) + (this.r + this.nr - 1);
  }
  getCell(r, c) { return new Range(this.sheet, this.r + r - 1, this.c + c - 1, 1, 1); }
  offset(r, c, nr, nc) { return new Range(this.sheet, this.r + r, this.c + c, nr || this.nr, nc || this.nc); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.valueAt(this.r + i, this.c + j)); out.push(row); } return out; }
  getValue() { return this.sheet.valueAt(this.r, this.c); }
  getDisplayValues() { return this.getValues().map((row) => row.map(displayValue)); }
  getDisplayValue() { return displayValue(this.getValue()); }
  getFormulas() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.cell(this.r + i, this.c + j).f || ''); out.push(row); } return out; }
  getFormula() { return this.getFormulas()[0][0]; }
  isBlank() { return this.getValues().every((row) => row.every((v) => v === '')); }
  setValues(v) {
    if (!v || v.length !== this.nr || v.some((row) => !row || row.length !== this.nc)) {
      throw new Error(`Exception: The number of rows or columns in the data does not match the range. The data has ${v ? v.length : 0} rows and ${v && v[0] ? v[0].length : 0} columns, the range has ${this.nr} and ${this.nc}.`);
    }
    this.each((r, c, i, j) => { const cl = this.cell(r, c); cl.v = sheetsParse(v[i][j]); cl.f = ''; });
    this.sheet.touch();
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setFormula(f) {
    if (!/^=/.test(f)) throw new Error('Exception: Formula must start with =');
    this.each((r, c) => { const cl = this.cell(r, c); cl.f = f; cl.v = ''; });
    this.sheet.touch();
    return this;
  }
  setFormulas(fs) { this.each((r, c, i, j) => { const cl = this.cell(r, c); cl.f = fs[i][j]; cl.v = ''; }); this.sheet.touch(); return this; }
  clearContent() { this.each((r, c) => { const cl = this.cell(r, c); cl.v = ''; cl.f = ''; }); this.sheet.touch(); return this; }
  clear() { return this.clearContent(); }
  clearFormat() { return this; }
  setNote(n) { return this.each((r, c) => { this.cell(r, c).note = n; }); }
  getNote() { return this.cell(this.r, this.c).note || ''; }
  setNumberFormat(f) { return this.each((r, c) => { this.cell(r, c).fmt = f; }); }
  setNumberFormats() { return this; }
  setDataValidation(dv) { return this.each((r, c) => { this.cell(r, c).dv = dv || undefined; }); }
  setDataValidations(dvs) { return this.each((r, c, i, j) => { this.cell(r, c).dv = dvs[i][j]; }); }
  getDataValidation() { return this.cell(this.r, this.c).dv || null; }
  clearDataValidations() { return this.each((r, c) => { delete this.cell(r, c).dv; }); }
  // Checkboxes keep TRUE / FALSE that is already there; anything else becomes unticked.
  insertCheckboxes() {
    this.each((r, c) => { const cl = this.cell(r, c); cl.dv = { kind: 'dv', calls: [['requireCheckbox', []]] }; if (cl.v !== true && cl.v !== false) cl.v = false; cl.f = ''; });
    this.sheet.touch();
    return this;
  }
  removeCheckboxes() { return this.clearDataValidations(); }
  check() { return this.setValues(this.getValues().map((row) => row.map(() => true))); }
  uncheck() { return this.setValues(this.getValues().map((row) => row.map(() => false))); }
  setBorder() { return this; }
  merge() { return this; } breakApart() { return this; }
  activate() { this.sheet.ss.active = this.sheet; return this; }
  protect() { const p = new Protection(this.sheet, this); this.sheet.protections.push(p); return p; }
}
for (const m of ['setBackground', 'setBackgrounds', 'setFontWeight', 'setFontWeights', 'setFontColor', 'setFontColors', 'setFontSize', 'setFontStyle', 'setFontFamily', 'setFontLine', 'setWrap', 'setWrapStrategy', 'setVerticalAlignment', 'setHorizontalAlignment', 'setTextRotation']) {
  Range.prototype[m] = function (v) { return this.each((r, c) => { this.cell(r, c)[m] = v; }); };
}

class Protection {
  constructor(sheet, range) { this.sheet = sheet; this.range = range; this.desc = ''; this.warn = false; this.type = range ? 'RANGE' : 'SHEET'; }
  setDescription(d) { this.desc = d; return this; } getDescription() { return this.desc; }
  setWarningOnly(w) { this.warn = w; return this; } isWarningOnly() { return this.warn; }
  canEdit() { return true; }
  addEditor() { return this; } removeEditors() { return this; } getEditors() { return []; }
  getRange() { return this.range; }
  remove() { this.sheet.protections = this.sheet.protections.filter((p) => p !== this); }
}

class Sheet {
  constructor(ss, name, gid) {
    Object.assign(this, { ss, name, gid: gid === undefined || gid === null || gid === '' ? 100000 + (seq++) * 7919 : Number(gid), maxRows: 1000, maxCols: 26, cells: new Map(), hidden: false, protections: [], rules: [], formUrl: null, widths: {}, frozenRows: 0, frozenCols: 0 });
  }
  touch() { if (this.ss.account) this.ss.account.dirty = true; }
  cell(r, c) { const k = r + ',' + c; if (!this.cells.has(k)) this.cells.set(k, { v: '', f: '' }); return this.cells.get(k); }
  rawAt(r, c) { const cl = this.cells.get(r + ',' + c); return cl ? cl.v : ''; }
  valueAt(r, c) {
    if (this.name === 'Inventory Master' && r >= 2) {
      const h1 = this.cells.get('1,' + c);
      if (h1 && h1.f && CALC_HEADERS.indexOf(headerKey(this, c)) >= 0) return this.ss.engine(this, r, headerKey(this, c));
    }
    const cl = this.cells.get(r + ',' + c);
    if (!cl) return '';
    if (cl.f) { const m = /^=\s*\{\s*"([^"]+)"/.exec(cl.f); return m && r === 1 ? m[1] : ''; }
    return cl.v === undefined || cl.v === null ? '' : cl.v;
  }
  getParent() { return this.ss; }
  getName() { return this.name; }
  setName(n) { if (this.ss.sheets.some((s) => s !== this && s.name === n)) throw new Error('Exception: A sheet with the name "' + n + '" already exists.'); this.name = n; this.touch(); return this; }
  getSheetId() { return this.gid; }
  getIndex() { return this.ss.sheets.indexOf(this) + 1; }
  getSheetName() { return this.name; }
  getMaxRows() { return this.maxRows; } getMaxColumns() { return this.maxCols; }
  getFrozenRows() { return this.frozenRows; } getFrozenColumns() { return this.frozenCols; }
  insertColumnsAfter(after, n) { this.maxCols += n || 1; return this; }
  insertColumnAfter(after) { return this.insertColumnsAfter(after, 1); }
  insertRowsAfter(after, n) { this.maxRows += n || 1; return this; }
  insertRowAfter(after) { return this.insertRowsAfter(after, 1); }
  insertRowBefore(before) { return this.insertRowsBefore(before, 1); }
  insertRowsBefore(before, n) {
    n = n || 1;
    const moved = new Map();
    for (const [k, cl] of this.cells) { const [rr, cc] = k.split(',').map(Number); moved.set((rr >= before ? rr + n : rr) + ',' + cc, cl); }
    this.cells = moved; this.maxRows += n; this.touch(); return this;
  }
  hasContent(cl) { return (cl.v !== '' && cl.v !== undefined && cl.v !== null) || !!cl.f; }
  getLastRow() { let m = 0; for (const [k, cl] of this.cells) if (this.hasContent(cl)) m = Math.max(m, +k.split(',')[0]); return m; }
  getLastColumn() { let m = 0; for (const [k, cl] of this.cells) if (this.hasContent(cl)) m = Math.max(m, +k.split(',')[1]); return m; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const x = a1ToRC(a); return new Range(this, x.r, x.c, x.nr === null ? Math.max(1, this.maxRows - x.r + 1) : x.nr, x.nc); }
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  // appendRow is the one write that grows the sheet by itself.
  appendRow(values) {
    const r = this.getLastRow() + 1;
    if (r > this.maxRows) this.maxRows = r;
    if (values.length > this.maxCols) this.maxCols = values.length;
    this.getRange(r, 1, 1, values.length).setValues([values]);
    return this;
  }
  setFrozenRows(n) { this.frozenRows = n; return this; } setFrozenColumns(n) { this.frozenCols = n; return this; }
  setRowHeight() { return this; } setRowHeights() { return this; } autoResizeColumns() { return this; } autoResizeColumn() { return this; }
  setColumnWidth(c, w) { this.widths[c] = w; return this; }
  setColumnWidths(c, n, w) { for (let i = 0; i < n; i++) this.widths[c + i] = w; return this; }
  setTabColor(c) { this.tab = c; return this; } setHiddenGridlines() { return this; }
  hideSheet() { if (this.ss.sheets.filter((s) => !s.hidden).length < 2) throw new Error("Exception: You can't hide all the sheets in a document."); this.hidden = true; return this; }
  showSheet() { this.hidden = false; return this; }
  isSheetHidden() { return this.hidden; }
  activate() { this.ss.active = this; return this; }
  clear() { this.cells = new Map(); this.touch(); return this; }
  clearContents() { for (const cl of this.cells.values()) { cl.v = ''; cl.f = ''; } this.touch(); return this; }
  clearFormats() { return this; }
  clearNotes() { for (const cl of this.cells.values()) delete cl.note; return this; }
  clearConditionalFormatRules() { this.rules = []; }
  setConditionalFormatRules(r) { this.rules = r; }
  getConditionalFormatRules() { return this.rules.slice(); }
  getProtections(type) { return this.protections.filter((p) => p.type === type); }
  protect() { const p = new Protection(this, null); this.protections.push(p); return p; }
  deleteRow(r) {
    if (r < 1 || r > this.maxRows) throw new Error('Exception: Those rows are out of bounds.');
    const moved = new Map();
    for (const [k, cl] of this.cells) { const [rr, cc] = k.split(',').map(Number); if (rr < r) moved.set(k, cl); else if (rr > r) moved.set((rr - 1) + ',' + cc, cl); }
    this.cells = moved; this.maxRows -= 1; this.touch(); return this;
  }
  deleteRows(start, n) { for (let i = 0; i < n; i++) this.deleteRow(start); return this; }
  deleteColumn(c) {
    if (c < 1 || c > this.maxCols) throw new Error('Exception: Those columns are out of bounds.');
    const moved = new Map();
    for (const [k, cl] of this.cells) { const [rr, cc] = k.split(',').map(Number); if (cc < c) moved.set(k, cl); else if (cc > c) moved.set(rr + ',' + (cc - 1), cl); }
    this.cells = moved;
    const w = {};
    Object.keys(this.widths).forEach((k) => { const n = Number(k); if (n < c) w[n] = this.widths[k]; else if (n > c) w[n - 1] = this.widths[k]; });
    this.widths = w;
    this.protections = this.protections.filter((pr) => !(pr.range && pr.range.c === c && pr.range.nc === 1));
    this.maxCols -= 1;
    this.touch();
    return this;
  }
  deleteColumns(start, n) { for (let i = 0; i < n; i++) this.deleteColumn(start); return this; }
  // Real semantics: destinationIndex is a column index from before the move.
  moveColumns(range, dest) {
    if (!(range instanceof Range) || range.nc !== 1) throw new Error('sandbox: moveColumns moves one column at a time');
    const from = range.c;
    const order = [];
    for (let c = 1; c <= this.maxCols; c++) order.push(c);
    order.splice(from - 1, 1);
    order.splice(dest > from ? dest - 2 : dest - 1, 0, from);
    const newCol = {};
    order.forEach((oldC, i) => { newCol[oldC] = i + 1; });
    const moved = new Map();
    for (const [k, cl] of this.cells) { const [rr, cc] = k.split(',').map(Number); moved.set(rr + ',' + newCol[cc], cl); }
    this.cells = moved;
    const w = {};
    Object.keys(this.widths).forEach((c) => { w[newCol[c]] = this.widths[c]; });
    this.widths = w;
    this.touch();
    return this;
  }
  getFormUrl() { return this.formUrl; }
  // Plain rows for the sandbox's own pages and fixtures.
  rows() { const lr = this.getLastRow(), lc = this.getLastColumn(); return lr && lc ? this.getRange(1, 1, lr, lc).getValues() : []; }
}

class Spreadsheet {
  constructor(account, opts) {
    opts = opts || {};
    this.account = account;
    this.id = opts.id || driveId();
    this.title = opts.name || 'Untitled spreadsheet';
    this.sheets = [];
    this.forms = [];
    this.editors = (opts.editors || []).slice();
    this.owner = opts.owner || (account && account.user) || 'you@sandbox.test';
    this.note = opts.note || '';
    if (!opts.empty) this.sheets.push(new Sheet(this, 'Sheet1', 0));   // a new spreadsheet's first tab is gid 0
    this.active = this.sheets[0] || null;
  }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getName() { return this.title; } rename(n) { this.title = n; if (this.account) this.account.dirty = true; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  getSheetById(gid) { return this.sheets.find((s) => String(s.gid) === String(gid)) || null; }
  getSheets() { return this.sheets.slice(); }
  getNumSheets() { return this.sheets.length; }
  getActiveSheet() { return this.active || this.sheets[0]; }
  insertSheet(name, index, gid) {
    if (typeof name !== 'string') name = 'Sheet' + (this.sheets.length + 1);
    if (this.getSheetByName(name)) throw new Error('Exception: A sheet with the name "' + name + '" already exists. Please enter another name.');
    const s = new Sheet(this, name, gid);
    if (typeof index === 'number') this.sheets.splice(index, 0, s); else this.sheets.push(s);
    this.active = s;
    if (this.account) this.account.dirty = true;
    return s;
  }
  deleteSheet(s) {
    if (s.formUrl) throw new Error('Exception: This sheet is linked to a form. Unlink the form first.');
    if (this.sheets.length < 2) throw new Error("Exception: You can't remove all the sheets in a document.");
    this.sheets = this.sheets.filter((x) => x !== s);
    if (this.active === s) this.active = this.sheets[0];
    if (this.account) this.account.dirty = true;
  }
  setActiveSheet(s) { this.active = s; return s; }
  moveActiveSheet(pos) { const s = this.active; this.sheets = this.sheets.filter((x) => x !== s); this.sheets.splice(pos - 1, 0, s); }
  getFormUrl() { return this.forms.length ? this.forms[0].getEditUrl() : null; }
  getSpreadsheetTimeZone() { return 'America/New_York'; }
  getEditors() { return this.editors.map((e) => user(e)); }
  getViewers() { return []; }
  getOwner() { return user(this.owner); }
  addEditor(e) { const email = typeof e === 'string' ? e : e.getEmail(); if (this.editors.indexOf(email) < 0) this.editors.push(email); if (this.account) this.account.dirty = true; return this; }
  removeEditor(e) {
    const email = typeof e === 'string' ? e : e.getEmail();
    if (email === this.owner) throw new Error('Exception: The owner of a file cannot be removed.');
    this.editors = this.editors.filter((x) => x.toLowerCase() !== String(email).toLowerCase());
    if (this.account) this.account.dirty = true;
    return this;
  }
  toast(msg, title) { const run = this.account && this.account.run; if (run) run.ui.push({ kind: 'toast', title: title || '', text: String(msg) }); }
  getRange(a1) { const m = /^'?([^!']+)'?!(.+)$/.exec(a1); if (!m) return this.getActiveSheet().getRange(a1); return this.getSheetByName(m[1]).getRange(m[2]); }
  // Stand-in for the inventory sheet's formulas: Status from Out of Service and the stock
  // counts, Next Maintenance from the latest Maintenance Log row for the item.
  engine(sheet, r, header) {
    const raw = (h) => { const col = colOf(sheet, h); return col ? sheet.rawAt(r, col) : ''; };
    const id = String(raw('Item ID') == null ? '' : raw('Item ID')).trim().toUpperCase();
    if (!id) return '';
    if (header === 'Status') {
      const oos = String(raw('Out of Service') || '').trim();
      const cat = String(raw('Category') || '').trim().toLowerCase();
      const qty = raw('Current Qty'), min = raw('Min Qty');
      if (oos) return oos;
      if (cat === 'consumable') return qty === '' ? '' : (Number(qty) <= 0 ? 'Out of Stock' : (min !== '' && Number(qty) < Number(min) ? 'Low Stock' : 'In Stock'));
      return 'In service';
    }
    if (header === 'Next Maintenance') {
      const log = this.getSheetByName('Maintenance Log');
      if (!log) return '';
      const col = (h) => colOf(log, h);
      let best = '', bestTs = -Infinity;
      for (let rr = 2; rr <= log.getLastRow(); rr++) {
        const item = String(log.rawAt(rr, col('Item')) || '').trim();
        if (!item || item.split(/\s+/)[0].toUpperCase() !== id) continue;
        const next = log.rawAt(rr, col('Next Due'));
        if (!(next instanceof Date)) continue;
        const ts = log.rawAt(rr, col('Timestamp'));
        const t = ts instanceof Date ? ts.getTime() : -1;
        if (t >= bestTs) { bestTs = t; best = next; }
      }
      return best;
    }
    return '';
  }
}

function user(email) { return { getEmail: () => email, getName: () => email.split('@')[0], toString: () => email }; }

// ---------------------------------------------------------------- forms
// Enough of FormApp to build a form, link it to a sheet and read it back. A linked form
// gets a "Form Responses N" tab with its question titles as headers, as on Google.

const ITEM_TYPES = { addTextItem: 'TEXT', addParagraphTextItem: 'PARAGRAPH_TEXT', addListItem: 'LIST', addMultipleChoiceItem: 'MULTIPLE_CHOICE', addCheckboxItem: 'CHECKBOX', addDateItem: 'DATE', addScaleItem: 'SCALE', addSectionHeaderItem: 'SECTION_HEADER' };
class FormItem {
  constructor(form, type) { Object.assign(this, { form, type, id: 100000 + (seq++), title: '', help: '', required: false, choices: [] }); }
  getId() { return this.id; } getType() { return this.type; } getIndex() { return this.form.items.indexOf(this); }
  setTitle(t) { this.title = String(t); return this; } getTitle() { return this.title; }
  setHelpText(t) { this.help = String(t); return this; } getHelpText() { return this.help; }
  setRequired(r) { this.required = !!r; return this; } isRequired() { return this.required; }
  setChoiceValues(v) { this.choices = v.map(String); return this; }
  setChoices(c) { this.choices = c.map((x) => (x && x.value !== undefined ? x.value : String(x))); return this; }
  createChoice(v) { return { value: String(v), getValue() { return this.value; } }; }
  getChoices() { return this.choices.map((v) => ({ getValue: () => v })); }
  showOtherOption() { return this; } setBounds() { return this; } setLabels() { return this; } setValidation() { return this; }
  duplicate() { const d = Object.assign(new FormItem(this.form, this.type), { title: this.title, help: this.help, required: this.required, choices: this.choices.slice() }); this.form.items.push(d); return d; }
  createResponse(v) { return { item: this, value: v, getItem: () => this, getResponse: () => v }; }
  asTextItem() { return this; } asParagraphTextItem() { return this; } asListItem() { return this; } asMultipleChoiceItem() { return this; } asCheckboxItem() { return this; } asDateItem() { return this; } asScaleItem() { return this; }
}

class Form {
  constructor(account, title) { Object.assign(this, { account, id: driveId(), title, desc: '', items: [], dest: null, trashed: false, collectEmail: false, accepting: true, confirmation: '' }); }
  getId() { return this.id; }
  getTitle() { return this.title; } setTitle(t) { this.title = String(t); return this; }
  getDescription() { return this.desc; } setDescription(d) { this.desc = String(d); return this; }
  setCollectEmail(c) { this.collectEmail = !!c; return this; } collectsEmail() { return this.collectEmail; }
  setEmailCollectionType() { this.collectEmail = true; return this; }
  setConfirmationMessage(m) { this.confirmation = String(m); return this; }
  setAcceptingResponses(a) { this.accepting = !!a; return this; } isAcceptingResponses() { return this.accepting; }
  setAllowResponseEdits() { return this; } setLimitOneResponsePerUser() { return this; } setRequireLogin() { return this; } setShowLinkToRespondAgain() { return this; }
  getEditUrl() { return 'https://docs.google.com/forms/d/' + this.id + '/edit'; }
  getPublishedUrl() { return 'https://docs.google.com/forms/d/e/' + this.id + '/viewform'; }
  shortenFormUrl(u) { return u; }
  getItems() { return this.items.slice(); }
  getItemById(id) { return this.items.find((i) => String(i.id) === String(id)) || null; }
  deleteItem(x) { const i = typeof x === 'number' ? x : this.items.indexOf(x); if (i >= 0) this.items.splice(i, 1); }
  getDestinationId() { return this.dest; }
  getDestinationType() { return this.dest ? 'SPREADSHEET' : null; }
  setDestination(type, id) {
    const ss = this.account.spreadsheets[id];
    if (!ss) throw new Error('Exception: No item with the given ID could be found.');
    this.dest = id;
    let n = 1;
    while (ss.getSheetByName('Form Responses ' + n)) n++;
    const sh = ss.insertSheet('Form Responses ' + n, 0);
    const heads = ['Timestamp'].concat(this.collectEmail ? ['Email Address'] : [], this.items.filter((i) => i.type !== 'SECTION_HEADER').map((i) => i.title));
    sh.getRange(1, 1, 1, heads.length).setValues([heads]);
    sh.formUrl = this.getEditUrl();
    ss.forms.push(this);
    return this;
  }
  removeDestination() {
    const ss = this.account.spreadsheets[this.dest];
    if (ss) { ss.forms = ss.forms.filter((f) => f !== this); ss.sheets.forEach((sh) => { if (sh.formUrl === this.getEditUrl()) sh.formUrl = null; }); }
    this.dest = null;
    return this;
  }
  getResponses() { return []; }
}
Object.keys(ITEM_TYPES).forEach((m) => { Form.prototype[m] = function () { const it = new FormItem(this, ITEM_TYPES[m]); this.items.push(it); return it; }; });

// ---------------------------------------------------------------- Drive

class Blob {
  constructor(bytes, mime, name) { this.bytes = bytes; this.mime = mime || 'application/octet-stream'; this.name = name || null; this.isBlob = true; }
  getBytes() { return this.bytes.slice(); }
  getContentType() { return this.mime; }
  setContentType(m) { this.mime = m; return this; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getDataAsString() { return Buffer.from(this.bytes.map((b) => b & 255)).toString('utf8'); }
  copyBlob() { return new Blob(this.bytes.slice(), this.mime, this.name); }
  getBlob() { return this; }
}
const toSigned = (buf) => Array.from(buf).map((b) => (b > 127 ? b - 256 : b));
const toBuffer = (bytes) => Buffer.from(bytes.map((b) => b & 255));

const iter = (list) => { let i = 0; return { hasNext: () => i < list.length, next: () => { if (i >= list.length) throw new Error('Exception: No more items.'); return list[i++]; }, getContinuationToken: () => null }; };

class DFolder {
  constructor(account, name, parentId, id) { Object.assign(this, { account, id: id || driveId(), name, parentId: parentId || null, trashed: false, access: null, perm: null }); account.drive.folders[this.id] = this; }
  getId() { return this.id; } getName() { return this.name; } setName(n) { this.name = n; return this; }
  getUrl() { return 'https://drive.google.com/drive/folders/' + this.id; }
  isTrashed() { return this.trashed; } setTrashed(t) { this.trashed = !!t; this.account.dirty = true; return this; }
  setSharing(access, perm) { if (!access || !perm) throw new Error('Exception: Invalid argument: permission'); this.access = access; this.perm = perm; return this; }
  createFile(blob, content, mime) {
    if (typeof blob === 'string') blob = new Blob(toSigned(Buffer.from(String(content || ''), 'utf8')), mime || 'text/plain', blob);
    if (!blob || !blob.isBlob) throw new Error('Exception: Invalid argument: blob');
    return new DFile(this.account, blob.name || 'Untitled', blob.mime, blob.bytes, this.id);
  }
  createFolder(name) { return new DFolder(this.account, name, this.id); }
  getFoldersByName(name) { return iter(Object.values(this.account.drive.folders).filter((f) => f.parentId === this.id && f.name === name && !f.trashed)); }
  getFolders() { return iter(Object.values(this.account.drive.folders).filter((f) => f.parentId === this.id && !f.trashed)); }
  getFiles() { return iter(Object.values(this.account.drive.files).filter((f) => f.parentId === this.id && !f.trashed)); }
  getFilesByName(name) { return iter(Object.values(this.account.drive.files).filter((f) => f.parentId === this.id && f.name === name && !f.trashed)); }
  getParents() { return iter(this.parentId ? [this.account.drive.folders[this.parentId]] : []); }
}
class DFile {
  constructor(account, name, mime, bytes, parentId, id) {
    Object.assign(this, { account, id: id || driveId(), name, mime, bytes, parentId, trashed: false, desc: '', access: null, created: new Date() });
    account.drive.files[this.id] = this;
    account.dirty = true;
  }
  getId() { return this.id; } getName() { return this.name; } setName(n) { this.name = n; return this; }
  getMimeType() { return this.mime; }
  getSize() { return this.bytes.length; }
  getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view'; }
  getDateCreated() { return this.created; }
  getParents() { return iter(this.parentId ? [this.account.drive.folders[this.parentId]] : []); }
  setTrashed(t) { this.trashed = !!t; this.account.dirty = true; return this; } isTrashed() { return this.trashed; }
  setDescription(d) { this.desc = d; return this; } getDescription() { return this.desc; }
  setSharing(access, perm) { if (!access || !perm) throw new Error('Exception: Invalid argument: permission'); this.access = access; this.perm = perm; return this; }
  getBlob() { return new Blob(this.bytes.slice(), this.mime, this.name); }
  getAs(mime) { return new Blob(this.bytes.slice(), mime, this.name); }
  moveTo(folder) { this.parentId = folder.getId(); return this; }
}

// ---------------------------------------------------------------- the fake Fabman API

// Speaks the parts of https://fabman.io/api/v1 that 05_fabman_members.gs uses. Pages of
// 50 with a Link header, like the real one, so fetchAll's paging gets exercised.
class FakeFabman {
  constructor(data) {
    data = data || {};
    this.members = data.members || [];
    this.trainings = data.trainings || {};
    this.packages = data.packages || {};
    this.invitations = data.invitations || [];
    this.nextId = data.nextId || 700001;
    this.calls = data.calls || [];
    // Switches on the Fake Fabman page, to see how the scripts cope:
    //   down        every request fails with 503, as in an outage
    //   refuseKey   every request fails with 401, as with a revoked API key
    this.down = !!data.down;
    this.refuseKey = !!data.refuseKey;
  }
  toJSON() { return { members: this.members, trainings: this.trainings, packages: this.packages, invitations: this.invitations, nextId: this.nextId, calls: this.calls.slice(-300), down: this.down, refuseKey: this.refuseKey }; }
  member(id) { return this.members.find((m) => String(m.id) === String(id)); }
  handle(method, url, payload, headers) {
    const u = new URL(url);
    const parts = u.pathname.replace(/^\/api\/v1\/?/, '').split('/').filter(Boolean);
    let body = null;
    try { body = payload ? JSON.parse(payload) : null; } catch (e) { body = null; }
    const out = (code, json, headers) => ({ code, text: json === undefined ? '' : JSON.stringify(json), headers: Object.assign({ 'Content-Type': 'application/json' }, headers || {}) });
    this.calls.push({ at: new Date().toISOString(), method, path: u.pathname + u.search });
    const auth = Object.keys(headers || {}).filter((k) => k.toLowerCase() === 'authorization').map((k) => String(headers[k]))[0] || '';
    if (this.refuseKey || !/^Bearer\s+\S+/.test(auth)) return out(401, { message: 'Unauthorized: missing or invalid API key' });
    if (this.down) return out(503, { message: 'Service Unavailable' });
    if (parts[0] !== 'members') return out(404, { error: 'Not found in the sandbox Fabman' });
    if (parts.length === 1) {
      if (method === 'get') {
        const q = String(u.searchParams.get('q') || '').toLowerCase();
        const list = q ? this.members.filter((m) => [m.firstName, m.lastName, m.emailAddress, m.company, JSON.stringify(m.metadata || {})].join(' ').toLowerCase().indexOf(q) >= 0) : this.members;
        const offset = Number(u.searchParams.get('offset') || 0), limit = Math.min(Number(u.searchParams.get('limit') || 50), 50);
        const page = list.slice(offset, offset + limit);
        const more = offset + limit < list.length ? { Link: '</api/v1/members?' + (q ? 'q=' + encodeURIComponent(q) + '&' : '') + 'offset=' + (offset + limit) + '>; rel="next"' } : {};
        return out(200, page, more);
      }
      if (method === 'post') {
        if (!body || !body.emailAddress) return out(422, { error: 'emailAddress is required' });
        if (this.members.some((m) => m.emailAddress.toLowerCase() === String(body.emailAddress).toLowerCase())) return out(422, { error: 'A member with this email address already exists' });
        const m = Object.assign({ id: this.nextId++, lockVersion: 1, state: 'active', metadata: {} }, body, { createdAt: new Date().toISOString() });
        this.members.push(m);
        return out(201, m);
      }
    }
    const m = this.member(parts[1]);
    if (!m) return out(404, { error: 'Member ' + parts[1] + ' not found' });
    if (parts.length === 2) {
      if (method === 'get') return out(200, m);
      if (method === 'put') {
        if (body && body.lockVersion !== undefined && body.lockVersion !== m.lockVersion) return out(409, { error: 'lockVersion is out of date' });
        Object.keys(body || {}).forEach((k) => { if (k !== 'lockVersion' && k !== 'id') m[k] = k === 'metadata' ? Object.assign({}, m.metadata, body.metadata) : body[k]; });
        m.lockVersion++;
        return out(200, m);
      }
      if (method === 'delete') {
        this.members = this.members.filter((x) => x !== m);
        return out(204);
      }
    }
    const list = parts[2] === 'trainings' ? (this.trainings[m.id] = this.trainings[m.id] || []) : parts[2] === 'packages' ? (this.packages[m.id] = this.packages[m.id] || []) : null;
    if (parts[2] === 'invitation' && method === 'post') { this.invitations.push({ member: m.id, email: m.emailAddress, at: new Date().toISOString() }); return out(200, { sent: true }); }
    if (!list) return out(404, { error: 'Not found in the sandbox Fabman' });
    if (parts.length === 3 && method === 'get') return out(200, list);
    if (parts.length === 3 && method === 'post') {
      const rec = Object.assign({ id: this.nextId++ }, body || {});
      list.push(rec);
      return out(201, rec);
    }
    if (parts.length === 4 && method === 'delete') {
      const i = list.findIndex((x) => String(x.id) === String(parts[3]));
      if (i < 0) return out(404, { error: 'Not found' });
      list.splice(i, 1);
      return out(204);
    }
    return out(405, { error: 'Not supported in the sandbox Fabman' });
  }
}

// ---------------------------------------------------------------- the account

class Account {
  constructor() {
    this.user = 'you@sandbox.test';
    this.spreadsheets = {};
    this.forms = {};
    this.drive = { folders: {}, files: {} };
    this.outbox = [];
    this.fabman = new FakeFabman();
    this.state = {};      // per project: props, triggers, logs, cache
    this.run = null;      // the execution in progress, for toasts and dialogs
    this.dirty = false;
    this.root = new DFolder(this, 'My Drive', null);
  }
  ss(id) { return this.spreadsheets[id] || null; }
  addSpreadsheet(opts) { const ss = new Spreadsheet(this, opts); this.spreadsheets[ss.id] = ss; this.dirty = true; return ss; }
  projectState(key) {
    if (!this.state[key]) this.state[key] = { props: {}, triggers: [], logs: [], cache: {} };
    return this.state[key];
  }
  fileById(id) { return this.drive.files[id] || null; }
  addFile(name, mime, buffer, folder) { return new DFile(this, name, mime, toSigned(buffer), (folder || this.root).id); }
  folderNamed(name, parent) {
    parent = parent || this.root;
    const hit = Object.values(this.drive.folders).find((f) => f.parentId === parent.id && f.name === name && !f.trashed);
    return hit || new DFolder(this, name, parent.id);
  }
  sendMail(project, kind, to, subject, body, opts) {
    opts = opts || {};
    const images = {};
    Object.keys(opts.inlineImages || {}).forEach((cid) => {
      const b = opts.inlineImages[cid];
      if (b && b.isBlob) images[cid] = { mime: b.mime, data: toBuffer(b.bytes).toString('base64') };
    });
    const mail = {
      id: 'm' + Date.now().toString(36) + (seq++).toString(36),
      at: new Date().toISOString(), project, kind,
      from: opts.from || opts.name ? (opts.name ? opts.name + ' <' + (opts.from || this.user) + '>' : opts.from) : this.user,
      to: String(to || ''), cc: String(opts.cc || ''), bcc: String(opts.bcc || ''), replyTo: String(opts.replyTo || ''),
      subject: String(subject || ''), body: String(body || ''), html: opts.htmlBody ? String(opts.htmlBody) : '',
      images, attachments: (opts.attachments || []).map((a) => (a && a.getName ? a.getName() : 'attachment')),
    };
    this.outbox.push(mail);
    this.dirty = true;
    if (this.run) this.run.mails.push(mail.id);
    return mail;
  }
}

// ---------------------------------------------------------------- Apps Script services

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Java SimpleDateFormat, the subset people use. The sandbox runs on New York time, like
// the real projects (server.js sets TZ), so the time zone argument is not consulted.
function formatDate(d, tz, pattern) {
  if (!(d instanceof Date)) throw new Error('Exception: The parameters (' + typeof d + ',String,String) don\'t match the method signature for Utilities.formatDate.');
  const p2 = (n) => String(n).padStart(2, '0');
  const out = [];
  const re = /'([^']*)'|y{4}|y{2}|M{4}|M{3}|M{2}|M|d{2}|d|E{4}|E{3}|E|H{2}|H|h{2}|h|m{2}|m|s{2}|s|S{3}|a|Z|z|./g;
  let m;
  while ((m = re.exec(pattern)) !== null) {
    const t = m[0];
    if (m[1] !== undefined) { out.push(m[1] || "'"); continue; }
    switch (t) {
      case 'yyyy': out.push(d.getFullYear()); break;
      case 'yy': out.push(String(d.getFullYear()).slice(-2)); break;
      case 'MMMM': out.push(MONTHS[d.getMonth()]); break;
      case 'MMM': out.push(MONTHS[d.getMonth()].slice(0, 3)); break;
      case 'MM': out.push(p2(d.getMonth() + 1)); break;
      case 'M': out.push(d.getMonth() + 1); break;
      case 'dd': out.push(p2(d.getDate())); break;
      case 'd': out.push(d.getDate()); break;
      case 'EEEE': out.push(DAYS[d.getDay()]); break;
      case 'EEE': case 'E': out.push(DAYS[d.getDay()].slice(0, 3)); break;
      case 'HH': out.push(p2(d.getHours())); break;
      case 'H': out.push(d.getHours()); break;
      case 'hh': out.push(p2(d.getHours() % 12 || 12)); break;
      case 'h': out.push(d.getHours() % 12 || 12); break;
      case 'mm': out.push(p2(d.getMinutes())); break;
      case 'm': out.push(d.getMinutes()); break;
      case 'ss': out.push(p2(d.getSeconds())); break;
      case 's': out.push(d.getSeconds()); break;
      case 'SSS': out.push(String(d.getMilliseconds()).padStart(3, '0')); break;
      case 'a': out.push(d.getHours() < 12 ? 'AM' : 'PM'); break;
      case 'Z': { const o = -d.getTimezoneOffset(); out.push((o >= 0 ? '+' : '-') + p2(Math.floor(Math.abs(o) / 60)) + p2(Math.abs(o) % 60)); break; }
      case 'z': out.push(/\((.*)\)/.test(d.toString()) ? d.toString().match(/\(([^)]*)\)/)[1].replace(/[^A-Z]/g, '') : 'ET'); break;
      default: out.push(t);
    }
  }
  return out.join('');
}

function makeBuilder(kind, allowed) {
  const calls = [];
  const b = {};
  // Ranges are kept as A1 text (also inside arrays, as in setRanges([range])), so a rule
  // can be saved and never holds on to a live sheet.
  const plain = (a) => (a instanceof Range ? a.getA1Notation() : Array.isArray(a) ? a.map(plain) : a);
  for (const m of allowed) b[m] = (...args) => { calls.push([m, args.map(plain)]); return b; };
  b.build = () => ({ kind, calls: calls.slice() });
  return b;
}

function textOutput(s) {
  return {
    content: String(s == null ? '' : s), mime: 'TEXT',
    setMimeType(m) { this.mime = m; return this; }, getMimeType() { return this.mime; },
    getContent() { return this.content; }, setContent(c) { this.content = String(c); return this; }, append(c) { this.content += String(c); return this; },
    downloadAsFile(name) { this.filename = name; return this; },
  };
}
function htmlOutput(html) {
  return {
    isHtmlOutput: true, html: String(html == null ? '' : html), title: '', metas: [], faviconUrl: '',
    setTitle(t) { this.title = String(t); return this; }, getTitle() { return this.title; },
    addMetaTag(name, content) { this.metas.push([name, content]); return this; }, getMetaTags() { return this.metas.slice(); },
    setXFrameOptionsMode() { return this; }, setSandboxMode() { return this; }, setFaviconUrl(u) { this.faviconUrl = u; return this; },
    setWidth() { return this; }, setHeight() { return this; }, getWidth() { return 600; }, getHeight() { return 400; },
    getContent() { return this.html; }, setContent(c) { this.html = String(c); return this; },
    append(c) { this.html += String(c); return this; }, appendUntrusted(c) { this.html += String(c).replace(/</g, '&lt;'); return this; },
    getBlob() { return new Blob(toSigned(Buffer.from(this.html, 'utf8')), 'text/html', 'page.html'); },
  };
}

// Everything a script sees as global. A fresh set is built for every execution, so
// state between runs lives only in sheets, Drive, properties and the cache, as on Google.
function makeServices(account, project, run) {
  const st = account.projectState(project.key);
  const bound = () => (project.boundId ? account.ss(project.boundId) : null);
  const log = (line) => { const s = String(line); run.logs.push(s); st.logs.push({ at: new Date().toISOString(), text: s }); if (st.logs.length > 400) st.logs.splice(0, st.logs.length - 400); };

  const Button = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' };
  const ButtonSet = { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' };
  const ui = {
    Button, ButtonSet,
    alert(a, b, c) {
      const hasTitle = typeof c === 'string' || (c === undefined && typeof b === 'string' && !ButtonSet[b]);
      const title = hasTitle ? a : '', text = hasTitle ? b : a, buttons = (hasTitle ? c : b) || ButtonSet.OK;
      const no = run.answer === 'no';
      const answer = buttons === ButtonSet.YES_NO || buttons === ButtonSet.YES_NO_CANCEL ? (no ? Button.NO : Button.YES) : (buttons === ButtonSet.OK_CANCEL ? (no ? Button.CANCEL : Button.OK) : Button.OK);
      run.ui.push({ kind: 'alert', title: String(title || ''), text: String(text || ''), buttons, answer });
      return answer;
    },
    prompt(a, b, c) {
      const hasTitle = typeof b === 'string' && !ButtonSet[b];
      const title = hasTitle ? a : '', text = hasTitle ? b : a;
      const no = run.answer === 'no';
      run.ui.push({ kind: 'prompt', title: String(title), text: String(text), answer: no ? 'CANCEL' : 'OK', typed: run.promptText || '' });
      return { getResponseText: () => run.promptText || '', getSelectedButton: () => (no ? Button.CANCEL : Button.OK) };
    },
    createMenu(name) {
      const menu = { name, items: [] };
      const api = {
        addItem(label, fn) { menu.items.push({ label, fn }); return api; },
        addSeparator() { menu.items.push({ sep: true }); return api; },
        addSubMenu(sub) { (sub.__menu.items || []).forEach((it) => menu.items.push(it)); return api; },
        addToUi() { run.menus.push(menu); },
        __menu: menu,
      };
      return api;
    },
    createAddonMenu() { return ui.createMenu('Add-ons'); },
    showModalDialog(out, title) { run.ui.push({ kind: 'dialog', title: String(title || ''), html: out && out.html !== undefined ? out.html : String(out) }); },
    showModelessDialog(out, title) { ui.showModalDialog(out, title); },
    showSidebar(out) { run.ui.push({ kind: 'dialog', title: out && out.title ? out.title : 'Sidebar', html: out && out.html !== undefined ? out.html : String(out) }); },
  };

  const SpreadsheetApp = {
    getActive: () => bound(), getActiveSpreadsheet: () => bound(),
    getActiveSheet: () => { const ss = bound(); return ss ? ss.getActiveSheet() : null; },
    openById: (id) => {
      const ss = account.ss(String(id));
      if (!ss) throw new Error('Exception: Unexpected error while getting the method or property openById on object SpreadsheetApp. (The sandbox has no spreadsheet with id "' + id + '".)');
      return ss;
    },
    openByUrl: (u) => { const m = /\/d\/([-\w]+)/.exec(String(u)); return SpreadsheetApp.openById(m ? m[1] : u); },
    create: (name) => account.addSpreadsheet({ name }),
    flush: () => {},
    getUi: () => {
      if (!project.boundId) throw new Error('Exception: Cannot call SpreadsheetApp.getUi() from this context.');
      return ui;
    },
    newDataValidation: () => makeBuilder('dv', ['requireValueInRange', 'requireValueInList', 'requireDate', 'requireDateAfter', 'requireDateBefore', 'requireNumberBetween', 'requireNumberGreaterThan', 'requireNumberGreaterThanOrEqualTo', 'requireNumberLessThan', 'requireCheckbox', 'requireTextIsEmail', 'requireTextIsUrl', 'requireFormulaSatisfied', 'setAllowInvalid', 'setHelpText']),
    newConditionalFormatRule: () => makeBuilder('cf', ['whenFormulaSatisfied', 'whenTextEqualTo', 'whenTextContains', 'whenCellEmpty', 'whenCellNotEmpty', 'whenNumberGreaterThan', 'whenNumberLessThan', 'whenNumberEqualTo', 'whenDateBefore', 'whenDateAfter', 'setBackground', 'setFontColor', 'setBold', 'setItalic', 'setStrikethrough', 'setRanges']),
    BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK', DOTTED: 'DOTTED', DASHED: 'DASHED' },
    ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
    WrapStrategy: { WRAP: 'WRAP', CLIP: 'CLIP', OVERFLOW: 'OVERFLOW' },
    DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST', VALUE_IN_RANGE: 'VALUE_IN_RANGE', CHECKBOX: 'CHECKBOX' },
  };

  const folderApi = (id) => { const f = account.drive.folders[id]; if (!f) throw new Error('Exception: No item with the given ID could be found. Possibly because you have not edited this item or you do not have permission to access it.'); return f; };
  const DriveApp = {
    getRootFolder: () => account.root,
    getFolderById: (id) => folderApi(String(id)),
    getFileById: (id) => {
      id = String(id);
      if (account.drive.files[id]) return account.drive.files[id];
      if (account.forms[id]) { const form = account.forms[id]; return { getId: () => id, getName: () => form.title, setTrashed(t) { form.trashed = !!t; return this; }, isTrashed: () => !!form.trashed, getParents: () => iter([account.root]) }; }
      if (account.spreadsheets[id]) {   // a spreadsheet is a Drive file too; it lives in My Drive
        const ss = account.spreadsheets[id];
        return { getId: () => id, getName: () => ss.title, getMimeType: () => 'application/vnd.google-apps.spreadsheet', getParents: () => iter([account.root]), getUrl: () => ss.getUrl(), setTrashed() { return this; }, isTrashed: () => false, setSharing() { return this; } };
      }
      throw new Error('Exception: No item with the given ID could be found. Possibly because you have not edited this item or you do not have permission to access it.');
    },
    getFoldersByName: (name) => iter(Object.values(account.drive.folders).filter((f) => f.name === name && !f.trashed)),
    getFilesByName: (name) => iter(Object.values(account.drive.files).filter((f) => f.name === name && !f.trashed)),
    createFolder: (name) => new DFolder(account, name, account.root.id),
    createFile: (blob, content, mime) => account.root.createFile(blob, content, mime),
    Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', DOMAIN: 'DOMAIN', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', PRIVATE: 'PRIVATE' },
    Permission: { VIEW: 'VIEW', EDIT: 'EDIT', COMMENT: 'COMMENT', OWNER: 'OWNER', NONE: 'NONE' },
  };

  const mailer = (kind) => ({
    sendEmail(a, subject, body, opts) {
      if (a && typeof a === 'object') { const m = a; return account.sendMail(project.key, kind, m.to, m.subject, m.body, m); }
      if (!a) throw new Error('Exception: Invalid argument: recipient');
      account.sendMail(project.key, kind, a, subject, body, opts);
    },
    getRemainingDailyQuota: () => 100,
  });

  const cacheStore = st.cache;
  const cache = {
    get: (k) => { const e = cacheStore[k]; if (!e) return null; if (e.exp < Date.now()) { delete cacheStore[k]; return null; } return e.v; },
    getAll: (keys) => { const out = {}; keys.forEach((k) => { const v = cache.get(k); if (v !== null) out[k] = v; }); return out; },
    put: (k, v, secs) => { v = String(v); if (v.length > 100 * 1024) throw new Error('Exception: Argument too large: value'); cacheStore[k] = { v, exp: Date.now() + 1000 * Math.min(secs || 600, 21600) }; },
    putAll: (obj, secs) => { Object.keys(obj).forEach((k) => cache.put(k, obj[k], secs)); },
    remove: (k) => { delete cacheStore[k]; },
    removeAll: (keys) => { keys.forEach((k) => { delete cacheStore[k]; }); },
  };
  const props = {
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(st.props, k) ? st.props[k] : null),
    setProperty: (k, v) => { st.props[k] = String(v); account.dirty = true; return props; },
    deleteProperty: (k) => { delete st.props[k]; account.dirty = true; return props; },
    getProperties: () => Object.assign({}, st.props),
    setProperties: (obj, deleteOthers) => { if (deleteOthers) st.props = {}; Object.keys(obj).forEach((k) => { st.props[k] = String(obj[k]); }); account.dirty = true; return props; },
    deleteAllProperties: () => { Object.keys(st.props).forEach((k) => delete st.props[k]); account.dirty = true; return props; },
    getKeys: () => Object.keys(st.props),
  };

  const lock = () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true });

  const triggerObj = (t) => ({ getHandlerFunction: () => t.fn, getEventType: () => t.type, getUniqueId: () => t.id, getTriggerSource: () => (t.type === 'time' ? 'CLOCK' : 'SPREADSHEETS'), fn: t.fn, type: t.type, id: t.id });
  const ScriptApp = {
    getService: () => ({ getUrl: () => project.prodUrl || '', isEnabled: () => !!project.prodUrl }),
    getScriptId: () => 'sandbox-' + project.key,
    getProjectTriggers: () => st.triggers.map(triggerObj),
    getUserTriggers: () => st.triggers.map(triggerObj),
    deleteTrigger: (t) => { st.triggers = st.triggers.filter((x) => x.id !== (t && t.id)); account.dirty = true; },
    newTrigger: (fn) => {
      const t = { fn, type: '', id: 't' + (seq++), source: project.boundId || '' };
      const chain = {
        forSpreadsheet: (ss) => { t.source = typeof ss === 'string' ? ss : ss.getId(); return chain; },
        forForm: () => chain,
        onEdit: () => { t.type = 'edit'; return chain; },
        onChange: () => { t.type = 'change'; return chain; },
        onOpen: () => { t.type = 'open'; return chain; },
        onFormSubmit: () => { t.type = 'submit'; return chain; },
        timeBased: () => { t.type = 'time'; return chain; },
        everyDays: (n) => { t.every = n + ' days'; return chain; }, everyHours: (n) => { t.every = n + ' hours'; return chain; }, everyMinutes: (n) => { t.every = n + ' minutes'; return chain; }, everyWeeks: (n) => { t.every = n + ' weeks'; return chain; },
        atHour: (h) => { t.hour = h; return chain; }, nearMinute: () => chain, onWeekDay: () => chain, inTimezone: () => chain, at: () => chain, after: () => chain,
        create: () => { st.triggers.push(t); account.dirty = true; return triggerObj(t); },
      };
      return chain;
    },
    AuthMode: { FULL: 'FULL', LIMITED: 'LIMITED', NONE: 'NONE' },
    EventType: { ON_EDIT: 'ON_EDIT', ON_CHANGE: 'ON_CHANGE', ON_FORM_SUBMIT: 'ON_FORM_SUBMIT', CLOCK: 'CLOCK', ON_OPEN: 'ON_OPEN' },
    TriggerSource: { SPREADSHEETS: 'SPREADSHEETS', CLOCK: 'CLOCK', FORMS: 'FORMS' },
    WeekDay: { MONDAY: 'MONDAY', TUESDAY: 'TUESDAY', WEDNESDAY: 'WEDNESDAY', THURSDAY: 'THURSDAY', FRIDAY: 'FRIDAY', SATURDAY: 'SATURDAY', SUNDAY: 'SUNDAY' },
  };

  const Utilities = {
    formatDate,
    formatString: (fmt, ...args) => { let i = 0; return String(fmt).replace(/%[sd]/g, () => String(args[i++])); },
    getUuid: () => crypto.randomUUID(),
    sleep: () => {},
    base64Decode: (s) => { s = String(s); if (/[^A-Za-z0-9+/=\s]/.test(s)) throw new Error('Exception: Could not decode string.'); return toSigned(Buffer.from(s, 'base64')); },
    base64DecodeWebSafe: (s) => toSigned(Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64')),
    base64Encode: (x) => (typeof x === 'string' ? Buffer.from(x, 'utf8') : toBuffer(x)).toString('base64'),
    base64EncodeWebSafe: (x) => Utilities.base64Encode(x).replace(/\+/g, '-').replace(/\//g, '_'),
    newBlob: (data, mime, name) => {
      if (typeof data === 'string') return new Blob(toSigned(Buffer.from(data, 'utf8')), mime || 'text/plain', name);
      if (!Array.isArray(data)) throw new Error('Exception: The parameters don\'t match the method signature for Utilities.newBlob.');
      return new Blob(data, mime, name);
    },
    computeDigest: (alg, value) => {
      const algo = { SHA_256: 'sha256', SHA_1: 'sha1', MD5: 'md5', SHA_512: 'sha512', SHA_384: 'sha384' }[alg] || 'sha256';
      return toSigned(crypto.createHash(algo).update(typeof value === 'string' ? Buffer.from(value, 'utf8') : toBuffer(value)).digest());
    },
    computeHmacSha256Signature: (value, key) => toSigned(crypto.createHmac('sha256', String(key)).update(String(value)).digest()),
    DigestAlgorithm: { SHA_256: 'SHA_256', SHA_1: 'SHA_1', MD5: 'MD5', SHA_512: 'SHA_512', SHA_384: 'SHA_384', MD2: 'MD2' },
    Charset: { UTF_8: 'UTF_8', US_ASCII: 'US_ASCII' },
    parseCsv: (csv) => String(csv).split(/\r?\n/).filter((l) => l !== '').map((l) => l.split(',')),
    jsonParse: (s) => JSON.parse(s),
    jsonStringify: (o) => JSON.stringify(o),
    unzip: () => { throw new Error('Exception: Utilities.unzip is not available in the sandbox (nothing here downloads real exports).'); },
    zip: () => { throw new Error('Exception: Utilities.zip is not available in the sandbox.'); },
  };

  const UrlFetchApp = {
    fetch: (url, params) => {
      params = params || {};
      url = String(url);
      const method = String(params.method || 'get').toLowerCase();
      if (!/^https:\/\/fabman\.io\/api\/v1\//.test(url)) {
        throw new Error('Exception: The sandbox does not reach the internet, so this request was stopped: ' + method.toUpperCase() + ' ' + url.replace(/\?.*$/, ''));
      }
      const res = account.fabman.handle(method, url, params.payload, params.headers);
      log('[fake Fabman] ' + method.toUpperCase() + ' ' + url.replace('https://fabman.io', '') + ' -> ' + res.code);
      account.dirty = true;
      if (res.code >= 400 && !params.muteHttpExceptions) {
        throw new Error('Exception: Request failed for https://fabman.io returned code ' + res.code + '. Truncated server response: ' + res.text.slice(0, 120) + ' (use muteHttpExceptions option to examine full response)');
      }
      return {
        getResponseCode: () => res.code,
        getContentText: () => res.text,
        getAllHeaders: () => Object.assign({}, res.headers),
        getHeaders: () => Object.assign({}, res.headers),
        getBlob: () => new Blob(toSigned(Buffer.from(res.text, 'utf8')), 'application/json', 'response.json'),
        getContent: () => toSigned(Buffer.from(res.text, 'utf8')),
      };
    },
    fetchAll: (reqs) => reqs.map((r) => UrlFetchApp.fetch(typeof r === 'string' ? r : r.url, r)),
  };

  const Session = {
    getScriptTimeZone: () => 'America/New_York',
    getActiveUser: () => user(account.user),
    getEffectiveUser: () => user(account.user),
    getTemporaryActiveUserKey: () => 'sandbox-user',
  };

  const FormApp = {
    openById: (id) => { const f = account.forms[id]; if (!f) throw new Error('Exception: No item with the given ID could be found.'); return f; },
    openByUrl: (u) => { const f = Object.values(account.forms).find((x) => x.getEditUrl() === u); if (!f) throw new Error('Exception: No form at ' + u); return f; },
    create: (title) => { const f = new Form(account, title); account.forms[f.id] = f; return f; },
    DestinationType: { SPREADSHEET: 'SPREADSHEET' },
    ItemType: Object.assign({}, ...Object.values(ITEM_TYPES).map((t) => ({ [t]: t }))),
    EmailCollectionType: { VERIFIED: 'VERIFIED', RESPONDER_INPUT: 'RESPONDER_INPUT', DO_NOT_COLLECT: 'DO_NOT_COLLECT' },
  };

  return {
    SpreadsheetApp, DriveApp, ScriptApp, Utilities, UrlFetchApp, Session, FormApp,
    GmailApp: Object.assign(mailer('Gmail'), { getAliases: () => [] }),
    MailApp: mailer('Mail'),
    HtmlService: {
      createHtmlOutput: (html) => htmlOutput(html && html.getDataAsString ? html.getDataAsString() : html),
      createTemplate: () => { throw new Error('sandbox: HtmlService templates are not used by the hub; add them to sandbox/lib/gas.js if you need them'); },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
      SandboxMode: { IFRAME: 'IFRAME', NATIVE: 'NATIVE', EMULATED: 'EMULATED' },
    },
    ContentService: { createTextOutput: textOutput, MimeType: { JAVASCRIPT: 'JAVASCRIPT', JSON: 'JSON', TEXT: 'TEXT', CSV: 'CSV', XML: 'XML', ICAL: 'ICAL', RSS: 'RSS', ATOM: 'ATOM', VCARD: 'VCARD' } },
    CacheService: { getScriptCache: () => cache, getUserCache: () => cache, getDocumentCache: () => cache },
    PropertiesService: { getScriptProperties: () => props, getUserProperties: () => props, getDocumentProperties: () => props },
    LockService: { getScriptLock: lock, getDocumentLock: lock, getUserLock: lock },
    Logger: { log: (fmt, ...args) => { let i = 0; log(args.length ? String(fmt).replace(/%s/g, () => String(args[i++])) : fmt); }, clear: () => {}, getLog: () => run.logs.join('\n') },
    console: { log: (...a) => log(a.join(' ')), info: (...a) => log(a.join(' ')), warn: (...a) => log('WARNING: ' + a.join(' ')), error: (...a) => log('ERROR: ' + a.join(' ')) },
    Date,
  };
}

module.exports = { Account, Spreadsheet, Sheet, Range, Blob, DFile, DFolder, FakeFabman, makeServices, formatDate, driveId, toSigned, toBuffer, colLetters, displayValue };
