'use strict';
// Runs the hub's Apps Script projects straight from the repo's apps-script folder.
//
// Like Apps Script itself, every execution starts from a fresh set of globals: the
// project's files run top to bottom, then the requested function is called. Anything
// that should last between executions has to live in a sheet, Drive, the properties or
// the cache, exactly as on Google. Edit a .gs file and the next execution uses it.

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeServices } = require('./gas');

const TIMEOUT_MS = 60000;   // Apps Script allows six minutes; nothing here should take one.

// The projects, where their code lives, and which sheet each one is bound to.
// ids and urls are filled in by server.js from config.js, app.js and the scripts.
function defineProjects(ids, urls) {
  return {
    space: {
      key: 'space', name: 'Space Status web app',
      blurb: 'Report emails, reminders, the issue dashboards and team pages, projects and proposals.',
      files: ['apps-script/02_notify_on_submit.gs', 'apps-script/04_tasks_projects.gs', 'apps-script/07_proposals.gs'],
      prodUrl: urls.space, local: '/sandbox/space/exec', webapp: true,
      editor: 'https://script.google.com/d/1DpKw0dXIQ5rH6_BQfIQnXpmTbTcPmfH3wq2Dzc6pJUA7lrmeFG7hh5P4/edit',
      dataSheet: ids.space,
    },
    inventory: {
      key: 'inventory', name: 'Inventory',
      blurb: 'Keeps the inventory sheet tidy and answers the inventory app.',
      files: ['apps-script/08_inventory.gs'], boundId: ids.inventory,
      prodUrl: urls.inventory, local: '/sandbox/inventory/exec', webapp: true,
    },
    command: {
      key: 'command', name: 'Command Center API',
      blurb: 'Reads the source sheets and hands the Command Center its numbers.',
      files: ['apps-script/command-center/Code.gs'],
      prodUrl: urls.command, local: '/sandbox/command/exec', webapp: true,
    },
    fabman: {
      key: 'fabman', name: 'Fabman members',
      blurb: 'The buttons in the Fabman members sheet. Talks to the sandbox\'s fake Fabman.',
      files: ['apps-script/05_fabman_members.gs'], boundId: ids.fabman,
      // The real sheet runs these from drawings with a script assigned. Grouped by tab.
      buttons: {
        Dashboard: ['fetchFabmanMembers', 'updateFabmanMembers'],
        Actions: ['postMachines', 'postMills', 'postRed', 'postGreen', 'postBlue', 'postSupervisor', 'deleteSupervisor', 'lockMembers', 'activateMembers', 'deleteMembers', 'invitation', 'disclaimer'],
        Graduates: ['flagGraduates', 'processGraduates'],
      },
    },
    fleet: {
      key: 'fleet', name: 'Fleet driver cleanup',
      blurb: 'The Driver cleanup menu in the student drivers sheet.',
      files: ['apps-script/fleet_driver_cleanup.gs'], boundId: ids.drivers,
    },
    contacts: {
      key: 'contacts', name: 'Team Contacts setup',
      blurb: 'One-off setup that built the Team Contacts tab. Safe to run again.',
      files: ['apps-script/01_setup_team_contacts.gs'], boundId: ids.space,
    },
  };
}

class Runner {
  constructor(repo, account, projects) {
    this.repo = repo;
    this.account = account;
    this.projects = projects;
    this.compiled = {};   // file -> { stamp, script }
    this.problems = [];   // recent errors, newest last
  }

  // Compile a file once per version. A syntax error is reported with its file and line.
  full(file) { return path.isAbsolute(file) ? file : path.join(this.repo, file); }

  script(file) {
    const full = this.full(file);
    const st = fs.statSync(full);
    const stamp = st.mtimeMs + ':' + st.size;
    const hit = this.compiled[file];
    if (hit && hit.stamp === stamp) return hit.script;
    const src = fs.readFileSync(full, 'utf8');
    try {
      const script = new vm.Script(src, { filename: file });
      this.compiled[file] = { stamp, script };
      return script;
    } catch (err) {
      const where = /:(\d+)\s*\n/.exec(String(err.stack)) || [];
      const e = new Error('SyntaxError: ' + err.message + ' (file ' + path.basename(file) + (where[1] ? ', line ' + where[1] : '') + ')');
      e.syntax = true;
      throw e;
    }
  }

  functionsOf(key) {
    const p = this.projects[key];
    const out = [];
    p.files.forEach((file) => {
      const src = fs.readFileSync(this.full(file), 'utf8');
      const re = /^function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
      let m;
      while ((m = re.exec(src)) !== null) out.push({ name: m[1], file, private: /_$/.test(m[1]) });
    });
    return out;
  }

  // Run one function in a fresh copy of the project.
  //   opts.answer      'yes' | 'no': how to answer the script's dialogs (alert, prompt)
  //   opts.promptText  what to type into a prompt
  //   opts.client      true for google.script.run calls (private_ functions are refused)
  exec(key, fn, args, opts) {
    opts = opts || {};
    const project = this.projects[key];
    if (!project) throw new Error('No project "' + key + '"');
    const run = { project: key, fn, logs: [], ui: [], menus: [], mails: [], answer: opts.answer || 'yes', promptText: opts.promptText || '', started: Date.now() };
    const result = { ok: false, value: undefined, run };
    const prev = this.account.run;
    this.account.run = run;
    try {
      const scripts = project.files.map((f) => this.script(f));
      const ctx = vm.createContext(makeServices(this.account, project, run));
      scripts.forEach((s) => s.runInContext(ctx, { timeout: TIMEOUT_MS }));
      if (opts.client && /_$/.test(fn)) throw new Error('Script function not found: ' + fn);
      if (!/^[A-Za-z_$][\w$]*$/.test(fn) || typeof ctx[fn] !== 'function') throw new Error('Script function not found: ' + fn);
      ctx.__sandboxArgs = args || [];
      result.value = vm.runInContext(fn + '.apply(null, __sandboxArgs)', ctx, { timeout: TIMEOUT_MS });
      result.ok = true;
    } catch (err) {
      result.error = this.describe(err);
      result.stack = String((err && err.stack) || '');
      this.problems.push({ at: new Date().toISOString(), project: key, fn, error: result.error });
      if (this.problems.length > 50) this.problems.shift();
    } finally {
      run.ms = Date.now() - run.started;
      this.account.run = prev;
    }
    return result;
  }

  // "Exception: ... (line 12, file 02_notify_on_submit)" like the Apps Script log.
  describe(err) {
    const msg = String((err && err.message) || err);
    if (err && err.syntax) return msg;
    const frame = /\(?((?:apps-script\/)[^():]+\.gs):(\d+):\d+\)?/.exec(String((err && err.stack) || ''));
    const timeout = /Script execution timed out/.test(msg) ? ' The sandbox stops a run after ' + TIMEOUT_MS / 1000 + ' seconds; look for a loop that never ends.' : '';
    return msg + timeout + (frame ? ' (line ' + frame[2] + ', file ' + path.basename(frame[1], '.gs') + ')' : '');
  }

  // Read a constant out of a project without calling anything, e.g. peek('space', 'CONFIG.spreadsheetId').
  peek(key, expr) {
    try {
      const project = this.projects[key];
      const run = { logs: [], ui: [], menus: [], mails: [], answer: 'yes' };
      const ctx = vm.createContext(makeServices(this.account, project, run));
      project.files.forEach((f) => this.script(f).runInContext(ctx, { timeout: 5000 }));
      return vm.runInContext(expr, ctx, { timeout: 1000 });
    } catch (err) {
      return undefined;
    }
  }

  // The event Apps Script hands doGet / doPost.
  static webEvent(query, body, contentType) {
    const parameter = {}, parameters = {};
    for (const [k, v] of query) { if (!(k in parameter)) parameter[k] = v; (parameters[k] = parameters[k] || []).push(v); }
    const e = { parameter, parameters, queryString: query.toString(), contextPath: '', contentLength: -1 };
    if (body !== undefined) {
      e.postData = { contents: body, length: Buffer.byteLength(body), type: contentType || 'text/plain', name: 'postData' };
      e.contentLength = e.postData.length;
      if (/application\/x-www-form-urlencoded/.test(contentType || '')) {
        for (const [k, v] of new URLSearchParams(body)) { if (!(k in parameter)) parameter[k] = v; (parameters[k] = parameters[k] || []).push(v); }
      }
    }
    return e;
  }

  // Fire what Google would fire for a change to a bound sheet: installable triggers of
  // that type, then the simple trigger (onEdit / onOpen) if the project has one.
  fire(ssId, type, event) {
    const results = [];
    Object.values(this.projects).forEach((p) => {
      if (p.boundId !== ssId) return;
      const st = this.account.projectState(p.key);
      st.triggers.filter((t) => t.type === type && (!t.source || t.source === ssId)).forEach((t) => {
        results.push(Object.assign({ project: p.key, fn: t.fn }, this.exec(p.key, t.fn, [event])));
      });
      const simple = { edit: 'onEdit', open: 'onOpen' }[type];
      if (simple && this.functionsOf(p.key).some((f) => f.name === simple)) {
        // A simple trigger runs only if no installable trigger already called the same function.
        if (!st.triggers.some((t) => t.type === type && t.fn === simple)) results.push(Object.assign({ project: p.key, fn: simple }, this.exec(p.key, simple, [event])));
      }
    });
    return results;
  }
}

module.exports = { Runner, defineProjects };
