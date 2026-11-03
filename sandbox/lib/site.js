'use strict';
// What the sandbox needs to know about the real platform, read from the repo itself so
// it follows along when someone changes a URL: the deployed web app URLs (to swap for
// local ones) and the spreadsheet ids the scripts open.

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const HUB_URL = 'https://cornellprojectteams.github.io/centralhub/';

// Used only if a script cannot be read (say, mid-edit with a syntax error).
const FALLBACK_IDS = {
  space: '1mZrlnA-GiVKB4_Um21aMH9jsSkf6TUJhauDWivnf7-I',
  ell: '1yazhxA8nwYWDRSdQEN6Cc3UKKnPn6PWF_b3zAe0KoQk',
  fabman: '1lDAA4z1YexkjaABSPuXJyEXfSxN2QAsKWVk88TQCRLI',
  master: '1k97Ey5aNcm58zaviaXxKvJtTgL1yeSI_fvUkXZFEyM4',
  drivers: '1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M',
  oldInventory: '1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU',
  workspaces: '1cKd-5815CafamF2AQpwTIqATjdkD4HgM404vQvwTvuM',
  summer: '1UrI5mBI4rqY9NeElobhzdd1lWsAlBDeYa1LZPVWHNnE',
};

const idFromUrl = (u) => { const m = /\/d\/([-\w]{20,})/.exec(String(u || '')); return m ? m[1] : ''; };

function readConfig(repo) {
  const window = {};
  try { vm.runInNewContext(fs.readFileSync(path.join(repo, 'config.js'), 'utf8'), { window }, { timeout: 1000 }); }
  catch (err) { throw new Error('Could not read config.js: ' + err.message); }
  const appJs = fs.readFileSync(path.join(repo, 'app.js'), 'utf8');
  const stats = /STATS_ENDPOINT\s*=\s*'([^']+)'/.exec(appJs);
  const inv = window.INVENTORY || {};
  return {
    spaceUrl: window.SPACE_STATUS_URL || '',
    statsUrl: stats ? stats[1] : '',
    inventoryUrl: inv.apiUrl || '',
    inventoryAppUrl: inv.appUrl || '',
    inventorySheetUrl: inv.sheetUrl || '',
    inventoryGids: { dash: inv.dashboardGid, master: inv.masterGid, maint: inv.maintenanceGid },
    people: window.PEOPLE_SHEET || {},
    schedule: window.SCHEDULE_SHEET || {},
  };
}

// peek(key, expr) reads a constant out of a project (see Runner.peek).
function readIds(cfg, peek) {
  const sources = (() => { try { return JSON.parse(peek('command', 'JSON.stringify(SOURCES)') || '{}'); } catch (e) { return {}; } })();
  return {
    space: peek('space', 'CONFIG.spreadsheetId') || sources.spaceStatus || FALLBACK_IDS.space,
    inventory: idFromUrl(cfg.inventorySheetUrl) || 'sandbox-inventory-sheet',
    inventoryGids: cfg.inventoryGids || {},   // so the admin page's links to a tab land on it
    ell: sources.ell || FALLBACK_IDS.ell,
    fabman: sources.fabman || FALLBACK_IDS.fabman,
    master: sources.master || FALLBACK_IDS.master,
    drivers: peek('fleet', 'FLEET_SS_ID') || sources.drivers || FALLBACK_IDS.drivers,
    oldInventory: sources.inventory || FALLBACK_IDS.oldInventory,
    workspaces: FALLBACK_IDS.workspaces,
    summer: FALLBACK_IDS.summer,
    people: (cfg.people && cfg.people.id) || 'sandbox-people-sheet',
    peopleGid: Number((cfg.people && cfg.people.gid) || 0) || 983316085,
    scheduleGid: Number((cfg.schedule && cfg.schedule.gid) || 0) || 1799869087,
  };
}

module.exports = { readConfig, readIds, idFromUrl, HUB_URL };
