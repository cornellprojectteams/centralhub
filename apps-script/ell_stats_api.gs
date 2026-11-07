/**
 * ELL Shift Stats API  (its own small, standalone Apps Script project)
 * =============================================================================
 * Returns every row of the ELL Shift Summary (Responses) sheet as JSONP for
 * ell-stats.html. It runs as whoever deploys it, so the sheet can stay shared
 * "Restricted" — the deployer just needs at least VIEW access to the sheet.
 * =============================================================================
 */

const ELL_SHEET_ID = '1yazhxA8nwYWDRSdQEN6Cc3UKKnPn6PWF_b3zAe0KoQk'; // ELL Shift Summary (Responses)
const ELL_TAB = 'Form Responses 1';
const ELL_CACHE_SECONDS = 180;

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const json = ellRowsJson_(p.fresh === '1');
  // Only allow a plain JS identifier as the callback name.
  if (p.callback && /^[A-Za-z_$][\w$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Rows shaped like the Google Sheets gviz response
 * ({ table: { cols:[{label}], rows:[{c:[{v}]}] } }) so ell-stats.html parses it
 * the same way as a link-shared sheet. Dates go out as "Date(y,m,d,h,mi,s)".
 */
function ellRowsJson_(fresh) {
  const cache = CacheService.getScriptCache();
  const key = 'ellrows_v1';
  if (!fresh) {
    const hit = cache.get(key);
    if (hit) return hit;
  }
  const ss = SpreadsheetApp.openById(ELL_SHEET_ID);
  const sh = ss.getSheetByName(ELL_TAB) || ss.getSheets()[0];
  const v = sh.getDataRange().getValues();
  const cell = x => {
    if (x instanceof Date && !isNaN(x.getTime())) {
      return { v: 'Date(' + [x.getFullYear(), x.getMonth(), x.getDate(), x.getHours(), x.getMinutes(), x.getSeconds()].join(',') + ')' };
    }
    return x === '' || x == null ? null : { v: x };
  };
  const json = JSON.stringify({
    table: {
      cols: (v[0] || []).map(h => ({ label: String(h) })),
      rows: v.slice(1)
        .filter(r => r.some(x => x !== '' && x != null))
        .map(r => ({ c: r.map(cell) })),
    },
  });
  // CacheService caps a value at 100 KB; past that, just serve uncached.
  if (json.length < 95000) {
    try { cache.put(key, json, ELL_CACHE_SECONDS); } catch (err) { /* ignore */ }
  }
  return json;
}

/** Run this from the editor to authorize and sanity-check. */
function testEllRows() {
  const out = JSON.parse(ellRowsJson_(true));
  Logger.log('OK: ' + out.table.rows.length + ' shift rows; columns: ' + out.table.cols.map(c => c.label).join(' | '));
}
