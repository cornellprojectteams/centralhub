'use strict';
// Keeps everything inside the sandbox. Every page, script and email the sandbox serves
// passes through here, and each address that would reach the real platform (the web
// apps, Drive, Sheets, the forms, Fabman) is swapped for its sandbox stand-in. The same
// list goes to the browser (public/guard.js) to catch addresses built while a page runs.
// The files on disk are never changed.

const { HUB_URL } = require('./site');

function makeRewriter(cfg) {
  const apps = [
    [cfg.spaceUrl, '/sandbox/space/exec'],
    [cfg.statsUrl, '/sandbox/command/exec'],
    [cfg.inventoryUrl, '/sandbox/inventory/exec'],
  ].filter((p) => p[0]);
  const google = [
    ['https://drive.google.com/thumbnail', '/sandbox/drive/thumbnail'],
    ['https://drive.google.com/file/d/', '/sandbox/drive/file/'],
    ['https://drive.google.com/open?id=', '/sandbox/drive/open?id='],
    ['https://drive.google.com/uc?', '/sandbox/drive/uc?'],
    ['https://docs.google.com/spreadsheets/d/', '/sandbox/sheets/'],
    ['https://forms.gle/6DHYsEPonVtdALps5', '/sandbox/forms/space-status'],
    ['https://docs.google.com/forms/d/1bVzxR-lGNDtll6wh7wsdZGsVuqMkaRavVcGxykLHMFw/edit', '/sandbox/forms/space-status?editor=1'],
    ['https://docs.google.com/forms/d/e/1FAIpQLSfUSbqVXJ8si5c-efXQl8k5LAvvhkTSrSVPWlNCaKUEbNGgog/viewform', '/sandbox/forms/ell'],
    ['https://fabman.io/manage/2366', '/sandbox/fabman'],
    ['https://script.google.com/d/1DpKw0dXIQ5rH6_BQfIQnXpmTbTcPmfH3wq2Dzc6pJUA7lrmeFG7hh5P4/edit', '/sandbox/code/space'],
    ['https://script.google.com/home', '/sandbox/code'],
  ];
  // Links the scripts build to the live hub (an inventory item, "back to the hub") open
  // the sandbox copy instead. The site's own pages keep their links to the live site,
  // which is public and read-only.
  const hub = [[HUB_URL, '/']];

  const swap = (pairs) => (text) => {
    let out = String(text);
    pairs.forEach(([from, to]) => { if (out.indexOf(from) >= 0) out = out.split(from).join(to); });
    return out;
  };
  return {
    pairs: apps.concat(google),
    files: swap(apps.concat(google)),        // the site's HTML, CSS and JS
    gas: swap(apps.concat(google, hub)),     // Apps Script pages, google.script.run results, emails
    data: swap(apps.concat(hub)),            // JSON from the inventory and Command Center APIs: photo links stay Google-shaped, the app reads them
  };
}

// Script and style tags that go at the very top of <head>, before the page's own code.
function headTags(rewriter, opts) {
  opts = opts || {};
  const map = JSON.stringify(rewriter.pairs).replace(/</g, '\\u003c');
  return '<script>window.__SANDBOX_MAP__=' + map + ';</script>'
    + '<script src="/sandbox/static/guard.js"></script>'
    + (opts.gas ? '<script src="/sandbox/static/gas-shim.js" data-run="' + opts.gas + '"></script>' : '');
}

function inject(html, tags) {
  html = String(html);
  const m = /<head[^>]*>/i.exec(html);
  if (m) return html.slice(0, m.index + m[0].length) + tags + html.slice(m.index + m[0].length);
  return tags + html;
}

module.exports = { makeRewriter, headTags, inject };
