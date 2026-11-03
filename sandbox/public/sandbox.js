/*
 * Buttons and forms on the sandbox's own pages (control panel, sheets, scripts, forms).
 * The server does the work; this file sends the clicks and shows what came back.
 *
 *   [data-api="/sandbox/api/..."]          POST, show the result (data-out) or reload (data-reload)
 *   [data-confirm="Press again to ..."]    needs a second press within a few seconds
 *   [data-run="project"][data-fn="name"]   run one Apps Script function
 *   table.sbx-grid                         click a cell to edit it
 *   form.sbx-form                          the sandbox copies of the Google Forms
 *   input[data-try]                        "Things to try" ticks, kept in this browser
 */
(function () {
  'use strict';

  function post(url, body) {
    return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) })
      .then(function (r) {
        return r.text().then(function (t) {
          try { return JSON.parse(t); } catch (e) { return { ok: false, error: 'The sandbox server answered with something unexpected (HTTP ' + r.status + ').', html: '' }; }
        });
      })
      .catch(function () { return { ok: false, error: 'Could not reach the sandbox server. Is it still running? (node sandbox/server.js)' }; });
  }
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  // Show a run result in its box, with a close button.
  function show(box, html) {
    if (!box) return;
    box.innerHTML = '<button type="button" class="sbx-result-close" aria-label="Close">&times;</button>' + html;
    box.hidden = false;
    box.querySelector('.sbx-result-close').addEventListener('click', function () { box.hidden = true; });
    var r = box.getBoundingClientRect();
    if (r.top < 0 || r.top > window.innerHeight - 80) box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  function failHtml(res) { return '<p class="sbx-run-head is-bad">Failed</p><pre class="sbx-err">' + esc(res.error || 'Unknown error') + '</pre>'; }
  function answer() { var a = document.querySelector('input[name="sbx-answer"]:checked'); return a ? a.value : 'yes'; }
  function promptText() { var t = document.querySelector('input[name="sbx-prompt"]'); return t ? t.value : ''; }

  // Swap parts of the page for fresh copies after something changed (the sheet grid,
  // the outbox count), without losing the result box or the scroll position.
  function refresh(selectors) {
    return fetch(location.href, { cache: 'no-store' }).then(function (r) { return r.text(); }).then(function (html) {
      var doc = new DOMParser().parseFromString(html, 'text/html');
      selectors.forEach(function (sel) {
        var now = document.querySelector(sel), next = doc.querySelector(sel);
        if (!now || !next) return;
        var keep = now.scrollLeft, keepTop = now.scrollTop;
        now.replaceWith(next);
        next.scrollLeft = keep; next.scrollTop = keepTop;
      });
    }).catch(function () {});
  }
  var SHEET_PARTS = ['.sbx-grid-wrap', '.sbx-tabs', '.sbx-sheet-buttons', '.sbx-nav'];
  var CODE_PARTS = ['.sbx-nav'];
  function partsForPage() { return document.querySelector('.sbx-grid') ? SHEET_PARTS : CODE_PARTS; }

  // Two-press buttons: the first press arms, the second does it.
  function confirmed(btn) {
    var text = btn.getAttribute('data-confirm');
    if (!text) return true;
    if (btn.classList.contains('is-armed')) { clearTimeout(btn._disarm); btn.classList.remove('is-armed'); btn.textContent = btn._label; return true; }
    btn._label = btn.textContent;
    btn.classList.add('is-armed');
    btn.textContent = text;
    btn._disarm = setTimeout(function () { btn.classList.remove('is-armed'); btn.textContent = btn._label; }, 4000);
    return false;
  }

  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('[data-api]');
    if (btn) {
      e.preventDefault();
      if (btn.classList.contains('is-busy') || !confirmed(btn)) return;
      btn.classList.add('is-busy'); btn.disabled = true;
      var payload = {};
      try { payload = JSON.parse(btn.getAttribute('data-body') || '{}'); } catch (err) { payload = {}; }
      post(btn.getAttribute('data-api'), payload).then(function (res) {
        btn.classList.remove('is-busy'); btn.disabled = false;
        if (btn.hasAttribute('data-reload') && res.ok !== false) { location.reload(); return; }
        var out = document.querySelector(btn.getAttribute('data-out') || '#act-out');
        show(out, res.html || (res.ok === false ? failHtml(res) : '<p class="sbx-run-head is-ok">Done</p>'));
        refresh(['.sbx-stats', '.sbx-nav', '.sbx-foot']);
      });
      return;
    }

    var run = e.target.closest && e.target.closest('[data-run][data-fn]');
    if (run) {
      e.preventDefault();
      if (run.classList.contains('is-busy')) return;
      var menu = run.closest('details'); if (menu && menu.classList.contains('sbx-menu')) menu.open = false;
      run.classList.add('is-busy');
      var grid = document.querySelector('.sbx-grid');
      post('/sandbox/api/run', { project: run.getAttribute('data-run'), fn: run.getAttribute('data-fn'), answer: answer(), promptText: promptText(),
        id: grid ? grid.getAttribute('data-ss') : '', sheet: grid ? grid.getAttribute('data-sheet') : '' }).then(function (res) {
        run.classList.remove('is-busy');
        show(document.getElementById('run-out'), res.html || failHtml(res));
        refresh(partsForPage());
      });
      return;
    }

    var del = e.target.closest && e.target.closest('.sbx-del');
    if (del) {
      e.preventDefault();
      if (!del.classList.contains('is-armed')) {
        del.classList.add('is-armed');
        setTimeout(function () { del.classList.remove('is-armed'); }, 3000);
        return;
      }
      var grid = del.closest('.sbx-grid');
      post('/sandbox/api/row', { id: grid.getAttribute('data-ss'), sheet: grid.getAttribute('data-sheet'), row: Number(del.getAttribute('data-row')) }).then(function (res) {
        if (res.ok === false) show(document.getElementById('run-out'), failHtml(res));
        else if (res.html) show(document.getElementById('run-out'), res.html);
        refresh(SHEET_PARTS);
      });
      return;
    }

    // Close any open sheet menu when clicking elsewhere.
    document.querySelectorAll('details.sbx-menu[open]').forEach(function (d) { if (!d.contains(e.target)) d.open = false; });

    var td = e.target.closest && e.target.closest('.sbx-grid td');
    if (td && !td.classList.contains('is-check') && !td.classList.contains('is-editing')) edit(td);
  });

  // ---------------------------------------------------------------- sheet cells

  function edit(td) {
    var grid = td.closest('.sbx-grid');
    var before = td.classList.contains('is-formula') ? td.getAttribute('title') : td.textContent;
    var input = document.createElement('input');
    input.type = 'text';
    input.value = before;
    input.setAttribute('aria-label', 'Cell ' + td.getAttribute('data-c') + ',' + td.getAttribute('data-r'));
    td.classList.add('is-editing');
    td.textContent = '';
    td.appendChild(input);
    input.focus();
    input.select();
    var done = false;
    function finish(save) {
      if (done) return; done = true;
      var value = input.value;
      td.classList.remove('is-editing');
      td.textContent = save ? value : before;
      if (!save || value === before) return;
      saveCell(grid, td, value);
    }
    input.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') { ev.preventDefault(); finish(true); }
      if (ev.key === 'Escape') { ev.preventDefault(); finish(false); }
      if (ev.key === 'Tab') { ev.preventDefault(); finish(true); var next = ev.shiftKey ? td.previousElementSibling : td.nextElementSibling; if (next && next.tagName === 'TD' && !next.classList.contains('is-check')) setTimeout(function () { var n = document.querySelector('.sbx-grid td[data-r="' + td.getAttribute('data-r') + '"][data-c="' + next.getAttribute('data-c') + '"]'); if (n) edit(n); }, 150); }
    });
    input.addEventListener('blur', function () { finish(true); });
  }

  function saveCell(grid, td, value, kind) {
    var r = td.getAttribute('data-r'), c = td.getAttribute('data-c');
    return post('/sandbox/api/cell', { id: grid.getAttribute('data-ss'), sheet: grid.getAttribute('data-sheet'), row: Number(r), col: Number(c), value: value, kind: kind || '' }).then(function (res) {
      if (res.ok === false) { show(document.getElementById('run-out'), failHtml(res)); return; }
      if (res.html) show(document.getElementById('run-out'), '<p class="sbx-ui"><b>Cell saved.</b> The sheet\'s edit triggers ran:</p>' + res.html);
      return refresh(SHEET_PARTS).then(function () {
        var cell = document.querySelector('.sbx-grid td[data-r="' + r + '"][data-c="' + c + '"]');
        if (cell) { cell.classList.add('is-saved'); setTimeout(function () { cell.classList.remove('is-saved'); }, 1300); }
      });
    });
  }

  document.addEventListener('change', function (e) {
    var box = e.target;
    if (box.matches && box.matches('.sbx-grid td.is-check input[type="checkbox"]')) {
      saveCell(box.closest('.sbx-grid'), box.closest('td'), box.checked, 'bool');
      return;
    }
    if (box.matches && box.matches('input[data-try]')) {
      var li = box.closest('li'); if (li) li.classList.toggle('is-done', box.checked);
      saveTicks();
    }
  });

  // ---------------------------------------------------------------- forms

  // Phone photos are large; shrink them the way a browser upload widget might, so the
  // sandbox's saved data stays small. SVG and GIF go up as they are.
  function readFile(file) {
    return new Promise(function (resolve) {
      var reader = new FileReader();
      reader.onload = function () {
        var url = reader.result;
        if (!/^image\/(jpeg|png|webp)/.test(file.type)) return resolve(url);
        var img = new Image();
        img.onload = function () {
          var max = 1600, w = img.naturalWidth, h = img.naturalHeight;
          if (w <= max && h <= max && file.size < 1.5e6) return resolve(url);
          var k = Math.min(1, max / Math.max(w, h));
          var cv = document.createElement('canvas'); cv.width = Math.round(w * k); cv.height = Math.round(h * k);
          cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
          resolve(cv.toDataURL('image/jpeg', 0.85));
        };
        img.onerror = function () { resolve(url); };
        img.src = url;
      };
      reader.onerror = function () { resolve(null); };
      reader.readAsDataURL(file);
    });
  }

  document.addEventListener('submit', function (e) {
    var form = e.target;
    if (!form.classList || !form.classList.contains('sbx-form')) return;
    e.preventDefault();
    var answers = {}, uploads = [], bad = null;
    form.querySelectorAll('.sbx-q').forEach(function (q) {
      var title = q.getAttribute('data-title'), kind = q.getAttribute('data-kind');
      q.classList.remove('is-invalid');
      if (kind === 'checks') {
        var picked = Array.prototype.map.call(q.querySelectorAll('input:checked'), function (i) { return i.value; });
        if (q.hasAttribute('data-required') && !picked.length && !bad) bad = q;
        if (q.hasAttribute('data-required') && !picked.length) q.classList.add('is-invalid');
        answers[title] = picked;
      } else if (kind === 'radio') {
        var on = q.querySelector('input:checked'); answers[title] = on ? on.value : '';
      } else if (kind === 'file') {
        var input = q.querySelector('input[type="file"]');
        Array.prototype.forEach.call((input && input.files) || [], function (f) { uploads.push({ question: title, file: f }); });
      } else {
        var field = q.querySelector('input, textarea, select'); answers[title] = field ? field.value : '';
      }
    });
    if (bad) { bad.scrollIntoView({ behavior: 'smooth', block: 'center' }); return; }
    var btn = form.querySelector('button[type="submit"]');
    btn.disabled = true; btn.classList.add('is-busy'); btn.textContent = 'Submitting...';
    Promise.all(uploads.map(function (u) { return readFile(u.file).then(function (url) { return url ? { question: u.question, name: u.file.name, dataUrl: url } : null; }); }))
      .then(function (files) { return post(location.pathname, { answers: answers, files: files.filter(Boolean) }); })
      .then(function (res) {
        btn.disabled = false; btn.classList.remove('is-busy'); btn.textContent = 'Submit';
        show(document.getElementById('form-out'), res.html || failHtml(res));
        if (res.ok !== false) {
          var email = form.querySelector('input[name="email"]'), keep = email ? email.value : '';
          form.reset();
          if (email) email.value = keep;
        }
        refresh(['.sbx-nav']);
      });
  });

  // Google links to a tab put it after a #: .../edit#gid=123. Open that tab.
  if (document.querySelector('.sbx-grid')) {
    var gid = /gid=(\d+)/.exec(location.hash);
    if (gid && !/[?&]gid=/.test(location.search)) location.replace(location.pathname + '?gid=' + gid[1]);
  }

  // ---------------------------------------------------------------- "Things to try" ticks

  var KEY = 'sandbox-try';
  function loadTicks() { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { return {}; } }
  function saveTicks() {
    var out = {};
    document.querySelectorAll('input[data-try]').forEach(function (i) { if (i.checked) out[i.getAttribute('data-try')] = 1; });
    try { localStorage.setItem(KEY, JSON.stringify(out)); } catch (e) { /* private window: ticks just are not kept */ }
  }
  var ticks = loadTicks();
  document.querySelectorAll('input[data-try]').forEach(function (i) {
    if (ticks[i.getAttribute('data-try')]) { i.checked = true; var li = i.closest('li'); if (li) li.classList.add('is-done'); }
  });
})();
