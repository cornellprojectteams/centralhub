(function () {
  function call(url, action, payload) {
    return new Promise(function (resolve, reject) {
      var cb = 'sb_' + Math.random().toString(36).slice(2);
      var timer = setTimeout(function () { cleanup(); reject(new Error('The web app did not answer.')); }, 20000);
      function cleanup() {
        clearTimeout(timer);
        try { delete window[cb]; } catch (e) { window[cb] = undefined; }
        if (script.parentNode) script.parentNode.removeChild(script);
      }
      window[cb] = function (data) { cleanup(); resolve(data || { ok: false }); };
      var script = document.createElement('script');
      var q = 'callback=' + encodeURIComponent(cb)
        + '&action=' + encodeURIComponent(action)
        + '&payload=' + encodeURIComponent(JSON.stringify(payload || {}));
      script.onerror = function () { cleanup(); reject(new Error('Could not reach the web app.')); };
      script.src = url + (url.indexOf('?') >= 0 ? '&' : '?') + q;
      document.head.appendChild(script);
    });
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function mount(host, opts) {
    var root = document.createElement('div');
    root.className = 'sb-root';
    host.appendChild(root);
    var bucket = 'open';
    var data = { open: [], done: [] };

    function paint() {
      var open = data.open.filter(function (it) { return !it.pending; });
      var pending = data.open.filter(function (it) { return it.pending; });
      var rows = bucket === 'pending' ? pending : bucket === 'done' ? data.done : open;
      var html = '<div class="sb-bar">'
        + chip('open', 'Open', open.length)
        + chip('pending', 'Pending', pending.length)
        + chip('done', 'Done', (data.resolved || data.done.length))
        + '</div>';
      if (!rows.length) html += '<div class="sb-empty">Nothing in this list.</div>';
      rows.forEach(function (it, i) { html += card(it, bucket + i); });
      root.innerHTML = html;
      root.querySelectorAll('[data-bucket]').forEach(function (btn) {
        btn.addEventListener('click', function () { bucket = btn.getAttribute('data-bucket'); paint(); });
      });
      root.querySelectorAll('.sb-head').forEach(function (btn) {
        btn.addEventListener('click', function () { btn.parentNode.classList.toggle('is-open'); });
      });
      bind(root);
    }

    function chip(id, label, n) {
      return '<button type="button" class="sb-chip' + (bucket === id ? ' is-on' : '') + '" data-bucket="' + id + '"><b>' + n + '</b>' + label + '</button>';
    }

    function card(it, rid) {
      var pending = !!it.pending;
      var sent = !pending && it.sentBackReason;
      var pill = pending ? 'Pending' : sent ? 'Sent back' : it.overdue ? 'Overdue' : (bucket === 'done' ? 'Completed' : 'Open');
      var cls = pending ? 'sb-pill--pending' : sent ? 'sb-pill--sent' : it.overdue ? 'sb-pill--late' : (bucket === 'done' ? 'sb-pill--done' : '');
      var sub = esc(it.team || opts.team || '');
      return '<article class="sb-card" data-token="' + esc(it.token || '') + '">'
        + '<button type="button" class="sb-head"><span><span class="sb-title">' + esc(it.issueType || 'Reported issue') + '</span>'
        + '<span class="sb-sub">' + sub + '</span></span><span class="sb-pill ' + cls + '">' + pill + '</span></button>'
        + '<div class="sb-body">'
        + (it.details ? '<div class="sb-details">' + esc(it.details) + '</div>' : '')
        + (sent ? '<div class="sb-note"><b>Sent back:</b> ' + esc(it.sentBackReason) + '</div>' : '')
        + (bucket === 'done' ? '' : fields(it, rid, pending))
        + '</div></article>';
    }

    function fields(it, rid, pending) {
      var note = '';
      if (!opts.admin && !pending) {
        note = '<label class="sb-label" for="' + rid + '-note">Note</label>'
          + '<textarea id="' + rid + '-note" class="sb-text" placeholder="What you did, or anything the reviewer should know"></textarea>'
          + '<span class="sb-help">Optional. This goes out with your completion.</span>';
      }
      return note + '<div class="sb-actions" data-act>' + buttons(it, pending) + '</div>';
    }

    function buttons(it, pending) {
      if (opts.admin && pending) {
        return '<span class="sb-tools"><button type="button" class="sb-btn sb-btn--ghost" data-do="edit">Edit</button>'
          + '<button type="button" class="sb-btn sb-btn--danger" data-do="delete">Delete</button></span>'
          + '<button type="button" class="sb-btn sb-btn--ok" data-do="approve">Approve</button>'
          + '<button type="button" class="sb-btn sb-btn--warn" data-do="send">Send back</button>';
      }
      if (opts.admin) {
        return '<span class="sb-tools"><button type="button" class="sb-btn sb-btn--ghost" data-do="edit">Edit</button>'
          + '<button type="button" class="sb-btn sb-btn--danger" data-do="delete">Delete</button></span>'
          + '<button type="button" class="sb-btn sb-btn--go" data-do="complete">Mark complete</button>';
      }
      if (pending) return '<span class="sb-help">Waiting for an admin to review this.</span>';
      return '<button type="button" class="sb-btn sb-btn--go" data-do="submit">' + (it.photoOptional ? 'Mark done' : 'Complete') + '</button>';
    }

    function bind(scope) {
      scope.querySelectorAll('[data-do]').forEach(function (btn) {
        btn.addEventListener('click', function () { onAction(btn); });
      });
    }

    function onAction(btn) {
      var cardEl = btn.closest('.sb-card');
      var act = cardEl.querySelector('[data-act]');
      var token = cardEl.getAttribute('data-token');
      var kind = btn.getAttribute('data-do');
      if (kind === 'send') return panel(act, 'Why send this back?', true, 'Send back', function (text) { run(act, 'sendBack', { token: token, reason: text }); });
      if (kind === 'complete') return panel(act, 'Mark this complete with no photo? It closes right away.', false, 'Yes, mark complete', function () { run(act, 'complete', { token: token }); });
      if (kind === 'delete') return panel(act, 'Delete this task? This cannot be undone.', false, 'Delete', function () { run(act, 'delete', { token: token }); }, true);
      if (kind === 'approve') return run(act, 'approve', { token: token });
      if (kind === 'submit') {
        var noteEl = cardEl.querySelector('.sb-text');
        var item = data.open.filter(function (it) { return it.token === token; })[0];
        if (item && !item.photoOptional) {
          act.innerHTML = '<span class="sb-err">A photo is still required. Use Open full page for that.</span>';
          return;
        }
        return run(act, 'submit', { token: token, note: noteEl ? noteEl.value : '' });
      }
      if (kind === 'edit') return edit(cardEl, act, token);
    }

    function panel(act, message, text, confirmLabel, done, danger) {
      act.innerHTML = '<div class="sb-panel' + (danger ? ' sb-panel--danger' : '') + '"><p>' + esc(message) + '</p>'
        + (text ? '<textarea class="sb-text" placeholder="What still needs to be fixed"></textarea><span class="sb-help">The team gets this in the email, with the original comment.</span>' : '')
        + '<div class="sb-row"><button type="button" class="sb-btn ' + (danger ? 'sb-btn--danger' : 'sb-btn--go') + '" data-yes>' + esc(confirmLabel) + '</button>'
        + '<button type="button" class="sb-btn sb-btn--ghost" data-no>Cancel</button></div></div>';
      act.querySelector('[data-yes]').addEventListener('click', function () {
        var area = act.querySelector('.sb-text');
        done(area ? area.value : '');
      });
      act.querySelector('[data-no]').addEventListener('click', function () { load(); });
    }

    function edit(cardEl, act, token) {
      var item = data.open.filter(function (it) { return it.token === token; })[0] || {};
      act.innerHTML = '<div class="sb-panel"><label class="sb-label">Team</label><input class="sb-field" data-team value="' + esc(item.team || '') + '">'
        + '<label class="sb-label">Issue type</label><input class="sb-field" data-type value="' + esc(item.issueType || '') + '">'
        + '<label class="sb-label">Details</label><textarea class="sb-text" data-det>' + esc(item.details || '') + '</textarea>'
        + '<div class="sb-row"><button type="button" class="sb-btn sb-btn--go" data-yes>Save</button><button type="button" class="sb-btn sb-btn--ghost" data-no>Cancel</button></div></div>';
      act.querySelector('[data-yes]').addEventListener('click', function () {
        run(act, 'edit', {
          token: token,
          team: act.querySelector('[data-team]').value,
          issueType: act.querySelector('[data-type]').value,
          details: act.querySelector('[data-det]').value
        });
      });
      act.querySelector('[data-no]').addEventListener('click', function () { load(); });
    }

    function run(act, action, payload) {
      act.innerHTML = '<span class="sb-help">Saving…</span>';
      call(opts.url, action, payload).then(function (r) {
        if (!r || r.ok === false) {
          act.innerHTML = '<span class="sb-err">' + esc((r && r.error) || 'Could not save') + '</span>';
          return;
        }
        load();
      }).catch(function (err) {
        act.innerHTML = '<span class="sb-err">' + esc(err.message || 'Could not save') + '</span>';
      });
    }

    function load() {
      return call(opts.url, 'issues', { team: opts.team || '' }).then(function (r) {
        if (!r || !r.ok) throw new Error((r && r.error) || 'No data');
        data = r;
        paint();
      });
    }

    return load();
  }

  window.SpaceBoard = { mount: mount };
})();
