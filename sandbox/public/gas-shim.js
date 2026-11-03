/*
 * google.script.run for Apps Script pages served by the sandbox.
 *
 * On Google, a web app page calls its server functions with
 *   google.script.run.withSuccessHandler(ok).withFailureHandler(fail).someFunction(args)
 * Here the same call posts to the sandbox server, which runs someFunction from the
 * repo's .gs files and sends back what it returned. Functions whose names end in an
 * underscore are private, as on Google, and cannot be called from a page.
 */
(function () {
  'use strict';
  var me = document.currentScript;
  var endpoint = (me && me.getAttribute('data-run')) || '/sandbox/space/run';

  function runner(opts) {
    return new Proxy({}, {
      get: function (target, name) {
        if (typeof name !== 'string' || name === 'then') return undefined;
        if (name === 'withSuccessHandler') return function (fn) { return runner(Object.assign({}, opts, { ok: fn })); };
        if (name === 'withFailureHandler') return function (fn) { return runner(Object.assign({}, opts, { fail: fn })); };
        if (name === 'withUserObject') return function (o) { return runner(Object.assign({}, opts, { user: o })); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var fail = function (message) {
            var err = new Error(message);
            setTimeout(function () { if (opts.fail) opts.fail(err, opts.user); else console.error('google.script.run.' + name + ' failed: ' + message); });
          };
          fetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: name, args: args }) })
            .then(function (r) { return r.json(); })
            .then(function (res) {
              if (!res.ok) return fail(res.error || 'Unknown error');
              setTimeout(function () { if (opts.ok) opts.ok(res.value, opts.user); });
            })
            .catch(function () { fail('NetworkError: Connection failure due to HTTP 0'); });
        };
      },
    });
  }

  window.google = window.google || {};
  window.google.script = {
    run: runner({}),
    host: { close: function () {}, setHeight: function () {}, setWidth: function () {}, editor: { focus: function () {} }, origin: location.origin },
    url: {
      getLocation: function (cb) {
        var q = new URLSearchParams(location.search), parameter = {}, parameters = {};
        q.forEach(function (v, k) { if (!(k in parameter)) parameter[k] = v; (parameters[k] = parameters[k] || []).push(v); });
        cb({ hash: location.hash.replace(/^#/, ''), parameter: parameter, parameters: parameters });
      },
    },
    history: { push: function () {}, replace: function () {}, setChangeHandler: function () {} },
  };
})();
