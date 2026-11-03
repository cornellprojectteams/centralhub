/*
 * Sandbox guard, loaded at the top of every page the sandbox serves.
 *
 * The server already swaps real addresses for sandbox ones in the files it sends. This
 * catches the ones a page builds while it runs (a JSONP call, an image added later),
 * using the same list (window.__SANDBOX_MAP__). Links that still point at real Google
 * files, forms or Fabman go to a "you are leaving the sandbox" page first.
 * It also adds the small "Sandbox" tab in the corner of top-level pages.
 */
(function () {
  'use strict';
  if (window.__sandboxGuard) return;
  window.__sandboxGuard = true;
  var MAP = window.__SANDBOX_MAP__ || [];

  function fix(u) {
    if (typeof u !== 'string' || u.indexOf('http') !== 0) return u;
    for (var i = 0; i < MAP.length; i++) {
      if (u.indexOf(MAP[i][0]) === 0) return MAP[i][1] + u.slice(MAP[i][0].length);
    }
    return u;
  }
  window.__sandboxFix = fix;

  // Addresses set from script: element.src = ..., setAttribute('href', ...), fetch, XHR, window.open.
  function patch(proto, prop) {
    var d = proto && Object.getOwnPropertyDescriptor(proto, prop);
    if (!d || !d.set) return;
    Object.defineProperty(proto, prop, {
      configurable: true, enumerable: d.enumerable,
      get: function () { return d.get.call(this); },
      set: function (v) { d.set.call(this, fix(String(v))); },
    });
  }
  patch(window.HTMLScriptElement && HTMLScriptElement.prototype, 'src');
  patch(window.HTMLImageElement && HTMLImageElement.prototype, 'src');
  patch(window.HTMLIFrameElement && HTMLIFrameElement.prototype, 'src');
  patch(window.HTMLSourceElement && HTMLSourceElement.prototype, 'src');
  patch(window.HTMLAnchorElement && HTMLAnchorElement.prototype, 'href');
  patch(window.HTMLFormElement && HTMLFormElement.prototype, 'action');
  var setAttr = Element.prototype.setAttribute;
  Element.prototype.setAttribute = function (name, value) {
    if (name === 'src' || name === 'href' || name === 'action') value = fix(String(value));
    return setAttr.call(this, name, value);
  };
  if (window.fetch) {
    var realFetch = window.fetch;
    window.fetch = function (input, init) { return realFetch.call(this, typeof input === 'string' ? fix(input) : input, init); };
  }
  var xhrOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method, url) { arguments[1] = fix(String(url)); return xhrOpen.apply(this, arguments); };
  var realOpen = window.open;
  window.open = function (url) { var a = Array.prototype.slice.call(arguments); if (a.length) a[0] = leaveUrl(fix(String(url || ''))); return realOpen.apply(window, a); };

  // HTML added with innerHTML skips the setters above; fix those attributes as they land.
  function scan(root) {
    if (!root || root.nodeType !== 1) return;
    var list = [root].concat(Array.prototype.slice.call(root.querySelectorAll('[src],[href],[action]')));
    list.forEach(function (el) {
      ['src', 'href', 'action'].forEach(function (a) {
        if (!el.hasAttribute || !el.hasAttribute(a)) return;
        var v = el.getAttribute(a), w = fix(v);
        if (w !== v) setAttr.call(el, a, w);
      });
    });
  }
  new MutationObserver(function (muts) {
    muts.forEach(function (m) { Array.prototype.forEach.call(m.addedNodes, scan); });
  }).observe(document.documentElement, { childList: true, subtree: true });

  // Real places the sandbox cannot stand in for. Reading them is fine; a click asks first.
  var REAL = /^https?:\/\/([^\/]*\.)?(google\.com|forms\.gle|fabman\.io|qualtrics\.com|box\.com|instructure\.com)(\/|$)/i;
  var SAFE = /^https?:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com)\//i;
  function leaveUrl(u) {
    if (!REAL.test(u) || SAFE.test(u)) return u;
    return '/sandbox/leaving?to=' + encodeURIComponent(u);
  }
  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest ? e.target.closest('a[href]') : null;
    if (!a || a.hasAttribute('data-leave')) return;
    var href = a.getAttribute('href') || '';
    var fixed = fix(href);
    if (fixed !== href) { a.setAttribute('href', fixed); return; }
    var out = leaveUrl(href);
    if (out !== href) {
      e.preventDefault();
      var target = a.getAttribute('target');
      if (target === '_blank') realOpen.call(window, out, '_blank', 'noopener');
      else (target === '_top' ? window.top : window).location.href = out;
    }
  }, true);
  document.addEventListener('submit', function (e) {
    var f = e.target, action = f && f.getAttribute && f.getAttribute('action') || '';
    var fixed = fix(action);
    if (fixed !== action) { setAttr.call(f, 'action', fixed); return; }
    if (REAL.test(action)) { e.preventDefault(); window.location.href = leaveUrl(action); }
  }, true);

  // The corner tab back to the control panel, on top-level pages only.
  function tab() {
    if (window.top !== window || document.getElementById('sandbox-tab') || /^\/sandbox\//.test(location.pathname) && !/\/exec$/.test(location.pathname)) return;
    var a = document.createElement('a');
    a.id = 'sandbox-tab';
    a.href = '/sandbox/';
    a.title = 'You are in the sandbox. Open the sandbox control panel.';
    a.textContent = 'Sandbox';
    a.setAttribute('style', 'position:fixed;left:14px;bottom:14px;z-index:2147483000;display:inline-flex;align-items:center;gap:7px;'
      + 'padding:7px 13px 7px 11px;border-radius:999px;background:#1b1f24;color:#fff;font:700 12px/1 system-ui,-apple-system,Segoe UI,sans-serif;'
      + 'letter-spacing:.06em;text-transform:uppercase;text-decoration:none;box-shadow:0 6px 18px rgba(0,0,0,.25);opacity:.92');
    var dot = document.createElement('span');
    dot.setAttribute('style', 'width:8px;height:8px;border-radius:50%;background:#f0c050;box-shadow:0 0 0 3px rgba(240,192,80,.25)');
    a.insertBefore(dot, a.firstChild);
    document.body.appendChild(a);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', tab); else tab();
})();
