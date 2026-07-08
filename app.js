/**
 * Project Teams Ops Hub client script
 *
 * Two responsibilities:
 *   1. View switching between #staff and #admin via header nav buttons
 *   2. Live search filtering across tool links and section keywords
 *
 * Markup conventions (see index.html):
 *   - .hub-zone          Top-level staff or admin panel
 *   - .category          Searchable block (task-block or admin-block)
 *   - .action            External tool link; may carry data-keywords
 *   - data-keywords       Extra terms matched by search (lowercase substring)
 *
 * URL hashes: #staff (default) and #admin (also used by admin.html redirect)
 */

(function () {
  'use strict';

  // Deployed Space Status web app /exec URL. The team picker opens a team's
  // live read-only open-issues page at SPACE_STATUS_URL?view=<team name>.
  const SPACE_STATUS_URL = 'https://script.google.com/macros/s/AKfycbwOnNmpSXc3biH14Fm9iLcUQ2X0UK-Gx5kQpNmrBsHd3K-l2u0GjsMblOumiY73drM_/exec';

  const search = document.getElementById('search');
  const categories = document.querySelectorAll('.category');
  const zones = document.querySelectorAll('.hub-zone');
  const noResults = document.getElementById('no-results');
  const navButtons = document.querySelectorAll('.hub-nav-btn');

  const adminNav = document.getElementById('admin-nav');

  /** @type {'staff' | 'admin'} */
  let currentView = 'staff';

  /**
   * Show one hub zone and sync nav + URL hash.
   * @param {'staff' | 'admin'} view
   */
  function setView(view) {
    if (view !== 'staff' && view !== 'admin') return;
    currentView = view;

    navButtons.forEach(btn => {
      btn.classList.toggle('is-active', btn.dataset.view === view);
    });

    zones.forEach(zone => {
      zone.classList.toggle('is-active', zone.id === view);
    });

    if (location.hash !== `#${view}`) {
      history.replaceState(null, '', `#${view}`);
    }

    if (search && !search.value.trim()) {
      document.body.classList.remove('is-searching');
    }
  }

  navButtons.forEach(btn => {
    btn.addEventListener('click', () => setView(btn.dataset.view));
  });

  window.addEventListener('hashchange', () => {
    const hash = location.hash.slice(1);
    if (hash === 'staff' || hash === 'admin') setView(hash);
  });

  const initialHash = location.hash.slice(1);
  setView(initialHash === 'admin' ? 'admin' : 'staff');

  // Scroll-spy: highlight the sidebar link for the section currently in view.
  if (adminNav && 'IntersectionObserver' in window) {
    const navLinks = [...adminNav.querySelectorAll('a')];
    const linkById = {};
    navLinks.forEach(a => { linkById[a.getAttribute('href').slice(1)] = a; });
    const setActive = id => {
      navLinks.forEach(a => a.classList.remove('active'));
      if (linkById[id]) linkById[id].classList.add('active');
    };
    const spy = new IntersectionObserver(entries => {
      const inView = entries.filter(e => e.isIntersecting)
        .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
      if (inView.length) setActive(inView[0].target.id);
    }, { rootMargin: '-96px 0px -65% 0px', threshold: 0 });
    document.querySelectorAll('#admin .admin-block').forEach(b => spy.observe(b));
    navLinks.forEach(a => a.addEventListener('click', () => setActive(a.getAttribute('href').slice(1))));
  }

  if (!search) return;

  const clearBtn = document.getElementById('search-clear');
  const kbdHint = document.getElementById('search-kbd');
  const meta = document.getElementById('search-meta');
  const countEl = document.getElementById('search-count');

  if (kbdHint) {
    const isMac = /mac|iphone|ipad/i.test(navigator.platform || navigator.userAgent || '');
    kbdHint.textContent = isMac ? '⌘K' : 'Ctrl K';
  }

  // Cache each label's original text so highlights can be re-rendered cleanly.
  document.querySelectorAll('.action').forEach(action => {
    const title = action.querySelector('.action-title');
    const sub = action.querySelector('.action-sub');
    if (title && title.dataset.text == null) title.dataset.text = title.textContent;
    if (sub && sub.dataset.text == null) sub.dataset.text = sub.textContent;
  });

  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /**
   * fzf-style subsequence scorer. Returns -1 when `needle` is not a
   * subsequence of `hay`, otherwise a score rewarding contiguous runs and
   * matches that land on a word boundary.
   * @param {string} hay
   * @param {string} needle
   * @returns {number}
   */
  function fuzzyScore(hay, needle) {
    if (!needle) return 0;
    let score = 0, at = 0, streak = 0;
    for (let i = 0; i < needle.length; i++) {
      const c = needle[i];
      let found = -1;
      for (let k = at; k < hay.length; k++) {
        if (hay[k] === c) { found = k; break; }
      }
      if (found === -1) return -1;
      if (found === at) { streak++; score += 4 + streak; }
      else { streak = 0; score += 1; }
      const prev = found > 0 ? hay[found - 1] : ' ';
      if (prev === ' ' || prev === '-' || prev === '/' || prev === '&') score += 10;
      at = found + 1;
    }
    return score;
  }

  /**
   * Build searchable text for an action link from cached labels + keywords.
   * @param {Element} action
   * @param {Element} section
   * @returns {string}
   */
  function actionSearchText(action, section) {
    const title = action.querySelector('.action-title');
    const sub = action.querySelector('.action-sub');
    const block = action.closest('.task-block');
    const footnote = (block || section).querySelector('.action-note');
    const footnoteText = footnote
      ? `${footnote.textContent} ${footnote.dataset.keywords || ''}`
      : '';
    const titleText = title ? title.dataset.text : action.textContent;
    const subText = sub ? sub.dataset.text : '';
    return `${titleText} ${subText} ${action.dataset.keywords || ''} ${footnoteText}`.toLowerCase();
  }

  /**
   * Wrap query-token occurrences in the title/sub labels with <mark>.
   * @param {Element} action
   * @param {string[]} tokens
   */
  function highlight(action, tokens) {
    ['.action-title', '.action-sub'].forEach(sel => {
      const el = action.querySelector(sel);
      if (!el) return;
      const orig = el.dataset.text != null ? el.dataset.text : el.textContent;
      if (!tokens.length) { el.textContent = orig; return; }
      const re = new RegExp('(' + tokens.map(escapeRe).join('|') + ')', 'ig');
      el.innerHTML = escapeHtml(orig).replace(re, '<mark>$1</mark>');
    });
  }

  // Keyboard navigation across the currently visible results.
  let results = [];
  let activeIdx = -1;

  function rebuildResults() {
    results = [...document.querySelectorAll('.action:not(.hidden)')]
      .filter(a => a.offsetParent !== null);
  }

  function clearActive() {
    if (activeIdx >= 0 && results[activeIdx]) {
      results[activeIdx].classList.remove('is-active-result');
    }
    activeIdx = -1;
  }

  function setActive(i) {
    if (activeIdx >= 0 && results[activeIdx]) {
      results[activeIdx].classList.remove('is-active-result');
    }
    if (!results.length) { activeIdx = -1; return; }
    activeIdx = Math.max(0, Math.min(results.length - 1, i));
    const el = results[activeIdx];
    el.classList.add('is-active-result');
    el.scrollIntoView({ block: 'nearest' });
  }

  function runSearch() {
    const raw = search.value;
    const q = raw.toLowerCase().trim();
    const tokens = q ? q.split(/\s+/).filter(Boolean) : [];
    const isSearching = tokens.length > 0;

    document.body.classList.toggle('is-searching', isSearching);
    if (clearBtn) clearBtn.hidden = !raw;
    if (kbdHint) kbdHint.hidden = Boolean(raw);

    let matchCount = 0;

    categories.forEach(section => {
      const sectionKeywords = (section.dataset.keywords || '').toLowerCase();
      const actions = section.matches('.action')
        ? [section]
        : [...section.querySelectorAll('.action')];
      let sectionHasMatch = false;

      actions.forEach(action => {
        const hay = actionSearchText(action, section);
        let matched = true;
        let score = 0;

        if (tokens.length) {
          for (const t of tokens) {
            const best = Math.max(fuzzyScore(hay, t), fuzzyScore(sectionKeywords, t));
            if (best < 0) { matched = false; break; }
            score += best;
          }
        }

        action.classList.toggle('hidden', tokens.length ? !matched : false);
        action.style.order = (tokens.length && matched) ? String(-score) : '';
        highlight(action, tokens);

        const block = action.closest('.task-block');
        if (block && block !== action) {
          const blockVisible = [...block.querySelectorAll('.action')].some(
            a => !a.classList.contains('hidden')
          ) || (block.querySelector('.action-note') && matched);
          block.classList.toggle('hidden', tokens.length && !blockVisible);
        }

        if (!tokens.length || matched) sectionHasMatch = true;
        if (tokens.length && matched) matchCount++;
      });

      if (!actions.length && (!tokens.length || fuzzyScore(sectionKeywords, q) >= 0)) {
        sectionHasMatch = true;
      }

      section.classList.toggle('hidden', tokens.length && !sectionHasMatch);
    });

    zones.forEach(zone => {
      if (!tokens.length) {
        zone.classList.remove('hidden');
        zone.classList.toggle('is-active', zone.id === currentView);
        return;
      }
      const hasVisible = zone.querySelector('.category:not(.hidden), .action:not(.hidden)');
      zone.classList.toggle('hidden', !hasVisible);
      if (hasVisible) zone.classList.add('is-active');
    });

    noResults.classList.toggle('visible', isSearching && matchCount === 0);

    rebuildResults();
    if (isSearching) setActive(0); else clearActive();

    if (meta) {
      meta.hidden = !isSearching;
      if (isSearching && countEl) {
        countEl.textContent = matchCount === 1 ? '1 result' : `${matchCount} results`;
      }
    }
  }

  search.addEventListener('input', runSearch);

  search.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive(activeIdx + 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(activeIdx - 1); }
    else if (e.key === 'Enter') {
      if (results[activeIdx]) { e.preventDefault(); results[activeIdx].click(); }
    } else if (e.key === 'Escape') {
      if (search.value) { e.preventDefault(); search.value = ''; runSearch(); }
      else { search.blur(); }
    }
  });

  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      search.value = '';
      runSearch();
      search.focus();
    });
  }

  // Open space issues dashboard, embedded under the Experiential Learning Lab section.
  const issuesFrame = document.getElementById('issues-frame');
  if (issuesFrame) {
    const issuesUrl = SPACE_STATUS_URL
      ? SPACE_STATUS_URL + (SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?') + 'view=all'
      : '';
    if (issuesUrl) {
      issuesFrame.src = issuesUrl;
      const link = document.getElementById('issues-fallback-link');
      if (link) link.href = issuesUrl;
    } else {
      const fb = document.getElementById('issues-fallback');
      if (fb) fb.hidden = false;
    }
  }

  // Global shortcuts: "/" or Cmd/Ctrl+K to jump to search from anywhere.
  document.addEventListener('keydown', e => {
    const tag = (e.target.tagName || '').toLowerCase();
    const typing = tag === 'input' || tag === 'textarea' || e.target.isContentEditable;
    if (!typing && e.key === '/') { e.preventDefault(); search.focus(); }
    else if ((e.metaKey || e.ctrlKey) && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      search.focus();
      search.select();
    }
  });
})();
