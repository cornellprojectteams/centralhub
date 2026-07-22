/**
 * Project Teams Ops Hub client script (shared by index.html and admin.html)
 *
 * Live search, the Command Center badge, the open-issues dashboard panel, and
 * the admin sidebar scroll-spy. Every feature is guarded by element presence,
 * so the same file runs on the staff page and the admin page unchanged.
 *
 * Markup conventions:
 *   - .hub-zone          The page's single panel (staff on index, admin on admin.html)
 *   - .category          Searchable block (task-block or admin-block)
 *   - .action            Tool link or button; may carry data-keywords
 */

(function () {
  'use strict';

  // Deployed Space Status web app /exec URL. Set in config.js (single source of
  // truth, shared with team.html); the literal below is only a fallback.
  const SPACE_STATUS_URL = window.SPACE_STATUS_URL || 'https://script.google.com/macros/s/AKfycbwNbGjVcBrcsMZiOl2nXzpqZHz04nvKLm9D_aC0VJDz7Xxxf_4kLKlNSOHubPXj1X74/exec';

  // Ops Command Center stats API. Used only to put a live attention count on the
  // Command button — the page still works fine if it never answers.
  const STATS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzSfUQgCfkOpXAmEExVh1bHIbIQ7LAipTzP_uW2x2XKKTmZuGvmxpPI6gS1cw7oLVOz/exec';

  /**
   * Badge the Command Center button with how many things need attention.
   * JSONP, because a plain fetch() to script.google.com is blocked by CORS.
   * Fails silently: a missing badge is better than a broken masthead.
   */
  function loadCommandBadge() {
    const badge = document.getElementById('cmd-badge');
    if (!badge || !STATS_ENDPOINT) return;

    const cb = '__hub' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    let settled = false;
    const cleanup = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    };
    const timer = setTimeout(cleanup, 12000);

    window[cb] = data => {
      const n = data && data.hero ? Number(data.hero.attention) : 0;
      cleanup();
      if (!Number.isFinite(n) || n <= 0) return;   // nothing wrong -> no badge
      badge.textContent = n > 99 ? '99+' : String(n);
      badge.hidden = false;
      const link = badge.closest('.cmd-link');
      if (link) link.setAttribute('aria-label', `Command Center — ${n} items need attention`);
    };
    script.onerror = cleanup;
    script.src = STATS_ENDPOINT + '?callback=' + cb;
    document.body.appendChild(script);
  }

  loadCommandBadge();

  // Registry viewer links (Equipment / Inventory) point at the web app.
  document.querySelectorAll('a[data-registry]').forEach(a => {
    if (!SPACE_STATUS_URL) return;
    const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
    a.href = SPACE_STATUS_URL + sep + 'registry=' + encodeURIComponent(a.dataset.registry);
  });

  // Module links (Projects) point at the same web app via ?module=.
  // data-admin links (only on the unlisted admin page) add &admin=1, which unlocks
  // the admin controls without a passcode. Regular staff links never carry it.
  document.querySelectorAll('a[data-module]').forEach(a => {
    if (!SPACE_STATUS_URL) return;
    const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
    a.href = SPACE_STATUS_URL + sep + 'module=' + encodeURIComponent(a.dataset.module)
      + (a.hasAttribute('data-admin') ? '&admin=1' : '');
  });

  const search = document.getElementById('search');
  const categories = document.querySelectorAll('.category');
  const zones = document.querySelectorAll('.hub-zone');
  const noResults = document.getElementById('no-results');

  const adminNav = document.getElementById('admin-nav');

  // Open space issues dashboard — slide-over panel under Space (not inline in ELL).
  const issuesPanel = document.getElementById('issues-panel');
  const issuesOpenBtn = document.getElementById('issues-open');
  const issuesCloseBtn = document.getElementById('issues-panel-close');
  const issuesBackdrop = document.getElementById('issues-panel-backdrop');
  const issuesPanelExt = document.getElementById('issues-panel-ext');
  const issuesFrame = document.getElementById('issues-frame');
  const issuesEmbedBody = document.getElementById('issues-embed-body');
  const issuesFallback = document.getElementById('issues-fallback');
  const issuesFallbackLink = document.getElementById('issues-fallback-link');
  const issuesLoadingText = document.getElementById('issues-loading-text');
  const issuesLoadingHint = document.getElementById('issues-loading-hint');
  let issuesFrameStarted = false;
  let issuesLoadTimer = null;
  let issuesLoadStepTimer = null;

  // Which view the panel embeds: admin = all teams, staff = one team (set via
  // data-embed-query on #issues-panel). Defaults to the all-teams dashboard.
  const embedQuery = (issuesPanel && issuesPanel.dataset.embedQuery) || 'view=all';
  const sep = SPACE_STATUS_URL.indexOf('?') >= 0 ? '&' : '?';
  const issuesUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + embedQuery + '&embed=1' : '';
  const issuesFullUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + embedQuery : '';

  if (issuesFallbackLink && issuesFullUrl) issuesFallbackLink.href = issuesFullUrl;
  if (issuesPanelExt && issuesFullUrl) issuesPanelExt.href = issuesFullUrl;

  // Combined Tasks/Projects segmented toggle. Present only on the student panel
  // (index.html); the admin dashboard panel has no toggle, so these are no-ops there.
  // The Projects iframe is lazy-loaded the first time its tab is opened.
  const segTasks = document.getElementById('seg-tasks');
  const segProjects = document.getElementById('seg-projects');
  const projectsFrame = document.getElementById('projects-frame');
  const projectsEmbedBody = document.getElementById('projects-embed-body');
  const projectsUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'module=projects&embed=1' : '';
  const projectsFullUrl = SPACE_STATUS_URL ? SPACE_STATUS_URL + sep + 'module=projects' : '';
  let projectsFrameStarted = false;

  function loadProjectsFrame() {
    if (projectsFrameStarted || !projectsFrame || !projectsUrl) return;
    projectsFrameStarted = true;
    projectsFrame.addEventListener('load', () => {
      if (projectsEmbedBody) projectsEmbedBody.classList.remove('is-loading');
    }, { once: true });
    projectsFrame.src = projectsUrl;
  }

  function showPanelView(view) {
    const tasks = view !== 'projects';
    if (issuesEmbedBody) issuesEmbedBody.hidden = !tasks;
    if (projectsEmbedBody) projectsEmbedBody.hidden = tasks;
    if (segTasks) { segTasks.classList.toggle('is-active', tasks); segTasks.setAttribute('aria-selected', String(tasks)); }
    if (segProjects) { segProjects.classList.toggle('is-active', !tasks); segProjects.setAttribute('aria-selected', String(!tasks)); }
    if (issuesPanelExt) issuesPanelExt.href = tasks ? issuesFullUrl : projectsFullUrl;
    if (!tasks) loadProjectsFrame();
  }

  if (segTasks) segTasks.addEventListener('click', () => showPanelView('tasks'));
  if (segProjects) segProjects.addEventListener('click', () => showPanelView('projects'));

  function setIssuesLoadingMessage(primary, hint) {
    if (issuesLoadingText && primary) issuesLoadingText.textContent = primary;
    if (issuesLoadingHint && hint) issuesLoadingHint.textContent = hint;
  }

  function startIssuesLoadingSteps() {
    setIssuesLoadingMessage('Connecting to Space Status…', 'Fetching open issues and photos from Google Sheets');
    stopIssuesLoadingSteps();
    const step2 = window.setTimeout(() => {
      setIssuesLoadingMessage('Loading issues…', 'This can take a few seconds on first open');
    }, 2800);
    const step3 = window.setTimeout(() => {
      setIssuesLoadingMessage('Almost ready…', 'Building filters and issue cards');
    }, 7000);
    issuesLoadStepTimer = [step2, step3];
  }

  function stopIssuesLoadingSteps() {
    if (!issuesLoadStepTimer) return;
    (Array.isArray(issuesLoadStepTimer) ? issuesLoadStepTimer : [issuesLoadStepTimer])
      .forEach(id => clearTimeout(id));
    issuesLoadStepTimer = null;
  }

  function showIssuesError() {
    stopIssuesLoadingSteps();
    if (issuesEmbedBody) issuesEmbedBody.classList.remove('is-loading');
    if (issuesFallback) issuesFallback.hidden = false;
  }

  function finishIssuesLoad() {
    if (issuesLoadTimer) {
      clearTimeout(issuesLoadTimer);
      issuesLoadTimer = null;
    }
    stopIssuesLoadingSteps();
    if (issuesEmbedBody) issuesEmbedBody.classList.remove('is-loading');
  }

  function loadIssuesFrame() {
    if (!issuesFrame || issuesFrameStarted) return;
    if (!issuesUrl) {
      showIssuesError();
      return;
    }

    issuesFrameStarted = true;
    if (issuesEmbedBody) issuesEmbedBody.classList.add('is-loading');
    if (issuesFallback) issuesFallback.hidden = true;
    startIssuesLoadingSteps();

    issuesFrame.addEventListener('load', finishIssuesLoad, { once: true });
    issuesFrame.addEventListener('error', showIssuesError, { once: true });

    issuesLoadTimer = window.setTimeout(() => {
      issuesLoadTimer = null;
      if (issuesEmbedBody && issuesEmbedBody.classList.contains('is-loading')) {
        showIssuesError();
      }
    }, 45000);

    issuesFrame.src = issuesUrl;
  }

  function openIssuesPanel() {
    if (!issuesPanel) return;
    issuesPanel.hidden = false;
    requestAnimationFrame(() => issuesPanel.classList.add('is-open'));
    document.body.classList.add('issues-panel-open');
    if (segTasks) showPanelView('tasks');   // always open on the Tasks tab
    loadIssuesFrame();
    if (issuesCloseBtn) issuesCloseBtn.focus();
  }

  function closeIssuesPanel() {
    if (!issuesPanel) return;
    issuesPanel.classList.remove('is-open');
    document.body.classList.remove('issues-panel-open');
    window.setTimeout(() => {
      if (!issuesPanel.classList.contains('is-open')) issuesPanel.hidden = true;
    }, 320);
    if (issuesOpenBtn) issuesOpenBtn.focus();
  }

  if (issuesOpenBtn) issuesOpenBtn.addEventListener('click', openIssuesPanel);
  if (issuesCloseBtn) issuesCloseBtn.addEventListener('click', closeIssuesPanel);
  if (issuesBackdrop) issuesBackdrop.addEventListener('click', closeIssuesPanel);

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && issuesPanel && issuesPanel.classList.contains('is-open')) {
      e.preventDefault();
      closeIssuesPanel();
    }
  });

  // Open the issues dashboard straight away via a shared admin.html#issues link.
  if (location.hash.slice(1) === 'issues') openIssuesPanel();

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
  document.querySelectorAll('.action, button.action').forEach(action => {
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
    results = [...document.querySelectorAll('.action:not(.hidden), button.action:not(.hidden)')]
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
      const actions = section.matches('.action, button.action')
        ? [section]
        : [...section.querySelectorAll('.action, button.action')];
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
        zone.classList.add('is-active');
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
      if (results[activeIdx]) {
        e.preventDefault();
        const el = results[activeIdx];
        if (el.id === 'issues-open') openIssuesPanel();
        else el.click();
      }
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

  // Deep-link from other pages: index.html?q=fleet
  const params = new URLSearchParams(location.search);
  const initialQ = params.get('q');
  if (initialQ) {
    search.value = initialQ;
    runSearch();
    search.focus();
    if (history.replaceState) {
      const url = location.pathname + location.hash;
      history.replaceState(null, '', url);
    }
  }
})();
