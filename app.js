/**
 * Project Teams Ops Hub — client script
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

  const search = document.getElementById('search');
  const categories = document.querySelectorAll('.category');
  const zones = document.querySelectorAll('.hub-zone');
  const noResults = document.getElementById('no-results');
  const navButtons = document.querySelectorAll('.hub-nav-btn');

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

  if (!search) return;

  /**
   * Build searchable text for an action link, including optional footnote copy.
   * @param {Element} action
   * @param {Element} section
   * @returns {string}
   */
  function actionSearchText(action, section) {
    const block = action.closest('.task-block');
    const footnote = (block || section).querySelector('.action-note');
    const footnoteText = footnote
      ? `${footnote.textContent} ${footnote.dataset.keywords || ''}`
      : '';
    return `${action.textContent} ${action.dataset.keywords || ''} ${footnoteText}`.toLowerCase();
  }

  search.addEventListener('input', () => {
    const q = search.value.toLowerCase().trim();
    const isSearching = Boolean(q);
    document.body.classList.toggle('is-searching', isSearching);
    let anyVisible = false;

    categories.forEach(section => {
      const sectionKeywords = (section.dataset.keywords || '').toLowerCase();
      const actions = section.matches('.action')
        ? [section]
        : [...section.querySelectorAll('.action')];
      let sectionHasMatch = false;

      actions.forEach(action => {
        const block = action.closest('.task-block');
        const match = !q
          || actionSearchText(action, section).includes(q)
          || sectionKeywords.includes(q);

        action.classList.toggle('hidden', !match);

        if (block && block !== action) {
          const blockVisible = [...block.querySelectorAll('.action')].some(
            a => !a.classList.contains('hidden')
          ) || (block.querySelector('.action-note') && match);
          block.classList.toggle('hidden', q && !blockVisible);
        }

        if (match) {
          sectionHasMatch = true;
          anyVisible = true;
        }
      });

      if (!actions.length && (!q || sectionKeywords.includes(q))) {
        sectionHasMatch = true;
        anyVisible = true;
      }

      section.classList.toggle('hidden', q && !sectionHasMatch);
    });

    zones.forEach(zone => {
      if (!q) {
        zone.classList.remove('hidden');
        zone.classList.toggle('is-active', zone.id === currentView);
        return;
      }
      const hasVisible = zone.querySelector('.category:not(.hidden), .action:not(.hidden)');
      zone.classList.toggle('hidden', !hasVisible);
      if (hasVisible) zone.classList.add('is-active');
    });

    noResults.classList.toggle('visible', q && !anyVisible);
  });
})();
