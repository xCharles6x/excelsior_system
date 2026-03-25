/**
 * Excelsior Aviation — Smart Autocomplete
 * Drop this into core/static/js/autocomplete.js
 *
 * Usage: add data attributes to any input:
 *   data-autocomplete="true"
 *   data-ac-field="customer_name"
 *   data-ac-model="loadtest"   (or "repair", "inventory", "all")
 *
 * The script will attach a dropdown to each matching input automatically.
 */

(function () {
  'use strict';

  const AC_URL = '/api/autocomplete/';

  // --- Styles injected once ---
  const STYLE = `
    .ac-wrapper { position: relative; }
    .ac-dropdown {
      position: absolute;
      top: calc(100% + 3px);
      left: 0; right: 0;
      background: #fff;
      border: 1.5px solid #f5a623;
      border-radius: 8px;
      box-shadow: 0 6px 24px rgba(0,0,0,0.13);
      z-index: 9999;
      max-height: 220px;
      overflow-y: auto;
      display: none;
    }
    .ac-dropdown.open { display: block; }
    .ac-item {
      padding: 0.5rem 0.85rem;
      font-family: 'Barlow', 'Rajdhani', sans-serif;
      font-size: 0.9rem;
      font-weight: 600;
      color: #1c1c2e;
      cursor: pointer;
      border-bottom: 1px solid #f5f5f5;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      transition: background 0.12s;
    }
    .ac-item:last-child { border-bottom: none; }
    .ac-item:hover, .ac-item.active {
      background: #fff8ec;
      color: #d4891a;
    }
    .ac-item i { color: #f5a623; font-size: 0.8rem; flex-shrink: 0; }
    .ac-match { font-weight: 800; color: #FF6B00; }
    .ac-empty {
      padding: 0.6rem 0.85rem;
      font-family: 'Barlow', 'Rajdhani', sans-serif;
      font-size: 0.82rem;
      color: #bbb;
      font-style: italic;
    }
  `;

  function injectStyles() {
    if (document.getElementById('ac-styles')) return;
    const s = document.createElement('style');
    s.id = 'ac-styles';
    s.textContent = STYLE;
    document.head.appendChild(s);
  }

  function highlightMatch(text, query) {
    if (!query) return text;
    const idx = text.toLowerCase().indexOf(query.toLowerCase());
    if (idx === -1) return text;
    return (
      text.slice(0, idx) +
      '<span class="ac-match">' + text.slice(idx, idx + query.length) + '</span>' +
      text.slice(idx + query.length)
    );
  }

  function initInput(input) {
    const field = input.dataset.acField;
    const model = input.dataset.acModel || 'all';

    if (!field) return;
    if (input._acInit) return;
    input._acInit = true;

    // Wrap input in a relative container
    const wrapper = document.createElement('div');
    wrapper.className = 'ac-wrapper';
    input.parentNode.insertBefore(wrapper, input);
    wrapper.appendChild(input);

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.className = 'ac-dropdown';
    wrapper.appendChild(dropdown);

    let activeIdx = -1;
    let currentResults = [];
    let debounceTimer = null;

    function openDropdown(results) {
      currentResults = results;
      activeIdx = -1;
      dropdown.innerHTML = '';

      if (!results.length) {
        dropdown.classList.remove('open');
        return;
      }

      results.forEach(function (text, i) {
        const item = document.createElement('div');
        item.className = 'ac-item';
        item.innerHTML = '<i class="bi bi-clock-history"></i><span>' +
          highlightMatch(text, input.value.trim()) + '</span>';
        item.addEventListener('mousedown', function (e) {
          e.preventDefault();
          selectItem(text);
        });
        dropdown.appendChild(item);
      });

      dropdown.classList.add('open');
    }

    function closeDropdown() {
      dropdown.classList.remove('open');
      activeIdx = -1;
    }

    function selectItem(text) {
      input.value = text;
      closeDropdown();
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function setActive(idx) {
      const items = dropdown.querySelectorAll('.ac-item');
      items.forEach(function (el) { el.classList.remove('active'); });
      if (idx >= 0 && idx < items.length) {
        items[idx].classList.add('active');
        items[idx].scrollIntoView({ block: 'nearest' });
      }
      activeIdx = idx;
    }

    function fetchSuggestions(q) {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(function () {
        const url = AC_URL + '?field=' + encodeURIComponent(field) +
          '&q=' + encodeURIComponent(q) +
          '&model=' + encodeURIComponent(model);
        fetch(url)
          .then(function (r) { return r.json(); })
          .then(function (data) {
            if (document.activeElement === input) {
              openDropdown(data.results || []);
            }
          })
          .catch(function () { closeDropdown(); });
      }, 180);
    }

    // Events
    input.addEventListener('input', function () {
      const q = input.value.trim();
      if (q.length === 0) {
        // Show recent (blank query = all)
        fetchSuggestions('');
      } else if (q.length >= 1) {
        fetchSuggestions(q);
      } else {
        closeDropdown();
      }
    });

    input.addEventListener('focus', function () {
      fetchSuggestions(input.value.trim());
    });

    input.addEventListener('keydown', function (e) {
      const items = dropdown.querySelectorAll('.ac-item');
      if (!dropdown.classList.contains('open')) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setActive(Math.min(activeIdx + 1, items.length - 1));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setActive(Math.max(activeIdx - 1, 0));
      } else if (e.key === 'Enter') {
        if (activeIdx >= 0 && currentResults[activeIdx]) {
          e.preventDefault();
          selectItem(currentResults[activeIdx]);
        }
      } else if (e.key === 'Escape') {
        closeDropdown();
      }
    });

    document.addEventListener('click', function (e) {
      if (!wrapper.contains(e.target)) {
        closeDropdown();
      }
    });
  }

  function initAll() {
    document.querySelectorAll('[data-autocomplete="true"]').forEach(initInput);
  }

  // Init on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () {
      injectStyles();
      initAll();
    });
  } else {
    injectStyles();
    initAll();
  }

})();